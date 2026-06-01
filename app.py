# -*- coding: utf-8 -*-
"""
Веб-платформа: ИНН-парсер salyk.kg + автозаполнение формы ZVN STI-222.
"""

import io
import os
import re
import time
import uuid
import zipfile
from pathlib import Path
from threading import Lock

from flask import (
    Flask, render_template, request, send_file, redirect,
    url_for, flash, jsonify, abort, after_this_request,
)

import pdf_filler
from pdf_filler import COLUMN_LABELS, fill_pdf_bytes, merge_pdfs, read_excel_rows, safe_filename
import inn_lookup


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PDF = BASE_DIR / "assets" / "form.pdf"
TEMPLATE_XLSX = BASE_DIR / "assets" / "data_template.xlsx"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB

# Простое in-memory хранилище готовых файлов на время одной сессии скачивания.
_FILES: dict[str, dict] = {}
_FILES_LOCK = Lock()
_FILE_TTL = 30 * 60  # 30 минут


def _cleanup_expired():
    now = time.time()
    with _FILES_LOCK:
        expired = [k for k, v in _FILES.items() if v["expires"] < now]
        for k in expired:
            _FILES.pop(k, None)


def _store_file(filename: str, data: bytes, mimetype: str) -> str:
    _cleanup_expired()
    token = uuid.uuid4().hex
    with _FILES_LOCK:
        _FILES[token] = {
            "filename": filename,
            "data": data,
            "mimetype": mimetype,
            "expires": time.time() + _FILE_TTL,
        }
    return token


@app.route("/")
def index():
    columns = [(letter, COLUMN_LABELS[letter]) for letter in "BCDEFGHIJKLMNOPQRSTU"]
    return render_template("index.html", columns=columns)


@app.route("/template")
def download_template():
    if not TEMPLATE_XLSX.exists():
        abort(404)
    return send_file(
        TEMPLATE_XLSX,
        as_attachment=True,
        download_name="data_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/form-template")
def download_form_template():
    if not TEMPLATE_PDF.exists():
        abort(404)
    return send_file(
        TEMPLATE_PDF,
        as_attachment=True,
        download_name="form.pdf",
        mimetype="application/pdf",
    )


@app.route("/generate", methods=["POST"])
def generate():
    if "excel" not in request.files or request.files["excel"].filename == "":
        flash("Загрузите Excel-файл", "error")
        return redirect(url_for("index"))

    f = request.files["excel"]
    try:
        excel_bytes = f.read()
        rows = read_excel_rows(excel_bytes)
    except Exception as e:
        flash(f"Не удалось прочитать Excel: {e}", "error")
        return redirect(url_for("index"))

    if not rows:
        flash("В файле нет строк с данными.", "error")
        return redirect(url_for("index"))

    mode = request.form.get("mode", "individual")
    group_col = request.form.get("group_column", "").strip().upper()

    # Сгенерировать PDF на каждую строку, в памяти.
    individual = []
    for idx, row in enumerate(rows, 1):
        pdf_bytes = fill_pdf_bytes(str(TEMPLATE_PDF), row)
        num = row.get("A") or idx
        name_part = safe_filename(row.get("C") or f"row{idx}")
        fname = f"{str(num).zfill(3)}_{name_part}.pdf"
        individual.append({"row": row, "filename": fname, "bytes": pdf_bytes})

    zip_buf = io.BytesIO()
    summary = []  # (label, filename, size)

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Всегда кладём отдельные файлы
        for item in individual:
            zf.writestr(f"individual/{item['filename']}", item["bytes"])

        # Группировка
        if mode == "individual":
            grouping_cols = []
        elif mode == "single":
            grouping_cols = [group_col] if group_col in COLUMN_LABELS else []
        elif mode == "all":
            # Группировка по всем содержательным колонкам (B..U), кроме A и пустых
            grouping_cols = [c for c in "BCDEFGHIJKLMNOPQRSTU"]
        else:
            grouping_cols = []

        for col in grouping_cols:
            groups = {}
            for item in individual:
                raw = item["row"].get(col)
                if raw is None or str(raw).strip() == "":
                    key = "_пусто_"
                elif hasattr(raw, "strftime"):
                    key = raw.strftime("%Y-%m-%d")
                else:
                    key = str(raw).strip()
                groups.setdefault(key, []).append(item["bytes"])

            label = COLUMN_LABELS[col]
            label_safe = safe_filename(label)
            for key, pdfs in groups.items():
                merged = merge_pdfs(pdfs)
                key_safe = safe_filename(key) or "_пусто_"
                zf.writestr(
                    f"grouped_by_{col}_{label_safe}/{key_safe}.pdf",
                    merged,
                )
                summary.append((label, key, len(pdfs)))

    zip_buf.seek(0)
    archive_name = f"zvn_pdfs_{int(time.time())}.zip"
    token = _store_file(archive_name, zip_buf.getvalue(), "application/zip")

    return render_template(
        "result.html",
        token=token,
        archive_name=archive_name,
        total_rows=len(rows),
        mode=mode,
        group_col=group_col,
        group_col_label=COLUMN_LABELS.get(group_col, ""),
        summary=summary,
    )


@app.route("/download/<token>")
def download(token):
    _cleanup_expired()
    with _FILES_LOCK:
        entry = _FILES.get(token)
    if not entry:
        abort(404)
    return send_file(
        io.BytesIO(entry["data"]),
        as_attachment=True,
        download_name=entry["filename"],
        mimetype=entry["mimetype"],
    )


@app.route("/inn")
def inn_page():
    return render_template("inn.html")


@app.route("/inn/lookup", methods=["POST"])
def inn_lookup_route():
    inns = []
    text_input = request.form.get("inns_text", "")
    if text_input:
        inns.extend(inn_lookup.parse_inn_list_from_text(text_input))

    if "excel" in request.files and request.files["excel"].filename:
        try:
            xls = request.files["excel"].read()
            inns.extend(inn_lookup.parse_inn_list_from_excel(xls))
        except Exception as e:
            flash(f"Не удалось прочитать Excel: {e}", "error")
            return redirect(url_for("inn_page"))

    # Удалить дубли с сохранением порядка
    seen = set()
    unique = []
    for x in inns:
        if x not in seen:
            seen.add(x)
            unique.append(x)
    inns = unique

    if not inns:
        flash("Не нашёл ни одного валидного ИНН (нужно 10+ цифр).", "error")
        return redirect(url_for("inn_page"))

    if len(inns) > 500:
        flash("Слишком много ИНН за раз (макс. 500). Разбейте на части.", "error")
        return redirect(url_for("inn_page"))

    results = inn_lookup.lookup_many(inns, delay=0.4)
    xlsx_bytes = inn_lookup.build_results_xlsx(results)
    fname = f"inn_results_{int(time.time())}.xlsx"
    token = _store_file(
        fname,
        xlsx_bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    found = sum(1 for r in results if r["name"])
    not_found = sum(1 for r in results if r["status"] == "не найдено")
    errors = len(results) - found - not_found

    return render_template(
        "inn_result.html",
        token=token,
        archive_name=fname,
        total=len(results),
        found=found,
        not_found=not_found,
        errors=errors,
        rows=results,
    )


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
