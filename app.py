# -*- coding: utf-8 -*-
"""Веб-интерфейс сверки деклараций НВОС с платежами."""

import os
import re
import tempfile
import uuid
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

import core

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100 МБ

# WORK_DIR настраивается через env SVERKA_WORK_DIR (в Docker — /data)
WORK_DIR = Path(os.environ.get('SVERKA_WORK_DIR') or (Path(tempfile.gettempdir()) / 'sverka-web'))
WORK_DIR.mkdir(parents=True, exist_ok=True)
JOB_ID_RE = re.compile(r'^[0-9a-f]{12}$')


@app.get('/')
def index():
    return render_template('index.html')


@app.post('/api/reconcile')
def reconcile():
    decl_file = request.files.get('decl')
    sbros_file = request.files.get('sbros')
    vybros_file = request.files.get('vybros')
    q_raw = (request.form.get('quarter') or '1').strip().lower()
    if q_raw in ('year', 'год', 'god'):
        quarter = 'year'
    else:
        try:
            quarter = int(q_raw)
        except ValueError:
            quarter = 1
        if quarter not in (1, 2, 3):
            quarter = 1

    pay_year = None
    year_raw = (request.form.get('year') or '').strip()
    if year_raw:
        try:
            y = int(year_raw)
            if 2000 <= y <= 2100:
                pay_year = y
        except ValueError:
            pass

    if decl_file is None or decl_file.filename == '':
        return jsonify({'error': 'Не загружен файл деклараций.'}), 400
    if (sbros_file is None or sbros_file.filename == '') and \
       (vybros_file is None or vybros_file.filename == ''):
        return jsonify({'error': 'Загрузите хотя бы один файл платежей (сбросы или выбросы).'}), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = WORK_DIR / job_id
    job_dir.mkdir()

    decl_path = job_dir / 'decl.xlsx'
    decl_file.save(decl_path)
    payment_paths = {'sbros': None, 'vybros': None}
    src_names = {'decl': decl_file.filename}
    if sbros_file and sbros_file.filename:
        p = job_dir / 'sbros.xlsx'
        sbros_file.save(p)
        payment_paths['sbros'] = p
        src_names['sbros'] = sbros_file.filename
    if vybros_file and vybros_file.filename:
        p = job_dir / 'vybros.xlsx'
        vybros_file.save(p)
        payment_paths['vybros'] = p
        src_names['vybros'] = vybros_file.filename

    period_label = 'ГОД' if quarter == 'year' else f'КВ{quarter}'
    if pay_year:
        period_label += f' {pay_year}'
    out_name = f'Сверка платежей {period_label} — {datetime.now():%d.%m.%Y %H-%M}.xlsx'
    out_path = job_dir / 'result.xlsx'
    try:
        stats = core.build_workbook(decl_path, payment_paths, quarter, out_path,
                                    year=pay_year)
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception as e:
        app.logger.exception('reconcile failed')
        return jsonify({'error': f'Не удалось обработать файлы: {e}'}), 500

    # Имя файла для скачивания сохраняем рядом — переживёт перезапуск,
    # будет видно всем воркерам gunicorn.
    (job_dir / 'name.txt').write_text(out_name, encoding='utf-8')
    return jsonify({
        'download_id': job_id,
        'filename': out_name,
        'sources': src_names,
        'quarter': quarter,
        'year': pay_year,
        'stats': stats,
    })


@app.get('/api/download/<job_id>')
def download(job_id):
    if not JOB_ID_RE.match(job_id):
        return jsonify({'error': 'Некорректный идентификатор.'}), 400
    job_dir = WORK_DIR / job_id
    path = job_dir / 'result.xlsx'
    name_file = job_dir / 'name.txt'
    if not path.exists() or not name_file.exists():
        return jsonify({'error': 'Результат не найден. Сформируйте сверку заново.'}), 404
    name = name_file.read_text(encoding='utf-8').strip() or 'result.xlsx'
    return send_file(path, as_attachment=True, download_name=name)


@app.get('/healthz')
def healthz():
    return {'ok': True}


if __name__ == '__main__':
    # Локальный запуск для разработки. В Docker слушает gunicorn.
    app.run(host=os.environ.get('SVERKA_HOST', '127.0.0.1'),
            port=int(os.environ.get('SVERKA_PORT', '8021')),
            debug=False)
