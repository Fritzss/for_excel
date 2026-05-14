import os
import subprocess
import tempfile
from flask import Flask, request, render_template_string, send_file

UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

HTML_FORM = '''
<!doctype html>
<html lang="ru">
<head>
    <title>Обработчик Excel | Mrio USA</title>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; background: #f4f7f6; margin: 0; padding: 20px; }
        .container { background: white; padding: 2rem; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); width: 450px; max-width: 100%; }
        h1 { color: #2c3e50; margin-bottom: 1.5rem; text-align: center; }
        label { display: block; margin-top: 1rem; font-weight: 600; color: #2c3e50; }
        input[type="file"], input[type="text"] { width: 100%; padding: 8px; margin-top: 5px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box; }
        input[type="submit"] { background-color: #3498db; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 1rem; margin-top: 1.5rem; width: 100%; }
        input[type="submit"]:hover { background-color: #2980b9; }
        .error { color: #e74c3c; margin-top: 1rem; text-align: center; }
        .help { font-size: 0.8rem; color: #7f8c8d; margin-top: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📁 Обработка Excel-файла</h1>
        <form method="post" enctype="multipart/form-data" action="/upload">
            <label>📎 Excel файл (.xlsx):</label>
            <input type="file" name="file" accept=".xlsx" required>

            <label>📄 Имя листа (sheet):</label>
            <input type="text" name="sheet" value="Лист1" required>
            <div class="help">Название листа, на котором находятся данные (например, "Лист1", "Sheet1").</div>

            <label>🏷️ Маркер окончания (tag) — опционально:</label>
            <input type="text" name="tag" placeholder="например, Итог">
            <div class="help">Если указать, скрипт остановит группировку на строке, содержащей этот маркер. Оставьте пустым, чтобы обработать весь лист.</div>

            <input type="submit" value="Обработать">
        </form>
        {% if error %}
            <div class="error">{{ error }}</div>
        {% endif %}
    </div>
</body>
</html>
'''

@app.route('/')
def upload_form():
    return render_template_string(HTML_FORM)

@app.route('/upload', methods=['POST'])
def upload_file():
    # Проверка наличия файла
    if 'file' not in request.files:
        return render_template_string(HTML_FORM, error='Файл не выбран'), 400

    file = request.files['file']
    if file.filename == '':
        return render_template_string(HTML_FORM, error='Файл не выбран'), 400

    if not allowed_file(file.filename):
        return render_template_string(HTML_FORM, error='Недопустимый формат. Загрузите .xlsx'), 400

    # Получаем параметры из формы
    sheet_name = request.form.get('sheet', '').strip()
    tag_value = request.form.get('tag', '').strip()

    if not sheet_name:
        return render_template_string(HTML_FORM, error='Не указано имя листа'), 400

    # Сохраняем загруженный файл во временную папку
    temp_input_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(temp_input_path)

    with tempfile.TemporaryDirectory() as work_dir:
        # Копируем скрипт groups2sheets.py
        script_src = '/app/groups2sheets.py'
        script_dst = os.path.join(work_dir, 'groups2sheets.py')
        subprocess.run(['cp', script_src, script_dst], check=True)

        # Формируем config.ini с переданными параметрами
        config_content = f"""[Settings]
src_file = {os.path.basename(temp_input_path)}
sheet = {sheet_name}
"""
        if tag_value:
            config_content += f"tag = {tag_value}\n"
        # Если tag не указан — не пишем строку tag, скрипт обработает весь лист

        config_path = os.path.join(work_dir, 'config.ini')
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(config_content)

        # Копируем загруженный файл в рабочую директорию
        work_input_path = os.path.join(work_dir, os.path.basename(temp_input_path))
        subprocess.run(['cp', temp_input_path, work_input_path], check=True)

        # Запускаем обработку
        result = subprocess.run(
            ['python3', script_dst],
            cwd=work_dir,
            capture_output=True,
            text=True
        )

        if result.returncode != 0:
            # Логируем stderr для отладки (можно вывести в консоль)
            print("STDERR:", result.stderr)
            return render_template_string(HTML_FORM, error=f'Ошибка обработки: {result.stderr}'), 500

        # Ищем выходной файл
        output_filename = f"groups2sheets-{os.path.basename(temp_input_path)}"
        output_path = os.path.join(work_dir, output_filename)

        if not os.path.exists(output_path):
            return render_template_string(HTML_FORM, error='Выходной файл не создан'), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
