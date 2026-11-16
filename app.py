import os, io, json, secrets
from pathlib import Path
from flask import Flask, request, redirect, url_for, render_template_string, send_file, session, flash
from docxtpl import DocxTemplate
import requests

APP_DIR = Path(__file__).parent
DATA_DIR = APP_DIR / "data"
TEMPLATE_DIR = DATA_DIR / "templates"
TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
CONFIG = DATA_DIR / "fields.json"

DEFAULT_FIELDS = [
    {"key":"nickname","label":"Никнейм","required":True},
    {"key":"organization","label":"Организация","required":False},
    {"key":"current_rank","label":"Текущий ранг","required":False},
    {"key":"target_rank","label":"Желаемый ранг","required":False},
    {"key":"work_done","label":"Выполненная работа","required":False},
    {"key":"evidence","label":"Ссылки на доказательства","required":False},
    {"key":"extra","label":"Дополнительная информация","required":False},
]
app = Flask(__name__)
app.secret_key = os.environ.get("APP_SECRET", secrets.token_hex(32))
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024

def load_fields():
    if CONFIG.exists():
        try: return json.loads(CONFIG.read_text(encoding="utf-8"))
        except Exception: pass
    return DEFAULT_FIELDS

def save_fields(fields):
    CONFIG.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")

def is_admin():
    return session.get("admin") is True

BASE = """
<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>RMRP Report Generator</title><style>
:root{color-scheme:dark}*{box-sizing:border-box}body{margin:0;background:#0c1220;color:#eef3ff;font:16px system-ui,Arial}
main{max-width:820px;margin:24px auto;padding:16px}.top{display:flex;align-items:center;justify-content:space-between;gap:12px}
h1{font-size:25px;margin:0}.sub{color:#a9b6cf;font-size:14px;line-height:1.5}.card{background:#172238;border:1px solid #30405d;border-radius:16px;padding:18px;margin:14px 0}
a{color:#9eb7ff}label{display:block;font-weight:650;margin:14px 0 6px}input,textarea,select{width:100%;padding:12px;border-radius:10px;border:1px solid #40516e;background:#0d1729;color:#fff;font-size:16px}
textarea{min-height:88px}button,.btn{display:inline-block;border:0;border-radius:10px;background:#6488f5;color:white;padding:12px 16px;font-size:15px;font-weight:700;text-decoration:none;cursor:pointer;margin-top:12px}
.secondary{background:#35445f}.notice{background:#253858;border-radius:10px;padding:12px;margin:12px 0}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.small{font-size:12px;color:#a9b6cf}.danger{background:#522c3a}
@media(max-width:600px){.grid{grid-template-columns:1fr}main{margin:8px auto;padding:12px}.top{align-items:flex-start;flex-direction:column}}
</style></head><body><main>
<div class="top"><div><h1>RMRP Report Generator</h1><div class="sub">Конструктор рапортов для RMRP · телефон и ПК</div></div><div><a href="/">Создать рапорт</a> · <a href="/admin">Настройки</a></div></div>
{% with msgs=get_flashed_messages() %}{% for m in msgs %}<div class="notice">{{m}}</div>{% endfor %}{% endwith %}
{{body|safe}}
</main></body></html>
"""

def page(body, **kwargs):
    return render_template_string(BASE, body=render_template_string(body, **kwargs))

HOME = """
<form class="card" method="post" enctype="multipart/form-data">
<h2>Новый рапорт</h2>
<label>Шаблон Word (.docx)</label><input type="file" name="template" accept=".docx" required>
<div class="small">Загрузи форму с метками в местах заполнения. Например: {{ '{{nickname}}' }} и {{ '{{work_done}}' }}.</div>
{% for f in fields %}
<label>{{f.label}}{% if f.required %} *{% endif %}</label>
{% if f.key in ['work_done','evidence','extra'] %}<textarea name="{{f.key}}" {% if f.required %}required{% endif %}></textarea>
{% else %}<input name="{{f.key}}" {% if f.required %}required{% endif %}>{% endif %}
{% endfor %}
<label>Действие</label><div class="grid"><button name="action" value="download">Скачать DOCX</button><button name="action" value="send">Отправить в Discord</button></div>
</form>
<div class="card sub"><b>Как подготовить DOCX:</b><br>Вставь в Word метки, например <code>{{ '{{nickname}}' }}</code>, <code>{{ '{{organization}}' }}</code>, <code>{{ '{{current_rank}}' }}</code>, <code>{{ '{{target_rank}}' }}</code>, <code>{{ '{{work_done}}' }}</code>, <code>{{ '{{evidence}}' }}</code>. Сами названия полей можно менять в разделе «Настройки».</div>
"""

ADMIN = """
<div class="card">
{% if not logged %}
<h2>Вход администратора</h2><form method="post"><input type="password" name="pin" placeholder="Пароль администратора" required><button>Войти</button></form>
{% else %}
<h2>Редактор полей</h2><div class="sub">Эти поля появятся в форме рапорта. Ключ — это метка, которую нужно поставить в Word: {{ '{{ключ}}' }}.</div>
<form method="post">
{% for f in fields %}
<div class="card">
<label>Название поля</label><input name="label_{{loop.index0}}" value="{{f.label}}" required>
<label>Ключ DOCX</label><input name="key_{{loop.index0}}" value="{{f.key}}" pattern="[a-zA-Z0-9_]+" required>
<label><input style="width:auto" type="checkbox" name="required_{{loop.index0}}" {% if f.required %}checked{% endif %}> Обязательное поле</label>
</div>
{% endfor %}
<button name="save" value="1">Сохранить поля</button>
</form>
<form method="post" action="/logout"><button class="secondary">Выйти из режима администратора</button></form>
{% endif %}
</div>
<div class="card sub">Для доступа администратора задай переменную <code>ADMIN_PIN</code> в настройках хостинга. Не используй простой пароль.</div>
"""

@app.get("/")
def home():
    return page(HOME, fields=load_fields())

@app.post("/")
def create():
    upload = request.files.get("template")
    if not upload or not upload.filename.lower().endswith(".docx"):
        flash("Нужен файл Word с расширением .docx.")
        return redirect(url_for("home"))
    try:
        raw = upload.read()
        doc = DocxTemplate(io.BytesIO(raw))
        fields = load_fields()
        context = {f["key"]: request.form.get(f["key"], "").strip() for f in fields}
        # Built-in aliases for common templates.
        aliases = {"никнейм":"nickname","ник":"nickname","организация":"organization",
                   "текущий_ранг":"current_rank","желаемый_ранг":"target_rank",
                   "выполненная_работа":"work_done","доказательства":"evidence"}
        for alias, source in aliases.items():
            context[alias] = context.get(source, "")
        doc.render(context, autoescape=True)
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        filename = "RMRP_raport.docx"
        if request.form.get("action") == "send":
            webhook = os.environ.get("DISCORD_WEBHOOK_URL", "").strip()
            if not webhook:
                flash("Отправка пока не настроена: администратор должен добавить DISCORD_WEBHOOK_URL на хостинге. DOCX скачивается.")
            else:
                summary = "**Новый рапорт RMRP**\n" + "\n".join(
                    f"**{f['label']}:** {context.get(f['key'],'') or '—'}" for f in fields
                )
                response = requests.post(webhook, data={"content":summary[:1900]},
                    files={"file":(filename,out.getvalue(),"application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                    timeout=20)
                if response.status_code in (200,204): flash("Рапорт отправлен в Discord.")
                else: flash(f"Discord вернул ошибку HTTP {response.status_code}. Проверь настройки Webhook.")
            out.seek(0)
        return send_file(out, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception:
        flash("Не удалось обработать документ. Проверь, что DOCX не повреждён, а метки не разбиты форматированием.")
        return redirect(url_for("home"))

@app.route("/admin", methods=["GET","POST"])
def admin():
    if request.method == "POST" and not is_admin():
        if request.form.get("pin") == os.environ.get("ADMIN_PIN", "SET_A_STRONG_ADMIN_PIN"):
            session["admin"] = True
            return redirect(url_for("admin"))
        flash("Неверный пароль.")
    elif request.method == "POST" and is_admin() and request.form.get("save"):
        new=[]
        for i in range(len(load_fields())):
            key=request.form.get(f"key_{i}","").strip()
            label=request.form.get(f"label_{i}","").strip()
            if not key or not label or not key.replace("_","").isalnum():
                flash("Проверь названия и ключи полей.")
                return redirect(url_for("admin"))
            new.append({"key":key,"label":label,"required":bool(request.form.get(f"required_{i}"))})
        save_fields(new)
        flash("Поля сохранены.")
        return redirect(url_for("admin"))
    return page(ADMIN, logged=is_admin(), fields=load_fields())

@app.post("/logout")
def logout():
    session.clear()
    return redirect(url_for("admin"))

@app.get("/health")
def health():
    return {"status":"ok"}

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)))
