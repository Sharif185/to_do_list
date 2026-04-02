import os
from flask import Flask, render_template, request, redirect, url_for
from flask_sqlalchemy import SQLAlchemy
from dotenv import load_dotenv

load_dotenv()

# FIX: Railway gives postgres:// but SQLAlchemy 2.x needs postgresql://
database_url = os.environ.get("DATABASE_URL", "sqlite:///tasks.db")
if database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql://", 1)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = database_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# ── Model ──────────────────────────────────────────────────
class Task(db.Model):
    id   = db.Column(db.Integer, primary_key=True)
    text = db.Column(db.String(500), nullable=False)

with app.app_context():
    db.create_all()

# ── Routes ─────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        text = request.form.get("task", "").strip()
        if text:
            db.session.add(Task(text=text))
            db.session.commit()
    tasks = Task.query.all()
    return render_template("index.html", tasks=tasks)

@app.route("/delete/<int:task_id>")
def delete(task_id):
    task = Task.query.get(task_id)
    if task:
        db.session.delete(task)
        db.session.commit()
    return redirect(url_for("index"))

@app.route("/edit/<int:task_id>", methods=["POST"])
def edit(task_id):
    task = Task.query.get(task_id)
    new_text = request.form.get("task_text", "").strip()
    if task and new_text:
        task.text = new_text
        db.session.commit()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
