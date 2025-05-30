from flask import Flask, render_template, request, redirect, url_for, session, Response, send_file
import json
import numpy as np
import os
from datetime import datetime
import random
import string

# Per Excel export:
import io
try:
    import xlsxwriter
except ImportError:
    import pip
    pip.main(['install', 'xlsxwriter'])
    import xlsxwriter

app = Flask(__name__)
app.secret_key = "test-bip"

# ========== PARAMETRI ADMIN ==========
ADMIN_USER = "go2badmin"
ADMIN_PASS = "SuperSegreta2024"
CODICE_MASTER = "GO2B-MASTER"

# ========== UTILITY CODICI SERIALI ==========
CODICI_FILE = "codici_seriali.json"

def carica_codici():
    if os.path.exists(CODICI_FILE):
        with open(CODICI_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return {CODICE_MASTER: {"usato": False, "email": "", "nome": "", "data": ""}}

def salva_codici(codici):
    if CODICE_MASTER not in codici:
        codici[CODICE_MASTER] = {"usato": False, "email": "", "nome": "", "data": ""}
    with open(CODICI_FILE, "w", encoding="utf-8") as f:
        json.dump(codici, f, ensure_ascii=False, indent=2)

def genera_codice(prefix="GO2B", lunghezza=6):
    lettere = string.ascii_uppercase + string.digits
    return f"{prefix}-" + ''.join(random.choices(lettere, k=lunghezza))

def genera_codici_batch(n=50, prefix="GO2B"):
    codici = carica_codici()
    nuovi = []
    while len(nuovi) < n:
        code = genera_codice(prefix)
        if code not in codici:
            codici[code] = {"usato": False, "email": "", "nome": "", "data": ""}
            nuovi.append(code)
    salva_codici(codici)
    with open("ultimi_codici_generati.json", "w", encoding="utf-8") as f:
        json.dump(nuovi, f, ensure_ascii=False, indent=2)
    return nuovi

def get_ultimi_codici():
    if os.path.exists("ultimi_codici_generati.json"):
        with open("ultimi_codici_generati.json", "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        return []

# ========== FINE UTILITY CODICI SERIALI ==========

# Carica domande e struttura dal file json
with open("data.json", "r", encoding="utf-8") as f:
    test_structure = json.load(f)

# Carica (o crea) il database storico dei risultati
if os.path.exists("database.json"):
    with open("database.json", "r", encoding="utf-8") as f:
        database_storico = json.load(f)
else:
    database_storico = []

def get_all_items():
    items = []
    for area in test_structure['areas']:
        for scala in area['scales']:
            for item in scala['items']:
                items.append({
                    "scala": scala['name'],
                    "text": item['text'],
                    "reverse": item['reverse']
                })
    return items

@app.route("/benvenuto")
def benvenuto():
    return render_template("benvenuto.html")

@app.route("/")
def root():
    return redirect(url_for("benvenuto"))

@app.route("/login", methods=["GET", "POST"])
def login():
    errore = None
    if request.method == "POST":
        nome = request.form["nome"].strip()
        email = request.form["email"].strip().lower()
        seriale = request.form["seriale"].strip().upper()
        codici = carica_codici()
        if seriale == CODICE_MASTER:
            session["nome"] = nome
            session["email"] = email
            session["seriale"] = seriale
            return redirect(url_for('start'))
        if not nome or not email or not seriale:
            errore = "Compila tutti i campi"
        elif seriale not in codici:
            errore = "Il codice seriale non è valido. Contatta il referente."
        elif codici[seriale]["usato"]:
            errore = "Questo codice seriale è già stato utilizzato."
        else:
            session["nome"] = nome
            session["email"] = email
            session["seriale"] = seriale
            codici[seriale]["usato"] = True
            codici[seriale]["email"] = email
            codici[seriale]["nome"] = nome
            codici[seriale]["data"] = datetime.now().strftime("%d/%m/%Y %H:%M")
            salva_codici(codici)
            return redirect(url_for('start'))
    return render_template("login.html", errore=errore)

@app.route("/start", methods=["GET", "POST"])
def start():
    if not session.get("nome") or not session.get("email") or not session.get("seriale"):
        return redirect(url_for('login'))
    session["answers"] = []
    items = get_all_items()
    session["items"] = items
    return redirect(url_for('question', idx=0))

@app.route("/question/<int:idx>", methods=["GET", "POST"])
def question(idx):
    items = session.get("items", get_all_items())
    if request.method == "POST":
        answer = int(request.form["answer"])
        answers = session.get("answers", [])
        answers.append(answer)
        session["answers"] = answers
        if idx+1 < len(items):
            return redirect(url_for('question', idx=idx+1))
        else:
            return redirect(url_for("result"))
    item = items[idx]
    return render_template("question.html", idx=idx+1, total=len(items), item=item)

@app.route("/result")
def result():
    items = session.get("items", get_all_items())
    answers = session.get("answers", [])
    scores = {}
    for i, ans in enumerate(answers):
        scala = items[i]['scala']
        rev = items[i]['reverse']
        val = 7-ans if rev else ans
        scores.setdefault(scala, []).append(val)
    sum_scores = {s: sum(v) for s, v in scores.items()}
    for scala, score in sum_scores.items():
        database_storico.append({"scala": scala, "score": score})
    with open("database.json", "w", encoding="utf-8") as f:
        json.dump(database_storico, f, ensure_ascii=False)
    report = {}
    for scala, score in sum_scores.items():
        scores_scala = [x["score"] for x in database_storico if x["scala"] == scala]
        percentile = int(round((np.sum(np.array(scores_scala) < score) / len(scores_scala))*100))
        sorted_scores = sorted(scores_scala)
        position = sorted_scores.index(score)
        stanina = int(np.ceil(((position + 1) / len(sorted_scores)) * 9))
        stanina = min(max(stanina, 1), 9)
        report[scala] = {
            "punteggio_grezzo": score,
            "percentile": percentile,
            "stanina": stanina
        }
    alert = False
    ds = report.get("Desiderabilità sociale", {})
    if ds and (ds.get("percentile",0) >= 85 or ds.get("stanina",0) >= 8):
        alert = True

    # ---- RISPOSTE DETTAGLIO PER IL REPORT
    risposte_dettaglio = []
    for i, ans in enumerate(answers):
        item = items[i]
        reverse = item['reverse']
        punteggio = 7-ans if reverse else ans
        risposte_dettaglio.append({
            "idx": i+1,
            "text": item['text'],
            "scala": item['scala'],
            "answer": ans,
            "punteggio": punteggio,
            "reverse": reverse
        })
    # Salva le risposte dettaglio come storico nel file codici_seriali.json
    codici = carica_codici()
    seriale = session.get("seriale")
    email = session.get("email")
    if seriale in codici:
        codici[seriale]["risposte_dettaglio"] = risposte_dettaglio
        codici[seriale]["report"] = report
        salva_codici(codici)

    return render_template(
        "result.html",
        report=report,
        alert=alert,
        data_test=datetime.now().strftime("%d/%m/%Y"),
        nome=session.get("nome"),
        email=session.get("email"),
        seriale=seriale,
        risposte_dettaglio=risposte_dettaglio
    )

# ================== AREA ADMIN =====================

@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    errore = None
    if request.method == "POST":
        user = request.form["user"]
        password = request.form["password"]
        if user == ADMIN_USER and password == ADMIN_PASS:
            session["admin_logged"] = True
            return redirect(url_for("admin_dashboard"))
        else:
            errore = "Credenziali errate."
    return render_template("admin_login.html", errore=errore)

@app.route("/admin_dashboard", methods=["GET", "POST"])
def admin_dashboard():
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))
    serials_list = []
    if request.method == "POST" and "genera" in request.form:
        serials_list = genera_codici_batch(50, "GO2B")
    else:
        serials_list = get_ultimi_codici()
    codici = carica_codici()
    utenti = []
    for s, info in codici.items():
        if info["usato"]:
            utenti.append({
                "nome": info["nome"],
                "email": info["email"],
                "seriale": s,
                "data": info.get("data", "")
            })
    utenti = sorted(utenti, key=lambda x: x.get("data", ""))
    return render_template("admin_dashboard.html", utenti=utenti, serials_list=serials_list)

@app.route("/admin/report/<email>/<seriale>")
def admin_report(email, seriale):
    codici = carica_codici()
    user = codici.get(seriale)
    if not user or user.get("email") != email:
        return "Utente non trovato", 404
    risposte_dettaglio = user.get("risposte_dettaglio", [])
    report = user.get("report", {})
    alert = False
    ds = report.get("Desiderabilità sociale", {})
    if ds and (ds.get("percentile",0) >= 85 or ds.get("stanina",0) >= 8):
        alert = True
    return render_template(
        "result.html",
        report=report,
        alert=alert,
        data_test=user.get("data", ""),
        nome=user.get("nome", ""),
        email=user.get("email", ""),
        seriale=seriale,
        risposte_dettaglio=risposte_dettaglio
    )

@app.route("/admin/codici_excel")
def codici_excel():
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))
    codici = get_ultimi_codici()
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Codici")
    worksheet.write(0, 0, "Codice seriale")
    for idx, code in enumerate(codici):
        worksheet.write(idx+1, 0, code)
    workbook.close()
    output.seek(0)
    return send_file(
        output,
        download_name="codici_seriali.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin/export")
def admin_export():
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))
    codici = carica_codici()
    rows = "Nome,Email,Seriale,Data\n"
    for s, info in codici.items():
        if info["usato"]:
            rows += f"{info['nome']},{info['email']},{s},{info.get('data','')}\n"
    return Response(
        rows,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=utenti_registrati.csv"}
    )

@app.route("/admin_logout")
def admin_logout():
    session.pop("admin_logged", None)
    return redirect(url_for("login"))

# ===================================================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=81, debug=True)
