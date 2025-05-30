from flask import Flask, render_template, request, redirect, url_for, session, Response, send_file, jsonify
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

def get_admin_stats():
    """Calcola statistiche per il dashboard admin"""
    codici = carica_codici()

    total_codes = len(codici)
    used_codes = sum(1 for c in codici.values() if c.get("usato", False))
    completed_tests = sum(1 for c in codici.values() if c.get("usato", False) and c.get("report"))

    usage_rate = round((used_codes / total_codes) * 100, 1) if total_codes > 0 else 0
    completion_rate = round((completed_tests / used_codes) * 100, 1) if used_codes > 0 else 0

    return {
        "total_codes": total_codes,
        "used_codes": used_codes,
        "completed_tests": completed_tests,
        "available_codes": total_codes - used_codes,
        "usage_rate": usage_rate,
        "completion_rate": completion_rate
    }

def get_usage_trend():
    """Analizza il trend di utilizzo negli ultimi giorni"""
    codici = carica_codici()
    usage_by_date = {}

    for seriale, info in codici.items():
        if info.get("usato") and info.get("data"):
            try:
                date_str = info["data"].split(" ")[0]  # Prende solo la data
                if date_str in usage_by_date:
                    usage_by_date[date_str] += 1
                else:
                    usage_by_date[date_str] = 1
            except:
                continue

    return usage_by_date

def get_scale_averages():
    """Calcola le medie per scala dai test completati"""
    codici = carica_codici()
    scale_scores = {}

    for seriale, info in codici.items():
        if info.get("report"):
            for scala, dati in info["report"].items():
                if scala not in scale_scores:
                    scale_scores[scala] = []
                scale_scores[scala].append(dati.get("punteggio_grezzo", 0))

    # Calcola medie
    scale_averages = {}
    for scala, scores in scale_scores.items():
        if scores:
            scale_averages[scala] = round(np.mean(scores), 1)

    return scale_averages

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
        if idx + 1 < len(items):
            return redirect(url_for('question', idx=idx + 1))
        else:
            return redirect(url_for("result"))
    item = items[idx]
    return render_template("question.html", idx=idx + 1, total=len(items), item=item)

@app.route("/result")
def result():
    items = session.get("items", get_all_items())
    answers = session.get("answers", [])
    scores = {}
    for i, ans in enumerate(answers):
        scala = items[i]['scala']
        rev = items[i]['reverse']
        val = 7 - ans if rev else ans
        scores.setdefault(scala, []).append(val)
    sum_scores = {s: sum(v) for s, v in scores.items()}
    for scala, score in sum_scores.items():
        database_storico.append({"scala": scala, "score": score})
    with open("database.json", "w", encoding="utf-8") as f:
        json.dump(database_storico, f, ensure_ascii=False)
    report = {}
    for scala, score in sum_scores.items():
        scores_scala = [x["score"] for x in database_storico if x["scala"] == scala]
        percentile = int(round((np.sum(np.array(scores_scala) < score) / len(scores_scala)) * 100))
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
    if ds and (ds.get("percentile", 0) >= 85 or ds.get("stanina", 0) >= 8):
        alert = True
    risposte_dettaglio = []
    for i, ans in enumerate(answers):
        item = items[i]
        reverse = item['reverse']
        punteggio = 7 - ans if reverse else ans
        risposte_dettaglio.append({
            "idx": i + 1,
            "text": item['text'],
            "scala": item['scala'],
            "answer": ans,
            "punteggio": punteggio,
            "reverse": reverse
        })
    codici = carica_codici()
    seriale = session.get("seriale")
    email = session.get("email")
    print(f"SALVO SU {seriale} - {email}")
    if seriale in codici:
        codici[seriale]["risposte_dettaglio"] = risposte_dettaglio
        codici[seriale]["report"] = report
        salva_codici(codici)
        print("SALVATO CORRETTAMENTE")
    else:
        print("SERIALE NON TROVATO, NON SALVO!")
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
    success_message = None

    if request.method == "POST" and "genera" in request.form:
        try:
            num_codici = int(request.form.get("num_codici", 50))
            serials_list = genera_codici_batch(num_codici, "GO2B")
            success_message = f"{num_codici} nuovi codici generati con successo!"
        except Exception as e:
            success_message = f"Errore nella generazione: {str(e)}"
    else:
        serials_list = get_ultimi_codici()

    # Ottieni statistiche
    stats = get_admin_stats()

    # Ottieni utenti
    codici = carica_codici()
    utenti = []
    for s, info in codici.items():
        if info["usato"]:
            # Calcola se ha alert di desiderabilità sociale
            has_alert = False
            if info.get("report"):
                ds = info["report"].get("Desiderabilità sociale", {})
                has_alert = ds.get("percentile", 0) >= 85 or ds.get("stanina", 0) >= 8

            utenti.append({
                "nome": info["nome"],
                "email": info["email"],
                "seriale": s,
                "data": info.get("data", ""),
                "completed": bool(info.get("report")),
                "has_alert": has_alert
            })

    # Ordina per data (più recenti prima)
    utenti = sorted(utenti, key=lambda x: x.get("data", ""), reverse=True)

    return render_template("admin_dashboard_enhanced.html", 
                         utenti=utenti, 
                         serials_list=serials_list,
                         stats=stats,
                         success_message=success_message)

@app.route("/admin/api/stats")
def admin_api_stats():
    """API endpoint per statistiche dashboard"""
    if not session.get("admin_logged"):
        return jsonify({"error": "Non autorizzato"}), 401

    stats = get_admin_stats()
    usage_trend = get_usage_trend()
    scale_averages = get_scale_averages()

    return jsonify({
        "stats": stats,
        "usage_trend": usage_trend,
        "scale_averages": scale_averages
    })

@app.route("/admin/api/generate_codes", methods=["POST"])
def admin_api_generate_codes():
    """API endpoint per generare codici"""
    if not session.get("admin_logged"):
        return jsonify({"error": "Non autorizzato"}), 401

    try:
        data = request.get_json()
        num_codici = data.get("num_codici", 50)
        prefix = data.get("prefix", "GO2B")

        nuovi_codici = genera_codici_batch(num_codici, prefix)

        return jsonify({
            "success": True,
            "message": f"{num_codici} codici generati con successo",
            "codes": nuovi_codici
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Errore: {str(e)}"
        }), 500

@app.route("/admin/report/<email>/<seriale>")
def admin_report(email, seriale):
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))

    codici = carica_codici()
    user = codici.get(seriale)
    if not user or user.get("email") != email:
        return "Utente non trovato", 404
    risposte_dettaglio = user.get("risposte_dettaglio", [])
    report = user.get("report", {})
    alert = False
    ds = report.get("Desiderabilità sociale", {})
    if ds and (ds.get("percentile", 0) >= 85 or ds.get("stanina", 0) >= 8):
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

@app.route("/admin/export/users")
def admin_export_users():
    """Esporta lista utenti in Excel"""
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))

    codici = carica_codici()
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Utenti")

    # Headers
    headers = ["Nome", "Email", "Codice Seriale", "Data Test", "Test Completato", "Alert Desiderabilità"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Data
    row = 1
    for seriale, info in codici.items():
        if info.get("usato"):
            has_alert = False
            if info.get("report"):
                ds = info["report"].get("Desiderabilità sociale", {})
                has_alert = ds.get("percentile", 0) >= 85 or ds.get("stanina", 0) >= 8

            worksheet.write(row, 0, info.get("nome", ""))
            worksheet.write(row, 1, info.get("email", ""))
            worksheet.write(row, 2, seriale)
            worksheet.write(row, 3, info.get("data", ""))
            worksheet.write(row, 4, "Sì" if info.get("report") else "No")
            worksheet.write(row, 5, "Sì" if has_alert else "No")
            row += 1

    workbook.close()
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"utenti_go2b_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin/export/codes")
def admin_export_codes():
    """Esporta ultimi codici generati in Excel"""
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))

    codici_recenti = get_ultimi_codici()
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Codici")

    # Headers
    worksheet.write(0, 0, "Codice Seriale")
    worksheet.write(0, 1, "Stato")

    # Data
    codici_db = carica_codici()
    for idx, code in enumerate(codici_recenti):
        worksheet.write(idx + 1, 0, code)
        stato = "Utilizzato" if codici_db.get(code, {}).get("usato") else "Disponibile"
        worksheet.write(idx + 1, 1, stato)

    workbook.close()
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"codici_go2b_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin/export/results")
def admin_export_results():
    """Esporta risultati completi in Excel"""
    if not session.get("admin_logged"):
        return redirect(url_for("admin_login"))

    codici = carica_codici()
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # Foglio riassuntivo
    summary_ws = workbook.add_worksheet("Riassunto")
    headers = ["Nome", "Email", "Data Test"]

    # Aggiungi headers per tutte le scale
    scale_names = []
    for seriale, info in codici.items():
        if info.get("report"):
            for scala in info["report"].keys():
                if scala not in scale_names:
                    scale_names.append(scala)

    for scala in scale_names:
        headers.extend([f"{scala} - Punteggio", f"{scala} - Percentile", f"{scala} - Stanina"])

    for col, header in enumerate(headers):
        summary_ws.write(0, col, header)

    # Dati riassuntivi
    row = 1
    for seriale, info in codici.items():
        if info.get("report"):
            col = 0
            summary_ws.write(row, col, info.get("nome", ""))
            col += 1
            summary_ws.write(row, col, info.get("email", ""))
            col += 1
            summary_ws.write(row, col, info.get("data", ""))
            col += 1

            for scala in scale_names:
                if scala in info["report"]:
                    dati = info["report"][scala]
                    summary_ws.write(row, col, dati.get("punteggio_grezzo", ""))
                    summary_ws.write(row, col + 1, dati.get("percentile", ""))
                    summary_ws.write(row, col + 2, dati.get("stanina", ""))
                col += 3
            row += 1

    # Foglio dettagli risposte
    detail_ws = workbook.add_worksheet("Dettaglio Risposte")
    detail_headers = ["Nome", "Email", "Domanda", "Scala", "Risposta", "Punteggio", "Inversa"]
    for col, header in enumerate(detail_headers):
        detail_ws.write(0, col, header)

    detail_row = 1
    for seriale, info in codici.items():
        if info.get("risposte_dettaglio"):
            for risposta in info["risposte_dettaglio"]:
                detail_ws.write(detail_row, 0, info.get("nome", ""))
                detail_ws.write(detail_row, 1, info.get("email", ""))
                detail_ws.write(detail_row, 2, risposta.get("text", ""))
                detail_ws.write(detail_row, 3, risposta.get("scala", ""))
                detail_ws.write(detail_row, 4, risposta.get("answer", ""))
                detail_ws.write(detail_row, 5, risposta.get("punteggio", ""))
                detail_ws.write(detail_row, 6, "Sì" if risposta.get("reverse") else "No")
                detail_row += 1

    workbook.close()
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"risultati_completi_go2b_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_logged", None)
    return redirect(url_for("admin_login"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=True)