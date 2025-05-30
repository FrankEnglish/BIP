import json
import random
import string

def genera_codice(prefix="GO2B", lunghezza=6):
    lettere = string.ascii_uppercase + string.digits
    return f"{prefix}-" + ''.join(random.choices(lettere, k=lunghezza))

def genera_codici(n=100, prefix="GO2B"):
    codici = {}
    usati = set()
    while len(codici) < n:
        codice = genera_codice(prefix)
        if codice not in usati:
            codici[codice] = {"usato": False, "email": "", "nome": ""}
            usati.add(codice)
    return codici

if __name__ == "__main__":
    NUM_CODICI = 150  # Numero di codici da generare (puoi cambiare)
    PREFIX = "GO2B"   # Prefisso dei codici (puoi cambiare)
    codici = genera_codici(n=NUM_CODICI, prefix=PREFIX)
    with open("codici_seriali.json", "w", encoding="utf-8") as f:
        json.dump(codici, f, ensure_ascii=False, indent=2)
    print(f"Creati {NUM_CODICI} codici seriali in 'codici_seriali.json'")