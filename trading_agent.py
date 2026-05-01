#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================
  AGENT DE TRADING CRYPTO — PAPER TRADING (ARGENT FICTIF)
=============================================================
Ce script surveille les nouvelles crypto toutes les 30 minutes,
analyse leur sentiment avec l'IA Claude (Anthropic), et simule
des décisions de trading BTC avec un capital fictif de 300 USDT.

⚠️  IMPORTANT : Aucun ordre réel n'est passé sur Binance.
               C'est uniquement de la simulation (paper trading).
"""

# ─────────────────────────────────────────────────────────────
#  IMPORTATION DES BIBLIOTHÈQUES
# ─────────────────────────────────────────────────────────────
import os           # Lire les variables d'environnement
import sys          # Lire les arguments passés au script (ex: --once)
import time         # Faire des pauses entre les cycles
import json         # Sauvegarder des données dans des fichiers texte
import requests     # Appeler l'API NewsData.io
import anthropic    # Utiliser l'IA Claude pour analyser les news
import openpyxl     # Lire et écrire des fichiers Excel
from datetime import datetime, timezone   # Gérer les dates et heures
from dotenv import load_dotenv  # Lire le fichier .env (clés secrètes)

# ─────────────────────────────────────────────────────────────
#  CHARGEMENT DES CLÉS API DEPUIS LE FICHIER .env
# ─────────────────────────────────────────────────────────────
# Le fichier .env contient tes clés secrètes.
# On ne les écrit JAMAIS directement dans le code !
load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")   # Clé pour l'IA Claude
BINANCE_API_KEY   = os.getenv("BINANCE_API_KEY")      # Clé Binance (lecture prix)
BINANCE_SECRET    = os.getenv("BINANCE_SECRET")       # Secret Binance
NEWSDATA_API_KEY  = os.getenv("NEWSDATA_API_KEY")     # Clé pour les news crypto

# ─────────────────────────────────────────────────────────────
#  PARAMÈTRES GÉNÉRAUX (tu peux les modifier ici)
# ─────────────────────────────────────────────────────────────
CAPITAL_DEPART_USDT  = 300.0         # Capital fictif de départ (en USDT ≈ EUR)
INTERVALLE_SECONDES  = 30 * 60       # 30 minutes entre chaque vérification
FICHIER_EXCEL        = "trading_journal.xlsx"   # Nom du fichier Excel
FICHIER_CACHE_NEWS   = "cache_news.json"         # Mémorisation des news déjà vues
FICHIER_PORTFOLIO    = "portfolio_state.json"    # Sauvegarde du portefeuille fictif
TAILLE_MAX_CACHE     = 300           # Nombre max d'IDs de news à mémoriser

# Modèles Claude
MODELE_HAIKU  = "claude-haiku-4-5-20251001"  # Rapide et économique
MODELE_SONNET = "claude-sonnet-4-6"           # Plus puissant (signaux forts)

# Seuil pour utiliser Sonnet au lieu de Haiku (signal très fort)
SEUIL_SIGNAL_FORT = 8  # Score > 8 ou < -8 → analyse approfondie avec Sonnet

# Paramètres de gestion du risque
FRAIS_TRADING    = 0.001   # 0.1% de frais Binance par trade (achat ET vente)
STOP_LOSS_PCT    = 0.05    # Stop-loss à -5% sous le prix d'achat
TAKE_PROFIT_PCT  = 0.08    # Take-profit à +8% au-dessus du prix d'achat

# Kill-switch budgétaire (coûts API Claude)
LIMITE_COUT_QUOTIDIEN_USD = 1.00          # Arrêt automatique au-delà de 1 $/jour
FICHIER_TOKEN_USAGE       = "token_usage.json"
COUTS_PAR_TOKEN = {
    # Prix Anthropic en $ par token (https://www.anthropic.com/pricing)
    "claude-haiku-4-5-20251001": {"input": 0.80e-6, "output": 4.00e-6},
    "claude-sonnet-4-6":         {"input": 3.00e-6, "output": 15.00e-6},
}


# ═════════════════════════════════════════════════════════════
#  FONCTIONS UTILITAIRES
# ═════════════════════════════════════════════════════════════

class LimiteQuotidienneAtteinte(Exception):
    """Levée quand le coût API Claude dépasse LIMITE_COUT_QUOTIDIEN_USD sur la journée UTC."""


def log(message: str):
    """Affiche un message dans la console avec la date et l'heure."""
    heure = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{heure}] {message}")


# ─────────────────────────────────────────────────────────────
#  GESTION DU PORTEFEUILLE FICTIF (sauvegarde entre redémarrages)
# ─────────────────────────────────────────────────────────────

def charger_portfolio() -> dict:
    """
    Charge l'état du portefeuille depuis le fichier JSON.
    Si le fichier n'existe pas (premier lancement), crée un portefeuille vide.
    """
    if os.path.exists(FICHIER_PORTFOLIO):
        with open(FICHIER_PORTFOLIO, "r") as f:
            data = json.load(f)
        data.setdefault("prix_initial_bh", 0.0)  # Rétrocompatibilité
        return data

    # Premier lancement : portefeuille initial
    return {
        "usdt_disponible": CAPITAL_DEPART_USDT,  # Capital en USDT (pas encore investi)
        "btc_en_stock":    0.0,                   # Quantité de BTC détenue
        "en_position":     False,                  # True si on a du BTC en cours
        "prix_achat_btc":  0.0,                   # Prix d'achat fictif du BTC
        "prix_initial_bh": 0.0,                   # Prix d'achat B&H (1er cycle valide)
    }


def sauvegarder_portfolio(portefeuille: dict):
    """Sauvegarde l'état du portefeuille dans le fichier JSON."""
    with open(FICHIER_PORTFOLIO, "w") as f:
        json.dump(portefeuille, f, indent=2)


# ─────────────────────────────────────────────────────────────
#  GESTION DU CACHE DES NEWS (mémoriser les news déjà vues)
# ─────────────────────────────────────────────────────────────

def charger_cache_news() -> list:
    """
    Charge la liste des IDs de news déjà analysées.
    Si le fichier cache n'existe pas, retourne une liste vide.
    """
    if not os.path.exists(FICHIER_CACHE_NEWS):
        return []
    with open(FICHIER_CACHE_NEWS, "r") as f:
        return json.load(f)


def sauvegarder_cache_news(ids: list):
    """
    Sauvegarde les IDs de news dans le fichier cache.
    Garde uniquement les TAILLE_MAX_CACHE derniers IDs pour éviter un fichier trop grand.
    """
    ids_a_garder = ids[-TAILLE_MAX_CACHE:]  # Ne garder que les plus récents
    with open(FICHIER_CACHE_NEWS, "w") as f:
        json.dump(ids_a_garder, f)


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 1 : RÉCUPÉRATION DES NEWS (NewsData.io)
# ─────────────────────────────────────────────────────────────

def recuperer_nouvelles_news() -> list:
    """
    Appelle l'API NewsData.io pour récupérer les dernières news crypto.
    Compare avec le cache pour ne retourner QUE les news pas encore vues.

    Retourne : liste de news (vide si aucune nouvelle depuis la dernière vérification).
    """
    log("Vérification des news sur NewsData.io...")

    url = "https://newsdata.io/api/1/news"
    parametres = {
        "apikey":   NEWSDATA_API_KEY,
        "q":        "bitcoin cryptocurrency BTC",  # Mots-clés de recherche
        "language": "en",                           # Anglais (meilleure couverture)
        "category": "business,technology",          # Catégories pertinentes
        "size":     10,                             # Max 10 articles par appel
    }

    try:
        reponse = requests.get(url, params=parametres, timeout=20)
        reponse.raise_for_status()   # Lève une erreur si le code HTTP n'est pas 200
        donnees = reponse.json()
    except requests.RequestException as e:
        log(f"⚠️  Erreur appel NewsData.io : {e}")
        return []

    if donnees.get("status") != "success":
        message_erreur = donnees.get("message", "inconnu")
        log(f"⚠️  Réponse inattendue NewsData.io : {message_erreur}")
        return []

    articles = donnees.get("results", [])

    # Charger les IDs des news déjà analysées
    ids_deja_vus = charger_cache_news()

    # Garder uniquement les news qu'on n'a pas encore vues
    nouvelles_news = []
    nouveaux_ids   = []
    for article in articles:
        article_id = article.get("article_id", "")
        if article_id and article_id not in ids_deja_vus:
            nouvelles_news.append(article)
            nouveaux_ids.append(article_id)

    # Mettre à jour le cache avec les nouveaux IDs
    if nouveaux_ids:
        sauvegarder_cache_news(ids_deja_vus + nouveaux_ids)
        log(f"✅ {len(nouvelles_news)} nouvelle(s) news détectée(s) !")
    else:
        log("Aucune nouvelle news depuis la dernière vérification. En attente...")

    return nouvelles_news


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 2 : RÉCUPÉRATION DU PRIX BTC (CoinGecko API publique)
# ─────────────────────────────────────────────────────────────

def recuperer_prix_btc() -> float:
    """
    Récupère le prix actuel du Bitcoin via l'API publique CoinGecko.
    Aucune clé API requise — fonctionne depuis n'importe quel serveur.

    Retourne : le prix en float (ex: 68234.50) ou 0.0 si erreur.
    """
    try:
        url = "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd"
        reponse = requests.get(url, timeout=10)
        reponse.raise_for_status()
        prix = float(reponse.json()["bitcoin"]["usd"])
        log(f"💰 Prix actuel du BTC : {prix:,.2f} USD (CoinGecko)")
        return prix

    except Exception as e:
        log(f"⚠️  Erreur récupération prix BTC : {e}")
        return 0.0


# ─────────────────────────────────────────────────────────────
#  COMPTEUR DE TOKENS / KILL-SWITCH BUDGÉTAIRE
# ─────────────────────────────────────────────────────────────

def charger_usage_quotidien() -> dict:
    """
    Charge le compteur de dépenses API du jour (UTC).
    Remet à zéro automatiquement si la date a changé.
    """
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    if os.path.exists(FICHIER_TOKEN_USAGE):
        with open(FICHIER_TOKEN_USAGE) as f:
            data = json.load(f)
        if data.get("date") == today:
            return data
    return {"date": today, "cout_total_usd": 0.0, "nb_appels": 0}


def enregistrer_et_verifier_tokens(modele: str, input_tokens: int, output_tokens: int) -> float:
    """
    Comptabilise les tokens du dernier appel Claude, sauvegarde, et déclenche le
    kill-switch si la limite quotidienne est atteinte.

    Retourne le coût de l'appel en USD.
    Lève LimiteQuotidienneAtteinte si le total dépasse LIMITE_COUT_QUOTIDIEN_USD.
    """
    tarifs   = COUTS_PAR_TOKEN.get(modele, {"input": 3.00e-6, "output": 15.00e-6})
    cout_appel = input_tokens * tarifs["input"] + output_tokens * tarifs["output"]

    usage = charger_usage_quotidien()
    usage["cout_total_usd"] += cout_appel
    usage["nb_appels"]      += 1

    with open(FICHIER_TOKEN_USAGE, "w") as f:
        json.dump(usage, f, indent=2)

    pct = usage["cout_total_usd"] / LIMITE_COUT_QUOTIDIEN_USD * 100
    log(f"💳 Appel : ${cout_appel:.4f} | Jour : ${usage['cout_total_usd']:.4f} "
        f"/ ${LIMITE_COUT_QUOTIDIEN_USD:.2f} ({pct:.1f}%)")

    if usage["cout_total_usd"] >= LIMITE_COUT_QUOTIDIEN_USD:
        raise LimiteQuotidienneAtteinte(
            f"Limite quotidienne de ${LIMITE_COUT_QUOTIDIEN_USD:.2f} atteinte "
            f"(total: ${usage['cout_total_usd']:.4f} en {usage['nb_appels']} appels). "
            "Agent arrêté pour éviter une surfacturation."
        )

    return cout_appel


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 3 & 4 : ANALYSE IA AVEC CLAUDE (Anthropic)
# ─────────────────────────────────────────────────────────────

def analyser_avec_claude(prix_btc: float, news: list) -> dict:
    """
    Envoie le prix BTC et les news à Claude pour une analyse de sentiment.

    Logique de choix du modèle :
    - Par défaut      → Haiku (rapide, économique)
    - Score fort ≥ 8  → Sonnet (plus puissant, analyse approfondie)

    Retourne un dictionnaire avec :
      - "score"          : de -10 (très baissier) à +10 (très haussier)
      - "recommandation" : "ACHETER", "VENDRE" ou "ATTENDRE"
      - "explication"    : 2-3 phrases d'explication
      - "modele_utilise" : nom du modèle Claude ayant produit le résultat
    """

    # Préparer un résumé lisible des news pour le prompt
    resume_news = ""
    for i, article in enumerate(news[:5], 1):  # Max 5 news pour limiter les coûts
        titre    = article.get("title", "Sans titre")
        source   = article.get("source_id", "Source inconnue")
        date_pub = article.get("pubDate", "")
        resume_news += f"\n{i}. [{date_pub}] {titre} (Source: {source})"

    # Prompt envoyé à Claude — instructions claires et format de réponse imposé
    prompt_analyse = f"""Tu es un analyste de trading crypto expérimenté et prudent.
Analyse les informations suivantes et donne une recommandation de trading BTC.

PRIX ACTUEL DU BITCOIN : {prix_btc:,.2f} USDT

DERNIÈRES NEWS CRYPTO :
{resume_news}

Réponds UNIQUEMENT avec un objet JSON valide dans ce format exact (rien d'autre autour) :
{{
  "score": <nombre entier entre -10 et 10>,
  "recommandation": "<ACHETER ou VENDRE ou ATTENDRE>",
  "explication": "<2-3 phrases expliquant ta décision en français>"
}}

Guide de scoring :
- +9 à +10 : signal haussier exceptionnel  → ACHETER
- +5 à +8  : signal haussier notable        → ACHETER
- -4 à +4  : signal neutre                  → ATTENDRE
- -5 à -8  : signal baissier notable        → VENDRE
- -9 à -10 : signal baissier exceptionnel  → VENDRE"""

    # ── Vérifier le budget avant tout appel ──
    usage = charger_usage_quotidien()
    if usage["cout_total_usd"] >= LIMITE_COUT_QUOTIDIEN_USD:
        raise LimiteQuotidienneAtteinte(
            f"Limite quotidienne de ${LIMITE_COUT_QUOTIDIEN_USD:.2f} déjà atteinte "
            f"(${usage['cout_total_usd']:.4f}). Aucun appel Claude émis."
        )

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # ── Première analyse avec Haiku (rapide et économique) ──
    log(f"🤖 Analyse en cours avec {MODELE_HAIKU}...")

    try:
        reponse = client.messages.create(
            model=MODELE_HAIKU,
            max_tokens=512,
            messages=[{"role": "user", "content": prompt_analyse}]
        )
        enregistrer_et_verifier_tokens(
            MODELE_HAIKU, reponse.usage.input_tokens, reponse.usage.output_tokens
        )
        contenu = reponse.content[0].text.strip()

        # Extraire le JSON de la réponse (Claude peut parfois ajouter du texte autour)
        debut = contenu.find("{")
        fin   = contenu.rfind("}") + 1
        contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu

        resultat = json.loads(contenu_json)
        resultat["modele_utilise"] = MODELE_HAIKU

    except LimiteQuotidienneAtteinte:
        raise  # Laisser remonter jusqu'au kill-switch dans main()
    except Exception as e:
        log(f"⚠️  Erreur avec Haiku : {e}")
        return {
            "score": 0,
            "recommandation": "ATTENDRE",
            "explication": "Analyse impossible suite à une erreur technique. Par prudence : ATTENDRE.",
            "modele_utilise": MODELE_HAIKU,
        }

    score = int(resultat.get("score", 0))
    log(f"📊 Score Haiku : {score:+d}/10 → {resultat.get('recommandation')}")

    # ── Si signal très fort, relancer avec Sonnet pour confirmer ──
    if abs(score) >= SEUIL_SIGNAL_FORT:
        log(f"🔥 Signal fort (score={score:+d}) → Analyse approfondie avec {MODELE_SONNET}...")

        prompt_approfondi = prompt_analyse + f"""

⚠️  ANALYSE APPROFONDIE DEMANDÉE
Une première analyse a donné un score de {score:+d}/10.
C'est un signal fort. Sois particulièrement rigoureux et prudent.
Confirme ou nuance cette évaluation avec tes propres conclusions."""

        try:
            reponse_sonnet = client.messages.create(
                model=MODELE_SONNET,
                max_tokens=1024,
                messages=[{"role": "user", "content": prompt_approfondi}]
            )
            enregistrer_et_verifier_tokens(
                MODELE_SONNET, reponse_sonnet.usage.input_tokens, reponse_sonnet.usage.output_tokens
            )
            contenu = reponse_sonnet.content[0].text.strip()
            debut   = contenu.find("{")
            fin     = contenu.rfind("}") + 1
            contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu

            resultat = json.loads(contenu_json)
            resultat["modele_utilise"] = MODELE_SONNET
            log(f"📊 Score Sonnet : {int(resultat.get('score', 0)):+d}/10 → {resultat.get('recommandation')}")

        except LimiteQuotidienneAtteinte:
            raise  # Laisser remonter jusqu'au kill-switch dans main()
        except Exception as e:
            # En cas d'erreur avec Sonnet, on garde le résultat Haiku
            log(f"⚠️  Erreur avec Sonnet, on conserve l'analyse Haiku : {e}")

    log(f"💬 Explication : {resultat.get('explication', '')}")
    return resultat


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 6 : SIMULATION DE TRADING (paper trading — argent fictif)
# ─────────────────────────────────────────────────────────────

def simuler_transaction(portefeuille: dict, recommandation: str, prix_btc: float) -> tuple:
    """
    Simule un achat ou une vente de BTC selon la recommandation de Claude.
    Applique les frais Binance (0.1%), le stop-loss (-5%) et le take-profit (+8%).
    ⚠️  AUCUN ordre réel n'est envoyé à Binance — 100% fictif.

    Retourne : (description_action, portefeuille_mis_à_jour)
    """
    action = "AUCUNE ACTION"
    raison_vente = ""  # Préfixe "STOP-LOSS" ou "TAKE-PROFIT" si déclenché automatiquement

    # ── Vérification Stop-Loss / Take-Profit (prioritaire sur la recommandation IA) ──
    if portefeuille["en_position"] and portefeuille["prix_achat_btc"] > 0:
        variation = prix_btc / portefeuille["prix_achat_btc"] - 1
        if variation <= -STOP_LOSS_PCT:
            log(f"🛑 STOP-LOSS déclenché ({variation*100:+.1f}%) → vente forcée")
            recommandation = "VENDRE"
            raison_vente = "[STOP-LOSS] "
        elif variation >= TAKE_PROFIT_PCT:
            log(f"🎯 TAKE-PROFIT déclenché ({variation*100:+.1f}%) → vente forcée")
            recommandation = "VENDRE"
            raison_vente = "[TAKE-PROFIT] "

    if recommandation == "ACHETER" and not portefeuille["en_position"]:
        # ── Simuler un achat avec frais ──
        capital = portefeuille["usdt_disponible"]
        frais = round(capital * FRAIS_TRADING, 4)
        btc_achete = (capital - frais) / prix_btc  # Frais déduits avant achat

        portefeuille["btc_en_stock"]   = btc_achete
        portefeuille["en_position"]     = True
        portefeuille["prix_achat_btc"]  = prix_btc
        portefeuille["usdt_disponible"] = 0.0

        action = (f"📈 ACHAT FICTIF : {btc_achete:.6f} BTC "
                  f"à {prix_btc:,.2f} USDT (investi: {capital:.2f} USDT, frais: {frais:.2f} USDT)")
        log(action)

    elif recommandation == "VENDRE" and portefeuille["en_position"]:
        # ── Simuler une vente avec frais ──
        valeur_brute  = portefeuille["btc_en_stock"] * prix_btc
        frais         = round(valeur_brute * FRAIS_TRADING, 4)
        valeur_vente  = valeur_brute - frais  # Frais déduits sur la vente
        cout_achat    = portefeuille["btc_en_stock"] * portefeuille["prix_achat_btc"]
        pnl_trade     = valeur_vente - cout_achat

        portefeuille["usdt_disponible"] = valeur_vente
        portefeuille["btc_en_stock"]    = 0.0
        portefeuille["en_position"]      = False
        portefeuille["prix_achat_btc"]   = 0.0

        signe = "+" if pnl_trade >= 0 else ""
        action = (f"{raison_vente}📉 VENTE FICTIVE : {valeur_vente:.2f} USDT récupérés "
                  f"(frais: {frais:.2f} USDT | P&L trade: {signe}{pnl_trade:.2f} USDT)")
        log(action)

    else:
        # Aucune action : soit ATTENDRE, soit impossible (ex: ACHETER mais déjà en position)
        valeur_actuelle = calculer_valeur_portfolio(portefeuille, prix_btc)
        if portefeuille["en_position"]:
            variation = (prix_btc / portefeuille["prix_achat_btc"] - 1) * 100
            action = (f"⏳ EN POSITION — Valeur: {valeur_actuelle:.2f} USDT "
                      f"({variation:+.1f}% depuis achat à {portefeuille['prix_achat_btc']:,.0f} USDT)")
        else:
            action = f"⏳ EN ATTENTE — Capital disponible : {valeur_actuelle:.2f} USDT"
        log(action)

    return action, portefeuille


def calculer_valeur_portfolio(portefeuille: dict, prix_btc: float) -> float:
    """
    Calcule la valeur totale du portfolio fictif en USDT au prix actuel.
    USDT disponible + valeur des BTC détenus converti en USDT.
    """
    valeur = portefeuille["usdt_disponible"]
    if portefeuille["en_position"] and portefeuille["btc_en_stock"] > 0:
        valeur += portefeuille["btc_en_stock"] * prix_btc
    return valeur


def _calculer_valeur_bh(portefeuille: dict, prix_btc: float) -> float:
    """Valeur du portfolio Buy & Hold de référence : achat fictif au 1er prix enregistré."""
    prix_initial = portefeuille.get("prix_initial_bh", 0.0)
    if prix_initial <= 0 or prix_btc <= 0:
        return CAPITAL_DEPART_USDT
    return CAPITAL_DEPART_USDT / prix_initial * prix_btc


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 5 : ENREGISTREMENT DANS EXCEL
# ─────────────────────────────────────────────────────────────

EN_TETES_EXCEL = [
    "Date",
    "Heure",
    "Prix BTC (USDT)",
    "Nb news analysées",
    "Titres des news",
    "Score sentiment",
    "Modèle Claude utilisé",
    "Recommandation",
    "Action simulée",
    "Capital fictif (USDT)",
    "P&L total (USDT)",
    "Valeur B&H (USDT)",
    "P&L B&H (USDT)",
]


def initialiser_excel():
    """
    Crée le fichier Excel avec les en-têtes si il n'existe pas encore.
    Si le fichier existe déjà, ajoute les colonnes B&H manquantes si nécessaire.
    """
    bold = openpyxl.styles.Font(bold=True)

    if not os.path.exists(FICHIER_EXCEL):
        classeur = openpyxl.Workbook()
        feuille  = classeur.active
        feuille.title = "Journal de Trading"
        feuille.append(EN_TETES_EXCEL)
        for cellule in feuille[1]:
            cellule.font = bold
        classeur.save(FICHIER_EXCEL)
        log(f"📁 Fichier Excel créé : {FICHIER_EXCEL}")
        return

    # Migration : ajouter les colonnes B&H si absentes
    classeur = openpyxl.load_workbook(FICHIER_EXCEL)
    feuille  = classeur.active
    en_tetes_actuels = [c.value for c in feuille[1]]
    if "Valeur B&H (USDT)" not in en_tetes_actuels:
        col = len(en_tetes_actuels) + 1
        cell_bh  = feuille.cell(row=1, column=col,   value="Valeur B&H (USDT)")
        cell_pnl = feuille.cell(row=1, column=col+1, value="P&L B&H (USDT)")
        cell_bh.font  = bold
        cell_pnl.font = bold
        classeur.save(FICHIER_EXCEL)
        log("📁 Colonnes Buy & Hold ajoutées au fichier Excel existant.")


def enregistrer_dans_excel(prix_btc: float, news: list, analyse: dict,
                            action: str, valeur_portfolio: float, valeur_bh: float):
    """
    Ajoute une nouvelle ligne dans le fichier Excel avec toutes les données du cycle,
    y compris la comparaison avec la stratégie Buy & Hold.
    """
    pnl_total = valeur_portfolio - CAPITAL_DEPART_USDT
    pnl_bh    = valeur_bh - CAPITAL_DEPART_USDT

    titres_news = " | ".join(
        article.get("title", "Sans titre")[:80]
        for article in news[:5]
    ) or "Aucune news"

    maintenant = datetime.now(timezone.utc)
    nouvelle_ligne = [
        maintenant.strftime("%Y-%m-%d"),
        maintenant.strftime("%H:%M:%S"),
        round(prix_btc, 2),
        len(news),
        titres_news,
        analyse.get("score", 0),
        analyse.get("modele_utilise", "N/A"),
        analyse.get("recommandation", "ATTENDRE"),
        action,
        round(valeur_portfolio, 2),
        round(pnl_total, 2),
        round(valeur_bh, 2),
        round(pnl_bh, 2),
    ]

    classeur = openpyxl.load_workbook(FICHIER_EXCEL)
    feuille  = classeur.active
    feuille.append(nouvelle_ligne)
    classeur.save(FICHIER_EXCEL)

    signe_agent = "+" if pnl_total >= 0 else ""
    signe_bh    = "+" if pnl_bh >= 0 else ""
    log(f"💾 Excel — P&L agent : {signe_agent}{pnl_total:.2f} USDT | P&L B&H : {signe_bh}{pnl_bh:.2f} USDT")


# ═════════════════════════════════════════════════════════════
#  CYCLE PRINCIPAL
# ═════════════════════════════════════════════════════════════

def executer_un_cycle(portefeuille: dict) -> dict:
    """
    Exécute UN cycle complet de l'agent :
    1. Vérifier les nouvelles news
    2. Si nouvelles news → récupérer le prix BTC
    3. Analyser avec Claude (Haiku ou Sonnet selon le signal)
    4. Simuler une transaction (paper trading)
    5. Enregistrer dans Excel
    6. Retourner le portefeuille mis à jour

    Si aucune nouvelle news, retourne le portefeuille inchangé.
    """
    separateur = "─" * 55
    log(separateur)
    log("🔄 Nouveau cycle d'analyse")
    log(separateur)

    # ── Étape 1 : Vérifier les nouvelles news ──
    nouvelles_news = recuperer_nouvelles_news()

    if not nouvelles_news:
        log("💤 Pas de nouvelles news ce cycle — écriture heartbeat dans Excel.")
        prix_btc = recuperer_prix_btc()
        valeur_portfolio = calculer_valeur_portfolio(portefeuille, prix_btc)
        valeur_bh = _calculer_valeur_bh(portefeuille, prix_btc)
        enregistrer_dans_excel(
            prix_btc, [],
            {"score": 0, "recommandation": "ATTENDRE", "modele_utilise": "—"},
            "ATTENDRE",
            valeur_portfolio,
            valeur_bh,
        )
        return portefeuille

    # ── Étape 2 : Récupérer le prix BTC ──
    prix_btc = recuperer_prix_btc()
    if prix_btc == 0.0:
        log("⚠️  Prix BTC indisponible — cycle annulé par sécurité.")
        return portefeuille

    # ── Initialiser le prix de référence Buy & Hold au premier cycle valide ──
    if portefeuille.get("prix_initial_bh", 0.0) == 0.0:
        portefeuille["prix_initial_bh"] = prix_btc
        log(f"📊 Prix de référence Buy & Hold initialisé : {prix_btc:,.2f} USDT")

    # ── Étapes 3 & 4 : Analyser avec Claude ──
    analyse = analyser_avec_claude(prix_btc, nouvelles_news)

    # ── Étape 6 : Simuler une transaction ──
    recommandation = analyse.get("recommandation", "ATTENDRE")
    action, portefeuille = simuler_transaction(portefeuille, recommandation, prix_btc)

    # ── Sauvegarder le portefeuille ──
    sauvegarder_portfolio(portefeuille)

    # ── Étape 5 : Enregistrer dans Excel ──
    valeur_portfolio = calculer_valeur_portfolio(portefeuille, prix_btc)
    valeur_bh = _calculer_valeur_bh(portefeuille, prix_btc)
    enregistrer_dans_excel(prix_btc, nouvelles_news, analyse, action, valeur_portfolio, valeur_bh)

    pnl = valeur_portfolio - CAPITAL_DEPART_USDT
    pnl_bh = valeur_bh - CAPITAL_DEPART_USDT
    signe    = "+" if pnl >= 0 else ""
    signe_bh = "+" if pnl_bh >= 0 else ""
    log(f"💼 Agent : {valeur_portfolio:.2f} USDT (P&L: {signe}{pnl:.2f}) | "
        f"B&H : {valeur_bh:.2f} USDT (P&L: {signe_bh}{pnl_bh:.2f})")

    return portefeuille


# ═════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE PRINCIPAL
# ═════════════════════════════════════════════════════════════

def main():
    """
    Lance l'agent de trading.

    Deux modes de fonctionnement :
    - Mode local (Mac)           : boucle infinie, attend 30 min entre chaque cycle
                                   → lance avec : python3 trading_agent.py
    - Mode GitHub Actions (--once) : exécute UN seul cycle puis s'arrête
                                   → GitHub Actions relance le script toutes les 30 min
                                   → lance avec : python3 trading_agent.py --once
    """
    # Détecter si on est en mode "un seul cycle" (GitHub Actions)
    mode_unique = "--once" in sys.argv

    log("=" * 55)
    log("🚀 DÉMARRAGE DE L'AGENT DE TRADING CRYPTO")
    if mode_unique:
        log("   Mode : GitHub Actions (un seul cycle)")
    else:
        log(f"   Mode : local (boucle toutes les {INTERVALLE_SECONDES // 60} min)")
    log(f"   Capital fictif de départ : {CAPITAL_DEPART_USDT:.2f} USDT")
    log(f"   Fichier Excel            : {FICHIER_EXCEL}")
    log(f"   Seuil signal fort        : ±{SEUIL_SIGNAL_FORT}/10 → Sonnet")
    log(f"   Frais trading            : {FRAIS_TRADING*100:.1f}% par trade")
    log(f"   Stop-loss                : -{STOP_LOSS_PCT*100:.0f}%")
    log(f"   Take-profit              : +{TAKE_PROFIT_PCT*100:.0f}%")
    log(f"   Limite coût API/jour     : ${LIMITE_COUT_QUOTIDIEN_USD:.2f}")
    usage_actuel = charger_usage_quotidien()
    log(f"   Dépense API aujourd'hui  : ${usage_actuel['cout_total_usd']:.4f} "
        f"({usage_actuel['nb_appels']} appels)")
    log("=" * 55)
    log("")

    # ── Vérification que les clés API sont bien renseignées ──
    cles_manquantes = []
    if not ANTHROPIC_API_KEY: cles_manquantes.append("ANTHROPIC_API_KEY")
    if not NEWSDATA_API_KEY:  cles_manquantes.append("NEWSDATA_API_KEY")

    if cles_manquantes:
        log(f"❌ ERREUR : Clés API manquantes :")
        for cle in cles_manquantes:
            log(f"   • {cle}")
        log("   En local : remplis le fichier .env")
        log("   Sur GitHub : ajoute les secrets dans Settings → Secrets → Actions")
        sys.exit(1)  # Quitter avec une erreur pour signaler l'échec à GitHub Actions

    if not BINANCE_API_KEY:
        log("ℹ️  Pas de clé Binance — le prix BTC sera quand même récupéré (lecture publique).")

    # ── Initialiser le fichier Excel et charger le portefeuille ──
    initialiser_excel()
    portefeuille = charger_portfolio()

    if portefeuille["en_position"]:
        log(f"📂 Portfolio chargé — En position BTC ({portefeuille['btc_en_stock']:.6f} BTC)")
    else:
        log(f"📂 Portfolio chargé — USDT disponible : {portefeuille['usdt_disponible']:.2f}")

    # ════════════════════════════════════════════
    #  MODE GITHUB ACTIONS : un seul cycle puis exit
    # ════════════════════════════════════════════
    if mode_unique:
        log("\n▶  Exécution du cycle unique (mode GitHub Actions)...")
        try:
            executer_un_cycle(portefeuille)
        except LimiteQuotidienneAtteinte as e:
            log(f"🚨 KILL-SWITCH BUDGÉTAIRE : {e}")
            sys.exit(1)
        except Exception as e:
            log(f"❌ Erreur : {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        log("✅ Cycle terminé. GitHub Actions sauvegardera les fichiers.")
        return

    # ════════════════════════════════════════════
    #  MODE LOCAL : boucle infinie toutes les 30 min
    # ════════════════════════════════════════════
    numero_cycle = 0
    while True:
        numero_cycle += 1
        log(f"\n{'═' * 55}")
        log(f"CYCLE N°{numero_cycle}")

        try:
            portefeuille = executer_un_cycle(portefeuille)
        except KeyboardInterrupt:
            log("\n⛔ Arrêt demandé (Ctrl+C). À bientôt !")
            break
        except LimiteQuotidienneAtteinte as e:
            log(f"🚨 KILL-SWITCH BUDGÉTAIRE : {e}")
            break
        except Exception as e:
            log(f"❌ Erreur inattendue : {e}")
            import traceback
            traceback.print_exc()
            log("   L'agent continue malgré l'erreur...")

        log(f"\n⏰ Prochain cycle dans {INTERVALLE_SECONDES // 60} minutes...")
        try:
            time.sleep(INTERVALLE_SECONDES)
        except KeyboardInterrupt:
            log("\n⛔ Arrêt demandé. À bientôt !")
            break


# Ce bloc s'exécute uniquement si on lance ce fichier directement
# (pas si on l'importe depuis un autre script)
if __name__ == "__main__":
    main()
