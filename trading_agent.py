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
import ccxt         # Se connecter à Binance pour lire les prix
import anthropic    # Utiliser l'IA Claude pour analyser les news
import openpyxl     # Lire et écrire des fichiers Excel
from datetime import datetime   # Gérer les dates et heures
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


# ═════════════════════════════════════════════════════════════
#  FONCTIONS UTILITAIRES
# ═════════════════════════════════════════════════════════════

def log(message: str):
    """Affiche un message dans la console avec la date et l'heure."""
    heure = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
            return json.load(f)

    # Premier lancement : portefeuille initial
    return {
        "usdt_disponible": CAPITAL_DEPART_USDT,  # Capital en USDT (pas encore investi)
        "btc_en_stock":    0.0,                   # Quantité de BTC détenue
        "en_position":     False,                  # True si on a du BTC en cours
        "prix_achat_btc":  0.0,                   # Prix d'achat fictif du BTC
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
#  ÉTAPE 2 : RÉCUPÉRATION DU PRIX BTC (Binance via ccxt)
# ─────────────────────────────────────────────────────────────

def recuperer_prix_btc() -> float:
    """
    Récupère le prix actuel du Bitcoin (BTC/USDT) sur Binance.
    Utilise ccxt en mode lecture seule (aucun ordre, aucun risque).

    Retourne : le prix en float (ex: 68234.50) ou 0.0 si erreur.
    """
    try:
        # Connexion à Binance (la lecture des prix est publique)
        exchange = ccxt.binance({
            "apiKey": BINANCE_API_KEY,
            "secret": BINANCE_SECRET,
            "options": {"defaultType": "spot"},  # Marché spot (pas les futures)
        })

        # Récupérer le ticker = les informations de prix actuelles
        ticker = exchange.fetch_ticker("BTC/USDT")
        prix   = float(ticker["last"])  # "last" = dernier prix de transaction

        log(f"💰 Prix actuel du BTC : {prix:,.2f} USDT")
        return prix

    except Exception as e:
        log(f"⚠️  Erreur récupération prix BTC : {e}")
        return 0.0


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

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # ── Première analyse avec Haiku (rapide et économique) ──
    log(f"🤖 Analyse en cours avec {MODELE_HAIKU}...")

    try:
        reponse = client.messages.create(
            model=MODELE_HAIKU,
            max_tokens=512,
            messages=[{"role": "user", "content": prompt_analyse}]
        )
        contenu = reponse.content[0].text.strip()

        # Extraire le JSON de la réponse (Claude peut parfois ajouter du texte autour)
        debut = contenu.find("{")
        fin   = contenu.rfind("}") + 1
        contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu

        resultat = json.loads(contenu_json)
        resultat["modele_utilise"] = MODELE_HAIKU

    except Exception as e:
        log(f"⚠️  Erreur avec Haiku : {e}")
        # En cas d'erreur, on retourne une réponse neutre par sécurité
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
            contenu = reponse_sonnet.content[0].text.strip()
            debut   = contenu.find("{")
            fin     = contenu.rfind("}") + 1
            contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu

            resultat = json.loads(contenu_json)
            resultat["modele_utilise"] = MODELE_SONNET
            log(f"📊 Score Sonnet : {int(resultat.get('score', 0)):+d}/10 → {resultat.get('recommandation')}")

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
    ⚠️  AUCUN ordre réel n'est envoyé à Binance — 100% fictif.

    Retourne : (description_action, portefeuille_mis_à_jour)
    """
    action = "AUCUNE ACTION"

    if recommandation == "ACHETER" and not portefeuille["en_position"]:
        # ── Simuler un achat ──
        # On investit tout le capital USDT disponible en BTC
        capital = portefeuille["usdt_disponible"]
        btc_achete = capital / prix_btc  # Quantité de BTC fictive achetée

        portefeuille["btc_en_stock"]   = btc_achete
        portefeuille["en_position"]     = True
        portefeuille["prix_achat_btc"]  = prix_btc
        portefeuille["usdt_disponible"] = 0.0  # Tout l'argent est "investi"

        action = (f"📈 ACHAT FICTIF : {btc_achete:.6f} BTC "
                  f"à {prix_btc:,.2f} USDT (investi: {capital:.2f} USDT)")
        log(action)

    elif recommandation == "VENDRE" and portefeuille["en_position"]:
        # ── Simuler une vente ──
        valeur_vente  = portefeuille["btc_en_stock"] * prix_btc
        cout_achat    = portefeuille["btc_en_stock"] * portefeuille["prix_achat_btc"]
        pnl_trade     = valeur_vente - cout_achat  # Profit ou perte sur ce trade

        portefeuille["usdt_disponible"] = valeur_vente  # On récupère l'argent
        portefeuille["btc_en_stock"]    = 0.0
        portefeuille["en_position"]      = False
        portefeuille["prix_achat_btc"]   = 0.0

        signe = "+" if pnl_trade >= 0 else ""
        action = (f"📉 VENTE FICTIVE : {valeur_vente:.2f} USDT récupérés "
                  f"(P&L trade: {signe}{pnl_trade:.2f} USDT)")
        log(action)

    else:
        # Aucune action : soit ATTENDRE, soit impossible (ex: ACHETER mais déjà en position)
        valeur_actuelle = calculer_valeur_portfolio(portefeuille, prix_btc)
        if portefeuille["en_position"]:
            action = (f"⏳ EN POSITION — Valeur portfolio: {valeur_actuelle:.2f} USDT "
                      f"(BTC en stock à {prix_btc:,.2f} USDT)")
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


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 5 : ENREGISTREMENT DANS EXCEL
# ─────────────────────────────────────────────────────────────

def initialiser_excel():
    """
    Crée le fichier Excel avec les en-têtes si il n'existe pas encore.
    Si le fichier existe déjà, ne fait rien (on continue à ajouter des lignes).
    """
    if os.path.exists(FICHIER_EXCEL):
        return  # Déjà existant, on ne le réinitialise pas

    classeur = openpyxl.Workbook()
    feuille  = classeur.active
    feuille.title = "Journal de Trading"

    en_tetes = [
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
    ]
    feuille.append(en_tetes)

    # Mettre les en-têtes en gras
    for cellule in feuille[1]:
        cellule.font = openpyxl.styles.Font(bold=True)

    classeur.save(FICHIER_EXCEL)
    log(f"📁 Fichier Excel créé : {FICHIER_EXCEL}")


def enregistrer_dans_excel(prix_btc: float, news: list, analyse: dict,
                            action: str, valeur_portfolio: float):
    """
    Ajoute une nouvelle ligne dans le fichier Excel avec toutes les données du cycle.
    """
    pnl_total = valeur_portfolio - CAPITAL_DEPART_USDT  # Gain ou perte depuis le départ

    # Concaténer les titres des news (séparés par " | ")
    titres_news = " | ".join(
        article.get("title", "Sans titre")[:80]  # Max 80 caractères par titre
        for article in news[:5]
    ) or "Aucune news"

    maintenant  = datetime.now()
    nouvelle_ligne = [
        maintenant.strftime("%Y-%m-%d"),           # Date
        maintenant.strftime("%H:%M:%S"),            # Heure
        round(prix_btc, 2),                         # Prix BTC
        len(news),                                  # Nombre de news
        titres_news,                                # Titres des news
        analyse.get("score", 0),                    # Score sentiment
        analyse.get("modele_utilise", "N/A"),        # Modèle Claude
        analyse.get("recommandation", "ATTENDRE"),  # Recommandation
        action,                                     # Action effectuée
        round(valeur_portfolio, 2),                 # Capital actuel
        round(pnl_total, 2),                        # P&L total depuis le départ
    ]

    classeur = openpyxl.load_workbook(FICHIER_EXCEL)
    feuille  = classeur.active
    feuille.append(nouvelle_ligne)
    classeur.save(FICHIER_EXCEL)

    signe = "+" if pnl_total >= 0 else ""
    log(f"💾 Excel mis à jour — P&L total : {signe}{pnl_total:.2f} USDT")


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
        log("💤 Pas de nouvelles news ce cycle.")
        prix_btc = recuperer_prix_btc()
        if prix_btc > 0:
            valeur_portfolio = calculer_valeur_portfolio(portefeuille, prix_btc)
            enregistrer_dans_excel(
                prix_btc, [],
                {"score": 0, "recommandation": "ATTENDRE", "modele_utilise": "—"},
                "ATTENDRE",
                valeur_portfolio,
            )
        return portefeuille

    # ── Étape 2 : Récupérer le prix BTC ──
    prix_btc = recuperer_prix_btc()
    if prix_btc == 0.0:
        log("⚠️  Prix BTC indisponible — cycle annulé par sécurité.")
        return portefeuille

    # ── Étapes 3 & 4 : Analyser avec Claude ──
    analyse = analyser_avec_claude(prix_btc, nouvelles_news)

    # ── Étape 6 : Simuler une transaction ──
    recommandation = analyse.get("recommandation", "ATTENDRE")
    action, portefeuille = simuler_transaction(portefeuille, recommandation, prix_btc)

    # ── Sauvegarder le portefeuille ──
    sauvegarder_portfolio(portefeuille)

    # ── Étape 5 : Enregistrer dans Excel ──
    valeur_portfolio = calculer_valeur_portfolio(portefeuille, prix_btc)
    enregistrer_dans_excel(prix_btc, nouvelles_news, analyse, action, valeur_portfolio)

    pnl = valeur_portfolio - CAPITAL_DEPART_USDT
    signe = "+" if pnl >= 0 else ""
    log(f"💼 Portfolio : {valeur_portfolio:.2f} USDT (départ: {CAPITAL_DEPART_USDT:.2f} | P&L: {signe}{pnl:.2f})")

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
