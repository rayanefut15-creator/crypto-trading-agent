#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================
  AGENT DE TRADING CRYPTO — PAPER TRADING (ARGENT FICTIF)
=============================================================
Ce script surveille les nouvelles crypto toutes les 15 minutes,
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
import feedparser   # Lire les flux RSS (CoinDesk, Cointelegraph)
import requests     # Récupérer le prix BTC (CoinGecko / Binance)
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

# ─────────────────────────────────────────────────────────────
#  PARAMÈTRES GÉNÉRAUX (tu peux les modifier ici)
# ─────────────────────────────────────────────────────────────
CAPITAL_DEPART_USDT  = 300.0         # Capital fictif de départ (en USDT ≈ EUR)
INTERVALLE_SECONDES  = 15 * 60       # 15 minutes entre chaque vérification
FICHIER_EXCEL        = "trading_journal.xlsx"   # Nom du fichier Excel
FICHIER_CACHE_NEWS   = "cache_news.json"         # Mémorisation des news déjà vues
FICHIER_PORTFOLIO    = "portfolio_state.json"    # Sauvegarde du portefeuille fictif
CACHE_EXPIRY_JOURS   = 7             # Expirer les news du cache après 7 jours

# Sources RSS
RSS_FEEDS = [
    "https://www.coindesk.com/arc/outboundfeeds/rss/",
    "https://cointelegraph.com/rss",
]

# Mots indicateurs de spam/publicité (filtre insensible à la casse)
MOTS_SPAM = {
    "casino", "presale", "pre-sale", "100x", "alphapepe", "pepeto",
    "gambling", "airdrop", "giveaway", "prize", "sponsor",
    "advertisement", "promoted",
}

# Modèles Claude
MODELE_HAIKU  = "claude-haiku-4-5-20251001"  # Rapide et économique
MODELE_SONNET = "claude-sonnet-4-6"           # Plus puissant (signaux forts)

# Seuil pour choisir Sonnet : nombre minimum de news pertinentes détectées par Haiku
SEUIL_NEWS_PERTINENTES = 2  # ≥ 2 news pertinentes → Sonnet ; sinon → Haiku

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

def charger_cache_news() -> dict:
    """
    Charge le cache des news : dict {url: timestamp_unix}.
    Expire automatiquement les entrées de plus de CACHE_EXPIRY_JOURS jours.
    Gère la migration depuis l'ancien format (liste d'IDs).
    """
    if not os.path.exists(FICHIER_CACHE_NEWS):
        return {}
    with open(FICHIER_CACHE_NEWS, "r") as f:
        try:
            cache = json.load(f)
        except json.JSONDecodeError:
            return {}
    if isinstance(cache, list):
        return {}  # Migration depuis l'ancien format liste → dict vide
    limite = time.time() - CACHE_EXPIRY_JOURS * 86400
    return {url: ts for url, ts in cache.items() if ts > limite}


def sauvegarder_cache_news(cache: dict):
    """Sauvegarde le cache {url: timestamp_unix} dans le fichier JSON."""
    with open(FICHIER_CACHE_NEWS, "w") as f:
        json.dump(cache, f)


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 1 : RÉCUPÉRATION DES NEWS (RSS CoinDesk + Cointelegraph)
# ─────────────────────────────────────────────────────────────

def _contient_spam(texte: str) -> bool:
    """Retourne True si le texte contient un mot de la liste MOTS_SPAM."""
    texte_lower = texte.lower()
    return any(mot in texte_lower for mot in MOTS_SPAM)


def recuperer_nouvelles_news() -> list:
    """
    Lit les flux RSS CoinDesk et Cointelegraph, filtre le spam,
    déduplique par titre normalisé et compare avec le cache (URL comme clé)
    pour ne retourner que les articles pas encore vus.
    """
    log("Vérification des news RSS (CoinDesk + Cointelegraph)...")

    tous_articles = []
    for feed_url in RSS_FEEDS:
        try:
            feed = feedparser.parse(feed_url)
            for entry in feed.entries:
                titre = entry.get("title", "").strip()
                lien  = entry.get("link", "").strip()
                if not titre or not lien:
                    continue
                tous_articles.append({
                    "title":     titre,
                    "link":      lien,
                    "source_id": feed.feed.get("title", feed_url),
                    "pubDate":   entry.get("published", ""),
                    "summary":   entry.get("summary", ""),
                })
        except Exception as e:
            log(f"⚠️  Erreur lecture RSS {feed_url} : {e}")

    # Filtre anti-spam sur titre et résumé
    articles_filtres = [
        a for a in tous_articles
        if not _contient_spam(a["title"]) and not _contient_spam(a.get("summary", ""))
    ]

    # Déduplique par titre normalisé (même article repris des deux sources)
    titres_vus: set = set()
    articles_uniques = []
    for article in articles_filtres:
        titre_norm = article["title"].lower().strip()
        if titre_norm not in titres_vus:
            titres_vus.add(titre_norm)
            articles_uniques.append(article)

    # Filtre anti-doublons via le cache (URL comme identifiant unique)
    cache = charger_cache_news()
    maintenant = time.time()

    nouvelles_news = []
    for article in articles_uniques:
        url = article["link"]
        if url not in cache:
            nouvelles_news.append(article)
            cache[url] = maintenant

    sauvegarder_cache_news(cache)

    if nouvelles_news:
        log(f"✅ {len(nouvelles_news)} nouvelle(s) news détectée(s) "
            f"({len(articles_filtres)} après filtre spam, {len(articles_uniques)} après dédup)")
    else:
        log("Aucune nouvelle news depuis la dernière vérification. En attente...")

    return nouvelles_news


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 2 : RÉCUPÉRATION DU PRIX BTC (CoinGecko → Binance en backup)
# ─────────────────────────────────────────────────────────────

def recuperer_prix_btc() -> float:
    """
    Récupère le prix actuel du Bitcoin.
    Source primaire  : CoinGecko (API publique, sans clé)
    Source de backup : Binance ticker public (sans clé)

    Retourne : le prix en float (ex: 68234.50) ou 0.0 si les deux sources échouent.
    """
    # ── Source 1 : CoinGecko ──
    try:
        url = "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=usd"
        reponse = requests.get(url, timeout=10)
        reponse.raise_for_status()
        prix = float(reponse.json()["bitcoin"]["usd"])
        log(f"💰 Prix actuel du BTC : {prix:,.2f} USD (CoinGecko)")
        return prix
    except Exception as e:
        log(f"⚠️  CoinGecko indisponible ({e}) — tentative Binance...")

    # ── Source 2 : Binance ticker public (backup) ──
    try:
        url = "https://api.binance.com/api/v3/ticker/price?symbol=BTCUSDT"
        reponse = requests.get(url, timeout=10)
        reponse.raise_for_status()
        prix = float(reponse.json()["price"])
        log(f"💰 Prix actuel du BTC : {prix:,.2f} USD (Binance backup)")
        return prix
    except Exception as e:
        log(f"⚠️  Binance indisponible ({e}) — prix BTC introuvable.")
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
#  ÉTAPE 3 : PRÉ-FILTRAGE DES NEWS PAR HAIKU
# ─────────────────────────────────────────────────────────────

def prefiltrer_news_haiku(articles: list, client: anthropic.Anthropic) -> tuple:
    """
    Haiku note chaque news de 0 à 3 :
      0 = pub / spam / hors sujet
      1 = news crypto générale (altcoins, NFT, DeFi sans lien BTC)
      2 = news crypto directement pertinente pour le prix du BTC
      3 = news macro importante (Fed, inflation, régulation, ETF Bitcoin)

    Retourne (articles_qualite_2_3, nb_qualite).
    En cas d'erreur, retourne (articles, len(articles)) en fallback.
    """
    if not articles:
        return [], 0

    liste_titres = ""
    for i, article in enumerate(articles, 1):
        liste_titres += f"\n{i}. {article.get('title', 'Sans titre')}"

    prompt_filtre = f"""Note chaque news de 0 à 3 selon son importance pour le trading BTC :
  0 = pub / spam / hors sujet crypto
  1 = news crypto générale (altcoins, NFT, DeFi sans lien BTC)
  2 = news crypto directement pertinente pour le prix du BTC
  3 = news macro importante (Fed, inflation, régulation, ETF Bitcoin, crise financière)

NEWS :{liste_titres}

Réponds UNIQUEMENT avec un objet JSON valide (rien d'autre autour) :
{{
  "scores": [2, 0, 3, 1, ...]
}}
L'ordre doit correspondre exactement à la liste. Chaque score est 0, 1, 2 ou 3."""

    try:
        reponse = client.messages.create(
            model=MODELE_HAIKU,
            max_tokens=256,
            messages=[{"role": "user", "content": prompt_filtre}]
        )
        enregistrer_et_verifier_tokens(
            MODELE_HAIKU, reponse.usage.input_tokens, reponse.usage.output_tokens
        )
        contenu = reponse.content[0].text.strip()
        debut = contenu.find("{")
        fin   = contenu.rfind("}") + 1
        contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu
        scores = json.loads(contenu_json).get("scores", [])

        news_qualite = [a for a, s in zip(articles, scores) if s >= 2]
        nb_qualite   = len(news_qualite)
        nb_niveau3   = sum(1 for s in scores if s == 3)
        log(f"🔍 Scoring Haiku : {nb_qualite}/{len(articles)} news qualité 2-3 "
            f"({nb_niveau3} niveau 3 macro)")
        return news_qualite, nb_qualite

    except LimiteQuotidienneAtteinte:
        raise
    except Exception as e:
        log(f"⚠️  Erreur scoring Haiku ({e}) — toutes les news conservées")
        return articles, len(articles)


# ─────────────────────────────────────────────────────────────
#  ÉTAPE 4 : DÉCISION FINALE (Haiku ou Sonnet selon le filtre)
# ─────────────────────────────────────────────────────────────

def analyser_avec_claude(prix_btc: float, news: list) -> dict:
    """
    Pipeline d'analyse en deux étapes :
    1. Haiku pré-filtre chaque news (oui/non : impact BTC dans 24h ?)
    2. Si ≥ SEUIL_NEWS_PERTINENTES news pertinentes → Sonnet décide
       Sinon → Haiku décide directement

    Retourne un dictionnaire avec :
      - "score"          : de -10 (très baissier) à +10 (très haussier)
      - "recommandation" : "ACHETER", "VENDRE" ou "ATTENDRE"
      - "explication"    : 2-3 phrases d'explication
      - "modele_utilise" : nom du modèle Claude ayant pris la décision finale
    """
    usage = charger_usage_quotidien()
    if usage["cout_total_usd"] >= LIMITE_COUT_QUOTIDIEN_USD:
        raise LimiteQuotidienneAtteinte(
            f"Limite quotidienne de ${LIMITE_COUT_QUOTIDIEN_USD:.2f} déjà atteinte "
            f"(${usage['cout_total_usd']:.4f}). Aucun appel Claude émis."
        )

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # ── Étape 1 : Scoring Haiku (0-3) ──
    news_pertinentes, nb_qualite = prefiltrer_news_haiku(news, client)
    nb_pertinentes = nb_qualite

    # ── Choix du modèle décisionnel ──
    if nb_pertinentes >= SEUIL_NEWS_PERTINENTES:
        modele_decision = MODELE_SONNET
        log(f"📡 {nb_pertinentes} news qualité 2-3 → décision confiée à {MODELE_SONNET}")
    else:
        modele_decision = MODELE_HAIKU
        log(f"📡 {nb_pertinentes} news qualité 2-3 → décision par {MODELE_HAIKU}")

    # Utiliser les news pertinentes si disponibles, sinon toutes (fallback 0 pertinente)
    news_pour_analyse = news_pertinentes if news_pertinentes else news

    resume_news = ""
    for i, article in enumerate(news_pour_analyse[:5], 1):
        titre    = article.get("title", "Sans titre")
        source   = article.get("source_id", "Source inconnue")
        date_pub = article.get("pubDate", "")
        resume_news += f"\n{i}. [{date_pub}] {titre} (Source: {source})"

    prompt_analyse = f"""Tu es un analyste de trading crypto expérimenté et prudent.
Analyse les informations suivantes et donne une recommandation de trading BTC.

PRIX ACTUEL DU BITCOIN : {prix_btc:,.2f} USDT

NEWS QUALITÉ 2-3 ({nb_qualite} sélectionnées sur {len(news)}) :
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

    log(f"🤖 Décision en cours avec {modele_decision}...")

    try:
        reponse = client.messages.create(
            model=modele_decision,
            max_tokens=512 if modele_decision == MODELE_HAIKU else 1024,
            messages=[{"role": "user", "content": prompt_analyse}]
        )
        enregistrer_et_verifier_tokens(
            modele_decision, reponse.usage.input_tokens, reponse.usage.output_tokens
        )
        contenu = reponse.content[0].text.strip()
        debut   = contenu.find("{")
        fin     = contenu.rfind("}") + 1
        contenu_json = contenu[debut:fin] if debut >= 0 and fin > debut else contenu

        resultat = json.loads(contenu_json)
        resultat["modele_utilise"] = modele_decision
        resultat["nb_qualite"]     = nb_qualite
        score = int(resultat.get("score", 0))
        nom_modele = "Sonnet" if modele_decision == MODELE_SONNET else "Haiku"
        log(f"📊 Score {nom_modele} : {score:+d}/10 → {resultat.get('recommandation')}")

    except LimiteQuotidienneAtteinte:
        raise
    except Exception as e:
        log(f"⚠️  Erreur avec {modele_decision} : {e}")
        return {
            "score": 0,
            "recommandation": "ATTENDRE",
            "explication": "Analyse impossible suite à une erreur technique. Par prudence : ATTENDRE.",
            "modele_utilise": modele_decision,
            "nb_qualite":     nb_qualite,
        }

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
    "News qualité 2-3",
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
        en_tetes_actuels = [c.value for c in feuille[1]]
        log("📁 Colonnes Buy & Hold ajoutées au fichier Excel existant.")
    if "News qualité 2-3" not in en_tetes_actuels:
        col = len(en_tetes_actuels) + 1
        cell = feuille.cell(row=1, column=col, value="News qualité 2-3")
        cell.font = bold
        classeur.save(FICHIER_EXCEL)
        log("📁 Colonne 'News qualité 2-3' ajoutée au fichier Excel existant.")


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
        analyse.get("nb_qualite", 0),
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
            {"score": 0, "recommandation": "ATTENDRE", "modele_utilise": "—", "nb_qualite": 0},
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
    - Mode local (Mac)           : boucle infinie, attend 15 min entre chaque cycle
                                   → lance avec : python3 trading_agent.py
    - Mode GitHub Actions (--once) : exécute UN seul cycle puis s'arrête
                                   → GitHub Actions relance le script toutes les 15 min
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
    log(f"   Seuil news qualité 2-3   : ≥{SEUIL_NEWS_PERTINENTES} → Sonnet décide")
    log(f"   Frais trading            : {FRAIS_TRADING*100:.1f}% par trade")
    log(f"   Stop-loss                : -{STOP_LOSS_PCT*100:.0f}%")
    log(f"   Take-profit              : +{TAKE_PROFIT_PCT*100:.0f}%")
    log(f"   Limite coût API/jour     : ${LIMITE_COUT_QUOTIDIEN_USD:.2f}")
    usage_actuel = charger_usage_quotidien()
    log(f"   Dépense API aujourd'hui  : ${usage_actuel['cout_total_usd']:.4f} "
        f"({usage_actuel['nb_appels']} appels)")
    log("=" * 55)
    log("")

    # ── Vérification que la clé API Anthropic est bien renseignée ──
    if not ANTHROPIC_API_KEY:
        log("❌ ERREUR : Clé API manquante : ANTHROPIC_API_KEY")
        log("   En local : remplis le fichier .env")
        log("   Sur GitHub : ajoute le secret dans Settings → Secrets → Actions")
        sys.exit(1)

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
    #  MODE LOCAL : boucle infinie toutes les 15 min
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
