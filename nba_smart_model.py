import subprocess
import sys
import time
import random
import numpy as np
import pandas as pd
import unicodedata
import requests
import json
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import os
import json
import io
from flask import Flask
import gspread
from oauth2client.service_account import ServiceAccountCredentials


# --- AUTO-INSTALLER ---
def install(package):
    try:
        __import__(package)
    except ImportError:
        pass
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except:
        pass

print("Initializing NBA Spread Pro v5.1 (Corrected Math & Advanced Stats)...")
install("pandas")
install("nba_api")
install("openpyxl")
install("numpy")
install("selenium")
install("webdriver_manager")
install("scipy")
install("requests")
install("fuzzywuzzy")              # <--- NEW LINE
install("python-levenshtein")      # <--- NEW LINE (for performance)
install("scikit-learn")

# --- IMPORTS ---
from nba_api.stats.endpoints import leaguedashteamstats, scoreboardv2, leaguedashplayerstats, leaguegamelog
from nba_api.stats.static import teams
#from selenium import webdriver
#from selenium.webdriver.chrome.service import Service
#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.common.by import By
#from webdriver_manager.chrome import ChromeDriverManager
from sklearn.linear_model import LinearRegression
from fuzzywuzzy import fuzz

# --- FLASK APP SETUP (Lisää tämä importtien jälkeen) ---
app = Flask(__name__)


# --- CONFIGURATION ---
ODDS_API_KEY = "ac7ac6b93d8c98983d9b2f09b87b1014" # Muista pitää API-avaimesi turvassa
SIMULATIONS = 10000 
HOME_ADVANTAGE_MAP = {
    'DEN': 3.5, 'UTA': 3.2, # Altitude Advantage
    'GSW': 2.8, 'BOS': 2.7, 'PHI': 2.7, 'NYK': 2.7, 'MIL': 2.6, # Strong Crowds
    'LAL': 2.3, 'LAC': 2.1, 'MIA': 2.3, 'TOR': 2.3, 
    'CHA': 1.8, 'WAS': 1.8, 'DET': 1.8, 'SAS': 2.0 # Weak/Rebuilding
}
DEFAULT_HCA = 2.3 # Käytetään jos joukkuetta ei löydy listalta
OUTPUT_FILE = 'NBA_Spread_Value_v5.xlsx'
AUDIT_FILE = 'model_audit_v5.json'

# FIX: B2B Penalty adjusted to realistic levels (approx 2.0 net rating swing total)
B2B_PENALTY = 1.0 


# GLOBAL AUDIT LOG
audit_log = {
    "timestamp": str(datetime.now()),
    "settings": {
        "simulations": SIMULATIONS,
        "home_advantage": HOME_ADVANTAGE_MAP,
        "b2b_penalty": B2B_PENALTY
    },
    "injuries_found": [],
    "team_stats_sample": {},
    "games_analyzed": []
}

# ========================================================
# 1. DATA ENGINE (INJURIES & ROSTERS)
# ========================================================

def get_current_season():
    now = datetime.now()
    if now.month >= 10: start = now.year
    else: start = now.year - 1
    season = f"{start}-{str(start+1)[-2:]}"
    audit_log["season"] = season
    return season

def normalize_name(name):
    # Poistetaan aksentit ja erikoismerkit (esim. Dončić -> Doncic)
    nfkd_form = unicodedata.normalize('NFKD', name)
    name_ascii = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    # Poistetaan pisteet (C.J. -> CJ) ja väliviivat
    clean_name = name_ascii.replace(".", "").replace("'", "").strip().upper()
    
    # Tunnetut aliakset (Korjaa nimet jos API ja Injury Report eroavat)
    aliases = {
        "NICOLAS CLAXTON": "NIC CLAXTON",
        "CAMERON THOMAS": "CAM THOMAS",
        "VICTOR WEMBANYAMA": "VICTOR WEMBANYAMA",
        "TIM HARDAWAY JR": "TIM HARDAWAY",
        "KENYON MARTIN JR": "KENYON MARTIN",
        "OG ANUNOBY": "OG ANUNOBY"
    }
    
    # Poistetaan suffixit vertailun helpottamiseksi
    suffixes = [" JR", " SR", " II", " III", " IV"]
    for s in suffixes:
        if clean_name.endswith(s):
            clean_name = clean_name.replace(s, "")
            
    return aliases.get(clean_name, clean_name)

def get_injured_players():
    print("Step 1: Fetching Injury Report (Source: ESPN Roster API)...")
    injured_data = {}
    
    # 1. Haetaan tiimilista ja ID:t
    teams_url = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams?limit=30"
    headers = {
        "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1"
    }

    try:
        print("   -> Fetching active Team IDs...")
        r_teams = requests.get(teams_url, headers=headers, timeout=5)
        teams_data = r_teams.json()
        
        team_id_map = {}
        if 'sports' in teams_data:
            for sport in teams_data['sports']:
                for league in sport['leagues']:
                    for team_entry in league['teams']:
                        t = team_entry['team']
                        team_id_map[t['id']] = normalize_name(t['displayName'])

        print(f"   -> Scanning rosters for {len(team_id_map)} teams...")

        count = 0
        total_players_scanned = 0
        
        with requests.Session() as s:
            for t_id, t_name in team_id_map.items():
                try:
                    url = f"https://site.api.espn.com/apis/site/v2/sports/basketball/nba/teams/{t_id}/roster"
                    resp = s.get(url, headers=headers, timeout=3)
                    
                    if resp.status_code == 200:
                        data = resp.json()
                        roster = data.get('athletes', data.get('items', []))
                        
                        for player in roster:
                            total_players_scanned += 1
                            injuries = player.get('injuries', [])
                            
                            if injuries:
                                raw_name = player.get('fullName', player.get('displayName', 'Unknown'))
                                name = normalize_name(raw_name)
                                
                                # Otetaan ensimmäinen status
                                inj_status_data = injuries[0]
                                status_desc = str(inj_status_data.get('status', '')).lower()
                                
                                if not status_desc:
                                    status_desc = str(inj_status_data.get('details', '')).lower()
                                
                                # --- KONVERTOIDAAN SPREADS-OHJELMAN YMMÄRTÄMÄÄN MUOTOON (Weight) ---
                                # Spreads käyttää "weight"-arvoa: 
                                # 1.0 = OUT/DOUBTFUL, 0.5 = GTD/QUESTIONABLE, 0.1 = PROBABLE
                                
                                impact_weight = 0.5 # Oletus (GTD)
                                status_text = "GTD"
                                
                                if "out" in status_desc:
                                    impact_weight = 1.0
                                    status_text = "OUT"
                                elif "doubtful" in status_desc:
                                    impact_weight = 1.0 # Kohdellaan poissaolona
                                    status_text = "DOUBTFUL"
                                elif "probable" in status_desc:
                                    impact_weight = 0.1
                                    status_text = "PROBABLE"
                                elif "day-to-day" in status_desc or "questionable" in status_desc:
                                    impact_weight = 0.5
                                    status_text = "GTD"

                                # Tallennetaan Spreads-ohjelman vaatimassa muodossa
                                injured_data[name] = {
                                    "weight": impact_weight,
                                    "team": t_name,
                                    "status_text": status_text # Debuggausta varten
                                }
                                count += 1
                                
                except Exception:
                    continue

        print(f"   -> Scanned {total_players_scanned} players. Found {count} injuries via API.")
        return injured_data

    except Exception as e:
        print(f"   -> CRITICAL ERROR: {e}")
        return {}

def get_all_player_stats(season):
    print("Step 2: Fetching Advanced Player Stats & Recency Info...")
    try:
        # 1. Haetaan perustilastot (PIE, Net Rating, Minuutit)
        stats = leaguedashplayerstats.LeagueDashPlayerStats(
            season=season, 
            measure_type_detailed_defense='Advanced', 
            per_mode_detailed='PerGame'
        ).get_data_frames()[0]
        
        # 2. Haetaan pelilokit, jotta nähdään MILLOIN pelaaja on viimeksi pelannut
        print("   -> Fetching Game Logs to identify long-term injuries...")
        logs = leaguegamelog.LeagueGameLog(season=season, player_or_team_abbreviation='P').get_data_frames()[0]
        logs['GAME_DATE'] = pd.to_datetime(logs['GAME_DATE'])
        
        # Otetaan jokaisen pelaajan viimeisin peli
        last_game_map = {}
        if not logs.empty:
            last_games = logs.sort_values('GAME_DATE').groupby('PLAYER_ID').tail(1)
            last_game_map = dict(zip(last_games['PLAYER_ID'], last_games['GAME_DATE']))

        player_map = {} 
        current_date = datetime.now()

        for _, row in stats.iterrows():
            tid = row['TEAM_ID']
            pid = row['PLAYER_ID']
            if tid not in player_map: player_map[tid] = []
            
            # Laske päivät viimeisestä pelistä
            days_inactive = 0
            if pid in last_game_map:
                last_played = last_game_map[pid]
                days_inactive = (current_date - last_played).days
            else:
                # Jos ei löydy lokeista (esim. ei pelejä tällä kaudella), merkitään pitkäksi poissaoloksi
                days_inactive = 999 

            player_map[tid].append({
                'NAME': normalize_name(row['PLAYER_NAME']),
                'MIN': row['MIN'],
                'NET_RTG': row['NET_RATING'],
                'USG_PCT': row['USG_PCT'], 
                'PIE': row['PIE'],
                'DAYS_INACTIVE': days_inactive  # <--- UUSI KENTTÄ
            })
            
        print(f"   -> Analyzed stats for {len(stats)} players.")
        return player_map
    except Exception as e:
        print(f"   -> Error fetching player stats: {e}")
        return {}

def calculate_smart_injury_impact(team_id, player_db, injured_list, team_stats):
    if team_id not in player_db: return 0.0, 0.0, []
    
    roster = player_db[team_id]
    current_team_name = normalize_name(team_stats[team_id]['NAME'])
    
    impact_score = 0.0
    missing_desc = []
    
    # Calculate Bench Baseline
    bench = [p for p in roster if 12.0 <= p['MIN'] < 26.0]
    bench_net_rtg = np.mean([p['NET_RTG'] for p in bench]) if len(bench) >= 2 else -2.0
    bench_net_rtg = max(-8.0, min(4.0, bench_net_rtg))

    injured_names_db = list(injured_list.keys())

    for p in roster:
        name = p['NAME']
        weight = 0.0
        
        # --- LONG TERM INJURY FILTER (UUSI OMINAISUUS) ---
        # Jos pelaaja on ollut pelaamatta yli 28 päivää (4 vkoa), markkina on jo korjannut.
        # Emme rankaise täysimääräisesti.
        days_inactive = p.get('DAYS_INACTIVE', 0)
        
        # Jos on ollut pois yli 4 viikkoa, ohitetaan kokonaan (Impact = 0)
        if days_inactive > 28:
            continue
            
        # --- TEAM-AWARE MATCHING ---
        match_found = False
        target_data = None
        
        # 1. Exact Match
        if name in injured_list:
            if fuzz.ratio(injured_list[name]['team'], current_team_name) > 70:
                target_data = injured_list[name]
                match_found = True
        
        # 2. Fuzzy Match
        if not match_found:
            best_score, best_name = 0, ""
            for inj in injured_names_db:
                score = fuzz.ratio(name, inj)
                if score > best_score: best_score, best_name = score, inj
            
            if best_score >= 80:
                cand_data = injured_list[best_name]
                if fuzz.ratio(cand_data['team'], current_team_name) > 70:
                    target_data = cand_data
                    match_found = True

        if match_found and target_data:
            weight = target_data['weight']
        
        # --- MATH ENGINE ---
        if weight > 0:
            # Pieni vaimennus, jos pelaaja on ollut pois 2-4 viikkoa (14-28 pv)
            # Markkina on osittain korjannut, joten vähennämme vaikutusta 50%
            recency_decay = 1.0
            if 14 < days_inactive <= 28:
                recency_decay = 0.5
            
            # 1. Minutes Factor
            min_factor = (p['MIN'] / 48.0)
            
            # 2. Base Talent Metric
            pie_score = p['PIE'] * 100 
            net_diff = max(0, p['NET_RTG'] - bench_net_rtg)
            
            base_talent = (pie_score * 0.70) + (net_diff * 0.30)
            
            damping = 0.55 
            
            # Lisätään recency_decay laskukaavaan
            final_impact = base_talent * min_factor * damping * weight * recency_decay
            
            if final_impact > 0.4:
                impact_score += final_impact
                status = "OUT"
                if weight == 0.5: status = "GTD"
                elif weight == 0.1: status = "PROB"
                
                # Merkitään listaan jos kyseessä "Old Injury" mutta vaikuttaa yhä vähän
                recency_tag = ""
                if recency_decay < 1.0: recency_tag = " (Old)"
                
                missing_desc.append(f"{name} ({status}{recency_tag} -{round(final_impact, 1)})")

    if impact_score > 14.0: impact_score = 14.0
    
    off_loss = impact_score * 0.60
    def_loss = impact_score * 0.40
    
    return off_loss, def_loss, missing_desc

def get_b2b_teams():
    print("Step 3: Checking Schedule for Back-to-Backs...")
    yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    b2b_teams = set()
    try:
        board = scoreboardv2.ScoreboardV2(game_date=yesterday)
        games = board.game_header.get_data_frame()
        for _, g in games.iterrows():
            b2b_teams.add(g['HOME_TEAM_ID'])
            b2b_teams.add(g['VISITOR_TEAM_ID'])
        audit_log["b2b_teams_count"] = len(b2b_teams)
    except:
        pass
    return b2b_teams

def get_blended_stats(season):
    print("Step 4: Fetching Team Stats (Advanced & Four Factors)...")
    try:
        # 1. ADVANCED
        adv = leaguedashteamstats.LeagueDashTeamStats(
            season=season, 
            measure_type_detailed_defense='Advanced', 
            per_mode_detailed='Per100Possessions'
        ).get_data_frames()[0]
        
        # 2. FOUR FACTORS
        ff = leaguedashteamstats.LeagueDashTeamStats(
            season=season, 
            measure_type_detailed_defense='Four Factors', 
            per_mode_detailed='PerGame'
        ).get_data_frames()[0]
        
        # Valitaan sarakkeet
        ff_cols = ['TEAM_ID', 'EFG_PCT', 'TM_TOV_PCT', 'OREB_PCT', 'FTA_RATE',
                   'OPP_EFG_PCT', 'OPP_TOV_PCT', 'OPP_OREB_PCT', 'OPP_FTA_RATE']
        
        # --- KORJAUS TÄSSÄ ---
        # Käytetään suffixes-parametria. 
        # (_adv) lisätään Advanced-taulun päällekkäisiin sarakkeisiin.
        # ('') eli tyhjä lisätään Four Factors -tauluun, jotta niiden nimet SÄILYVÄT ennallaan.
        merged = pd.merge(adv, ff[ff_cols], on='TEAM_ID', suffixes=('_adv', ''))

        stats = {}
        
        for _, row in merged.iterrows():
            tid = row['TEAM_ID']
            stats[tid] = {
                'NAME': row['TEAM_NAME'],
                'OFF_RTG': row['OFF_RATING'],
                'DEF_RTG': row['DEF_RATING'],
                'PACE': row['PACE'],
                
                # Nyt nämä löytyvät varmasti, koska merge ei nimennyt niitä uudelleen
                'FF_EFG': row['EFG_PCT'],       
                'FF_TOV': row['TM_TOV_PCT'],    
                'FF_ORB': row['OREB_PCT'],      
                'FF_FTA': row['FTA_RATE'],      
                
                'DEF_EFG': row['OPP_EFG_PCT'],  
                'DEF_TOV': row['OPP_TOV_PCT'],  
                'DEF_ORB': row['OPP_OREB_PCT'], 
                'DEF_FTA': row['OPP_FTA_RATE']  
            }
        return stats

    except Exception as e:
        print(f"Error fetching stats: {e}")
        # Debug-tuloste auttaa yhä, jos jotain outoa tapahtuu
        if 'ff' in locals():
            print("Debug - Four Factors Columns:", ff.columns.tolist())
        return {}

def calibrate_four_factors(season):
    print("Step X: Calibrating Physics Engine (Linear Regression)...")
    try:
        # Haetaan joukkueiden tilastot (Advanced)
        team_stats = leaguedashteamstats.LeagueDashTeamStats(
            season=season,
            measure_type_detailed_defense='Four Factors',
            per_mode_detailed='Per100Possessions' # Tärkeä: Käytämme Per100Possessions dataa regressioon
        ).get_data_frames()[0]

        # Haetaan myös Offensive Rating (Advanced-näkymästä, tai lasketaan pisteistä)
        # Huom: Four Factors endpointissa ei välttämättä ole OFF_RATING suoraan samassa skaalassa,
        # joten varmistetaan hakemalla Advanced-data.
        adv_stats = leaguedashteamstats.LeagueDashTeamStats(
            season=season,
            measure_type_detailed_defense='Advanced',
            per_mode_detailed='Per100Possessions'
        ).get_data_frames()[0]
        
        # Yhdistetään datat
        merged = pd.merge(team_stats, adv_stats[['TEAM_ID', 'OFF_RATING']], on='TEAM_ID')
        
        # Määritellään muuttujat regressiolle
        # X = Four Factors (eFG%, TOV%, ORB%, FT Rate)
        # Y = Offensive Rating (Pisteet per 100 hallintaa)
        
        X = merged[['EFG_PCT', 'TM_TOV_PCT', 'OREB_PCT', 'FTA_RATE']]
        y = merged['OFF_RATING']
        
        # Ajetaan lineaarinen regressio
        model = LinearRegression()
        model.fit(X, y)
        
        # Nämä ovat uudet "Fysiikkakertoimet"
        coeffs = {
            'BASE': model.intercept_,       # Vakiotermi (ns. pohjataso)
            'EFG': model.coef_[0],          # Ennen: 175.0
            'TOV': model.coef_[1],          # Ennen: -115.0
            'ORB': model.coef_[2],          # Ennen: 55.0
            'FTA': model.coef_[3]           # Ennen: 68.0
        }
        
        print(f"   -> Calibration Complete. New Coefficients:")
        print(f"      Base: {coeffs['BASE']:.2f}")
        print(f"      eFG%: {coeffs['EFG']:.2f} (Impact of 1% shooting)")
        print(f"      TOV%: {coeffs['TOV']:.2f} (Cost of turnover)")
        print(f"      ORB%: {coeffs['ORB']:.2f} (Value of off rebound)")
        print(f"      FTA : {coeffs['FTA']:.2f} (Value of FT rate)")
        
        return coeffs
        
    except Exception as e:
        print(f"   -> WARNING: Calibration failed ({e}). Using fallback weights.")
        # Palautetaan vanhat arvot varalle, jos API kaatuu
        return {'BASE': 5.0, 'EFG': 175.0, 'TOV': -115.0, 'ORB': 55.0, 'FTA': 68.0}

def get_schedule():
    print("Step 5: Fetching Upcoming Games...")
    today = datetime.now()
    date_str = today.strftime('%Y-%m-%d')
    try:
        board = scoreboardv2.ScoreboardV2(game_date=date_str)
        games = board.game_header.get_data_frame()
        
        # --- KORJAUS 1: Poistetaan tuplat heti haun jälkeen ---
        # Tämä varmistaa, ettei sama GAME_ID esiinny listalla kahdesti
        if not games.empty:
            games = games.drop_duplicates(subset=['GAME_ID']) 

        if games.empty:
            print("   -> No games today. Checking tomorrow...")
            date_str = (today + timedelta(days=1)).strftime('%Y-%m-%d')
            board = scoreboardv2.ScoreboardV2(game_date=date_str)
            games = board.game_header.get_data_frame()
            
            # --- KORJAUS 2: Sama varmistus, jos haetaan huomisen pelit ---
            if not games.empty:
                games = games.drop_duplicates(subset=['GAME_ID'])

        return games, date_str
    except:
        return pd.DataFrame(), None

# ========================================================
# 2. ODDS API
# ========================================================

def fetch_spread_odds():
    print("Step 6: Fetching Spread & Moneyline Odds from API...")
    # LISÄTTY: markets=spreads,h2h
    url = f"https://api.the-odds-api.com/v4/sports/basketball_nba/odds/?apiKey={ODDS_API_KEY}&regions=eu&markets=spreads,h2h&oddsFormat=decimal"
    try:
        response = requests.get(url)
        if response.status_code != 200: return {}
        data = response.json()
        odds_map = {} 
        for event in data:
            home, away = event['home_team'], event['away_team']
            # Alustetaan arvot
            h_p, h_o, a_p, a_o = 0, 0, 0, 0 # Spread
            h_ml, a_ml = 0, 0              # Moneyline (UUSI)
            
            found_spread = False
            found_ml = False
            
            for bookie in event['bookmakers']:
                for market in bookie['markets']:
                    # 1. Spreadit
                    if market['key'] == 'spreads':
                        for outcome in market['outcomes']:
                            if outcome['name'] == home: h_p, h_o = outcome['point'], outcome['price']
                            elif outcome['name'] == away: a_p, a_o = outcome['point'], outcome['price']
                        found_spread = True
                    
                    # 2. Moneyline (h2h)
                    elif market['key'] == 'h2h':
                        for outcome in market['outcomes']:
                            if outcome['name'] == home: h_ml = outcome['price']
                            elif outcome['name'] == away: a_ml = outcome['price']
                        found_ml = True
                
                # Jos molemmat löytyivät samalta bookkerilta, break. (Tai jatketaan parhaan etsimistä)
                if found_spread and found_ml: break
            
            if found_spread or found_ml:
                key = f"{normalize_team_name(away)} @ {normalize_team_name(home)}"
                odds_map[key] = {
                    'H_Pt': h_p, 'H_Od': h_o, 
                    'A_Pt': a_p, 'A_Od': a_o,
                    'H_ML': h_ml, 'A_ML': a_ml # UUDET KENTÄT
                }
        return odds_map
    except Exception as e: 
        print(f"Odds Error: {e}")
        return {}

def normalize_team_name(name):
    return name.replace("Los Angeles", "L.A.")

def get_game_spreads(home, away, odds_data):
    h_s, a_s = normalize_team_name(home), normalize_team_name(away)
    key = f"{a_s} @ {h_s}"
    if key in odds_data: return odds_data[key]
    for k, v in odds_data.items():
        if a_s in k and h_s in k: return v
    return None

# ========================================================
# 3. PRO SIMULATION ENGINE (AUDITABLE)
# ========================================================

# ========================================================
# 3. PRO SIMULATION ENGINE (AUDITABLE)
# ========================================================

def simulate_spread_pro(home_id, away_id, stats_db, b2b_set, player_db, injured_list, spread_line_home, coeffs):
    """
    PURE FOUR FACTORS MODEL (v6.1 Updated)
    Ennustaa tehokkuuden suoraan neljästä faktorista (eFG, TOV, ORB, FT)
    ja lisää sen jälkeen HCA, B2B ja loukkaantumisvaikutukset.
    """
    if home_id not in stats_db or away_id not in stats_db:
        return 0, 0, 0, [], [], [] 
    
    h = stats_db[home_id]
    a = stats_db[away_id]
    notes = []
    
    # --- 1. MATCHUP ENGINE: ENNUSTETAAN PELIN TILASTOT ---
    # Logiikka: (Oma Hyökkäys + Vastustajan Puolustus) / 2
    
    # Heittotarkkuus (eFG%)
    h_proj_efg = (h['FF_EFG'] + a['DEF_EFG']) / 2
    a_proj_efg = (a['FF_EFG'] + h['DEF_EFG']) / 2
    
    # Menetykset (TOV%)
    h_proj_tov = (h['FF_TOV'] + a['DEF_TOV']) / 2
    a_proj_tov = (a['FF_TOV'] + h['DEF_TOV']) / 2
    
    # Hyökkäyslevypallot (OREB%)
    h_proj_orb = (h['FF_ORB'] + a['DEF_ORB']) / 2
    a_proj_orb = (a['FF_ORB'] + h['DEF_ORB']) / 2
    
    # Vapaaheitot (FTA Rate)
    h_proj_fta = (h['FF_FTA'] + a['DEF_FTA']) / 2
    a_proj_fta = (a['FF_FTA'] + h['DEF_FTA']) / 2

    # --- 2. MUUNNETAAN TILASTOT PISTEIKSI (Efficiency per 100 poss) ---
    # Käytetään regressiokertoimia muuttamaan %-luvut pisteodotusarvoksi.
    
    def calculate_implied_ortg(efg, tov, orb, fta):
        # --- DYNAAMINEN PURE FOUR FACTORS ---
        # Käytetään regressiosta saatuja kertoimia
        
        base = coeffs['BASE']
        pts_efg = efg * coeffs['EFG']
        pts_tov = tov * coeffs['TOV']
        pts_orb = orb * coeffs['ORB']
        pts_fta = fta * coeffs['FTA']
        
        return base + pts_efg + pts_tov + pts_orb + pts_fta

    # Lasketaan "Raaka" hyökkäysteho (ilman loukkaantumisia/rasitusta)
    h_raw_eff = calculate_implied_ortg(h_proj_efg, h_proj_tov, h_proj_orb, h_proj_fta)
    a_raw_eff = calculate_implied_ortg(a_proj_efg, a_proj_tov, a_proj_orb, a_proj_fta)

    # --- 3. KOTIETU (HCA) ---
    current_hca = DEFAULT_HCA
    if "NUGGETS" in h['NAME'].upper(): current_hca = 3.5
    elif "JAZZ" in h['NAME'].upper(): current_hca = 3.2
    elif "CELTICS" in h['NAME'].upper() or "KNICKS" in h['NAME'].upper(): current_hca = 2.7
    
    # Lisätään kotietu kotijoukkueen tehokkuuteen
    h_raw_eff += current_hca

    # --- 4. B2B ADJUSTMENT (Rasitus) ---
    # Vähennetään rasittuneen joukkueen tehokkuutta
    if home_id in b2b_set:
        h_raw_eff -= 2.0  # ~2 pistettä per 100 hallintaa
        notes.append("Home B2B")
    if away_id in b2b_set:
        a_raw_eff -= 2.0
        notes.append("Away B2B")
        
    # --- 5. LOUKKAANTUMISET (Smart Injury Impact) ---
    # Haetaan vaikutus: off_loss (oma hyökkäys huononee) ja def_loss (oma puolustus huononee)
    h_off_loss, h_def_loss, h_missing = calculate_smart_injury_impact(home_id, player_db, injured_list, stats_db)
    a_off_loss, a_def_loss, a_missing = calculate_smart_injury_impact(away_id, player_db, injured_list, stats_db)
    
    # A. Vähennetään OMASTA hyökkäyksestä poissaolijat
    h_final_eff = h_raw_eff - h_off_loss
    a_final_eff = a_raw_eff - a_off_loss
    
    # B. Lisätään hyökkäystehoa, jos VASTUSTAJAN puolustus on heikentynyt
    # (Jos kotijoukkueelta puuttuu Rudy Gobert -> Vierasjoukkueen hyökkäys paranee)
    h_final_eff += a_def_loss
    a_final_eff += h_def_loss

    # --- 6. VARIANSSI JA SIMULAATIO ---
    
    # Lisätään huomiot (Notes) Exceliä varten
    if h_proj_efg > a_proj_efg + 0.025: notes.append("Home Shooting Edge")
    if a_proj_efg > h_proj_efg + 0.025: notes.append("Away Shooting Edge")
    if h_proj_orb > 0.30 and a['DEF_ORB'] > 0.30: notes.append("Home High REB Potential")

    # Pelinopeus (Pace)
    pace = (h['PACE'] + a['PACE']) / 2
    
    # Muutetaan tehokkuus (per 100) odotetuiksi pisteiksi (per Pace)
    # TÄMÄ ON TÄRKEÄ KORJAUS: Nyt vertaamme oikeita pisteitä spreadiin
    h_proj_pts = (h_final_eff * pace) / 100.0
    a_proj_pts = (a_final_eff * pace) / 100.0
    
    # Keskihajonta skaalattuna pelinopeudella
    base_std_dev = 13.5 # Hieman korkeampi, koska Four Factors -mallissa on enemmän liikkuvia osia
    adjusted_std_dev = base_std_dev * (pace / 100.0)
    
    # Monte Carlo -ajo
    h_sims = np.random.normal(h_proj_pts, adjusted_std_dev, SIMULATIONS)
    a_sims = np.random.normal(a_proj_pts, adjusted_std_dev, SIMULATIONS)
    
    # 1. SPREAD CALCULATION
    sim_margins = h_sims - a_sims
    covers = (sim_margins > -spread_line_home).sum()
    cover_prob = (covers / SIMULATIONS) * 100
    
    # 2. MONEYLINE CALCULATION (UUSI)
    # Lasketaan kuinka monta kertaa kotijoukkue teki enemmän pisteitä kuin vieras
    wins = (h_sims > a_sims).sum()
    win_prob = (wins / SIMULATIONS) * 100
    
    # Palautetaan myös win_prob (LISÄTTY TOISEKSI ARVOKSI)
    return cover_prob, win_prob, h_proj_pts, a_proj_pts, notes, h_missing, a_missing
# ========================================================
# 4. MAIN RUNNER
# ========================================================

def run_spread_pro():
    season = get_current_season()
    print(f"--- NBA SMART MODEL v6.2 (Spread & Moneyline) ---\nSeason: {season}")
    
    # 1. Fetch ALL Data
    injured = get_injured_players()
    b2b_teams = get_b2b_teams()
    player_stats = get_all_player_stats(season)
    stats = get_blended_stats(season)
    coeffs = calibrate_four_factors(season)
    games, date = get_schedule()
    
    if games.empty: print("No games."); return

    # 2. Get Odds (Nyt sisältää ML)
    spreads = fetch_spread_odds()
    
    print(f"Simulating {len(games)} games for {date} with {SIMULATIONS} runs...")
    
    # Listat tuloksille
    results_spread = []
    results_ml = []
    
    nba_teams = teams.get_teams()
    team_map = {t['id']: t['full_name'] for t in nba_teams}
    
    for _, game in games.iterrows():
        hid, aid = game['HOME_TEAM_ID'], game['VISITOR_TEAM_ID']
        h_name, a_name = team_map.get(hid, "Home"), team_map.get(aid, "Away")
        
        odds = get_game_spreads(h_name, a_name, spreads)
        line_h = odds['H_Pt'] if odds else 0
        
        # --- KUTSU SIMULAATIOTA (NYT 7 PALUUARVOA) ---
        h_cov_pct, h_win_pct, h_pts, a_pts, notes, h_inj_list, a_inj_list = simulate_spread_pro(
            hid, aid, stats, b2b_teams, player_stats, injured, line_h, coeffs
        )
        
        # Formatoidaan datat
        h_inj_str = ", ".join(h_inj_list) if h_inj_list else "-"
        a_inj_str = ", ".join(a_inj_list) if a_inj_list else "-"
        note_str = " | ".join(notes) if notes else ""
        proj_score = f"{int(a_pts)} - {int(h_pts)}"
        
        # --- A. SPREAD DATA ---
        if odds:
            a_cov_pct = 100 - h_cov_pct
            rec_spread = "-"
            if h_cov_pct >= 62.5: rec_spread = "STRONG HOME"
            elif h_cov_pct >= 57.0: rec_spread = "BET HOME"
            elif a_cov_pct >= 62.5: rec_spread = "STRONG AWAY"
            elif a_cov_pct >= 57.0: rec_spread = "BET AWAY"
            
            results_spread.append({
                'Match': f"{a_name} @ {h_name}",
                'Spread': f"{line_h}",
                'Proj Score': proj_score,
                'Context': note_str,
                'Home Injuries': h_inj_str, 'Away Injuries': a_inj_str,
                'Home Cover %': round(h_cov_pct, 1), 'Home Odds': odds['H_Od'],
                'Away Cover %': round(a_cov_pct, 1), 'Away Odds': odds['A_Od'],
                'RECOMMENDATION': rec_spread,
                'Sort_Key': abs(h_cov_pct - 50.0)
            })
        else:
            results_spread.append({
                'Match': f"{a_name} @ {h_name}", 'Spread': "N/A",
                'Proj Score': proj_score, 'Context': note_str,
                'Home Injuries': h_inj_str, 'Away Injuries': a_inj_str,
                'Home Cover %': "-", 'RECOMMENDATION': "No Odds", 'Sort_Key': 0
            })
            
        # --- B. MONEYLINE DATA (SMART FILTER) ---
        if odds and odds.get('H_ML', 0) > 0:
            a_win_pct = 100 - h_win_pct
            
            # 1. Lasketaan Odotusarvo (ROI)
            h_ev = (h_win_pct / 100 * odds['H_ML']) - 1
            a_ev = (a_win_pct / 100 * odds['A_ML']) - 1
            
            rec_ml = "-"
            
            # --- UUSI SUODATUSLOGIIKKA ---
            # Hylätään "roskavedot" (liian pieni osumatodennäköisyys), vaikka kerroin olisi korkea.
            # Asetetaan rajaksi esim. 35% tai 40%.
            
            MIN_PROBABILITY = 38.0  # Älä suosittele, jos mallin mielestä voittosanssi on alle 38%
            
            # Kotijoukkueen analyysi
            if h_ev > 0.03 and h_win_pct >= MIN_PROBABILITY:
                if h_ev > 0.10 and h_win_pct > 50:
                    rec_ml = "TOP PICK (HOME)"   # <--- Lisätty (HOME)
                elif h_ev > 0.05 and h_win_pct > 40:
                    rec_ml = "STRONG VALUE (HOME)" # <--- Lisätty (HOME)
                elif h_ev > 0.03:
                    rec_ml = "VALUE (HOME)"        # <--- Lisätty (HOME)
            
            # Vierasjoukkueen analyysi
            elif a_ev > 0.03 and a_win_pct >= MIN_PROBABILITY:
                if a_ev > 0.10 and a_win_pct > 50:
                    rec_ml = "TOP PICK (AWAY)"     # <--- Lisätty (AWAY)
                elif a_ev > 0.05 and a_win_pct > 40:
                    rec_ml = "STRONG VALUE (AWAY)" # <--- Lisätty (AWAY)
                elif a_ev > 0.03:
                    rec_ml = "VALUE (AWAY)"        # <--- Lisätty (AWAY)

            # Jos on kova ylikerroin mutta pieni todennäköisyys, merkitään se varoituksella
            if rec_ml == "-" and (h_ev > 0.15 or a_ev > 0.15):
                rec_ml = "(High Risk / Longshot)"

            # Määritetään lajitteluavain (Sort Key) uusien tekstien perusteella
            # Jos suosituksessa lukee HOME, käytetään h_ev. Jos AWAY, käytetään a_ev.
            sort_val = -1
            if "HOME" in rec_ml:
                sort_val = h_ev
            elif "AWAY" in rec_ml:
                sort_val = a_ev
            elif "Risk" in rec_ml:
                sort_val = max(h_ev, a_ev) # Longshoteissa otetaan isompi

            results_ml.append({
                'Match': f"{a_name} @ {h_name}",
                'Proj Score': proj_score,
                'Context': note_str,
                'Home Injuries': h_inj_str, 'Away Injuries': a_inj_str,
                # Moneyline Specifics
                'Home Win %': round(h_win_pct, 1),
                'Home Odds': odds['H_ML'],
                'Home EV': round(h_ev * 100, 1), 
                'Away Win %': round(a_win_pct, 1),
                'Away Odds': odds['A_ML'],
                'Away EV': round(a_ev * 100, 1), 
                'RECOMMENDATION': rec_ml,
                # Lajitellaan nyt korjatun logiikan mukaan
                'Sort_Key': sort_val
            })

    # --- GOOGLE SHEETS UPLOAD ---
    if results_spread:
        # Create DataFrames
        df_spread = pd.DataFrame(results_spread).sort_values('Sort_Key', ascending=False).drop(columns=['Sort_Key'])
        
        df_ml = pd.DataFrame()
        if results_ml:
            df_ml = pd.DataFrame(results_ml).sort_values('Sort_Key', ascending=False).drop(columns=['Sort_Key'])
        
        print("   -> Connecting to Google Sheets (Render Mode)...")
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            
            # --- KORJAUS: Luetaan tunnukset suoraan Renderin ympäristömuuttujasta ---
            # Tämä vaatii, että Renderissä on asetettu muuttuja: GOOGLE_CREDENTIALS_JSON
            creds_dict = json.loads(os.environ['GOOGLE_CREDENTIALS_JSON'])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
            # ------------------------------------------------------------------------

            client = gspread.authorize(creds)
            
            SHEET_NAME = "NBA_Spread_Model_Output"
            sh = client.open(SHEET_NAME)

            # --- VÄLILEHTI 1: SPREADS ---
            try:
                ws1 = sh.get_worksheet(0)
            except:
                ws1 = sh.add_worksheet(title="Spreads", rows="100", cols="20")
            
            ws1.clear()
            ws1.update([df_spread.columns.values.tolist()] + df_spread.values.tolist())
            ws1.format('A1:O1', {'textFormat': {'bold': True}, 'backgroundColor': {'red': 0.12, 'green': 0.28, 'blue': 0.49}, 'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}})
            print("   -> Updated Sheet 1 (Spreads)")

            # --- VÄLILEHTI 2: MONEYLINE ---
            if not df_ml.empty:
                try:
                    # Yritetään hakea toinen välilehti (index 1)
                    ws2 = sh.get_worksheet(1)
                    if ws2 is None: # Jos index 1 ei ole olemassa
                         ws2 = sh.add_worksheet(title="Moneyline", rows="100", cols="20")
                except:
                    # Jos haku epäonnistuu, luodaan uusi
                    ws2 = sh.add_worksheet(title="Moneyline", rows="100", cols="20")

                ws2.clear()
                ws2.update([df_ml.columns.values.tolist()] + df_ml.values.tolist())
                # Eri väri otsikoille jotta erottuu (Vihreä sävy)
                ws2.format('A1:O1', {'textFormat': {'bold': True}, 'backgroundColor': {'red': 0.1, 'green': 0.35, 'blue': 0.15}, 'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}}})
                print("   -> Updated Sheet 2 (Moneyline)")

            return "Success: Sheets Updated!"

        except Exception as e:
            print(f"Google Sheets Error: {e}")
            return f"Error: {e}"
    else:
        print("No games.")
        return "No games analyzed."

# ========================================================
# 5. WEB TRIGGER (FLASK)
# ========================================================

@app.route('/')
def index():
    return "<h1>NBA Spread Model Ready</h1><p><a href='/run'>Click here to RUN MODEL</a></p>"

@app.route('/run')
def trigger_run():
    # 1. Create a memory buffer to catch text
    log_capture = io.StringIO()
    
    # 2. Redirect standard output (print) to our buffer
    original_stdout = sys.stdout
    sys.stdout = log_capture
    
    status = "Unknown"
    
    try:
        # 3. Run the actual analysis
        status = run_spread_pro()
        
    except Exception as e:
        print(f"\nCRITICAL CRASH: {e}")
        status = "Failed"
        
    finally:
        # 4. Restore standard output
        sys.stdout = original_stdout

    # 5. Get the text from the buffer
    full_logs = log_capture.getvalue()
    log_capture.close()

    # 6. Return HTML page with the logs
    return f"""
    <html>
        <head>
            <title>NBA Spread Model Status</title>
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <style>
                body {{ font-family: -apple-system, sans-serif; padding: 20px; max-width: 800px; margin: 0 auto; }}
                .status {{ font-size: 1.2em; font-weight: bold; margin-bottom: 20px; }}
                .success {{ color: green; }}
                .error {{ color: red; }}
                .logs {{ 
                    background-color: #f5f5f5; 
                    padding: 15px; 
                    border-radius: 8px; 
                    font-family: monospace; 
                    font-size: 0.9em; 
                    white-space: pre-wrap; 
                    overflow-x: auto;
                    border: 1px solid #ddd;
                }}
                .btn {{ 
                    display: inline-block; 
                    padding: 10px 20px; 
                    background-color: #007bff; 
                    color: white; 
                    text-decoration: none; 
                    border-radius: 5px; 
                    margin-top: 20px;
                }}
            </style>
        </head>
        <body>
            <h1>Process Finished</h1>
            <div class="status">
                Result: <span class="{ 'success' if 'Success' in str(status) else 'error' }">{status}</span>
            </div>
            
            <h3>Execution Logs:</h3>
            <div class="logs">{full_logs}</div>
            
            <a href='/' class="btn">Back to Home</a>
        </body>
    </html>
    """

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))

    app.run(host='0.0.0.0', port=port)







