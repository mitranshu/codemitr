import time
import random
import re
import logging
from typing import List, Dict, Any, Optional, Tuple
from functools import lru_cache
import pandas as pd
import streamlit as st
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
import spacy
import nltk
from nltk.corpus import stopwords
import plotly.express as px

from google_play_scraper import search, reviews, reviews_all, Sort, app
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from sklearn.feature_extraction.text import TfidfVectorizer

from urllib.parse import quote_plus

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- NLTK stopwords (download if needed) ---
try:
    nltk.data.find("corpora/stopwords")
except LookupError:
    nltk.download("stopwords")

SCAM_PATTERN = re.compile(r"\bscam(?:med|mer|mers|ming)?\b", flags=re.I)

STOPWORDS = set(stopwords.words("english"))

_IRREGULAR_ADJ = {
    "better": "good",
    "best": "good",
    "worse": "bad",
    "worst": "bad",
}

# --- Streamlit page config ---
st.set_page_config(page_title="Play Store Reviews Explorer", layout="wide", page_icon="icon.png")
st.markdown(
    """
    <style>
    .heading-pill {
      display: block;                      
      width: 100%;                         
      text-align: center;
      font-weight: 700;
      font-size: 1.4rem;
      padding: 0.8rem 0;                   
      border-radius: 12px;
      background: #003049;
      border: 1px solid rgba(99,110,250,0.25);
      box-shadow: 0 1px 2px rgba(0,0,0,0.04) inset;
      margin: 0.5rem 0 1rem 0;
      color: inherit;                      
    }

    .heading-sub {
      display: block;
      text-align: center;
      color: #6b7280;
      font-weight: 500;
      margin-top: -0.25rem;
      margin-bottom: 1rem;
      font-size: 0.95rem;
    }
    
    .heading-main {
      display: block;
      text-align: center;
      color: #6b7280;
      font-weight: 500;
      margin-top: -0.25rem;
      margin-bottom: 1rem;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04) inset;
      border: 1px solid rgba(99,110,250,0.25);
      background: #F3F5FF;
      border-radius: 12px;
      font-size: 2rem;
    }

    /* Optional: adjust for dark mode automatically */
    @media (prefers-color-scheme: dark) {
      .heading-pill {
        background: #1e293b);
        border: 1px solid rgba(255,255,255,0.1);
        color: #f1f5f9;
      }
      .heading-sub {
        color: #9ca3af;
      }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <h1 class="heading-pill" style="color:#FFFFFF">Play Store Reviews Explorer</h1>
    <div class="heading-sub"><br>Analyze, visualize, and explore sentiment trends from Google Play reviews</div>
    """,
    unsafe_allow_html=True,
)

@st.cache_data(show_spinner=False)
def search_and_enrich(app_name: str, lang: str = "en", country: str = "us", top_n: Optional[int] = 50, max_workers: int = 8, max_retries: int = 3) -> pd.DataFrame:
    """
    Search Play Store for `app_name`, build summary df, and enrich with app() metadata in parallel.
    Results are cached to avoid repeated searches on reruns.
    """
    results = []
    try:
        results = list(search(app_name, lang=lang, country=country))
        if top_n:
            results = results[:top_n]

        rows: List[Dict[str, Any]] = []
        for r in results:
            rows.append({
                "appId": r.get("appId") or r.get("app_id") or r.get("packageName"),
                "title": r.get("title"),
                "genre": r.get("genre"),
                "installs": r.get("installs"),
                "score": round(r.get("score"), 1) if r.get("score") is not None else None,
                "free": r.get("free"),
                "icon": r.get("icon"),
                "description":r.get("description")
            })
        df_search = pd.DataFrame(rows).dropna(subset=["appId"]).drop_duplicates("appId").reset_index(drop=True)

        pkg_ids = df_search['appId'].tolist()

        info_dicts: List[Dict[str, Any]] = []
        if pkg_ids:
            with ThreadPoolExecutor(max_workers=max_workers) as exe:
                futures = {
                    exe.submit(fetch_app_info_with_retries, pkg, lang, country, max_retries): pkg
                    for pkg in pkg_ids
                }
                for fut in as_completed(futures):
                    info = fut.result()
                    info_dicts.append(info)

        df_info = pd.DataFrame(info_dicts) if info_dicts else pd.DataFrame(columns=[
            "appId", "Number of ratings", "Number of reviews", "URL", "containsAds", "developer", "inAppProductPrice", "developerWebsite", "_error"
        ])

        merged = df_search.merge(df_info, on='appId', how='left')
        return merged
    except:
        return pd.DataFrame()

# --- Review fetching (cached) ---
@st.cache_data(show_spinner=False)
def fetch_reviews(pkg: str, lang: str = 'en', country: str = 'us', sort: Sort = Sort.NEWEST, count: int = 200, fetch_all: bool = False) -> List[Dict[str, Any]]:
    """
    Fetch reviews for a given package. When fetch_all True, uses reviews_all (may be slow).
    This function is cached by Streamlit keyed on (pkg, lang, country, sort, count, fetch_all).
    """
    try:
        if fetch_all:
            # reviews_all returns generator-like list of reviews — careful with very large apps
            return list(reviews_all(pkg, lang=lang, country=country))
        else:
            revs, _ = reviews(pkg, lang=lang, country=country, sort=sort, count=int(count))
            return revs
    except Exception as e:
        logger.exception("Failed to fetch reviews for %s: %s", pkg, e)
        return []   
    
    
@st.cache_data(show_spinner=False)
def fetch_app_info_with_retries(package_id: str, lang: str = 'en', country: str = 'us',
                                max_retries: int = 3, base_backoff: float = 1.0) -> Dict[str, Any]:
    """
    Fetch metadata for a single package with simple retry + exponential backoff.
    Returns a dict with expected keys (fills None on failure).
    Cached by Streamlit to avoid repeated network calls across reruns.
    """
    last_exc = None
    for attempt in range(max_retries):
        try:
            info = app(package_id, lang=lang, country=country)
            st.write(info.keys())
            return {
                "appId": info.get('appId') or package_id,
                "Number of ratings": info.get('ratings'),
                "Number of reviews": info.get('reviews'),
                "URL": info.get('url'),
                "containsAds": info.get('containsAds'),
                "developer": info.get('developer'),
                "inAppProductPrice": info.get('inAppProductPrice'),
                "developerWebsite": info.get('developerWebsite')
            }
        except Exception as e:
            last_exc = e
            wait = base_backoff * (2 ** attempt) + random.uniform(0, 0.5)
            logger.warning("fetch_app_info failed for %s attempt %d: %s — retry in %.1f s", package_id, attempt+1, e, wait)
            time.sleep(wait)
    return {
        "appId": package_id,
        "Title": None,
        "Number of ratings": None,
        "Number of reviews": None,
        "URL": None,
        "containsAds": None,
        "developer": None,
        "inAppProductPrice": None,
        "developerWebsite": None,
        "_error": str(last_exc) if last_exc else "Unknown error"
    }
    
def clean_text(txt: Optional[str]) -> str:
    """Basic cleaning: lowercase, remove urls, non-alnum (keep apostrophe), collapse whitespace."""
    if not isinstance(txt, str):
        return ""
    txt = txt.lower()
    txt = re.sub(r"https?://\S+", " ", txt)
    txt = re.sub(r"[^a-z0-9\s']", " ", txt)
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt

@lru_cache(maxsize=1)
def get_spacy_model(model_name: str = "en_core_web_sm"):
    """
    Load spaCy model once (cached). Use en_core_web_sm by default.
    """
    try:
        return spacy.load(model_name, disable=["parser", "ner"])
    except Exception as e:
        # re-raise with friendly guidance
        raise RuntimeError(f"Failed to load spaCy model '{model_name}'. Run 'python -m spacy download {model_name}'. Original: {e}")

def to_df_from_reviews(reviews_list: List[dict]) -> pd.DataFrame:
    """Normalize google_play_scraper review dicts into a DataFrame and add cleaned and lemmatized columns."""
    rows = []
    analyzer = SentimentIntensityAnalyzer()
    for r in reviews_list:
        rows.append({
            "userName": r.get("userName"),
            "content": r.get("content") or "",
            "score": r.get("score"),
            "thumbsUpCount": r.get("thumbsUpCount"),
            "reviewCreatedVersion": r.get("reviewCreatedVersion"),
            "at": r.get("at")
        })
    df = pd.DataFrame(rows)
    if "at" in df.columns:
        try:
            df["at"] = pd.to_datetime(df["at"])
        except Exception:
            pass

    # basic cleaned text (useful for sentiment + further processing)
    df["cleaned"] = df["content"].apply(clean_text)
    
    # cleaned + lemmatized (useful for TF-IDF / grouping morphological variants)
    # keep lemmatization optional in UI; compute here for simplicity (cached by process; spaCy cached internally)
    try:
        df["cleaned_lemma"] = df["cleaned"].apply(normalize_text_with_lemmatizer)
    except RuntimeError as e:
        # spaCy model not loaded; fallback to cleaned only
        st.warning(f"spaCy lemmatizer unavailable: {e}. Using cleaned text only.")
        df["cleaned_lemma"] = df["cleaned"]
    
    df["compound"] = df["cleaned_lemma"].apply(lambda x: analyzer.polarity_scores(x)['compound'])
    df["label"] = df["compound"].apply(label_from_compound)
    
    return df

def normalize_token(token_text: str) -> str:
    """
    Normalize a single token: irregular map -> spaCy lemma -> lower() fallback.
    """
    if not token_text:
        return ""
    low = token_text.lower()
    if low in _IRREGULAR_ADJ:
        return _IRREGULAR_ADJ[low]
    nlp = get_spacy_model()
    doc = nlp(token_text)
    if len(doc) == 0:
        return low
    lemma = doc[0].lemma_.lower()
    if lemma == "-pron-":
        return doc[0].text.lower()
    return lemma or low


def normalize_text_with_lemmatizer(text: str) -> str:
    """
    Normalize (lemmatize) whitespace-split tokens in `text`. Assumes text is pre-cleaned.
    """
    if not isinstance(text, str):
        return ""
    parts = text.split()
    normalized = [normalize_token(p) for p in parts]
    return " ".join([p for p in normalized if p])

# --- Sentiment + keywords ---
def analyze_sentiments(texts: List[str], analyzer: Optional[SentimentIntensityAnalyzer] = None) -> List[Dict[str, Any]]:
    if analyzer is None:
        analyzer = SentimentIntensityAnalyzer()
    results = []
    for t in texts:
        cleaned = clean_text(t)
        compound = analyzer.polarity_scores(cleaned)["compound"]
        label = label_from_compound(compound)
        results.append({"text": t, "cleaned": cleaned, "compound": compound, "label": label})
    return results

def top_tfidf_terms(texts: List[str], top_n: int = 20, ngram_range=(1,2)) -> List[Tuple[str, float]]:
    """Return top_n terms by average TF-IDF score (excludes stopwords)."""
    if not texts:
        return []
    vect = TfidfVectorizer(stop_words=list(STOPWORDS), ngram_range=ngram_range, max_features=2000)
    X = vect.fit_transform(texts)
    scores = X.mean(axis=0).A1  # average tf-idf across docs
    terms = vect.get_feature_names_out()
    term_scores = sorted(zip(terms, scores), key=lambda x: x[1], reverse=True)[:top_n]
    return term_scores

def label_from_compound(c: float, pos_thresh: float = 0.5, neg_thresh: float = -0.5) -> str:
    if c >= pos_thresh:
        return "positive"
    if c <= neg_thresh:
        return "negative"
    return "neutral"

def get_spacy_stopwords(model_name: str = "en_core_web_sm"):
    """
    Return a set of stopwords from the spaCy model Defaults.
    Uses the same model as get_spacy_model() so lists align.
    """
    nlp = get_spacy_model(model_name)
    return set(nlp.Defaults.stop_words)

def remove_stopwords_from_text(text: str, stopword_set: set) -> str:
    """
    Remove stopwords from whitespace-tokenized text (expects lemmatized tokens).
    Keeps token order, filters by membership in stopword_set (case-insensitive).
    """
    if not isinstance(text, str) or text.strip() == "":
        return ""
    tokens = text.split()
    filtered = [t for t in tokens if t.lower() not in stopword_set]
    return " ".join(filtered)


with st.sidebar:
    st.header("Search & Options")
    app_name = st.text_input("App name (search)", value="WhatsApp")
    lang = st.selectbox("Language", ["en", "in"], index=0)    
    country = st.selectbox("Country", ["us", "in", "gb"], index=1)
    top_n = st.number_input("Search top N results to show", min_value=1, max_value=200, value=20, step=1)
    max_workers = st.slider("Parallel workers (metadata fetch)", min_value=1, max_value=20, value=6)

    # 1) Search + metadata
    with st.spinner("Searching Play Store..."):
        df_apps = search_and_enrich(app_name, lang=lang, country=country, top_n=top_n, max_workers=max_workers)
        if df_apps.empty:
            st.warning("No search results found. Try a different app name or expand top N.")
            st.stop()
        else:
            pkg_choice = st.selectbox("Choose package to fetch reviews", options=df_apps['appId'].tolist())
            mode = st.selectbox("Fetch mode", ["Latest N reviews", "All reviews"])
            if mode == "Latest N reviews":
                n = st.number_input("N (how many reviews you want)", min_value=1, max_value=5000, value=100, step=10)
            
            st.caption("Notes: `All reviews` can be slow for apps with many reviews — use with caution. Metadata and review fetches are cached during the Streamlit session to reduce repeat network calls.")

if list(search(app_name, lang=lang, country=country)):
    print("A")
else:
    print("B")
    

results = list(search(pkg_choice, lang=lang, country=country))

df_apps['installs'] = df_apps['installs'].str.extract(r'([-+]?\d[\d,\.]*)', expand=False)
df_apps['installs']  = pd.to_numeric(df_apps['installs'].str.replace(',', ''), errors='coerce')

data = df_apps[df_apps['appId'] == pkg_choice].iloc[0].to_dict()

# helper
def fmt_int(n):
    if n >= 1_000_000:
        return f"{n//1_000_000}M"
    if n >= 1_000:
        return f"{n//1_000}K"
    return str(n)

st.markdown(
    f"""
    <table width="100%" style="
        position: relative;
        background: rgba(5,10,15,0.86);
        background-size: cover;
        background-position: right center;
        color: #ffffff;
        padding: 40px 36px;
        border-radius: 10px;
        box-shadow: 0 8px 28px rgba(10,20,30,0.35);
        margin-bottom: 18px;
    ">
        <tr>
            <th width="150px">
                <img style=" width: 96px; height: 96px; border-radius: 18px; box-shadow: 0 8px 20px rgba(0,0,0,0.4);" src="{data['icon']}" alt="app icon" />
            </th>
            <th>
                <table style="border-radius: 10px; width: 100%;">
                    <tr>
                        <th style="font-size: 44px; font-weight: 800; margin: 6px;">
                            {data['title']}
                        </th>
                    </tr>
                    <tr>
                        <th style="color: #7fd08f; margin-bottom: 18px; font-weight: 600; font-size: 14px;">
                            {data['developer']}
                        </th>
                    </tr>
                    <tr>
                        <th>
                            <table width="100%">
                                <tr>
                                    <th style="font-weight:800; font-size:20px; text-align:center; width:33%; background: rgba(255,255,255,0.06); padding:8px 12px; border-radius:10px;">
                                        {data['score']:.1f}★
                                    </th>
                                    <th>|</th>
                                    <th style="font-weight:800; font-size:20px; text-align:center; width:33%; background: rgba(255,255,255,0.06); padding:8px 12px; border-radius:10px;">
                                        {fmt_int(data['Number of reviews'])}<br>reviews
                                    </th>
                                    <th>|</th>
                                    <th style="font-weight:800; font-size:20px; text-align:center; width:33%; background: rgba(255,255,255,0.06); padding:8px 12px; border-radius:10px;">
                                        {fmt_int(data['installs'])}+<br>Downloads
                                    </th>
                                </tr>
                            </table>
                        </th>
                    </tr>
                    <tr>
                        <th colspan="5">
                            <a href="https://play.google.com/store/search?q={quote_plus(data['title'])}" target="_blank"
                                style="text-decoration:none; display:inline-block; padding:10px 16px; border-radius:12px; background: linear-gradient(90deg,#22c55e,#16a34a); color:white; font-weight:700;">
                                View in Play Store
                            </a>
                        </th>
                    </tr>
                </table>
            </th>
        </tr>
    </table>
    """,
    unsafe_allow_html=True,
)

#st.subheader("Search results (top)")
with st.expander("Search Results", expanded=False):
    st.dataframe(df_apps.drop(columns={"icon","description","URL"}), hide_index = True)


with st.spinner("Fetching reviews..."):
    if mode == "Latest N reviews":
        reviews_list = fetch_reviews(pkg_choice, lang=lang, country=country, sort=Sort.NEWEST, count=int(n))
    else:
        reviews_list = fetch_reviews(pkg_choice, lang=lang, country=country, sort=Sort.NEWEST, fetch_all = True)

if not reviews_list:
    st.error("No reviews returned for the selected package.")
    st.stop()

df_reviews = to_df_from_reviews(reviews_list)


# 3) Sentiment analysis
with st.spinner("Analyzing sentiment..."):
    stop_set = get_spacy_stopwords()
    df_reviews["cleaned_lemma_nostop"] = df_reviews["cleaned_lemma"].apply(lambda t: remove_stopwords_from_text(t, stop_set))
    
    counts = Counter(df_reviews["label"])
    total = len(df_reviews) or 1
    percent_summary = {k: f"{v} ({v/total*100:.1f}%)" for k, v in counts.items()}

st.markdown("""<div class="heading-main">Sentiment Summary</div>""", unsafe_allow_html=True,)

col1, col2, col3 = st.columns(3)
with col1:
    st.info(f"Neutral: {percent_summary.get('neutral', '0 (0.0%)')}")
with col2:
    st.success(f"Positive: {percent_summary.get('positive', '0 (0.0%)')}")
with col3:
    st.error(f"Negative: {percent_summary.get('negative', '0 (0.0%)')}")

df = df_reviews['score'].value_counts().sort_index().reset_index()
df['score'] = df['score'].astype(str)
df['pct'] = (df['count'] / df['count'].sum() * 100).round(1)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""<div class="heading-main">Review Score Distribution</div>""", unsafe_allow_html=True,)

    color_map = {
        "1": "#d9534f",   # red
        "2": "#f7a8a8",   # light red
        "3": "#f7d86b",   # yellow
        "4": "#a8e6a1",   # light green
        "5": "#5cb85c"    # green
    }

    fig = px.bar(
        df,
        x="count",
        y="score",
        orientation="h",
        text="count",
        labels={"count": "", "score": "Score"},
        height=320,
        color="score",
        color_discrete_map=color_map,
        category_orders={"score": ["1", "2", "3", "4", "5"]}  # keep explicit order
    )

    # show percent in hover (customdata expects 2D array-like)
    fig.update_traces(
        hovertemplate="Score: %{y}<br>Count: %{x:,}<br>Percent: %{customdata[0]}%",
        customdata=df[["pct"]].values,
        marker_line_width=0,
        texttemplate="%{text}",       
        textposition="outside"     
    )

    fig.update_layout(
        yaxis=dict(autorange="reversed", showticklabels=False),   # hide y ticks
        xaxis=dict(showticklabels=False),                         # hide x ticks
        margin=dict(l=60, r=20, t=20, b=20),
        bargap=0.15
    )

    st.plotly_chart(fig, use_container_width=True)

with col2:
    # 4) Keyword extraction
    with st.spinner("Computing top TF-IDF terms..."):
        top_terms = top_tfidf_terms(df_reviews["cleaned_lemma_nostop"].fillna("").tolist(), top_n=30)
    if top_terms:
        st.markdown("""<div class="heading-main">Top keywords/phrases</div>""", unsafe_allow_html=True,)
        df_terms = pd.DataFrame(top_terms, columns=["term", "avg_tfidf"])
        st.dataframe(df_terms.head(5), hide_index = True)
    else:
        st.info("No keywords found (empty text).")


# 5) Scam / search term finder
# Custom CSS for expander styling
st.markdown("""
<style>
div[data-testid="stExpander"] {
    position: relative;
    background: rgba(5,10,15,0.86);
    background-size: cover;
    background-position: right center;
    color: #ffffff;
    padding: 20px;
    border-radius: 10px;
    box-shadow: 0 8px 28px rgba(10,20,30,0.35);
    margin-bottom: 18px;
}

/* Expander header (title) styling */
div[data-testid="stExpander"] > div:first-child {
    font-weight: 600;
    font-size: 1.1rem;
    color: #f0f0f0;
}

/* Optional: hover effect */
div[data-testid="stExpander"]:hover {
    box-shadow: 0 12px 32px rgba(10,20,30,0.5);
}
</style>
""", unsafe_allow_html=True)

# Expander component
with st.expander("🔍 Search within reviews", expanded=True):
    search_term = st.text_input("Enter term to search in reviews (regex allowed)", value="scam")
    if search_term:
        try:
            pattern = re.compile(search_term, flags=re.I)
            mask = df_reviews["content"].fillna("").str.contains(pattern)
            matched = df_reviews[mask].drop(columns = ["cleaned", "compound", "label","cleaned_lemma","cleaned_lemma_nostop"])
            st.markdown(f"Found {len(matched)} matching reviews for `{search_term}`.")
            if not matched.empty:
                st.dataframe(matched.head(50))
        except re.error:
            st.error("Invalid regex pattern. Try a simpler term.")


with st.expander("Reviews", expanded=True):
    st.dataframe(df_reviews.drop(columns = ["cleaned", "compound", "label","cleaned_lemma","cleaned_lemma_nostop"]), hide_index = True)

import base64

# ---------- FOOTER CONFIG ----------
USER_NAME = "CodeMitr"
LOGO_PATH = "codemitr.png"  # your local logo file
LOGO_MAX_HEIGHT = 100  # in px

# Convert logo to base64 (so it displays even when app is deployed)
with open(LOGO_PATH, "rb") as f:
    data = f.read()
encoded_logo = base64.b64encode(data).decode()
logo_data_url = f"data:image/png;base64,{encoded_logo}"


# ---------- STYLES ----------
footer_css = f"""
<style>
/* ===== FOOTER CONTAINER ===== */
.app-footmark-bar {{
  position: fixed;
  bottom: 0;
  left: 0;
  width: 100%;
  background: rgba(5, 10, 15, 0.85);
  backdrop-filter: blur(8px);
  border-top: 1px solid rgba(255, 255, 255, 0.1);
  box-shadow: 0 -2px 18px rgba(0, 0, 0, 0.25);
  padding: 5px 0;
  display: flex;
  justify-content: center;
  z-index: 9999;
}}

.app-footmark {{
  width: 90%;
  max-width: 1100px;
  display: flex;
  align-items: center;
  justify-content: flex-start;
  gap: 5px;
  color: #ffffff;
  font-family: "Inter", sans-serif;
  font-size: 14px;
  pointer-events: none;
}}

.app-footmark img {{
  height: {LOGO_MAX_HEIGHT}px;
  width: auto;
  border-radius: 6px;
  pointer-events: auto;
}}

.app-footmark .text {{
  display: flex;
  flex-direction: column;
  pointer-events: auto;
}}

.app-footmark .name {{
  font-weight: 600;
  font-size: 15px;
  letter-spacing: 0.3px;
  color: #f3f4f6;
}}

.app-footmark .meta {{
  font-size: 12px;
  color: rgba(243,244,246,0.8);
}}

.app-footmark:hover .name {{
  color: #8ab4f8;
  transition: color 0.3s ease-in-out;
}}
</style>
"""

# ---------- HTML FOOTER ----------
footer_html = f"""
{footer_css}
<div class="app-footmark-bar">
  <div class="app-footmark">
    <img src="{logo_data_url}" alt="logo" />
    <div class="text">
      <div class="name">{USER_NAME}</div>
      <div class="meta">Built with Streamlit</div>
    </div>
  </div>
</div>
"""

# ---------- RENDER ----------
with st.sidebar:
    st.markdown(footer_html, unsafe_allow_html=True)

# Add bottom padding so footer doesn’t overlap content
st.markdown("<div style='height:70px'></div>", unsafe_allow_html=True)


