import streamlit as st
import pandas as pd
import numpy as np
import re, io, os, time, datetime as dt, unicodedata, textwrap
from urllib.parse import unquote
import matplotlib.pyplot as plt

# ================== AYAR ==================
DEFAULT_EXCEL_PATH = "BÜTÇE ÇALIŞMA GÜNCEL.xlsx"

st.set_page_config(page_title="Bütçe Uygulaması", page_icon="💰")
st.title("Bütçe Uygulaması 💰")

# ================== YARDIMCI ==================
def speak(text: str):
    st.components.v1.html(
        "<script>try{const u=new SpeechSynthesisUtterance("
        + repr(str(text)) +
        ");u.lang='tr-TR';speechSynthesis.cancel();speechSynthesis.speak(u);}catch(e){}</script>", height=0
    )

def get_query_param(name: str):
    try:
        qp = st.query_params
        val = qp.get(name)
        if isinstance(val, list):
            return val[0] if val else None
        return val
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def read_excel_path(path: str, mtime: float) -> pd.DataFrame:
    # Tüm sütunları METİN olarak oku (Excel'in sayı formatına karışmasına izin verme)
    return pd.read_excel(path, dtype=str, keep_default_na=False)

def _strip_accents(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def _canon(s: str) -> str:
    s = (s or "").strip().lower()
    s = _strip_accents(s)
    s = re.sub(r"[^a-z0-9çğıöşü]+", "", s)
    return s

def norm(x: str) -> str:
    s = (x or "").strip().lower()
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")

def _likely_personref_col(cols):
    KEYS = [
        "personref","person ref","personel ref","personelref",
        "personel no","personelno","sicil","sicil no","sicilno",
        "ref","ref no","refno","employee id","employeeid","id"
    ]
    cn = {_canon(c): c for c in cols}
    for k in KEYS:
        if _canon(k) in cn:
            return cn[_canon(k)]
    for c in cols:
        cc = _canon(c)
        if ("ref" in cc) or ("sicil" in cc):
            return c
    return None

def tl(x):
    try: return f"{x:,.2f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception: return f"{x} TL"

def pct(x):
    try:
        return f"{x*100:,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "-"

def get_numeric(val, default=0.0):
    try:
        if pd.isna(val): return float(default)
        return float(val)
    except Exception: return float(default)

# TR/US & parantezli muhasebe sayılarını güvenli parse
def parse_number_tr(s):
    if s is None:
        return 0.0
    t = str(s).strip().replace("\u00A0", "")
    if t == "" or t.lower() in {"nan","none","-"}:
        return 0.0
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1].strip()
    # TR biçimi
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})*(,\d+)?", t) or re.fullmatch(r"-?\d+,\d+", t):
        t = t.replace(".", "").replace(",", ".")
        try: v = float(t); return -v if neg else v
        except: return 0.0
    # US biçimi
    if re.fullmatch(r"-?\d{1,3}(,\d{3})*(\.\d+)?", t):
        t = t.replace(",", "")
        try: v = float(t); return -v if neg else v
        except: return 0.0
    # Düz sayı
    try:
        v = float(t.replace(",", "."))
        return -v if neg else v
    except:
        return 0.0

# ==== TR sayı kelimeleri (ses için) ====
TR1={"sıfır":0,"sifir":0,"bir":1,"iki":2,"üç":3,"uc":3,"dört":4,"dort":4,"beş":5,"bes":5,"altı":6,"alti":6,"yedi":7,"sekiz":8,"dokuz":9}
TR10={"on":10,"yirmi":20,"otuz":30,"kırk":40,"kirk":40,"elli":50,"altmış":60,"altmis":60,"yetmiş":70,"yetmis":70,"seksen":80,"doksan":90}
TRM={"yüz":100,"yuz":100,"bin":1000}
def parse_tr_words(words):
    total=0; cur=0; used=False
    for w in words:
        w=w.lower()
        if w in TR1: cur+=TR1[w]; used=True
        elif w in TR10: cur+=TR10[w]; used=True
        elif w in TRM:
            mul=TRM[w]
            if mul==100: cur=(cur or 1)*100
            else: cur=(cur or 1)*mul; total+=cur; cur=0
            used=True
        else:
            if used: break
    total+=cur
    return total if used and total>0 else None

def splitw(txt):
    return [re.sub(r"[^a-zçğıöşü0-9]", "", w.lower()) for w in txt.split()]

# ==== PersonRef çıkarımı (metinden) ====
def extract_personref(txt):
    txt = txt or ""
    m = re.search(r"(?:person|ref|sicil|kişi|kisi)\D*([0-9][0-9\s]{3,})", txt, re.I)
    if m:
        d = re.sub(r"\D", "", m.group(1))
        if d.isdigit() and len(d) >= 4:
            return d, d
    m2 = re.search(r"\b(\d[ \d]{3,})\b", txt)
    if m2:
        d = re.sub(r"\D", "", m2.group(1))
        if d.isdigit() and len(d) >= 4:
            return d, d
    return None, None

def extract_amount(txt, pref_digits):
    m=re.search(r"(\d[\d\.\,]*)\s*(tl|lira)?\b", txt, re.I)
    if m:
        raw=m.group(1).replace(".","").replace(",",".")
        try:
            val=float(raw); return val if val>0 else None
        except: pass
    ws=splitw(txt)
    if "tl" in ws or "lira" in ws:
        idxs=[i for i,w in enumerate(ws) if w in ("tl","lira")]
        for idx in reversed(idxs):
            val=parse_tr_words(ws[max(0,idx-6):idx])
            if val: return float(val)
    toks=re.findall(r"\d[\d\.\,]*", txt)
    if pref_digits: toks=[t for t in toks if re.sub(r"\D","",t)!=pref_digits]
    if toks:
        raw=toks[-1].replace(".","").replace(",",".")
        try: val=float(raw); return val if val>0 else None
        except: pass
    val=parse_tr_words(ws[::-1])
    return float(val) if val else None

# ==== İsimden kişi bulma (digits döndürür) ====
def build_fullname_columns(df: pd.DataFrame) -> pd.DataFrame:
    out=df.copy()
    full = None
    for c in out.columns:
        if _canon(c) in {"adsoyad","adsoyadi","ad soyad","ad soyadi"}:
            full = out[c].astype(str).fillna("").str.strip(); break
    if full is None:
        ad_col=None; soyad_col=None
        for c in out.columns:
            if _canon(c) in {"ad","adi","isim"}: ad_col=c
            if _canon(c) in {"soyad","soyadi"}: soyad_col=c
        if ad_col and soyad_col:
            full = (out[ad_col].astype(str).fillna("") + " " + out[soyad_col].astype(str).fillna("")).str.strip()
    out["FULLNAME"] = full if full is not None else ""
    out["FULLNAME_NORM"]=out["FULLNAME"].astype(str).map(_canon)
    return out

@st.cache_data(show_spinner=False)
def normalize_all(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()

    # ---- PersonRef kolonunu bul/oluştur ----
    if "PersonRef" not in df.columns:
        guess = _likely_personref_col(df.columns)
        if guess:
            df.rename(columns={guess: "PersonRef"}, inplace=True)
        else:
            df["PersonRef"] = ""

    raw = df["PersonRef"].astype(str).str.strip().str.replace(r"\s+", "", regex=True)

    # --- EK TEMİZLEME (ör. ...160000 ya da ...0000) ---
    def _clean_ref(s: str) -> str:
        d = re.sub(r"\D", "", s)
        if re.fullmatch(r"\d+160000", d):
            d = re.sub(r"160000$", "", d)
        elif re.fullmatch(r"\d+0000", d) and len(d) >= 6:
            d = re.sub(r"0{4}$", "", d)
        s2 = re.sub(r"^(\d+)[\.,]0+$", r"\1", s)  # "12345.0" vb.
        d2 = re.sub(r"\D", "", s2)
        out = d or d2
        return out or s

    df["PersonRef_raw"] = raw.map(_clean_ref)                      # ekranda gösterilecek
    df["PersonRef_digits"] = df["PersonRef_raw"].str.replace(r"\D", "", regex=True)  # eşleştirme için

    # Diğer alanlar
    c2orig = {_canon(c): c for c in df.columns}
    def need(name, alts, default):
        if name in df.columns: return
        src = None
        if _canon(name) in c2orig: src = c2orig[_canon(name)]
        else:
            for a in alts:
                if _canon(a) in c2orig: src = c2orig[_canon(a)]; break
        if src: df.rename(columns={src: name}, inplace=True)
        else: df[name] = default

    need("CurrentSalary", ["mevcut maaş","mevcut ucret","salary","maas"], "0")
    need("BÜTÇE DIŞI TALEPLER İLE", ["butce disi","budget extra","ekstra","toplam butce disi"], "0")
    need("DEPARTMAN", ["departman","bölüm","bolum","department","birim"], "")

    for y in ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]:
        if y not in df.columns: df[y] = ""
        df[y] = df[y].fillna("").astype(str)

    # Sayısal kolonları (TR/US/paren formatlarını da okuyacak şekilde) çevir
    for c in ["CurrentSalary","BÜTÇE DIŞI TALEPLER İLE"]:
        df[c] = df[c].apply(parse_number_tr)

    df = build_fullname_columns(df)

    # === Metrikler ===
    cur = df["CurrentSalary"].fillna(0.0).astype(float)
    bdt = df["BÜTÇE DIŞI TALEPLER İLE"].fillna(0.0).astype(float)

    df["KALAN TL"] = cur*1.40 - bdt
    with np.errstate(divide="ignore", invalid="ignore"):
        df["KULLANILAN BÜTÇE ORANI"] = (bdt/cur) - 1.0
        df["SİSTEM DIŞI KALAN TL"]   = cur*1.41 - bdt
        df["SİSTEM DIŞI İLE KULLANILAN ORAN"] = np.where(bdt>0, (cur*1.41)/bdt - 1.0, 0.0)

    return df

def find_personref_by_name(df: pd.DataFrame, text: str):
    norm_t=_canon(text)
    best_len=0; best_digits=None; best_name=None
    for _,row in df.iterrows():
        fn=str(row.get("FULLNAME","") or "")
        fnn=str(row.get("FULLNAME_NORM","") or "")
        if fnn and fnn in norm_t:
            digits=str(row.get("PersonRef_digits","") or "")
            if digits:
                L=len(fnn)
                if L>best_len:
                    best_len=L; best_digits=digits; best_name=fn
    return (best_digits, best_name) if best_digits else (None, None)

def manager_chain(row):
    mans = [str(row.get(k,"")).strip() for k in ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]]
    mans = [m for m in mans if m]
    return " > ".join(mans) if mans else ""

# ================== STATE ==================
defaults = {
    "_last_voice": "",
    "history": [],
    "unsaved_ops": [],
    "selected_ref": None,          # digits string
    "force_listen": True,
    "listening": True,
    "last_final_text": "",
    "sticky_amount": None,
    "sticky_amount_ts": 0.0,
    "auto_apply": True,
}
for k,v in defaults.items():
    if k not in st.session_state: st.session_state[k]=v

def set_sticky_amount(val: float):
    st.session_state.sticky_amount = float(val)
    st.session_state.sticky_amount_ts = time.time()
def get_sticky_amount():
    if st.session_state.sticky_amount and (time.time()-st.session_state.sticky_amount_ts)<=30.0:
        return float(st.session_state.sticky_amount)
    return None

st.session_state.listening = bool(st.session_state.get("force_listen", True))

# ================== SİDEBAR - AYARLAR ==================
with st.sidebar:
    st.header("⚙️ Ayarlar")
    st.session_state.auto_apply = st.toggle("🎤 Sesle otomatik uygula", value=st.session_state.get("auto_apply", True))

# ================== VERİ YÜKLEME ==================
with st.sidebar:
    st.header("📄 Veri Kaynağı")
    use_default = st.toggle("Varsayılan dosya (BÜTÇE ÇALIŞMA GÜNCEL.xlsx)", value=True)

try:
    file_mtime = os.path.getmtime(DEFAULT_EXCEL_PATH) if use_default and os.path.exists(DEFAULT_EXCEL_PATH) else 0.0
    base_df = read_excel_path(DEFAULT_EXCEL_PATH, file_mtime) if use_default else st.stop()
except FileNotFoundError:
    st.error(f"'{DEFAULT_EXCEL_PATH}' bulunamadı."); st.stop()
except Exception as e:
    st.error(f"Excel okunamadı: {e}"); st.stop()

if "df" not in st.session_state or st.session_state.df is None:
    st.session_state.df = normalize_all(base_df)
else:
    st.session_state.df = normalize_all(st.session_state.df)
df = st.session_state.df

# ================== FİLTRE ==================
with st.sidebar:
    st.header("🎛️ Filtreler & İşlemler")
    mgr_cols = ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]

    managers_raw = pd.concat([df[c].fillna("").astype(str) for c in mgr_cols], ignore_index=True)
    managers_raw = managers_raw[managers_raw.str.strip() != ""]
    opts = sorted(managers_raw.unique().tolist())
    selected_manager = st.selectbox("İşlem yapılacak yönetici", opts if opts else ["(yok)"])

selected_key = norm(selected_manager) if opts else ""
mask = pd.Series(False, index=df.index)
for c in ["1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ"]:
    mask = mask | (df[c].astype(str).map(norm) == selected_key)

if opts and selected_manager != "(yok)":
    df_filtered = df[mask].copy()
    if not mask.any():
        st.warning("Bu yöneticiye ait kayıt bulunamadı (yazım farkı olabilir).")
else:
    df_filtered = df.copy()

# =========== PERSONREF SEÇİMİ: LİSTEDEN ===========
with st.sidebar:
    st.markdown("—")
    st.subheader("👤 Kişi Seç (PersonRef)")
    df_sel = df_filtered[df_filtered["PersonRef_digits"].str.len() > 0].copy()
    if len(df_sel)==0:
        st.info("Bu yönetici için kişi bulunamadı.")
        selected_ref = None
    else:
        df_sel["LABEL"] = df_sel["PersonRef_raw"].astype(str) + " — " + df_sel["FULLNAME"].astype(str)
        df_sel = df_sel.reset_index(drop=True)
        default_index = 0
        if st.session_state.get("selected_ref"):
            prev = str(st.session_state["selected_ref"])
            hit = df_sel.index[df_sel["PersonRef_digits"] == prev]
            if len(hit) > 0:
                default_index = int(hit[0])
        choice = st.selectbox("Kişi", df_sel["LABEL"].tolist(), index=default_index)
        row = df_sel.loc[df_sel["LABEL"] == choice].iloc[0]
        st.session_state.selected_ref = str(row["PersonRef_digits"])
        selected_ref = st.session_state.selected_ref
        st.success(f"Seçili PersonRef: {row['PersonRef_raw']}")

# ================== KPI (4 başlık) ==================
cur_sum = df_filtered["CurrentSalary"].fillna(0).astype(float).sum()
bdt_sum = df_filtered["BÜTÇE DIŞI TALEPLER İLE"].fillna(0).astype(float).sum()

kalan_tl_sum = (df_filtered["KALAN TL"].fillna(0)).sum()
kullanilan_oran = (bdt_sum/cur_sum - 1.0) if cur_sum>0 else 0.0
sistem_disi_kalan_tl_sum = (cur_sum*1.41) - bdt_sum
sistem_disi_kullanilan_oran = ((cur_sum*1.41)/bdt_sum - 1.0) if bdt_sum>0 else 0.0

c1,c2,c3,c4 = st.columns(4)
c1.metric("KALAN TL", tl(kalan_tl_sum))
c2.metric("KULLANILAN BÜTÇE ORANI", pct(kullanilan_oran))
c3.metric("SİSTEM DIŞI KALAN TL", tl(sistem_disi_kalan_tl_sum))
c4.metric("SİSTEM DIŞI İLE KULLANILAN ORAN", pct(sistem_disi_kullanilan_oran))

# Doğrulama satırı: Excel pivot ile birebir kıyas için
st.caption(
    ("Seçili yönetici satır sayısı: "
     f"{len(df_filtered):,} • Toplam CurrentSalary: {cur_sum:,.2f} • Toplam BDT: {bdt_sum:,.2f}")
    .replace(',', 'X').replace('.', ',').replace('X','.')
)

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# ================== TABLO (salt görüntü) ==================
view_cols = [
    "PersonRef_raw","FULLNAME","DEPARTMAN",
    "1.YÖNETİCİSİ","2.YÖNETİCİSİ","3.YÖNETİCİSİ","4.YÖNETİCİSİ",
    "CurrentSalary","BÜTÇE DIŞI TALEPLER İLE",
    "KALAN TL","KULLANILAN BÜTÇE ORANI","SİSTEM DIŞI KALAN TL",
    "SİSTEM DIŞI İLE KULLANILAN ORAN"
]
df_show = df_filtered[view_cols].copy()
df_show.rename(columns={"PersonRef_raw":"PersonRef"}, inplace=True)
for oc in ["KULLANILAN BÜTÇE ORANI","SİSTEM DIŞI İLE KULLANILAN ORAN"]:
    df_show[oc] = df_filtered[oc].apply(pct)
st.dataframe(df_show, use_container_width=True, height=420)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ================== SİDEBAR İŞLEM ALANLARI ==================
with st.sidebar:
    st.subheader("🛠️ İşlem")
    manuel_ref = st.text_input("Veya Manuel PersonRef", key="manuel_ref_input")
    if manuel_ref:
        manuel_ref_digits = re.sub(r"\D", "", manuel_ref)
        if manuel_ref_digits:
            st.session_state.selected_ref = manuel_ref_digits
    tutar = st.number_input("Tutar (TL) — (istersen boş bırak)", step=100.0, min_value=0.0, value=0.0, key="tutar_input")
    islem = st.radio("İşlem Türü", ["Kalan TL’den Düş","Kalan TL’ye Ekle"], index=0, key="islem_radio")

# ================== İŞLEM (KALAN TL) ==================
def _apply_kalan_change(dff, idx, delta_sign, tutar):
    """
    delta_sign=+1 -> Kalan TL’ye Ekle  → BDT -= tutar
    delta_sign=-1 -> Kalan TL’den Düş  → BDT += tutar
    KALAN = cur*1.40 - BDT  ==>  BDT += (-delta_sign) * tutar
    """
    bd = get_numeric(dff.at[idx, "BÜTÇE DIŞI TALEPLER İLE"], 0.0)
    yeni_bdt = bd + (-delta_sign) * float(tutar)
    if yeni_bdt < 0:
        yeni_bdt = 0.0
    dff.at[idx, "BÜTÇE DIŞI TALEPLER İLE"] = float(yeni_bdt)
    return dff

def islem_yap(person_ref_digits: str, tutar: float, islem_tipi: str, announce=True, do_rerun=True):
    dff = st.session_state.df.copy()

    ser = dff["PersonRef_digits"].astype(str)
    idxs = dff.index[ser == str(person_ref_digits)]
    if len(idxs) == 0:
        st.warning("Girilen PersonRef ile eşleşen kişi bulunamadı.")
        if announce: speak("Girilen kişi bulunamadı.")
        return
    i = int(idxs[0])

    pre_kalan = get_numeric(dff.at[i, "KALAN TL"], 0.0)

    if islem_tipi == "Kalan TL’den Düş":
        dff = _apply_kalan_change(dff, i, delta_sign=-1, tutar=tutar)
        verb = "Kalan TL’den düşüldü"
    elif islem_tipi == "Kalan TL’ye Ekle":
        dff = _apply_kalan_change(dff, i, delta_sign=+1, tutar=tutar)
        verb = "Kalan TL’ye eklendi"
    else:
        st.warning("Bilinmeyen işlem tipi.")
        return

    dff = normalize_all(dff)
    try: st.cache_data.clear()
    except: pass
    st.session_state.df = dff

    mask = dff["PersonRef_digits"].astype(str) == str(person_ref_digits)
    j = int(dff.index[mask][0]) if mask.any() else i

    post_kalan = get_numeric(dff.at[j, "KALAN TL"], 0.0)
    row = dff.loc[j]

    st.session_state.unsaved_ops.append({
        "Zaman": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "PersonRef": row.get("PersonRef_raw", ""),
        "AdSoyad": str(row.get("FULLNAME", "") or ""),
        "Departman": str(row.get("DEPARTMAN", "") or ""),
        "Yöneticiler": manager_chain(row),
        "Tür": islem_tipi,
        "Havuz": "KalanTL",
        "Tutar": float(tutar),
        "Önce_KalanTL": float(pre_kalan),
        "Sonra_KalanTL": float(post_kalan),
    })

    st.success(f"İşlem uygulandı: {tutar:.2f} TL {verb}. (Geçmiş: Kaydet ile)")
    if announce:
        speak(f"{int(round(float(tutar)))} lira {verb}. Kaydet tuşuyla geçmişe eklenecek.")
    if do_rerun:
        st.rerun()

# ================== SES / KOMUT PARSE ==================
def parse_op_from_text(text: str, fallback_ui_op: str | None = None) -> str | None:
    t=(text or "").lower()
    if re.search(r"\b(düş|dus|düşür|çıkar|cikar|eksilt|azalt)\b", t): return "Kalan TL’den Düş"
    if re.search(r"\b(ekle|arttır|artır|yükselt|yukselt)\b", t):    return "Kalan TL’ye Ekle"
    return fallback_ui_op

def handle_command(text: str, ui_amount: float, ui_islem: str, ui_selected_ref: str|None, auto_apply: bool = True):
    df = st.session_state.df
    t = text.lower()
    trigger = any(k in t for k in ["işlem yap","islem yap","hemen uygula","uygula","onayla","tamam"])

    pref, pref_digits = extract_personref(t)
    if pref is None and ui_selected_ref is not None:
        pref = ui_selected_ref
    if pref is None:
        pref_by_name, name_found = find_personref_by_name(df, t)
        if pref_by_name is not None:
            pref = pref_by_name
            speak(f"{name_found} bulundu.")

    amt_voice = extract_amount(t, pref_digits)
    if amt_voice:
        set_sticky_amount(amt_voice)
    amt = float(amt_voice) if (amt_voice and amt_voice>0) else (float(ui_amount) if ui_amount and float(ui_amount)>0 else (get_sticky_amount() or None))
    op = parse_op_from_text(t, fallback_ui_op=ui_islem)

    if op and amt is not None and pref is not None:
        if auto_apply or trigger:
            islem_yap(str(pref), float(amt), op); return
        else:
            st.info("Komut çözüldü. 'İşlem Yap' butonuyla uygulayabilirsiniz.")
            speak("Komut hazır. İşlem Yap'a basın."); return

    if trigger:
        missing=[]
        if pref is None: missing.append("kişi (seçin/PersonRef söyleyin)")
        if amt is None:  missing.append("tutar (söyleyin ya da girin)")
        if op is None:   missing.append("işlem türü (düş/ekle)")
        if not missing:
            islem_yap(str(pref), float(amt), op); return
        msg=" , ".join(missing) + "."
        st.warning(msg); speak(msg)
    else:
        if amt is None:
            st.warning("Tutar algılanamadı. Cümlede tutarı söyleyin (ör. 'seksen beş', '85 TL') ya da soldan girin.")
            speak("Tutar algılanamadı. Lütfen tutarı söyleyin veya girin.")
        else:
            st.warning("Komut eksik. 'Bu kişinin kalanından 85 TL düş' gibi söyleyin.")
            speak("Komut eksik. Lütfen düş mü ekle mi olduğunu da söyleyin.")

# ================== BUTONLAR ==================
cA,cB,cC=st.columns([1,1,1])
with cA:
    if st.button("İşlem Yap", use_container_width=True):
        last_text = st.session_state.get("last_final_text", "")
        pref=None; pref_digits=None
        man = st.session_state.get("manuel_ref_input","")
        if isinstance(man, str) and man.strip():
            pref = re.sub(r"\D","",man)
            if not pref: pref=None
        if pref is None and st.session_state.get("selected_ref"):
            pref=str(st.session_state.selected_ref)
        if pref is None and last_text:
            pnum, pdig = extract_personref(last_text)
            if pnum: pref=str(pnum); pref_digits=pdig
        if pref is None and last_text:
            pbyname,_=find_personref_by_name(st.session_state.df,last_text)
            if pbyname: pref=str(pbyname)

        amt=None
        try:
            tval = st.session_state.get("tutar_input", 0.0)
            if tval and float(tval)>0: amt=float(tval)
        except: pass
        if amt is None:
            stick=get_sticky_amount()
            if stick: amt=float(stick)
        if amt is None and last_text:
            a=extract_amount(last_text, pref_digits)
            if a: amt=float(a)

        op=parse_op_from_text(last_text, fallback_ui_op=st.session_state.get("islem_radio","Kalan TL’den Düş"))

        if pref is None:
            st.warning("Kişi bulunamadı. Soldaki listeden seçin ya da ad/PersonRef içeren komut söyleyin.")
            speak("Kişi bulunamadı.")
        elif not amt or float(amt)<=0:
            st.warning("Tutar yok. Soldan girin veya komutta söyleyin (örn. 80 TL).")
            speak("Tutar algılanmadı.")
        elif not op:
            st.warning("İşlem türü anlaşılmadı. 'düş' veya 'ekle' deyin.")
            speak("İşlem türü anlaşılmadı.")
        else:
            islem_yap(str(pref), float(amt), op, do_rerun=True)

with cB:
    if st.session_state.unsaved_ops: st.info(f"Kaydedilmemiş işlem: {len(st.session_state.unsaved_ops)}")
    if st.button("Kaydet", type="primary", use_container_width=True):
        out=DEFAULT_EXCEL_PATH
        st.session_state.df.to_excel(out, index=False)
        st.cache_data.clear()
        st.session_state.history = st.session_state.get("history", [])
        st.session_state.history.extend(st.session_state.unsaved_ops)
        st.session_state.unsaved_ops=[]
        st.success("Veriler kaydedildi — diğer kullanıcılar da aynı şekilde görecek.")
        speak("Veriler kaydedildi ve geçmişe işlendi.")
        st.rerun()

with cC:
    if st.button("Komut Örnekleri", use_container_width=True):
        st.info("Örnek: 'Bu kişinin kalanından 85 TL düş' | 'Ayşegül Ünal’ın kalanına 5 TL ekle'")
        speak("Örnek komutlar ekranınızda.")

# ================== 🎧 CANLI YAZIM (JS) ==================
st.markdown("### 🎧 Canlı Yazım")
st.components.v1.html(f"""
<div style="border:1px dashed #bbb;padding:8px;border-radius:8px;background:#fbfbfb">
  <div><b>Canlı:</b> <span id="stt_live">{'Dinleniyor…' if st.session_state.listening else 'Kapalı'}</span></div>
  <div style="margin-top:6px"><b>Son:</b> <span id="stt_final">{st.session_state.get('last_final_text') or ''}</span></div>
</div>
<script>
(function(){{
  const PY_SHOULD = {str(st.session_state.listening).lower()};
  const SR   = window.SpeechRecognition || window.webkitSpeechRecognition;
  const live = document.getElementById('stt_live');
  const fin  = document.getElementById('stt_final');
  const setLive = (t)=>{{ if(live) live.textContent = t; localStorage.setItem('stt_live_last', t||''); }};
  const setFinal= (t)=>{{ if(fin)  fin.textContent  = t; localStorage.setItem('stt_final_last', t||''); }};

  try {{
    const lastL = localStorage.getItem('stt_live_last'); if(lastL && live) live.textContent=lastL;
    const lastF = localStorage.getItem('stt_final_last'); if(lastF && fin) fin.textContent=lastF;
  }} catch(_ ){{}}

  if (!SR) {{ setLive('Tarayıcıda Ses Tanıma yok'); return; }}

  function ensureHandlers(rec){{
    if (rec.__handlersAttached) return;
    rec.__handlersAttached = true;
    rec.onresult = (e) => {{
      let interim = '', finalTxt = '';
      for (let i=e.resultIndex; i<e.results.length; i++) {{
        const t = e.results[i][0].transcript;
        if (e.results[i].isFinal) finalTxt += t; else interim += t;
      }}
      if (interim && interim.trim()) {{ setLive(interim.trim()); }}
      if (finalTxt && finalTxt.trim()) {{
        setFinal(finalTxt.trim());
        try {{ if (window.__stt_rec) window.__stt_rec.stop(); }} catch(_ ){{}}
        const u = new URL(location.href);
        u.searchParams.set('voice', finalTxt.trim());
        setTimeout(()=>{{ location.assign(u.toString()); }}, 10);
      }}
    }};
    rec.onerror = () => {{
      if (!window.__stt_should_listen) return;
      setTimeout(()=>{{ try{{ rec.start(); }}catch(_ ){{}} }}, 200);
    }};
    rec.onstart = ()=>{{ window.__stt_running=true; setLive('Dinleniyor…'); }};
    rec.onend   = ()=>{{ window.__stt_running=false; if (window.__stt_should_listen) setTimeout(()=>{{ try{{ rec.start(); }}catch(_ ){{}} }}, 200); else setLive('Kapalı'); }};
  }}

  async function askMic(){{
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return true;
    try {{
      const s = await navigator.mediaDevices.getUserMedia({{audio:true}});
      try{{ s.getTracks().forEach(t=>t.stop()); }}catch(_ ){{}}
      return true;
    }} catch(_ ) {{ setLive('Mikrofon izni yok'); return false; }}
  }}

  window.sttStart = async function(){{
    window.__stt_should_listen = true;
    if (!window.__stt_rec) {{
      window.__stt_rec = new SR();
      window.__stt_rec.lang='tr-TR';
      window.__stt_rec.continuous=true;
      window.__stt_rec.interimResults=true;
      window.__stt_rec.maxAlternatives=1;
      ensureHandlers(window.__stt_rec);
    }} else {{
      ensureHandlers(window.__stt_rec);
    }}
    const ok = await askMic(); if (!ok) return;
    try {{ window.__stt_rec.start(); setLive('Dinleniyor…'); }} catch(_ ){{}}
  }}
  window.sttStop = function(){{ window.__stt_should_listen = false; try {{ window.__stt_rec && window.__stt_rec.stop(); }} catch(_ ){{}}; setLive('Kapalı'); }}
  if (PY_SHOULD) window.sttStart(); else window.sttStop();
}})();
</script>
""", height=150)

# ================== BAŞLAT / DURDUR ==================
with st.sidebar:
    st.markdown("---"); st.subheader("🎧 Dinleme")
    colS, colT = st.columns(2)
    with colS:
        if st.button("🎤 Başlat", use_container_width=True, disabled=st.session_state.listening):
            st.session_state.force_listen = True
            st.session_state.listening = True
            st.components.v1.html("<script>try{window.sttStart && window.sttStart();}catch(e){}</script>", height=0)
            st.rerun()
    with colT:
        if st.button("⏹️ Durdur", use_container_width=True, disabled=not st.session_state.listening):
            st.session_state.force_listen = False
            st.session_state.listening = False
            st.components.v1.html("<script>try{window.sttStop && window.sttStop();}catch(e){}</script>", height=0)
            st.rerun()

# ================== SES PARAM → KOMUT ==================
voice_param = get_query_param("voice")
if voice_param:
    if st.session_state.force_listen:
        st.session_state.listening = True
    vtxt=unquote(voice_param).strip()
    if vtxt!=st.session_state._last_voice:
        st.session_state._last_voice=vtxt
        st.session_state.last_final_text=vtxt
        st.info(f"🎤 Algılanan komut: **{vtxt}**")
        speak("Komut alındı.")
        st.components.v1.html("<script>const u=new URL(location.href);u.searchParams.delete('voice');history.replaceState({},'',u);</script>", height=0)
        handle_command(vtxt, st.session_state.get("tutar_input",0.0), st.session_state.get("islem_radio","Kalan TL’den Düş"), st.session_state.get("selected_ref"), st.session_state.get("auto_apply", True))

# ================== 🧾 GEÇMİŞ ==================
st.markdown("## 🧾 İşlem Geçmişi")
if not st.session_state.get("history"):
    st.info("Henüz geçmiş kaydı yok. İşlem yap → Kaydet’e bas.")
else:
    hd=pd.DataFrame(st.session_state.history)
    try:
        hd["Zaman_dt"]=pd.to_datetime(hd["Zaman"]); hd=hd.sort_values("Zaman_dt",ascending=False).drop(columns=["Zaman_dt"])
    except: pass
    st.dataframe(hd, use_container_width=True, height=280)
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w: hd.to_excel(w,index=False,sheet_name="Islem_Gecmisi")
    st.download_button("⬇️ İşlem Geçmişini İndir (Excel)", data=buf.getvalue(),
        file_name=f"Islem_Gecmisi_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ================== 📊 Seçili Yönetici Oran Grafiği ==================
st.markdown("## 📊 Seçili Yönetici Oran Grafiği")
oran_full = ["KULLANILAN BÜTÇE ORANI", "SİSTEM DIŞI İLE KULLANILAN ORAN"]
oran_vals = [kullanilan_oran*100.0, sistem_disi_kullanilan_oran*100.0]
wrapped = ["\n".join(textwrap.wrap(x, width=22)) for x in oran_full]

fig, ax = plt.subplots(figsize=(7.2,4.2))
ax.bar(wrapped, oran_vals)
ax.set_ylabel("%")
ax.set_title(f"{selected_manager if opts and selected_manager!='(yok)' else 'Tümü'} — Oranlar")
ax.set_ylim(bottom=min(0, min(oran_vals)*1.15), top=max(0.1, max(oran_vals)*1.15))
for i,v in enumerate(oran_vals):
    ax.text(i, v if v>=0 else 0, f"{v:.2f}%", ha='center', va='bottom' if v>=0 else 'top', fontsize=9)
plt.tight_layout()
st.pyplot(fig, use_container_width=True)

# PNG indir
buf_png = io.BytesIO()
fig.savefig(buf_png, format="png", dpi=150, bbox_inches="tight")
st.download_button("⬇️ Oran Grafiği (PNG)", data=buf_png.getvalue(), file_name="yonetici_oran_grafik.png")

# Excel indir (oranlar)
oran_df = pd.DataFrame({
    "Yönetici": [selected_manager if opts and selected_manager!="(yok)" else "Tümü"]*len(oran_full),
    "Oran": oran_full,
    "Değer (%)": oran_vals
})
buf_xlsx = io.BytesIO()
with pd.ExcelWriter(buf_xlsx, engine="openpyxl") as w:
    oran_df.to_excel(w, index=False, sheet_name="Yonetici_Oranlari")
st.download_button("⬇️ Oranlar (Excel)", data=buf_xlsx.getvalue(),
    file_name=f"Oranlar_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ================== ⬇️ GÜNCEL VERİ İNDİR ==================
st.markdown("## ⬇️ Güncel Veriyi İndir (Excel)")
only = st.checkbox("Sadece seçili yönetici filtresi", value=False)
export = df_filtered.copy() if only else st.session_state.df.copy()
buf2=io.BytesIO()
with pd.ExcelWriter(buf2, engine="openpyxl") as w: export.to_excel(w,index=False,sheet_name="Veri")
st.download_button("⬇️ Veriyi İndir", data=buf2.getvalue(),
    file_name=f"Veri_{'filtreli_' if only else ''}{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
