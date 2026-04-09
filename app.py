import streamlit as st
import json, os, uuid
from datetime import datetime, timedelta
from io import BytesIO
import openpyxl

# ─── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="StreamVault — Plataformas Premium",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DB_FILE = "database.json"
ADMIN_PASSWORD = "admin2024"
YAPE_NUMERO = "930 734 823"
YAPE_TITULAR = "Saul"

# ─── DATABASE ──────────────────────────────────────────────────────────────────
def load_db():
    if not os.path.exists(DB_FILE):
        db = default_db()
        save_db(db)
        return db
    with open(DB_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_db(db):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

def default_db():
    return {
        "platforms": [
            {"id":"netflix","name":"Netflix","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/0/08/Netflix_2015_logo.svg","description":"Películas, series y documentales en 4K. Perfil compartido con PIN exclusivo.","delivery":"perfil","stock":5,"prices":{"30":15,"60":25,"90":35},"active":True},
            {"id":"prime","name":"Prime Video","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/1/11/Amazon_Prime_Video_logo.svg","description":"Series y películas Amazon Originals. Acceso por perfil individual.","delivery":"perfil","stock":8,"prices":{"30":10,"60":18,"90":25},"active":True},
            {"id":"hbo","name":"HBO Max","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/1/17/HBO_Max_Logo.svg","description":"Series HBO, DC y Warner Bros. Perfil con PIN protegido.","delivery":"perfil","stock":6,"prices":{"30":12,"60":20,"90":30},"active":True},
            {"id":"disney","name":"Disney+","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/3/3e/Disney%2B_logo.svg","description":"Marvel, Star Wars, Pixar y National Geographic. Perfil familiar.","delivery":"perfil","stock":7,"prices":{"30":10,"60":18,"90":26},"active":True},
            {"id":"disneypremium","name":"Disney Premium","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/3/3e/Disney%2B_logo.svg","description":"Cuenta exclusiva Disney+ sin compartir.","delivery":"cuenta","stock":3,"prices":{"30":20,"60":35,"90":50},"active":True},
            {"id":"vix","name":"VIX Premium","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/thumb/b/bc/VIX_logo.svg/1200px-VIX_logo.svg.png","description":"Cine y series en español. Deportes y telenovelas premium.","delivery":"perfil","stock":4,"prices":{"30":8,"60":14,"90":20},"active":True},
            {"id":"crunchyroll","name":"Crunchyroll","category":"streaming","image":"https://upload.wikimedia.org/wikipedia/commons/0/08/Crunchyroll_Logo.svg","description":"Anime en simulcast, doblado y subtitulado. La mayor biblioteca anime.","delivery":"perfil","stock":6,"prices":{"30":8,"60":14,"90":20},"active":True},
            {"id":"spotify","name":"Spotify Premium","category":"musica","image":"https://upload.wikimedia.org/wikipedia/commons/2/26/Spotify_logo_with_text.svg","description":"Música sin anuncios, descargas offline. Activación a tu correo.","delivery":"correo","stock":10,"prices":{"30":8,"60":14,"90":20},"active":True},
            {"id":"youtube","name":"YouTube Premium","category":"musica","image":"https://upload.wikimedia.org/wikipedia/commons/b/b8/YouTube_Logo_2017.svg","description":"Sin anuncios en YouTube y YouTube Music. Activación por correo.","delivery":"correo","stock":8,"prices":{"30":10,"60":18,"90":25},"active":True},
            {"id":"tidal","name":"Tidal HiFi","category":"musica","image":"https://upload.wikimedia.org/wikipedia/commons/thumb/8/84/Tidal_logo.svg/1200px-Tidal_logo.svg.png","description":"Audio en máxima calidad HiFi y MQA. Exclusivas de artistas.","delivery":"correo","stock":5,"prices":{"30":12,"60":20,"90":28},"active":True},
            {"id":"deezer","name":"Deezer Premium","category":"musica","image":"https://upload.wikimedia.org/wikipedia/commons/e/e3/Deezer_logo.svg","description":"90 millones de canciones en alta calidad. Activación por correo.","delivery":"correo","stock":6,"prices":{"30":8,"60":14,"90":20},"active":True},
            {"id":"canva","name":"Canva Pro","category":"musica","image":"https://upload.wikimedia.org/wikipedia/commons/b/bb/Canva_Logo.svg","description":"Diseño profesional con plantillas premium. Invitación a tu correo. Válido 1 año.","delivery":"correo_manual","stock":15,"prices":{"365":25},"active":True},
            {"id":"gemini","name":"Gemini AI","category":"ia","image":"https://www.gstatic.com/lamda/images/gemini_sparkle_v002_d4735304ff6292a690345.svg","description":"IA de Google con capacidades multimodales avanzadas. Por dispositivo/correo.","delivery":"dispositivo","stock":5,"prices":{"30":15,"60":25,"90":35},"active":True},
            {"id":"chatgpt","name":"ChatGPT Plus","category":"ia","image":"https://upload.wikimedia.org/wikipedia/commons/0/04/ChatGPT_logo.svg","description":"GPT-4o con plugins, DALL-E y análisis avanzado. Por correo personal.","delivery":"dispositivo","stock":4,"prices":{"30":18,"60":30,"90":42},"active":True},
            {"id":"perplexity","name":"Perplexity Pro","category":"ia","image":"https://upload.wikimedia.org/wikipedia/commons/thumb/1/1d/Perplexity_AI_logo.svg/1200px-Perplexity_AI_logo.svg.png","description":"Búsqueda con IA en tiempo real, sin límites.","delivery":"dispositivo","stock":6,"prices":{"30":12,"60":20,"90":28},"active":True},
            {"id":"claude","name":"Claude AI Pro","category":"ia","image":"https://upload.wikimedia.org/wikipedia/commons/thumb/8/8a/Claude_AI_logo.svg/1200px-Claude_AI_logo.svg.png","description":"Claude Pro de Anthropic con uso prioritario y todos los modelos.","delivery":"dispositivo","stock":5,"prices":{"30":15,"60":25,"90":35},"active":True}
        ],
        "accounts": [],
        "orders": [],
        "clients": [],
        "finances": {"ingresos": [], "egresos": []},
        "payment_codes": {}
    }

# ─── SESSION STATE ──────────────────────────────────────────────────────────────
defaults = {"page":"store","admin_logged":False,"buy_platform":None,"buy_days":None,"order_result":None,"cat_filter":"todos","admin_nav":"📊 Dashboard","show_add_acc":False}
for k,v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Sans:wght@300;400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif!important}
.stApp{background:#08090d!important}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding:0!important;max-width:100%!important}

.sv-nav{background:rgba(8,9,13,0.97);border-bottom:1px solid rgba(255,255,255,0.07);padding:0 5%;display:flex;align-items:center;justify-content:space-between;height:58px}
.sv-logo{font-family:'Syne',sans-serif;font-weight:800;font-size:1.3rem;background:linear-gradient(135deg,#6c3fff,#00e5b4);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.sv-hero{text-align:center;padding:60px 5% 44px;background:radial-gradient(ellipse 70% 50% at 50% 0%,rgba(108,63,255,0.1) 0,transparent 70%)}
.sv-badge{display:inline-flex;align-items:center;gap:8px;background:rgba(108,63,255,0.1);border:1px solid rgba(108,63,255,0.3);border-radius:100px;padding:5px 16px;font-size:.78rem;color:#a78bfa;margin-bottom:20px}
.sv-h1{font-family:'Syne',sans-serif;font-size:clamp(1.9rem,4vw,3.2rem);font-weight:800;line-height:1.1;color:#f0f0f8;margin-bottom:14px}
.sv-h1 em{font-style:normal;background:linear-gradient(135deg,#6c3fff,#00e5b4);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.sv-sub{color:#6b7280;font-size:.95rem;max-width:500px;margin:0 auto 32px}
.sv-stats{display:flex;gap:36px;justify-content:center;flex-wrap:wrap;margin-bottom:8px}
.sv-stat-num{font-family:'Syne',sans-serif;font-size:1.7rem;font-weight:800;color:#00e5b4}
.sv-stat-label{font-size:.72rem;color:#6b7280}
.sv-section-title{font-family:'Syne',sans-serif;font-size:.9rem;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:.1em;margin:32px 0 12px}
.sv-section-title span{color:#00e5b4;margin-left:8px;font-size:.78rem}
.sv-card{background:#0f1117;border:1px solid rgba(255,255,255,0.07);border-radius:16px;padding:18px;position:relative;overflow:hidden;transition:.3s;margin-bottom:4px}
.sv-card:hover{border-color:rgba(108,63,255,0.35);box-shadow:0 14px 40px rgba(108,63,255,0.1)}
.sv-card-name{font-family:'Syne',sans-serif;font-weight:700;font-size:.97rem;color:#f0f0f8;margin-bottom:5px}
.sv-card-desc{font-size:.8rem;color:#6b7280;line-height:1.6;margin-bottom:10px}
.sv-ribbon{position:absolute;top:10px;right:10px;background:#00e5b4;color:#08090d;font-size:.6rem;font-weight:700;padding:2px 8px;border-radius:100px}
.sv-badge-perfil{background:rgba(108,63,255,0.12);color:#a78bfa;border:1px solid rgba(108,63,255,0.2);padding:2px 8px;border-radius:100px;font-size:.67rem;font-weight:600;display:inline-block;margin-bottom:8px}
.sv-badge-correo{background:rgba(0,229,180,0.08);color:#00e5b4;border:1px solid rgba(0,229,180,0.15);padding:2px 8px;border-radius:100px;font-size:.67rem;font-weight:600;display:inline-block;margin-bottom:8px}
.sv-badge-dispositivo{background:rgba(255,165,0,0.08);color:#ffa500;border:1px solid rgba(255,165,0,0.15);padding:2px 8px;border-radius:100px;font-size:.67rem;font-weight:600;display:inline-block;margin-bottom:8px}
.sv-price-box{background:#161921;border-radius:12px;padding:14px 18px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}
.sv-price-val{font-family:'Syne',sans-serif;font-weight:800;font-size:1.25rem;color:#00e5b4}
.sv-qr-box{background:#161921;border-radius:14px;padding:20px;text-align:center;margin:16px 0}
.sv-yape-info{background:rgba(108,63,255,0.06);border:1px solid rgba(108,63,255,0.2);border-radius:10px;padding:12px 16px;font-size:.83rem;margin-bottom:12px;color:#c4b5fd}
.sv-cred-box{background:#161921;border:1px dashed rgba(108,63,255,0.4);border-radius:12px;padding:16px;margin:16px 0}
.sv-cred-row{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.05);font-size:.84rem}
.sv-cred-key{color:#6b7280}
.sv-cred-val{font-family:monospace;color:#00e5b4;font-weight:600}
.adm-stat{background:#0e0f14;border:1px solid rgba(255,255,255,0.07);border-radius:14px;padding:18px;text-align:center;margin-bottom:12px}
.adm-stat-label{font-size:.72rem;color:#6b7280;margin-bottom:5px;text-transform:uppercase;letter-spacing:.06em}
.adm-stat-value{font-family:'Syne',sans-serif;font-size:1.6rem;font-weight:800}
.stButton>button{background:linear-gradient(135deg,#6c3fff,#8b5cf6)!important;color:white!important;border:none!important;border-radius:10px!important;font-family:'Syne',sans-serif!important;font-weight:700!important}
.stButton>button:hover{opacity:.9!important}
.stTextInput>div>div>input,.stNumberInput>div>div>input,.stTextArea>div>div>textarea,.stSelectbox>div>div{background:#161921!important;border:1px solid rgba(255,255,255,0.1)!important;color:#f0f0f8!important;border-radius:10px!important}
div[data-testid="metric-container"]{background:#0e0f14!important;border:1px solid rgba(255,255,255,0.07)!important;border-radius:14px!important;padding:16px!important}
div[data-testid="metric-container"] label{color:#6b7280!important}
.stDataFrame{background:#0f1117!important;border-radius:12px!important}
label{color:#9ca3af!important;font-size:.82rem!important}
.stSuccess{background:rgba(0,229,180,0.08)!important;border:1px solid rgba(0,229,180,0.3)!important;border-radius:10px!important}
.stError{background:rgba(255,71,87,0.08)!important;border:1px solid rgba(255,71,87,0.3)!important;border-radius:10px!important}
.stInfo{background:rgba(108,63,255,0.06)!important;border:1px solid rgba(108,63,255,0.2)!important;border-radius:10px!important}
.stWarning{background:rgba(245,158,11,0.06)!important;border:1px solid rgba(245,158,11,0.2)!important;border-radius:10px!important}
hr{border-color:rgba(255,255,255,0.07)!important}
</style>
""", unsafe_allow_html=True)

# ─── HELPERS ───────────────────────────────────────────────────────────────────
DELIVERY_LABELS = {"perfil":"👤 Por Perfil","correo":"📧 Por Correo","dispositivo":"💻 Por Dispositivo","cuenta":"🔑 Cuenta Completa","correo_manual":"📨 Invitación"}

def add_or_update_client(db, name, whatsapp, email, order):
    existing = next((c for c in db["clients"] if c.get("whatsapp") == whatsapp), None)
    if existing:
        if order["id"] not in existing["orders"]:
            existing["orders"].append(order["id"])
        existing["last_order"] = datetime.now().isoformat()
    else:
        db["clients"].append({"id":str(uuid.uuid4())[:8],"name":name,"whatsapp":whatsapp,"email":email,"orders":[order["id"]],"created_at":datetime.now().isoformat(),"last_order":datetime.now().isoformat()})

def process_payment(db, platform, days_key, name, whatsapp, email, method, code):
    days = int(days_key)
    if platform.get("stock", 0) <= 0:
        return False, "❌ Sin stock disponible en este momento."
    if method == "yape":
        code_str = str(code).strip()
        if not code_str or len(code_str) != 3 or not code_str.isdigit():
            return False, "❌ El código Yape debe ser exactamente 3 dígitos numéricos."
        if code_str in db.get("payment_codes", {}):
            return False, "❌ Este código Yape ya fue utilizado anteriormente."
        db["payment_codes"][code_str] = {"used_at": datetime.now().isoformat(), "platform": platform["id"]}
    if platform["delivery"] == "correo_manual":
        order_id = str(uuid.uuid4())[:8].upper()
        order = {"id":order_id,"platform_id":platform["id"],"platform_name":platform["name"],"client_name":name,"client_whatsapp":whatsapp,"client_email":email,"days":days_key,"price":platform["prices"].get(days_key,0),"payment_method":method,"payment_code":str(code),"status":"pending_manual","created_at":datetime.now().isoformat(),"expires_at":(datetime.now()+timedelta(days=days)).isoformat()}
        db["orders"].append(order)
        add_or_update_client(db, name, whatsapp, email, order)
        save_db(db)
        return True, {"manual":True,"order_id":order_id,"email":email}
    available = [a for a in db["accounts"] if a["platform_id"] == platform["id"] and a["status"] == "available"]
    if not available:
        return False, "❌ Sin cuentas disponibles. Contáctanos por WhatsApp: " + YAPE_NUMERO
    account = available[0]
    account["status"] = "assigned"
    account["assigned_to"] = name
    account["assigned_email"] = email
    account["assigned_at"] = datetime.now().isoformat()
    account["expires_at"] = (datetime.now() + timedelta(days=days)).isoformat()
    for p in db["platforms"]:
        if p["id"] == platform["id"]:
            p["stock"] = max(0, p.get("stock",1) - 1)
    order_id = str(uuid.uuid4())[:8].upper()
    order = {"id":order_id,"platform_id":platform["id"],"platform_name":platform["name"],"client_name":name,"client_whatsapp":whatsapp,"client_email":email,"days":days_key,"price":platform["prices"].get(days_key,0),"payment_method":method,"payment_code":str(code),"account_id":account["id"],"status":"completed","created_at":datetime.now().isoformat(),"expires_at":account["expires_at"]}
    db["orders"].append(order)
    db["finances"]["ingresos"].append({"date":datetime.now().isoformat(),"amount":float(platform["prices"].get(days_key,0)),"description":f"{platform['name']} - {name}","order_id":order_id})
    add_or_update_client(db, name, whatsapp, email, order)
    save_db(db)
    return True, {"manual":False,"order_id":order_id,"account":account,"expires_at":account["expires_at"]}

def export_excel(db):
    wb = openpyxl.Workbook()
    ws1 = wb.active; ws1.title = "Clientes"
    ws1.append(["ID","Nombre","WhatsApp","Email","Órdenes","Registro","Último Pedido"])
    for c in db["clients"]:
        ws1.append([c.get("id"),c.get("name"),c.get("whatsapp"),c.get("email"),len(c.get("orders",[])),c.get("created_at","")[:10],c.get("last_order","")[:10]])
    ws2 = wb.create_sheet("Órdenes")
    ws2.append(["Orden","Plataforma","Cliente","WhatsApp","Email","Días","Precio","Método","Estado","Vencimiento"])
    for o in db["orders"]:
        ws2.append([o.get("id"),o.get("platform_name"),o.get("client_name"),o.get("client_whatsapp"),o.get("client_email"),o.get("days"),o.get("price"),o.get("payment_method"),o.get("status"),o.get("expires_at","")[:10] if o.get("expires_at") else ""])
    ws3 = wb.create_sheet("Cuentas")
    ws3.append(["ID","Plataforma","Email","Contraseña","Perfil","PIN","Estado","Asignado a","Vence"])
    for a in db["accounts"]:
        ws3.append([a.get("id"),a.get("platform_id"),a.get("email"),a.get("password"),a.get("profile_name",""),a.get("profile_pin",""),a.get("status"),a.get("assigned_to",""),a.get("expires_at","")[:10] if a.get("expires_at") else ""])
    ws4 = wb.create_sheet("Finanzas")
    ws4.append(["Tipo","Fecha","Descripción","Monto"])
    for i in db["finances"]["ingresos"]:
        ws4.append(["Ingreso",i.get("date","")[:10],i.get("description"),i.get("amount")])
    for e in db["finances"]["egresos"]:
        ws4.append(["Egreso",e.get("date","")[:10],e.get("description"),e.get("amount")])
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ══════════════════════════════════════════════════════════════════════════════
# STORE
# ══════════════════════════════════════════════════════════════════════════════
def page_store():
    db = load_db()
    platforms = [p for p in db["platforms"] if p.get("active")]

    st.markdown(f"""
    <div class="sv-nav">
      <div class="sv-logo">🎬 StreamVault</div>
      <div style="display:flex;gap:6px;font-size:.83rem">
        <span style="color:#6b7280;padding:6px 10px;">📺 Streaming</span>
        <span style="color:#6b7280;padding:6px 10px;">🎵 Música</span>
        <span style="color:#6b7280;padding:6px 10px;">🤖 IA</span>
        <a href="https://wa.me/51930734823" target="_blank" style="color:#00e5b4;text-decoration:none;padding:6px 12px;background:rgba(0,229,180,0.08);border-radius:8px;">💬 Soporte</a>
      </div>
    </div>""", unsafe_allow_html=True)

    if st.session_state.buy_platform:
        page_checkout(db)
        return

    st.markdown("""
    <div class="sv-hero">
      <div class="sv-badge">🔥 Precios desde S/ 8 · Entrega automática</div>
      <div class="sv-h1">Plataformas Premium<br><em>al mejor precio</em></div>
      <div class="sv-sub">Streaming, música e inteligencia artificial. Paga con Yape o Binance y recibe acceso al instante.</div>
      <div class="sv-stats">
        <div><div class="sv-stat-num">16+</div><div class="sv-stat-label">Plataformas</div></div>
        <div><div class="sv-stat-num">24/7</div><div class="sv-stat-label">Entrega automática</div></div>
        <div><div class="sv-stat-num">100%</div><div class="sv-stat-label">Garantizado</div></div>
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("<div style='padding:0 3%'>", unsafe_allow_html=True)
    cols_f = st.columns(4)
    filters = [("todos","🌐 Todos"),("streaming","📺 Streaming"),("musica","🎵 Música & Apps"),("ia","🤖 Inteligencia Artificial")]
    for i,(key,label) in enumerate(filters):
        with cols_f[i]:
            if st.button(label, key=f"filt_{key}", use_container_width=True):
                st.session_state.cat_filter = key
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='padding:0 3%'>", unsafe_allow_html=True)
    cat_names = {"streaming":"📺 Streaming","musica":"🎵 Música & Apps","ia":"🤖 Inteligencia Artificial"}
    for cat in ["streaming","musica","ia"]:
        if st.session_state.cat_filter not in ("todos", cat):
            continue
        cat_platforms = [p for p in platforms if p["category"] == cat]
        if not cat_platforms:
            continue
        st.markdown(f'<div class="sv-section-title">{cat_names[cat]} <span>{len(cat_platforms)} plataformas</span></div>', unsafe_allow_html=True)
        n_cols = min(len(cat_platforms), 4)
        cols = st.columns(n_cols, gap="small")
        for i, p in enumerate(cat_platforms):
            with cols[i % n_cols]:
                render_platform_card(p)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("""
    <div style='text-align:center;padding:36px 5%;background:#0f1117;border-top:1px solid rgba(255,255,255,0.07);margin-top:40px;'>
      <p style='color:#6b7280;font-size:.84rem;'>StreamVault © 2024 · Pagos seguros por Yape y Binance ·
      <a href="https://wa.me/51930734823" style="color:#00e5b4;text-decoration:none;">WhatsApp: 930 734 823</a></p>
    </div>
    <a href="https://wa.me/51930734823" target="_blank"
       style="position:fixed;bottom:22px;right:22px;z-index:999;width:52px;height:52px;border-radius:50%;
       background:#25d366;display:flex;align-items:center;justify-content:center;
       box-shadow:0 4px 18px rgba(37,211,102,0.4);text-decoration:none;font-size:1.4rem;">💬</a>""",
    unsafe_allow_html=True)

def render_platform_card(p):
    is_canva = p["id"] == "canva"
    prices = p.get("prices", {})
    stock = p.get("stock", 0)
    delivery = p.get("delivery", "perfil")
    badge_class = {"correo":"sv-badge-correo","correo_manual":"sv-badge-correo","dispositivo":"sv-badge-dispositivo"}.get(delivery,"sv-badge-perfil")
    stock_color = "#00e5b4" if stock > 5 else ("#f59e0b" if stock > 0 else "#6b7280")
    stock_text = f"✓ {stock} disponibles" if stock > 5 else (f"⚡ Solo {stock}" if stock > 0 else "✗ Sin stock")
    ribbon = '<div class="sv-ribbon">¡Últimas!</div>' if 0 < stock <= 3 else ""

    st.markdown(f"""
    <div class="sv-card">
      {ribbon}
      <div style="display:flex;align-items:center;gap:9px;margin-bottom:8px;">
        <img src="{p['image']}" style="width:34px;height:34px;object-fit:contain;background:#fff;border-radius:7px;padding:3px"
             onerror="this.src='https://placehold.co/34x34/6c3fff/ffffff?text={p['name'][0]}'">
        <div class="sv-card-name">{p['name']}</div>
      </div>
      <span class="{badge_class}">{DELIVERY_LABELS.get(delivery, delivery)}</span>
      <div class="sv-card-desc">{p['description']}</div>
      <div style="font-size:.74rem;color:{stock_color};margin-bottom:6px;">{stock_text}</div>
    </div>""", unsafe_allow_html=True)

    price_keys = list(prices.keys())
    if price_keys:
        labels = [f"{'1 Año' if is_canva else k+' días'} — S/ {prices[k]}" for k in price_keys]
        sel_label = st.selectbox("Plan", labels, key=f"plan_{p['id']}", label_visibility="collapsed")
        sel_days = price_keys[labels.index(sel_label)]
        price_val = prices[sel_days]
        if stock > 0:
            if st.button(f"🛒 Comprar · S/ {price_val}", key=f"buy_{p['id']}", use_container_width=True):
                st.session_state.buy_platform = p
                st.session_state.buy_days = sel_days
                st.session_state.order_result = None
                st.rerun()
        else:
            st.markdown('<div style="text-align:center;padding:9px;background:#161921;border-radius:9px;color:#6b7280;font-size:.83rem;">Sin Stock · WhatsApp para notificación</div>', unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────────────────
def page_checkout(db):
    p = st.session_state.buy_platform
    days_key = st.session_state.buy_days
    price = p["prices"].get(days_key, 0)
    is_canva = p["id"] == "canva"
    days_label = "1 Año" if is_canva else f"{days_key} días"

    if st.button("← Volver a la tienda", key="back_btn"):
        st.session_state.buy_platform = None
        st.session_state.buy_days = None
        st.session_state.order_result = None
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Show result
    if st.session_state.order_result:
        r = st.session_state.order_result
        if r.get("success"):
            d = r["data"]
            if d.get("manual"):
                st.success(f"🎉 **¡Pedido registrado! Orden #{d['order_id']}**")
                st.info(f"📨 Recibirás la invitación de **Canva Pro** en `{d['email']}` en pocas horas. ¿Dudas? Escríbenos al WhatsApp.")
            else:
                st.success(f"🎉 **¡Acceso activado! Orden #{d['order_id']}**")
                acc = d.get("account", {})
                creds = []
                if acc.get("email"): creds.append(("📧 Email / Usuario", acc["email"]))
                if acc.get("password"): creds.append(("🔑 Contraseña", acc["password"]))
                if acc.get("profile_name"): creds.append(("👤 Perfil", acc["profile_name"]))
                if acc.get("profile_pin"): creds.append(("🔐 PIN Perfil", acc["profile_pin"]))
                if acc.get("extra"): creds.append(("ℹ️ Extra", acc["extra"]))
                if d.get("expires_at"): creds.append(("📅 Vence", d["expires_at"][:10]))
                st.markdown('<div class="sv-cred-box">', unsafe_allow_html=True)
                for k,v in creds:
                    st.markdown(f'<div class="sv-cred-row"><span class="sv-cred-key">{k}</span><span class="sv-cred-val">{v}</span></div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
                st.info("💡 Copia y guarda tus credenciales en un lugar seguro.")
                for k,v in creds:
                    st.code(f"{k}: {v}")
            st.markdown(f'<a href="https://wa.me/51930734823" target="_blank" style="color:#00e5b4;font-size:.85rem;">💬 Soporte WhatsApp: {YAPE_NUMERO}</a>', unsafe_allow_html=True)
        else:
            st.error(r.get("message","Error al procesar el pago."))
            if st.button("🔄 Intentar nuevamente"):
                st.session_state.order_result = None
                st.rerun()
        return

    col_form, col_info = st.columns([3,2], gap="large")

    with col_form:
        st.markdown(f"""
        <div style="font-family:'Syne',sans-serif;font-weight:800;font-size:1.2rem;color:#f0f0f8;margin-bottom:16px;">
          🛒 Checkout — {p['name']}
        </div>
        <div class="sv-price-box">
          <div>
            <div style="font-size:.82rem;color:#6b7280;">{p['name']} · {days_label}</div>
            <div style="font-size:.74rem;color:#6b7280;">{DELIVERY_LABELS.get(p['delivery'],'')}</div>
          </div>
          <div class="sv-price-val">S/ {price}</div>
        </div>""", unsafe_allow_html=True)

        st.markdown("**📋 Tus datos**")
        name = st.text_input("Nombre completo *", placeholder="Ej: Juan Pérez", key="co_name")
        whatsapp = st.text_input("WhatsApp *", placeholder="Ej: 51930734823 (con código de país)", key="co_wa")
        email = ""
        if p["delivery"] in ("correo","dispositivo","correo_manual"):
            lbl = "Tu correo (recibirás la invitación Canva Pro) *" if is_canva else "Tu correo para activar el servicio *"
            email = st.text_input(lbl, placeholder="tucorreo@gmail.com", key="co_email")

        st.markdown("<br>**💳 Método de pago**", unsafe_allow_html=True)
        method_choice = st.radio("", ["💜 Yape","🟡 Binance Pay"], horizontal=True, key="co_method", label_visibility="collapsed")
        method_key = "yape" if "Yape" in method_choice else "binance"

        code = ""
        uploaded = None
        if method_key == "yape":
            st.markdown(f"""
            <div class="sv-yape-info">
              <strong style="color:#00e5b4;">📱 Pagar con Yape</strong><br>
              Número: <strong>{YAPE_NUMERO}</strong> &nbsp;·&nbsp; Titular: <strong>{YAPE_TITULAR}</strong><br>
              <span style="color:#f59e0b;font-size:.78rem;">⚠️ Envía exactamente: <strong>S/ {price}</strong></span>
            </div>""", unsafe_allow_html=True)
            code = st.text_input("Código de operación Yape (3 dígitos) *", placeholder="847", max_chars=3, key="co_ycode")
            st.caption("📍 Encuéntralo en el comprobante dentro de tu app Yape")
            uploaded = st.file_uploader("Captura de pantalla del pago *", type=["png","jpg","jpeg","webp"], key="co_cap")
            if uploaded:
                st.image(uploaded, width=170, caption="✅ Captura recibida")
        else:
            st.markdown(f"""
            <div class="sv-yape-info" style="border-color:rgba(240,185,11,0.3);background:rgba(240,185,11,0.04);color:#fcd34d;">
              <strong style="color:#f0b90b;">🟡 Pagar con Binance Pay</strong><br>
              Escanea el QR en la pestaña de la derecha y paga <strong>S/ {price}</strong> en USDT.<br>
              <span style="font-size:.78rem;color:#6b7280;">Luego copia el ID de transacción aquí.</span>
            </div>""", unsafe_allow_html=True)
            code = st.text_input("ID de transacción Binance *", placeholder="Ej: TX-123456789", key="co_bcode")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button(f"✅ Confirmar Pago y Obtener Acceso — S/ {price}", key="co_confirm", use_container_width=True):
            errors = []
            if not name.strip(): errors.append("Ingresa tu nombre completo.")
            if not whatsapp.strip(): errors.append("Ingresa tu número de WhatsApp.")
            if p["delivery"] in ("correo","dispositivo","correo_manual") and not email.strip():
                errors.append("Ingresa tu correo electrónico.")
            if method_key == "yape":
                cs = str(code).strip()
                if not cs or len(cs)!=3 or not cs.isdigit():
                    errors.append("El código Yape debe ser exactamente 3 dígitos numéricos.")
                if not uploaded:
                    errors.append("Adjunta la captura de pantalla del pago Yape.")
            else:
                if not str(code).strip():
                    errors.append("Ingresa el ID de transacción de Binance.")
            if errors:
                for e in errors:
                    st.error(f"⚠️ {e}")
            else:
                db2 = load_db()
                ok, data = process_payment(db2, p, days_key, name.strip(), whatsapp.strip(), email.strip(), method_key, str(code).strip())
                st.session_state.order_result = {"success":ok,"data":data} if ok else {"success":False,"message":data}
                st.rerun()

    with col_info:
        st.markdown("**📋 Resumen**")
        st.markdown(f"""
        <div style="background:#0f1117;border:1px solid rgba(255,255,255,0.07);border-radius:14px;padding:20px;">
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
            <img src="{p['image']}" style="width:38px;height:38px;object-fit:contain;background:#fff;border-radius:8px;padding:3px"
                 onerror="this.src='https://placehold.co/38/6c3fff/fff?text={p['name'][0]}'">
            <div>
              <div style="font-family:'Syne',sans-serif;font-weight:700;color:#f0f0f8;">{p['name']}</div>
              <div style="font-size:.74rem;color:#6b7280;">{days_label}</div>
            </div>
          </div>
          <div style="border-top:1px solid rgba(255,255,255,0.07);padding-top:12px;">
            <div style="display:flex;justify-content:space-between;font-size:.83rem;color:#6b7280;margin-bottom:5px;"><span>Subtotal</span><span style="color:#f0f0f8;">S/ {price}</span></div>
            <div style="display:flex;justify-content:space-between;font-size:.83rem;color:#6b7280;margin-bottom:5px;"><span>Entrega</span><span style="color:#00e5b4;">Automática ⚡</span></div>
            <div style="display:flex;justify-content:space-between;font-family:'Syne',sans-serif;font-weight:800;font-size:1.1rem;margin-top:10px;border-top:1px solid rgba(255,255,255,0.07);padding-top:10px;"><span>Total</span><span style="color:#00e5b4;">S/ {price}</span></div>
          </div>
        </div>""", unsafe_allow_html=True)

        st.markdown("<br>**🔲 QR de Pago**")
        t1,t2 = st.tabs(["💜 Yape","🟡 Binance"])
        with t1:
            try:
                st.image("static/yape_qr.png", use_container_width=True, caption=f"Yape · {YAPE_NUMERO} · {YAPE_TITULAR}")
            except:
                st.markdown(f"""<div class="sv-qr-box">
                  <div style="font-size:2.5rem;">💜</div>
                  <div style="font-family:'Syne',sans-serif;font-weight:700;font-size:1.1rem;color:#a78bfa;margin:8px 0;">{YAPE_NUMERO}</div>
                  <div style="color:#6b7280;font-size:.8rem;">Titular: {YAPE_TITULAR}</div>
                  <div style="color:#6b7280;font-size:.74rem;margin-top:6px;">⚠️ Agrega tu QR en static/yape_qr.png</div>
                </div>""", unsafe_allow_html=True)
        with t2:
            try:
                st.image("static/binance_qr.png", use_container_width=True, caption="Escanea con la app de Binance para pagar")
            except:
                st.markdown("""<div class="sv-qr-box">
                  <div style="font-size:2.5rem;">🟡</div>
                  <div style="font-family:'Syne',sans-serif;font-weight:700;font-size:1.1rem;color:#f0b90b;margin:8px 0;">Binance Pay</div>
                  <div style="color:#6b7280;font-size:.74rem;margin-top:6px;">⚠️ Agrega tu QR en static/binance_qr.png</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("""
        <div style="margin-top:14px;background:rgba(0,229,180,0.05);border:1px solid rgba(0,229,180,0.15);border-radius:10px;padding:12px;font-size:.77rem;color:#6b7280;line-height:1.8;">
          ✅ Entrega automática al confirmar<br>
          🔐 Acceso con PIN de perfil exclusivo<br>
          💬 Soporte WhatsApp 24/7<br>
          📅 Renovación antes de vencer
        </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ADMIN
# ══════════════════════════════════════════════════════════════════════════════
def page_admin():
    if not st.session_state.admin_logged:
        st.markdown("<br><br>", unsafe_allow_html=True)
        _,col,_ = st.columns([2,2,2])
        with col:
            st.markdown("""
            <div style="text-align:center;background:#0f1117;border:1px solid rgba(255,255,255,0.07);border-radius:20px;padding:36px;">
              <div style="font-size:2.5rem;margin-bottom:12px;">🔐</div>
              <div style="font-family:'Syne',sans-serif;font-weight:800;font-size:1.5rem;color:#f0f0f8;margin-bottom:6px;">Admin Panel</div>
              <div style="color:#6b7280;font-size:.84rem;margin-bottom:20px;">StreamVault · Acceso exclusivo</div>
            </div>""", unsafe_allow_html=True)
            pwd = st.text_input("Contraseña", type="password", key="adm_pwd")
            if st.button("→ Entrar al Panel", key="adm_login_btn", use_container_width=True):
                if pwd == ADMIN_PASSWORD:
                    st.session_state.admin_logged = True
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta")
        return

    db = load_db()
    st.markdown("""
    <div style="background:#0e0f14;border-bottom:1px solid rgba(255,255,255,0.07);padding:14px 24px;display:flex;align-items:center;gap:16px;margin-bottom:4px;">
      <div style="font-family:'Syne',sans-serif;font-weight:800;font-size:1.1rem;background:linear-gradient(135deg,#6c3fff,#00e5b4);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">🎬 StreamVault Admin</div>
      <span style="background:rgba(0,229,180,0.1);border:1px solid rgba(0,229,180,0.2);color:#00e5b4;padding:3px 12px;border-radius:100px;font-size:.72rem;">● Conectado</span>
    </div>""", unsafe_allow_html=True)

    col_nav, col_content = st.columns([1,5], gap="small")
    with col_nav:
        st.markdown("<br>", unsafe_allow_html=True)
        nav_items = ["📊 Dashboard","📦 Órdenes","👥 Clientes","🔑 Cuentas","📺 Plataformas","💰 Finanzas","⏳ Manuales"]
        for item in nav_items:
            if st.button(item, key=f"nav_{item}", use_container_width=True):
                st.session_state.admin_nav = item
                st.rerun()
        st.markdown("---")
        if st.button("🏪 Ver Tienda", key="nav_store", use_container_width=True):
            st.session_state.page = "store"
            st.session_state.admin_logged = False
            st.rerun()
        if st.button("🚪 Salir", key="nav_logout", use_container_width=True):
            st.session_state.admin_logged = False
            st.rerun()

    with col_content:
        nav = st.session_state.admin_nav
        if "Dashboard" in nav:     admin_dashboard(db)
        elif "Órdenes" in nav:     admin_orders(db)
        elif "Clientes" in nav:    admin_clients(db)
        elif "Cuentas" in nav:     admin_accounts(db)
        elif "Plataformas" in nav: admin_platforms(db)
        elif "Finanzas" in nav:    admin_finances(db)
        elif "Manuales" in nav:    admin_pending(db)

def admin_dashboard(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">📊 Dashboard</div>', unsafe_allow_html=True)
    today = datetime.now().date().isoformat()
    total_in = sum(i["amount"] for i in db["finances"]["ingresos"])
    total_eg = sum(e["amount"] for e in db["finances"]["egresos"])
    today_orders = [o for o in db["orders"] if o["created_at"][:10]==today]
    expiring = [o for o in db["orders"] if o.get("expires_at") and o["status"]=="completed" and 0<=(datetime.fromisoformat(o["expires_at"])-datetime.now()).days<=5]
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("👥 Clientes",len(db["clients"]))
    c2.metric("📦 Órdenes",len(db["orders"]))
    c3.metric("📦 Hoy",len(today_orders))
    c4.metric("⚡ Por vencer",len(expiring))
    c5,c6,c7,c8 = st.columns(4)
    c5.metric("💚 Ingresos",f"S/ {total_in:.2f}")
    c6.metric("🔴 Egresos",f"S/ {total_eg:.2f}")
    c7.metric("🟡 Ganancia",f"S/ {(total_in-total_eg):.2f}")
    c8.metric("📺 Plataformas",len(db["platforms"]))
    if expiring:
        st.warning(f"⚠️ **{len(expiring)} cuenta(s)** vencen en los próximos 5 días.")
    st.markdown("<br>**⚡ Últimas órdenes**", unsafe_allow_html=True)
    orders = sorted(db["orders"],key=lambda x:x["created_at"],reverse=True)[:10]
    if orders:
        rows = [{"Orden":o["id"],"Plataforma":o["platform_name"],"Cliente":o["client_name"],"Precio":f"S/ {o['price']}","Estado":o["status"],"Fecha":o["created_at"][:10]} for o in orders]
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    else:
        st.info("Aún no hay órdenes.")

def admin_orders(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">📦 Órdenes</div>', unsafe_allow_html=True)
    search = st.text_input("🔍 Buscar por cliente, plataforma, orden...", key="ord_search")
    orders = sorted(db["orders"],key=lambda x:x["created_at"],reverse=True)
    if search:
        orders = [o for o in orders if search.lower() in json.dumps(o).lower()]
    col_exp, _ = st.columns([1,4])
    with col_exp:
        st.download_button("📥 Exportar Excel", export_excel(db), "ordenes.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if orders:
        rows = [{"#":o["id"],"Plataforma":o["platform_name"],"Cliente":o["client_name"],"WhatsApp":o["client_whatsapp"],"Email":o.get("client_email",""),"Días":o["days"],"Precio":f"S/ {o['price']}","Pago":o["payment_method"],"Estado":o["status"],"Vence":o.get("expires_at","")[:10] if o.get("expires_at") else "—"} for o in orders]
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    else:
        st.info("Sin órdenes.")

def admin_clients(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">👥 Clientes</div>', unsafe_allow_html=True)
    search = st.text_input("🔍 Buscar por nombre o WhatsApp", key="cli_search")
    clients = db["clients"]
    if search:
        clients = [c for c in clients if search.lower() in c.get("name","").lower() or search in c.get("whatsapp","")]
    col_exp, _ = st.columns([1,4])
    with col_exp:
        st.download_button("📥 Exportar Excel", export_excel(db), "clientes.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if clients:
        rows = [{"Nombre":c["name"],"WhatsApp":c.get("whatsapp",""),"Email":c.get("email",""),"Órdenes":len(c.get("orders",[])),"Registro":c.get("created_at","")[:10],"Último Pedido":c.get("last_order","")[:10]} for c in clients]
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    else:
        st.info("Sin clientes aún.")

def admin_accounts(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">🔑 Cuentas & Perfiles</div>', unsafe_allow_html=True)
    plat_names = {p["id"]:p["name"] for p in db["platforms"]}
    plat_options_filter = ["Todas"] + [p["name"] for p in db["platforms"]]
    c_filter, c_add = st.columns([2,1])
    with c_filter:
        plat_filter = st.selectbox("Filtrar por plataforma",plat_options_filter,key="acc_filter")
    with c_add:
        st.markdown("<br>",unsafe_allow_html=True)
        if st.button("➕ Agregar Cuenta",key="add_acc_btn",use_container_width=True):
            st.session_state.show_add_acc = not st.session_state.get("show_add_acc",False)

    if st.session_state.get("show_add_acc"):
        with st.expander("➕ Nueva Cuenta",expanded=True):
            plat_map = {p["name"]:p["id"] for p in db["platforms"]}
            sel_p = st.selectbox("Plataforma *",list(plat_map.keys()),key="nacc_plat")
            c1,c2 = st.columns(2)
            with c1: ne = st.text_input("Email / Usuario",key="nacc_email")
            with c2: np = st.text_input("Contraseña",key="nacc_pass")
            c3,c4 = st.columns(2)
            with c3: npr = st.text_input("Nombre Perfil",key="nacc_profile")
            with c4: npin = st.text_input("PIN Perfil",key="nacc_pin")
            nex = st.text_input("Info Extra (opcional)",key="nacc_extra")
            if st.button("💾 Guardar",key="save_nacc",use_container_width=True):
                pid = plat_map[sel_p]
                new_acc = {"id":str(uuid.uuid4())[:8],"platform_id":pid,"email":ne,"password":np,"profile_name":npr,"profile_pin":npin,"extra":nex,"status":"available","created_at":datetime.now().isoformat()}
                db["accounts"].append(new_acc)
                for pl in db["platforms"]:
                    if pl["id"]==pid:
                        pl["stock"] = len([a for a in db["accounts"] if a["platform_id"]==pid and a["status"]=="available"])+1
                save_db(db); st.success("✅ Cuenta agregada"); st.session_state.show_add_acc=False; st.rerun()

    accounts = db["accounts"]
    if plat_filter != "Todas":
        pid_f = next((p["id"] for p in db["platforms"] if p["name"]==plat_filter),None)
        accounts = [a for a in accounts if a.get("platform_id")==pid_f]
    if accounts:
        rows = [{"ID":a["id"],"Plataforma":plat_names.get(a["platform_id"],a["platform_id"]),"Email":a.get("email",""),"Contraseña":a.get("password",""),"Perfil":a.get("profile_name",""),"PIN":a.get("profile_pin",""),"Estado":a["status"],"Asignado a":a.get("assigned_to","—"),"Vence":a.get("expires_at","")[:10] if a.get("expires_at") else "—"} for a in accounts]
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
        st.markdown("**✏️ Editar/Eliminar cuenta por ID**")
        acc_ids = [a["id"] for a in accounts]
        sel_id = st.selectbox("Seleccionar ID",["—"]+acc_ids,key="edit_acc_sel")
        if sel_id != "—":
            acc = next(a for a in db["accounts"] if a["id"]==sel_id)
            c1,c2 = st.columns(2)
            with c1: new_e = st.text_input("Email",value=acc.get("email",""),key="edit_email")
            with c2: new_p = st.text_input("Contraseña",value=acc.get("password",""),key="edit_pass")
            c3,c4 = st.columns(2)
            with c3: new_pr = st.text_input("Perfil",value=acc.get("profile_name",""),key="edit_profile")
            with c4: new_pin = st.text_input("PIN",value=acc.get("profile_pin",""),key="edit_pin")
            new_st = st.selectbox("Estado",["available","assigned"],index=0 if acc["status"]=="available" else 1,key="edit_status")
            cs,cd = st.columns(2)
            with cs:
                if st.button("💾 Actualizar",key="upd_acc",use_container_width=True):
                    acc.update({"email":new_e,"password":new_p,"profile_name":new_pr,"profile_pin":new_pin,"status":new_st})
                    save_db(db); st.success("✅ Actualizado"); st.rerun()
            with cd:
                if st.button("🗑 Eliminar",key="del_acc",use_container_width=True):
                    db["accounts"]=[a for a in db["accounts"] if a["id"]!=sel_id]
                    save_db(db); st.success("✅ Eliminada"); st.rerun()
    else:
        st.info("No hay cuentas. Agrega una arriba.")

def admin_platforms(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">📺 Plataformas</div>', unsafe_allow_html=True)
    rows = [{"ID":p["id"],"Nombre":p["name"],"Cat":p["category"],"Entrega":p["delivery"],"Stock":p["stock"],"Precios":str(p.get("prices",{})),"Activo":"✅" if p.get("active") else "❌"} for p in db["platforms"]]
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
    t_edit,t_new = st.tabs(["✏️ Editar Plataforma","➕ Nueva Plataforma"])
    with t_new:
        c1,c2 = st.columns(2)
        with c1: np_id = st.text_input("ID slug único",placeholder="mi_plataforma",key="np_id")
        with c2: np_name = st.text_input("Nombre",placeholder="Mi Plataforma",key="np_name")
        c3,c4 = st.columns(2)
        with c3: np_cat = st.selectbox("Categoría",["streaming","musica","ia"],key="np_cat")
        with c4: np_del = st.selectbox("Tipo entrega",["perfil","cuenta","correo","correo_manual","dispositivo"],key="np_del")
        np_img = st.text_input("URL Imagen/Logo",key="np_img")
        np_desc = st.text_area("Descripción",key="np_desc")
        np_stock = st.number_input("Stock",min_value=0,value=5,key="np_stock")
        c5,c6,c7,c8 = st.columns(4)
        with c5: np30 = st.number_input("Precio 30d",min_value=0.0,value=0.0,key="np30")
        with c6: np60 = st.number_input("Precio 60d",min_value=0.0,value=0.0,key="np60")
        with c7: np90 = st.number_input("Precio 90d",min_value=0.0,value=0.0,key="np90")
        with c8: np365 = st.number_input("Precio 365d",min_value=0.0,value=0.0,key="np365")
        np_active = st.checkbox("Activa en tienda",value=True,key="np_active")
        if st.button("✅ Crear Plataforma",key="create_plat",use_container_width=True):
            prices={}
            if np30>0: prices["30"]=np30
            if np60>0: prices["60"]=np60
            if np90>0: prices["90"]=np90
            if np365>0: prices["365"]=np365
            db["platforms"].append({"id":np_id,"name":np_name,"category":np_cat,"delivery":np_del,"image":np_img,"description":np_desc,"stock":int(np_stock),"prices":prices,"active":np_active})
            save_db(db); st.success("✅ Plataforma creada"); st.rerun()
    with t_edit:
        plat_map = {p["name"]:p["id"] for p in db["platforms"]}
        sel = st.selectbox("Seleccionar",["—"]+list(plat_map.keys()),key="ep_sel")
        if sel != "—":
            p = next(pl for pl in db["platforms"] if pl["id"]==plat_map[sel])
            prices = p.get("prices",{})
            c1,c2 = st.columns(2)
            with c1: ep_name = st.text_input("Nombre",value=p["name"],key="ep_name")
            with c2: ep_stock = st.number_input("Stock",min_value=0,value=p["stock"],key="ep_stock")
            ep_desc = st.text_area("Descripción",value=p.get("description",""),key="ep_desc")
            ep_img = st.text_input("URL Imagen",value=p.get("image",""),key="ep_img")
            c3,c4 = st.columns(2)
            with c3: ep_cat = st.selectbox("Categoría",["streaming","musica","ia"],index=["streaming","musica","ia"].index(p.get("category","streaming")),key="ep_cat")
            with c4: ep_del = st.selectbox("Entrega",["perfil","cuenta","correo","correo_manual","dispositivo"],index=["perfil","cuenta","correo","correo_manual","dispositivo"].index(p.get("delivery","perfil")),key="ep_del")
            c5,c6,c7,c8 = st.columns(4)
            with c5: ep30 = st.number_input("P.30d",value=float(prices.get("30",0)),key="ep30")
            with c6: ep60 = st.number_input("P.60d",value=float(prices.get("60",0)),key="ep60")
            with c7: ep90 = st.number_input("P.90d",value=float(prices.get("90",0)),key="ep90")
            with c8: ep365 = st.number_input("P.365d",value=float(prices.get("365",0)),key="ep365")
            ep_active = st.checkbox("Activa",value=p.get("active",True),key="ep_active")
            cu,cd = st.columns(2)
            with cu:
                if st.button("💾 Actualizar",key="upd_plat",use_container_width=True):
                    np2={}
                    if ep30>0: np2["30"]=ep30
                    if ep60>0: np2["60"]=ep60
                    if ep90>0: np2["90"]=ep90
                    if ep365>0: np2["365"]=ep365
                    p.update({"name":ep_name,"stock":int(ep_stock),"description":ep_desc,"image":ep_img,"category":ep_cat,"delivery":ep_del,"prices":np2,"active":ep_active})
                    save_db(db); st.success("✅ Actualizada"); st.rerun()
            with cd:
                if st.button("🗑 Eliminar",key="del_plat",use_container_width=True):
                    db["platforms"]=[pl for pl in db["platforms"] if pl["id"]!=plat_map[sel]]
                    save_db(db); st.success("✅ Eliminada"); st.rerun()

def admin_finances(db):
    import pandas as pd
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">💰 Finanzas</div>', unsafe_allow_html=True)
    ingresos = db["finances"].get("ingresos",[])
    egresos = db["finances"].get("egresos",[])
    total_in = sum(i["amount"] for i in ingresos)
    total_eg = sum(e["amount"] for e in egresos)
    c1,c2,c3 = st.columns(3)
    c1.metric("💚 Ingresos Totales",f"S/ {total_in:.2f}")
    c2.metric("🔴 Egresos Totales",f"S/ {total_eg:.2f}")
    c3.metric("🟡 Ganancia Neta",f"S/ {(total_in-total_eg):.2f}")
    with st.expander("➕ Registrar Egreso"):
        eg_desc = st.text_input("Descripción",placeholder="Ej: Renovación Netflix",key="eg_desc")
        eg_amt = st.number_input("Monto (S/)",min_value=0.0,key="eg_amt")
        if st.button("Registrar",key="save_eg",use_container_width=True):
            db["finances"]["egresos"].append({"date":datetime.now().isoformat(),"amount":float(eg_amt),"description":eg_desc})
            save_db(db); st.success("✅ Egreso registrado"); st.rerun()
    col_e,_ = st.columns([1,3])
    with col_e:
        st.download_button("📥 Excel Finanzas",export_excel(db),"finanzas.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    t1,t2 = st.tabs(["💚 Ingresos","🔴 Egresos"])
    with t1:
        if ingresos:
            st.dataframe(pd.DataFrame([{"Fecha":i["date"][:10],"Descripción":i["description"],"Monto":f"S/ {i['amount']}"} for i in reversed(ingresos)]),use_container_width=True,hide_index=True)
        else: st.info("Sin ingresos.")
    with t2:
        if egresos:
            st.dataframe(pd.DataFrame([{"Fecha":e["date"][:10],"Descripción":e["description"],"Monto":f"S/ {e['amount']}"} for e in reversed(egresos)]),use_container_width=True,hide_index=True)
        else: st.info("Sin egresos.")

def admin_pending(db):
    st.markdown('<div style="font-family:Syne,sans-serif;font-weight:800;font-size:1.35rem;color:#f0f0f8;margin-bottom:16px;">⏳ Pedidos Manuales</div>', unsafe_allow_html=True)
    pending = [o for o in db["orders"] if o["status"]=="pending_manual"]
    if not pending:
        st.success("✅ Sin pedidos pendientes. ¡Todo al día!")
        return
    st.warning(f"⚠️ Tienes **{len(pending)}** pedido(s) pendiente(s) de activación manual.")
    for o in pending:
        with st.expander(f"📋 #{o['id']} — {o['platform_name']} — {o['client_name']}"):
            c1,c2 = st.columns(2)
            with c1:
                st.write(f"**Cliente:** {o['client_name']}")
                st.write(f"**WhatsApp:** {o['client_whatsapp']}")
                st.markdown(f"[💬 Abrir WhatsApp](https://wa.me/{o['client_whatsapp']})")
            with c2:
                st.write(f"**Email:** `{o.get('client_email','—')}`")
                st.write(f"**Plan:** {o['days']} días — **S/ {o['price']}**")
                st.write(f"**Código pago:** `{o['payment_code']}`")
            st.write(f"**Fecha:** {o['created_at'][:16].replace('T',' ')}")
            if st.button(f"✅ Marcar completado",key=f"comp_{o['id']}",use_container_width=True):
                for ord2 in db["orders"]:
                    if ord2["id"]==o["id"]:
                        ord2["status"]="completed"
                        ord2["completed_at"]=datetime.now().isoformat()
                        db["finances"]["ingresos"].append({"date":datetime.now().isoformat(),"amount":float(o.get("price",0)),"description":f"{o['platform_name']} - {o['client_name']}","order_id":o["id"]})
                save_db(db); st.success(f"✅ Orden #{o['id']} completada"); st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🎬 StreamVault")
    if st.session_state.page == "store":
        if st.button("🔐 Panel Admin", use_container_width=True):
            st.session_state.page = "admin"
            st.rerun()
    else:
        if st.button("🏪 Ver Tienda", use_container_width=True):
            st.session_state.page = "store"
            st.session_state.admin_logged = False
            st.rerun()

if st.session_state.page == "admin":
    page_admin()
else:
    page_store()