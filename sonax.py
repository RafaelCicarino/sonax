# -*- coding: utf-8 -*-

import re
import time
import socket
from dataclasses import dataclass
from io import BytesIO
from typing import List, Optional

import pandas as pd
import streamlit as st
from docx import Document
import streamlit.components.v1 as components

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://chat.sonax.net.br/app/omnichannel/chat"


@dataclass
class Cliente:
    nome: str
    placa: str
    telefone: str
    endereco: str = ""
    horario: str = ""  # data


PHONE_RE = re.compile(r"(?:\+?55)?\s*\(?\d{2}\)?\s*\d{4,5}-?\d{4}")
PLATE_CANDIDATE_RE = re.compile(r"\b[A-Z]{3}[A-Z0-9]{4,5}\b", re.IGNORECASE)


def normalize_phone(raw: str) -> Optional[str]:
    if not raw:
        return None
    digits = re.sub(r"\D+", "", raw)
    if len(digits) == 11:
        return digits
    if len(digits) == 13 and digits.startswith("55"):
        return digits
    return None


def phone_variations(phone: str) -> List[str]:
    d = re.sub(r"\D+", "", phone)
    out = []
    if len(d) == 11:
        out += [d, "55" + d]
    elif len(d) == 13 and d.startswith("55"):
        out += [d, d[2:]]
    else:
        out += [d]
    seen, res = set(), []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            res.append(x)
    return res


def normalize_plate(raw: str) -> str:
    if not raw:
        return ""
    return re.sub(r"[^A-Za-z0-9]+", "", raw).upper()


def is_valid_plate_relaxed(s: str) -> bool:
    s = normalize_plate(s)
    if not (7 <= len(s) <= 8):
        return False
    if not re.match(r"^[A-Z]{3}[A-Z0-9]+$", s):
        return False
    if not any(ch.isdigit() for ch in s):
        return False
    return True


def find_plate_in_text(text: str) -> str:
    m = re.search(r"(?i)\bplaca\s*:\s*([A-Z0-9-]+)", text)
    if m:
        p = normalize_plate(m.group(1))
        if is_valid_plate_relaxed(p):
            return p
    for cand in PLATE_CANDIDATE_RE.findall(text):
        p = normalize_plate(cand)
        if is_valid_plate_relaxed(p):
            return p
    return ""


def docx_lines_preserve_blanks(doc: Document) -> List[str]:
    out: List[str] = []
    for p in doc.paragraphs:
        out.append((p.text or "").strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [((c.text or "").strip()) for c in row.cells]
            out.append("" if not any(cells) else " | ".join(cells))
    return out


def parse_record_block(block_lines: List[str]) -> Optional[Cliente]:
    lines = [ln.strip() for ln in block_lines if ln and ln.strip()]
    if not lines:
        return None

    joined = "\n".join(lines)

    mph = PHONE_RE.search(joined)
    phone = normalize_phone(mph.group(0)) if mph else None
    plate = find_plate_in_text(joined)

    nome = ""
    for ln in lines[:6]:
        if phone and phone[-11:] in re.sub(r"\D+", "", ln):
            continue
        if "placa" in ln.lower():
            continue
        if re.search(r"(?i)\b(rua|avenida|av\.|rodovia|estrada)\b", ln):
            continue
        nome = ln
        break
    if not nome:
        nome = "Sem nome"

    endereco = ""
    for ln in lines:
        if re.search(r"(?i)\b(rua|avenida|av\.|rodovia|estrada)\b", ln):
            endereco = ln
            break

    data = ""
    mdate = re.search(r"\b\d{2}/\d{2}/\d{4}\b", joined)
    if mdate:
        data = mdate.group(0)

    if phone and plate:
        return Cliente(nome=nome, telefone=phone, placa=plate, endereco=endereco, horario=data)

    return None


def parse_docx_to_clients(file_bytes: bytes) -> List[Cliente]:
    doc = Document(BytesIO(file_bytes))
    lines = docx_lines_preserve_blanks(doc)

    blocks, cur = [], []
    for ln in lines:
        if ln.strip() == "":
            if cur:
                blocks.append(cur)
                cur = []
        else:
            cur.append(ln)
    if cur:
        blocks.append(cur)

    out = []
    seen = set()
    for b in blocks:
        c = parse_record_block(b)
        if c:
            k = (c.telefone, c.placa)
            if k not in seen:
                seen.add(k)
                out.append(c)
    return out


def port_open(host: str, port: int, timeout_s: float = 0.35) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout_s):
            return True
    except Exception:
        return False


def make_driver_attach(debug_port: int) -> webdriver.Chrome:
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    return webdriver.Chrome(options=opts)


def make_driver_new() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")
    return webdriver.Chrome(options=opts)


def maybe_close_popup(driver) -> bool:
    try:
        btn = WebDriverWait(driver, 1.2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#botao-fechar"))
        )
        btn.click()
        time.sleep(0.12)
        return True
    except Exception:
        return False


def click_retry(driver, by, value, tries=3, timeout=25):
    last = None
    for _ in range(tries):
        try:
            el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.click()
            return el
        except Exception as e:
            last = e
            maybe_close_popup(driver)
            time.sleep(0.25)
    raise last


def type_retry(driver, by, value, text, clear=True, press_enter=False, tries=3, timeout=25):
    last = None
    for _ in range(tries):
        try:
            el = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, value)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
            el.click()
            if clear:
                el.send_keys(Keys.CONTROL, "a")
                el.send_keys(Keys.BACKSPACE)
            el.send_keys(text)
            if press_enter:
                el.send_keys(Keys.ENTER)
            return el
        except Exception as e:
            last = e
            maybe_close_popup(driver)
            time.sleep(0.25)
    raise last


def ensure_sonax_tab(driver):
    target = "chat.sonax.net.br/app/omnichannel/chat"
    for h in driver.window_handles:
        driver.switch_to.window(h)
        try:
            if target in (driver.current_url or ""):
                return
        except Exception:
            pass

    try:
        driver.get(URL)
    except Exception:
        driver.execute_script(f"window.open('{URL}','_blank');")
        driver.switch_to.window(driver.window_handles[-1])

    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
    )


# Seleção Sonax
SEL_CONTATOS = (By.XPATH, "//a[contains(@class,'nav-link')][contains(.,'Contatos')]")
SEL_BUSCA = (By.CSS_SELECTOR, "input.form-control.input-search")
SEL_CONVERSAR = (By.CSS_SELECTOR, "button#dropdownBasic1")
SEL_LWSIMAPP = (By.XPATH, "//*[contains(@class,'ml-2') and contains(.,'LWSIMAPP')]")
SEL_COMBOBOX = (By.CSS_SELECTOR, "input[role='combobox']")
SEL_TEMPLATE = (
    By.XPATH,
    "//span[contains(@class,'ng-option-label') and normalize-space(.)='abertura_de_diagnostico_tagpro']",
)
SEL_VAR_INPUTS = (By.CSS_SELECTOR, "input.form-control[placeholder='Insira a variável aqui']")
SEL_ENVIAR = (By.XPATH, "//button[contains(@class,'btn-primary') and normalize-space(.)='Enviar']")


def click_card_contact(driver, phone_digits: str) -> bool:
    digits = re.sub(r"\D+", "", phone_digits)
    xpath = (
        "//*[contains(@class,'kt-widget__info')]"
        f"[.//span[contains(@class,'kt-widget__desc')][contains(translate(., ' ()+-', ''), '{digits}')]]"
    )
    try:
        click_retry(driver, By.XPATH, xpath, tries=2, timeout=8)
        return True
    except Exception:
        try:
            click_retry(driver, By.CSS_SELECTOR, ".kt-widget__info", tries=1, timeout=3)
            return True
        except Exception:
            return False


def fill_template_variables_in_order(driver, placa: str, data: str, endereco: str):
    inputs = WebDriverWait(driver, 25).until(lambda d: d.find_elements(*SEL_VAR_INPUTS))
    if len(inputs) < 3:
        raise RuntimeError(f"Esperava 3 campos de variável, mas encontrei {len(inputs)}.")

    values = [placa or "", data or "", endereco or ""]
    for i in range(3):
        el = inputs[i]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
        el.send_keys(values[i])
        time.sleep(0.1)


def run_one_client(driver, client: Cliente, log):
    maybe_close_popup(driver)

    log(f"➡️ {client.nome}: Contatos")
    click_retry(driver, *SEL_CONTATOS, tries=3, timeout=30)
    maybe_close_popup(driver)

    found = False
    for ph in phone_variations(client.telefone):
        log(f"🔎 {client.nome}: Buscar {ph}")
        click_retry(driver, *SEL_BUSCA, tries=3, timeout=30)
        type_retry(driver, *SEL_BUSCA, ph, clear=True, press_enter=True, tries=3, timeout=30)
        time.sleep(0.9)
        if click_card_contact(driver, ph):
            found = True
            break

    if not found:
        log(f"❌ {client.nome}: NÃO ENCONTRADO")
        return {"nome": client.nome, "telefone": client.telefone, "placa": client.placa, "status": "NÃO ENCONTRADO"}

    maybe_close_popup(driver)

    log(f"💬 {client.nome}: Conversar")
    click_retry(driver, *SEL_CONVERSAR, tries=3, timeout=30)
    maybe_close_popup(driver)

    log(f"📲 {client.nome}: LWSIMAPP")
    click_retry(driver, *SEL_LWSIMAPP, tries=3, timeout=30)
    maybe_close_popup(driver)

    log(f"🧾 {client.nome}: Template")
    click_retry(driver, *SEL_COMBOBOX, tries=3, timeout=30)
    click_retry(driver, *SEL_TEMPLATE, tries=3, timeout=30)
    maybe_close_popup(driver)

    log(f"⌨️ {client.nome}: preenchendo variáveis (placa/data/endereço)")
    fill_template_variables_in_order(driver, client.placa, client.horario, client.endereco)
    maybe_close_popup(driver)

    log(f"📨 {client.nome}: Enviar")
    click_retry(driver, *SEL_ENVIAR, tries=3, timeout=30)

    log(f"✅ {client.nome}: OK")
    return {"nome": client.nome, "telefone": client.telefone, "placa": client.placa, "status": "OK"}


# =========================
# UI (PT-BR)
# =========================

st.set_page_config(page_title="Sonax Automação Kezia", layout="wide")

# CSS: sem [class*="st-"] (corrige o bug do texto duplicado)
st.markdown(
    """
<style>
section[data-testid="stSidebar"] { display: none !important; }
div[data-testid="collapsedControl"] { display: none !important; }
.block-container { padding-top: 1.2rem; }

html, body {
  font-family: "Segoe UI", "Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji", sans-serif !important;
}

input, textarea, [contenteditable="true"] {
  text-shadow: none !important;
  -webkit-text-stroke: 0 !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# JS: traduz textos do file_uploader (Browse files, Drag and drop..., Limit...)
components.html(
    """
<script>
(function() {
  const PT = {
    browse: "Selecionar arquivo",
    drop: "Arraste e solte o arquivo aqui",
    limit: "Limite de 200MB por arquivo • DOCX"
  };

  function applyPT() {
    // tenta localizar todas as áreas de uploader na página
    const uploaders = parent.document.querySelectorAll('[data-testid="stFileUploader"]');
    uploaders.forEach(u => {
      // botão
      const btn = u.querySelector('button');
      if (btn) {
        const span = btn.querySelector('span') || btn;
        if (span && span.innerText && span.innerText.toLowerCase().includes('browse')) {
          span.innerText = PT.browse;
        }
      }

      // dropzone
      const drop = u.querySelector('[data-testid="stFileUploaderDropzone"]');
      if (drop) {
        // texto principal (p/div)
        const nodes = drop.querySelectorAll('p, div, span');
        nodes.forEach(n => {
          const t = (n.innerText || '').trim().toLowerCase();
          if (t === 'drag and drop file here' || t.includes('drag and drop')) {
            n.innerText = PT.drop;
          }
        });

        // linha do limite (small)
        const smalls = drop.querySelectorAll('small');
        smalls.forEach(s => {
          const t = (s.innerText || '').trim().toLowerCase();
          if (t.includes('limit') && (t.includes('per file') || t.includes('file'))) {
            s.innerText = PT.limit;
          }
        });
      }
    });
  }

  // aplica agora e fica observando mudanças (Streamlit re-renderiza)
  applyPT();

  const obs = new MutationObserver(() => applyPT());
  obs.observe(parent.document.body, { childList: true, subtree: true });
})();
</script>
""",
    height=0,
)

st.title("Sonax • Automação da KEZIA")

st.session_state.setdefault("attach", True)
st.session_state.setdefault("debug_port", 9222)
st.session_state.setdefault("max_items", 50)

with st.expander("Configurar", expanded=False):
    st.session_state.max_items = st.number_input(
        "Processar quantos clientes?",
        1,
        500,
        int(st.session_state.max_items),
        1,
    )

st.markdown("---")

uploaded = st.file_uploader("📄 Envie o arquivo DOCX com os clientes", type=["docx"])
if not uploaded:
    st.info("Envie o arquivo para carregar os clientes.")
    st.stop()

clients = parse_docx_to_clients(uploaded.getvalue())
clients = clients[: int(st.session_state.max_items)]

st.success(f"Clientes carregados: {len(clients)}")

with st.expander("Ver clientes identificados", expanded=False):
    st.dataframe(pd.DataFrame([c.__dict__ for c in clients]), use_container_width=True, hide_index=True)

start = st.button("▶ Iniciar automação", type="primary")

if start:
    status_box = st.empty()
    log_box = st.empty()
    prog = st.progress(0)
    logs = []

    def log(msg: str):
        logs.append(msg)
        log_box.write("\n".join(logs[-35:]))

    results = []
    driver = None

    try:
        if st.session_state.attach:
            status_box.info("Testando porta do Chrome...")
            if not port_open("127.0.0.1", int(st.session_state.debug_port)):
                status_box.warning("Não encontrei Chrome na porta informada. Vou abrir um Chrome novo.")
                driver = make_driver_new()
                driver.get(URL)
            else:
                status_box.info("Conectando ao Chrome existente...")
                driver = make_driver_attach(int(st.session_state.debug_port))
        else:
            status_box.info("Abrindo Chrome novo...")
            driver = make_driver_new()
            driver.get(URL)

        status_box.info("Indo para a aba do Sonax...")
        ensure_sonax_tab(driver)

        status_box.success("Executando automação...")
        for i, c in enumerate(clients, start=1):
            status_box.info(f"Processando {i}/{len(clients)}: {c.nome}")
            r = run_one_client(driver, c, log)
            results.append(r)
            prog.progress(i / len(clients))

        status_box.success("Finalizado!")
    except Exception as e:
        status_box.error(f"Erro na automação: {e}")
        log("⚠️ Se o template não tiver 3 variáveis ou a ordem for diferente, me avise que eu mapeio pelo label.")
    finally:
        pass

    if results:
        rdf = pd.DataFrame(results)
        st.subheader("Resultado")
        st.dataframe(rdf, use_container_width=True, hide_index=True)
        st.download_button(
            "Baixar resultado (.csv)",
            data=rdf.to_csv(index=False).encode("utf-8-sig"),
            file_name="resultado_sonax.csv",
            mime="text/csv",
        )