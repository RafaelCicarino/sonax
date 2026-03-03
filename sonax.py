import re
import time
from dataclasses import dataclass
from typing import List, Optional

import pandas as pd
import streamlit as st
from docx import Document

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://chat.sonax.net.br/app/omnichannel/chat"


# -----------------------------
# DOCX -> Clientes
# -----------------------------
@dataclass
class Cliente:
    nome: str
    placa: str
    telefone: str
    endereco: str = ""
    horario: str = ""


def normalize_phone(raw: str) -> Optional[str]:
    if not raw:
        return None
    digits = re.sub(r"\D+", "", raw)

    # Aceita: 21999999999 (11) ou 5521999999999 (13)
    if len(digits) == 11:
        return digits
    if len(digits) == 13 and digits.startswith("55"):
        return digits
    return None


def normalize_plate(raw: str) -> str:
    if not raw:
        return ""
    return re.sub(r"[^A-Za-z0-9]+", "", raw).upper()


def parse_docx_to_clients(file_bytes: bytes) -> List[Cliente]:
    doc = Document(file_bytes)
    lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    text = "\n".join(lines)

    blocks = re.split(r"\n{2,}|-{3,}|={3,}", text)

    clientes: List[Cliente] = []

    phone_re = re.compile(r"(?:\+?55)?\s*\(?\d{2}\)?\s*\d{4,5}-?\d{4}")
    plate_re = re.compile(r"\b[A-Z]{3}\d[A-Z0-9]\d{2}\b|\b[A-Z]{3}\d{4}\b", re.IGNORECASE)

    for b in blocks:
        b = b.strip()
        if not b:
            continue

        mphone = phone_re.search(b)
        phone = normalize_phone(mphone.group(0)) if mphone else None

        mplate = plate_re.search(b)
        plate = normalize_plate(mplate.group(0)) if mplate else ""

        nome = ""
        mname = re.search(r"(?im)\b(nome|cliente)\s*:\s*(.+)$", b)
        if mname:
            nome = mname.group(2).strip()
        else:
            candidates = [ln.strip() for ln in b.split("\n") if ln.strip()]
            for c in candidates[:4]:
                if phone and (re.sub(r"\D+", "", c).find(phone[-11:]) != -1):
                    continue
                if plate and (normalize_plate(c).find(plate) != -1):
                    continue
                if re.search(r"(?i)\b(placa|telefone|celular|endereco|endereço|horario|horário)\b", c):
                    continue
                nome = c
                break

        endereco = ""
        horario = ""
        mend = re.search(r"(?im)\b(endereco|endereço)\s*:\s*(.+)$", b)
        if mend:
            endereco = mend.group(2).strip()
        mhor = re.search(r"(?im)\b(horario|horário)\s*:\s*(.+)$", b)
        if mhor:
            horario = mhor.group(2).strip()

        if phone and plate:
            clientes.append(
                Cliente(
                    nome=nome or "Sem nome",
                    placa=plate,
                    telefone=phone,
                    endereco=endereco,
                    horario=horario,
                )
            )

    # Fallback (caso o DOCX venha diferente)
    if not clientes:
        phones = list({normalize_phone(p) for p in phone_re.findall(text)} - {None})
        for ph in phones:
            idx = text.find(ph[-11:])
            window = text[max(0, idx - 200) : idx + 200] if idx != -1 else text
            mp = plate_re.search(window)
            plate = normalize_plate(mp.group(0)) if mp else ""
            if plate:
                clientes.append(Cliente(nome="Sem nome", placa=plate, telefone=ph))

    return clientes


# -----------------------------
# Selenium
# -----------------------------
def make_driver(
    attach: bool,
    debug_port: int,
    headless: bool,
    chromedriver_path: str,
) -> webdriver.Chrome:
    opts = Options()

    if attach:
        # Anexa no Chrome já aberto via remote debugging
        opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    else:
        opts.add_argument("--start-maximized")

    if headless and not attach:
        opts.add_argument("--headless=new")

    opts.add_argument("--disable-notifications")
    opts.add_argument("--disable-popup-blocking")

    service = Service(chromedriver_path)
    return webdriver.Chrome(service=service, options=opts)


def wwait(driver, seconds=20):
    return WebDriverWait(driver, seconds)


def safe_click(driver, by, value, timeout=20):
    el = wwait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    el.click()
    return el


def safe_type(driver, by, value, text, clear_first=True, timeout=20, press_enter=False):
    el = wwait(driver, timeout).until(EC.visibility_of_element_located((by, value)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    if clear_first:
        el.click()
        el.send_keys(Keys.CONTROL, "a")
        el.send_keys(Keys.BACKSPACE)
    el.send_keys(text)
    if press_enter:
        el.send_keys(Keys.ENTER)
    return el


def maybe_close_popup(driver) -> bool:
    try:
        btn = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#botao-fechar"))
        )
        btn.click()
        time.sleep(0.2)
        return True
    except Exception:
        return False


def open_contacts_tab(driver):
    # Clica em "Contatos"
    safe_click(
        driver,
        By.XPATH,
        "//a[contains(@class,'nav-link')][contains(.,'Contatos')]",
        timeout=25,
    )


def search_contact_by_phone(driver, phone: str):
    # Campo de busca
    safe_type(
        driver,
        By.CSS_SELECTOR,
        "input.form-control.input-search",
        phone,
        clear_first=True,
        timeout=25,
        press_enter=True,
    )


def click_contact_card_if_found(driver, phone: str) -> bool:
    digits = re.sub(r"\D+", "", phone)

    # tenta achar card com "Celular:" contendo os dígitos
    xpath = (
        "//*[contains(@class,'kt-widget__info')]"
        f"[.//span[contains(@class,'kt-widget__desc')][contains(translate(., ' ()+-', ''), '{digits}')]]"
    )

    try:
        el = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        el.click()
        return True
    except Exception:
        # fallback: clica no primeiro item da lista se existir
        try:
            el2 = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".kt-widget__info"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el2)
            el2.click()
            return True
        except Exception:
            return False


def click_conversar(driver):
    safe_click(driver, By.CSS_SELECTOR, "button#dropdownBasic1", timeout=20)


def click_channel_lwsimapp(driver):
    safe_click(
        driver,
        By.XPATH,
        "//*[contains(@class,'ml-2') and contains(.,'LWSIMAPP')]",
        timeout=20,
    )


def select_template_abertura(driver):
    # Abre o combobox e seleciona a opção
    safe_click(driver, By.CSS_SELECTOR, "input[role='combobox']", timeout=20)
    safe_click(
        driver,
        By.XPATH,
        "//span[contains(@class,'ng-option-label') and normalize-space(.)='abertura_de_diagnostico_tagpro']",
        timeout=20,
    )


def fill_variable_plate(driver, plate: str):
    safe_type(
        driver,
        By.CSS_SELECTOR,
        "input.form-control[placeholder='Insira a variável aqui']",
        plate,
        clear_first=True,
        timeout=20,
        press_enter=False,
    )


def process_client(driver, client: Cliente) -> dict:
    maybe_close_popup(driver)

    open_contacts_tab(driver)
    maybe_close_popup(driver)

    search_contact_by_phone(driver, client.telefone)
    time.sleep(0.8)  # tempo pro filtro renderizar

    found = click_contact_card_if_found(driver, client.telefone)
    if not found:
        return {"telefone": client.telefone, "placa": client.placa, "status": "NÃO ENCONTRADO"}

    maybe_close_popup(driver)

    click_conversar(driver)
    maybe_close_popup(driver)

    click_channel_lwsimapp(driver)
    maybe_close_popup(driver)

    select_template_abertura(driver)
    maybe_close_popup(driver)

    fill_variable_plate(driver, client.placa)

    return {"telefone": client.telefone, "placa": client.placa, "status": "OK"}


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="Sonax Omnichannel • Automação", layout="wide")
st.title("Automação Sonax Omnichannel (DOCX ➜ Telefone ➜ Template ➜ Placa)")

with st.sidebar:
    st.subheader("Chrome / Selenium")

    attach = st.checkbox("Anexar ao Chrome já aberto (remote debugging)", value=True)
    debug_port = st.number_input("Porta remote debugging", min_value=1, max_value=65535, value=9222, step=1)

    headless = st.checkbox("Headless (só se abrir Chrome novo)", value=False)

    st.markdown("---")
    st.subheader("ChromeDriver")
    chromedriver_path = st.text_input(
        "Caminho do chromedriver.exe",
        value="chromedriver.exe",
        help="Se estiver na mesma pasta do script, deixe 'chromedriver.exe'. Senão, coloque o caminho completo.",
    )

    st.markdown("---")
    max_items = st.number_input("Processar quantos clientes do DOCX", min_value=1, max_value=500, value=50, step=1)

uploaded = st.file_uploader("Envie o arquivo .docx com os clientes", type=["docx"])

if uploaded:
    clients = parse_docx_to_clients(uploaded)
    if not clients:
        st.error("Não consegui identificar clientes no DOCX. Precisa ter pelo menos TELEFONE e PLACA em cada registro.")
        st.stop()

    clients = clients[: int(max_items)]

    df = pd.DataFrame([{
        "nome": c.nome,
        "telefone": c.telefone,
        "placa": c.placa,
        "endereco": c.endereco,
        "horario": c.horario
    } for c in clients])

    st.subheader("Clientes identificados")
    st.dataframe(df, use_container_width=True, hide_index=True)

    col1, col2 = st.columns([1, 2], vertical_alignment="center")
    with col1:
        start = st.button("▶ Iniciar automação", type="primary")
    with col2:
        st.info(
            "Se for anexar no Chrome já aberto, abra o Chrome assim:\n\n"
            "chrome.exe --remote-debugging-port=9222 --user-data-dir=\"%TEMP%\\chrome_sonax_profile\"\n\n"
            "Aí faça login e deixe o omnichannel aberto."
        )

    if start:
        st.warning("Durante a execução, evite mexer no Chrome para não atrapalhar os cliques.")
        log_box = st.empty()
        prog = st.progress(0)

        driver = None
        results = []
        try:
            driver = make_driver(
                attach=attach,
                debug_port=int(debug_port),
                headless=headless,
                chromedriver_path=chromedriver_path.strip(),
            )

            if not attach:
                driver.get(URL)

            # garante carregamento inicial
            WebDriverWait(driver, 30).until(
                lambda d: d.execute_script("return document.readyState") in ("interactive", "complete")
            )

            for i, c in enumerate(clients, start=1):
                log_box.write(f"[{i}/{len(clients)}] Telefone={c.telefone} • Placa={c.placa}")
                r = process_client(driver, c)
                results.append(r)
                prog.progress(i / len(clients))

            st.success("Finalizado!")
        except Exception as e:
            st.error(f"Erro na automação: {e}")
        finally:
            # Se anexou no Chrome do usuário, não fecha.
            if driver and not attach:
                try:
                    driver.quit()
                except Exception:
                    pass

        if results:
            st.subheader("Resultado")
            rdf = pd.DataFrame(results)
            st.dataframe(rdf, use_container_width=True, hide_index=True)

            st.download_button(
                "Baixar resultado (.csv)",
                data=rdf.to_csv(index=False).encode("utf-8"),
                file_name="resultado_sonax.csv",
                mime="text/csv",
            )
else:
    st.caption("Envie um DOCX para o app identificar telefone/placa e iniciar.")


st.markdown("---")
st.caption("Se algum seletor mudar no site (classes/ids), me mande um print do HTML do elemento que eu ajusto os seletores.")