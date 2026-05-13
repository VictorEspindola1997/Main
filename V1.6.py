import tkinter as tk
from tkinter import messagebox, simpledialog
import win32com.client as win32
import pythoncom
import os
import time
import pandas as pd
import re
import unicodedata
from datetime import datetime
import threading
import sys
import zipfile
import win32api
import win32con
import shutil
import glob
import subprocess
import webbrowser
import tempfile
from pywinauto import Application
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import customtkinter as ctk
from PIL import Image

# ================== CONFIGURAÇÕES ==================
TXT_ULTIMO_ENVIO = "ultimo_salvamento.txt"
TXT_ENVIO = "ultimos_envios.txt"
TXT_COBRANCA = "ultimas_cobrancas.txt"
TXT_PROCURACOES = "Procuracoes.txt"
TXT_ULTIMO_AJUSTE = "ultimo_ajuste.txt"
TXT_ULTIMO_GERAR = "ultimo_gerar.txt"

# Configurações padrão (usuário Victor)
BASE_DIR_VICTOR = r"C:\Users\TI Medicina\Desktop\ESOCIAL - PYTHON"
CAMINHO_CONTROLE_VICTOR = os.path.join(
    BASE_DIR_VICTOR, "Controle do E-Social.xlsx")
EMAIL_CONTA_PADRAO_VICTOR = "ti@medicinaoswaldocruz.com.br"
PASTA_PLANILHAS_VICTOR = os.path.join(BASE_DIR_VICTOR, "Planilhas")
PASTA_DOWNLOAD_VICTOR = r"C:\Users\TI Medicina\Downloads"

# Configurações Thailany
BASE_DIR_THAILANY = r"D:\Desktop\Oswaldo Cruz Gestão Ocupacional\E-Social"
CAMINHO_CONTROLE_THAILANY = os.path.join(
    BASE_DIR_THAILANY, "Controle do E-Social.xlsx")
EMAIL_CONTA_PADRAO_THAILANY = "esocial@medicinaoswaldocruz.com.br"
EMAIL_PLANILHAS_THAILANY = "planilhasesocial@outlook.com"
PASTA_PLANILHAS_THAILANY = os.path.join(
    BASE_DIR_THAILANY, "Planilhas Mensais - Conferência")
PASTA_DOWNLOAD_THAILANY = r"D:\Downloads"

CAMINHO_PATINHO_1 = os.path.join(BASE_DIR_VICTOR, "patinho.gif")
CAMINHO_PATINHO_2 = os.path.join(BASE_DIR_THAILANY, "patinho.gif")

# --- Novas definições de fonte ---
FONT_FAMILY = "Magsen-Regular"
FONT_NORMAL = (FONT_FAMILY, 12)
FONT_BOLD = (FONT_FAMILY, 12, "bold")
FONT_ENTRY = (FONT_FAMILY, 13)
FONT_LABEL = (FONT_FAMILY, 15)
FONT_BUTTON = (FONT_FAMILY, 14, "bold")
FONT_FORGOT_PASSWORD = (FONT_FAMILY, 14, "underline")
FONT_TITLE = (FONT_FAMILY, 24, "bold")

MAX_LINHAS_COLUNA = 5
IDX_AK_LABEL = 36
IDX_AL_ANEXOS = 37
IDX_AM_ASSUNTO = 38
IDX_EMAIL_TO = 32
IDX_EMAIL_CC = 33
IDX_F_BLOQUEADA = 5  # Coluna F para verificar "BLOQUEADA"
IDX_AA_STATUS = 26  # Coluna AA para verificar "OK"
IDX_AC_VALIDADE = 28  # Coluna AC para validade da procuração
IDX_AE_PROCURACAO_TO = 30  # Coluna AE para email de procuração
IDX_AF_PROCURACAO_CC = 31  # Coluna AF para email de procuração em cópia
EXCLUIR_NA_COBRANCA = {"GRUPO OSWALDO CRUZ"}

# Configurações de Aparência Corporativa
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# ================== UTILITÁRIOS ==================

def hide_file(file_path):
    try:
        if os.path.exists(file_path):
            win32api.SetFileAttributes(file_path, win32con.FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass

def unhide_file(file_path):
    try:
        if os.path.exists(file_path):
            win32api.SetFileAttributes(file_path, win32con.FILE_ATTRIBUTE_NORMAL)
    except Exception:
        pass

def normalize_key(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = unicodedata.normalize("NFD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch)).lower()

def safe_email_field(val) -> str:
    if val is None: return ""
    try:
        if isinstance(val, float) and pd.isna(val): return ""
    except Exception: pass
    s = str(val).strip()
    if s.lower() in {"nan", "none", "null", ""}: return ""
    return s

def limpar_caminho(p: str) -> str:
    if p is None: return ""
    return str(p).strip().strip('"').strip("'")

def extrair_nome_empresa_do_arquivo(caminho: str) -> str:
    nome = os.path.basename(caminho)
    nome = re.sub(r"(?i)listagem de funcionarios\s*-\s*", "", nome)
    return re.sub(r"\.(xlsx|xlsm|xls)$", "", nome, flags=re.I).strip()

def centralizar_janela(janela, largura=850, altura=700):
    janela.update_idletasks()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

def carregar_controle(usuario):
    if usuario == "thailany":
        caminho_controle = CAMINHO_CONTROLE_THAILANY
    else:
        caminho_controle = CAMINHO_CONTROLE_VICTOR
    if not os.path.isfile(caminho_controle):
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_controle}")
    df = pd.read_excel(caminho_controle, sheet_name=0, engine="openpyxl")
    if df.shape[1] <= IDX_AK_LABEL:
        raise ValueError(f"A planilha tem {df.shape[1]} colunas, mas o código espera pelo menos {IDX_AK_LABEL + 1}.")
    return df

def get_row_by_label(df, label):
    mask = df.iloc[:, IDX_AK_LABEL].astype(str) == label
    if not mask.any():
        return None
    return df[mask].iloc[0]

def get_row_by_name(df, name):
    mask = df.iloc[:, 0].astype(str).str.upper() == name.upper()
    if not mask.any():
        return None
    return df[mask].iloc[0]

def carregar_txt(nome_arquivo):
    registros = {}
    if os.path.isfile(nome_arquivo):
        with open(nome_arquivo, "r", encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if "|" in linha:
                    parts = linha.split("|", 1)
                    if len(parts) == 2:
                        registros[parts[0]] = parts[1]
    return registros

def registrar_txt(nome_arquivo, empresa):
    hoje = datetime.now().strftime("%d/%m/%Y")
    if not os.path.isfile(nome_arquivo):
        with open(nome_arquivo, "w", encoding="utf-8") as f:
            f.write("")
        hide_file(nome_arquivo)
    with open(nome_arquivo, "a", encoding="utf-8") as f:
        f.write(f"{empresa}|{hoje}\n")
    hide_file(nome_arquivo)

def get_anexos_from_row(row, usuario):
    try:
        raw = str(row.iloc[IDX_AL_ANEXOS])
    except Exception:
        raw = ""
    if not raw or raw.lower() in {"nan", "none"}: return []
    parts = re.split(r"[;,]", raw)
    anexos = []
    for p in parts:
        p_clean = limpar_caminho(p)
        if p_clean and os.path.isfile(p_clean):
            anexos.append(p_clean)
    return anexos

def get_assunto_from_row(row):
    try:
        if IDX_AM_ASSUNTO >= len(row): return ""
        return safe_email_field(row.iloc[IDX_AM_ASSUNTO])
    except Exception:
        return ""

def listar_labels(df):
    labels = list(df.iloc[2:, IDX_AK_LABEL].dropna().astype(str))
    labels.sort(key=normalize_key)
    return labels

def abrir_link_ajuda(event=None):
    webbrowser.open("https://wa.me/551221239207?text=Encontrei%20um%20erro%20no%20sistema%20E.A.I.%20Pode%20me%20ajudar%3F")

# ================== LÓGICA DE NEGÓCIO ORIGINAIS (Selenium / Excel / Outlook) ==================

def executar_gerar_planilha(app_usuario, soc_usuario, soc_senha, soc_id, app_instance, btn_widget, status_banner):
    if app_usuario == "thailany":
        download_dir = PASTA_DOWNLOAD_THAILANY
        desktop_dir = os.path.dirname(os.path.dirname(BASE_DIR_THAILANY))
    else:
        download_dir = PASTA_DOWNLOAD_VICTOR
        desktop_dir = os.path.dirname(BASE_DIR_VICTOR)

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.popups": 2,
        "download.prompt_for_download": False
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")

    temp_profile_dir = tempfile.mkdtemp(prefix="chrome_profile_eai_")
    chrome_options.add_argument(f"--user-data-dir={temp_profile_dir}")

    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        pythoncom.CoInitialize()
        driver.get("https://sistema.soc.com.br/WebSoc/")

        wait.until(EC.presence_of_element_located((By.NAME, "usu"))).send_keys(soc_usuario)
        driver.find_element(By.NAME, "senha").send_keys(soc_senha)

        teclado_div = wait.until(EC.presence_of_element_located((By.ID, "pteclado")))
        botoes = teclado_div.find_elements(By.TAG_NAME, "input")
        botoes_map = {btn.get_attribute("value"): btn for btn in botoes}

        for digito in soc_id:
            if digito not in botoes_map:
                raise ValueError(f"Dígito {digito} não encontrado no teclado virtual.")
            botoes_map[digito].click()
            time.sleep(0.25)

        wait.until(EC.element_to_be_clickable((By.ID, "bt_entrar"))).click()

        campo_cod_prog = wait.until(EC.element_to_be_clickable((By.ID, "cod_programa")))
        campo_cod_prog.click()
        campo_cod_prog.clear()
        campo_cod_prog.send_keys("267")
        wait.until(EC.element_to_be_clickable((By.ID, "btn_programa"))).click()

        excel_icon_found = False
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(iframe)
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'excel.png')]"))).click()
                excel_icon_found = True
                driver.switch_to.default_content()
                break
            except Exception:
                driver.switch_to.default_content()
        if not excel_icon_found:
            raise Exception("Ícone do Excel para gerar pedido não encontrado.")

        consultar_btn_found = False
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(iframe)
                WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.botaoT.btn-consultar-pedido"))).click()
                consultar_btn_found = True
                driver.switch_to.default_content()
                break
            except Exception:
                driver.switch_to.default_content()
        if not consultar_btn_found:
            raise Exception("Botão 'Consultar Pedidos' não encontrado.")

        start_time, timeout, poll_interval, download_found = time.time(), 300, 3, False
        while time.time() - start_time < timeout:
            for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.frame(iframe)
                    WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.div-download")))
                    maior_id, div_maior = -1, None
                    divs = driver.find_elements(By.CSS_SELECTOR, "div.div-download")
                    for div in divs:
                        match = re.search(r"(\d+)-download", div.get_attribute("id"))
                        if match and int(match.group(1)) > maior_id:
                            maior_id, div_maior = int(match.group(1)), div
                    if div_maior:
                        div_maior.find_element(By.TAG_NAME, "a").click()
                        download_found = True
                    driver.switch_to.default_content()
                    if download_found: break
                except Exception:
                    driver.switch_to.default_content()

            if download_found: break

            try:
                for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                    try:
                        driver.switch_to.frame(iframe)
                        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//a[@href=\"javascript:doAcao('browse');\"]"))).click()
                        driver.switch_to.default_content()
                        break
                    except Exception:
                        driver.switch_to.default_content()
            except Exception: pass
            time.sleep(poll_interval)

        if not download_found:
            raise Exception("Tempo esgotado. O link de download não foi encontrado.")

        latest_zip = None
        zips_before = set(glob.glob(os.path.join(download_dir, '*.zip')))
        for _ in range(120):
            zips_after = set(glob.glob(os.path.join(download_dir, '*.zip')))
            new_zips = zips_after - zips_before
            if new_zips:
                latest_zip = max(new_zips, key=os.path.getctime)
                break
            time.sleep(0.5)

        if not latest_zip:
            raise FileNotFoundError("❌ Download do arquivo .zip falhou.")

        with zipfile.ZipFile(latest_zip, 'r') as zip_ref:
            xls_files = [f for f in zip_ref.namelist() if f.lower().endswith('.xls')]
            if not xls_files: raise FileNotFoundError("❌ Nenhum arquivo .xls encontrado no .zip.")
            xls_file = xls_files[0]
            zip_ref.extract(xls_file, download_dir)

        xls_path = os.path.join(download_dir, xls_file)
        excel = win32.Dispatch("Excel.Application")
        excel.Visible, excel.DisplayAlerts = False, False
        wb = excel.Workbooks.Open(xls_path)
        xlsx_filename = f"Planilha de pendentes_{int(time.time())}.xlsx"
        dest_path = os.path.join(desktop_dir, xlsx_filename)
        wb.SaveAs(dest_path, FileFormat=51)
        wb.Close(False)
        wb = excel.Workbooks.Open(dest_path)
        ws = wb.Worksheets(1)

        whitelist_empresas = set(app_instance.df.iloc[:, 0].dropna().astype(str).str.upper())
        max_row = ws.UsedRange.Rows.Count
        for i in range(max_row, 4, -1):
            empresa_cell = ws.Cells(i, 1)
            empresa = str(empresa_cell.Value).strip().upper() if empresa_cell.Value else ""
            if empresa not in whitelist_empresas:
                ws.Rows(i).Delete()

        used_range = ws.UsedRange
        used_range.Borders.LineStyle, used_range.Borders.Weight = 1, 2
        used_range.HorizontalAlignment, used_range.VerticalAlignment = -4108, -4108
        used_range.WrapText = True
        ws.Columns.AutoFit(); ws.Rows.AutoFit()

        for r in range(1, ws.UsedRange.Rows.Count + 1):
            if ws.Rows(r).RowHeight < 23: ws.Rows(r).RowHeight = 23
        wb.Save()
        excel.Visible = True
        try:
            app_pwa = Application(backend="uia").connect(process=excel.ProcessID)
            app_pwa.top_window().set_focus().maximize()
        except Exception: pass
        app_instance.root.after(0, app_instance.root.iconify)
        with open(TXT_ULTIMO_GERAR, "w", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%d/%m/%Y"))
        hide_file(TXT_ULTIMO_GERAR)
        if btn_widget: app_instance.root.after(0, btn_widget.pack_forget)

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao gerar planilha: {e}")
    finally:
        driver.quit(); pythoncom.CoUninitialize()
        if status_banner: app_instance.root.after(0, status_banner.destroy)
        shutil.rmtree(temp_profile_dir, ignore_errors=True)

# ================== INTERFACE (UI CUSTOMTKINTER) ==================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("E.A.I. - E-SOCIAL ARTIFICIAL INTELLIGENCE")
        self.root.withdraw()

        for f in [TXT_ULTIMO_ENVIO, TXT_ENVIO, TXT_COBRANCA, TXT_PROCURACOES, TXT_ULTIMO_AJUSTE, TXT_ULTIMO_GERAR]:
            unhide_file(f)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.usuario_atual, self.df = None, pd.DataFrame()
        self.caps_lock_labels, self.password_visible, self.show_envio_cobranca_button = [], False, True

        self.main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True)

        self.check_caps_lock()
        self.tela_login()

    def on_close(self):
        for f in [TXT_ULTIMO_ENVIO, TXT_ENVIO, TXT_COBRANCA, TXT_PROCURACOES, TXT_ULTIMO_AJUSTE, TXT_ULTIMO_GERAR]:
            hide_file(f)
        self.root.destroy()

    def check_caps_lock(self):
        try:
            state = win32api.GetKeyState(win32con.VK_CAPITAL) & 1
            for lbl in self.caps_lock_labels:
                try: lbl.configure(text="CAPS LOCK ATIVADA" if state else "")
                except: pass
        except Exception: pass
        self.root.after(500, self.check_caps_lock)

    def clear_screen(self):
        for w in self.main_container.winfo_children(): w.destroy()
        self.caps_lock_labels = []

    def tela_login(self):
        self.clear_screen()
        self.root.geometry("850x700"); centralizar_janela(self.root, 850, 700); self.root.deiconify()

        container = ctk.CTkFrame(self.main_container, width=600, height=580, corner_radius=25)
        container.pack_propagate(False)
        container.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(container, text="E.A.I.", font=FONT_TITLE).pack(pady=(50, 10))
        ctk.CTkLabel(container, text="E-Social Artificial Intelligence", font=FONT_LABEL).pack(pady=(0, 50))

        self.u_ent = ctk.CTkEntry(container, placeholder_text="Usuário", font=FONT_ENTRY, width=420, height=55)
        self.u_ent.pack(pady=15)

        sf = ctk.CTkFrame(container, fg_color="transparent"); sf.pack(pady=15)
        self.s_ent = ctk.CTkEntry(sf, placeholder_text="Senha", show="*", font=FONT_ENTRY, width=420, height=55)
        self.s_ent.pack(side="left")

        self.eye = ctk.CTkLabel(sf, text="🔒", font=("Segoe UI Emoji", 22), cursor="hand2")
        self.eye.place(relx=0.92, rely=0.5, anchor="center")
        self.eye.bind("<Button-1>", self.toggle_pass)

        cl = ctk.CTkLabel(container, text="", font=(FONT_FAMILY, 13), text_color="#E74C3C")
        cl.pack(pady=5); self.caps_lock_labels.append(cl)

        forgot = ctk.CTkLabel(container, text="Esqueci a minha senha", font=FONT_FORGOT_PASSWORD, text_color="#3498DB", cursor="hand2")
        forgot.pack(pady=15); forgot.bind("<Button-1>", self.abrir_janela_palavra_seguranca)

        ctk.CTkButton(container, text="ENTRAR", font=FONT_BUTTON, width=420, height=65, corner_radius=12, command=self.validar).pack(pady=35)

        self.u_ent.bind("<Return>", lambda e: self.s_ent.focus_set())
        self.s_ent.bind("<Return>", lambda e: self.validar())
        self.u_ent.focus_set()

    def toggle_pass(self, e=None):
        self.password_visible = not self.password_visible
        self.s_ent.configure(show="" if self.password_visible else "*")
        self.eye.configure(text="🔓" if self.password_visible else "🔒")

    def validar(self):
        u, s = self.u_ent.get().strip().lower(), self.s_ent.get().strip()
        users = {"victor": "1001", "thailany": "1001"}
        if u in users and s == users[u]:
            self.usuario_atual = u
            try:
                self.df = carregar_controle(u); self.labels = listar_labels(self.df)
            except Exception as e: messagebox.showerror("Erro", str(e)); return
            messagebox.showinfo("Sucesso", f"Bem-vindo, {u.upper()}!"); self.tela_inicial()
        else: self._show_login_error_dialog()

    def _show_login_error_dialog(self):
        win = ctk.CTkToplevel(self.root); win.title("Erro"); win.geometry("500x350"); centralizar_janela(win, 500, 350)
        ctk.CTkLabel(win, text="DADOS INCORRETOS", font=FONT_TITLE, text_color="#E74C3C").pack(pady=40)
        btns = ctk.CTkFrame(win, fg_color="transparent"); btns.pack(pady=30)
        ctk.CTkButton(btns, text="TENTAR NOVAMENTE", width=180, height=50, command=win.destroy).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="REDEFINIR SENHA", width=180, height=50, fg_color="#F39C12", command=lambda: (win.destroy(), self.abrir_janela_palavra_seguranca())).pack(side="left", padx=10)

    def _has_ok_status(self, label):
        row = get_row_by_label(self.df, label)
        return str(row.iloc[IDX_AA_STATUS]).strip().upper() == "OK" if row is not None else False

    def tela_inicial(self):
        self.clear(); self.root.geometry("1100x900"); centralizar_janela(self.root, 1100, 900); self.root.deiconify()
        ctk.CTkLabel(self.main_container, text="MENU DE OPÇÕES", font=FONT_TITLE, text_color="#2C3E50").pack(pady=50)
        btns_frame = ctk.CTkFrame(self.main_container, fg_color="transparent"); btns_frame.pack(expand=True)

        def add_btn(text, cmd, color="#2980B9", hvr="#21618C"):
            ctk.CTkButton(btns_frame, text=text, font=FONT_BUTTON, width=500, height=65, corner_radius=15, fg_color=color, hover_color=hvr, command=cmd).pack(pady=15)

        if not self._check_monthly(TXT_ULTIMO_GERAR): add_btn("GERAR PENDENTES", self.abrir_gerar)
        if not self._check_monthly(TXT_ULTIMO_ENVIO): add_btn("SALVAR PLANILHAS", lambda: threading.Thread(target=self.salvar_planilhas_outlook, daemon=True).start(), "#27AE60", "#1E8449")
        if not self._check_monthly(TXT_ULTIMO_AJUSTE): add_btn("AJUSTAR PLANILHAS", lambda: threading.Thread(target=self.ajustar_planilhas_thread, daemon=True).start(), "#F39C12", "#D68910")
        if self.show_envio_cobranca_button: add_btn("ENVIO / COBRANÇA", self.tela_envio_cobranca, "#8E44AD", "#71368A")
        add_btn("PROCURAÇÕES", self.tela_procuracoes, "#34495E", "#283747")

        self.status_frame = ctk.CTkFrame(self.main_container, fg_color="transparent"); self.status_frame.pack(fill="x", pady=25)
        rodape = ctk.CTkFrame(self.main_container, height=70, fg_color="transparent"); rodape.pack(side="bottom", fill="x", padx=40, pady=30)
        ctk.CTkButton(rodape, text="SAIR", fg_color="#E74C3C", width=140, height=50, corner_radius=10, command=self.confirmar_saida).pack(side="left")
        h = ctk.CTkLabel(rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, text_color="#3498DB", cursor="hand2")
        h.pack(side="right"); h.bind("<Button-1>", abrir_link_ajuda)

    def _check_monthly(self, f):
        if os.path.isfile(f):
            try:
                with open(f, "r") as file:
                    d = datetime.strptime(file.read().strip(), "%d/%m/%Y")
                    return d.month == datetime.now().month and d.year == datetime.now().year
            except: pass
        return False

    def confirmar_saida(self):
        if messagebox.askyesno("Sair", f"{self.usuario_atual.upper()}, deseja realmente sair?"): self.tela_login()

    def abrir_gerar(self):
        win = ctk.CTkToplevel(self.root); win.title("SOC"); win.geometry("650x600"); centralizar_janela(win, 650, 600); win.after(200, lambda: win.focus())
        f = ctk.CTkFrame(win, fg_color="transparent"); f.pack(expand=True, padx=50, pady=50, fill="both")
        ctk.CTkLabel(f, text="AUTENTICAÇÃO SOC", font=FONT_TITLE).pack(pady=(0, 40))
        u = ctk.CTkEntry(f, placeholder_text="Usuário SOC", font=FONT_ENTRY, width=450, height=50); u.pack(pady=12)
        s = ctk.CTkEntry(f, placeholder_text="Senha SOC", show="*", font=FONT_ENTRY, width=450, height=50); s.pack(pady=12)
        i = ctk.CTkEntry(f, placeholder_text="ID SOC", font=FONT_ENTRY, width=450, height=50); i.pack(pady=12)
        def go():
            if not (u.get() and s.get() and i.get()): return
            win.destroy(); b = self.mostrar_banner("PROCESSO DE GERAÇÃO INICIADO...")
            threading.Thread(target=executar_gerar_planilha, args=(self.usuario_atual, u.get(), s.get(), i.get(), self, None, b), daemon=True).start()
        ctk.CTkButton(f, text="GERAR PLANILHA", font=FONT_BUTTON, width=450, height=60, corner_radius=12, command=go).pack(pady=45)

    def mostrar_banner(self, text, color="#D6EAF8", t_color="#1B4F72"):
        b = ctk.CTkFrame(self.status_frame, fg_color=color, corner_radius=15); b.pack(fill="x", pady=15, padx=50)
        l = ctk.CTkLabel(b, text=text, font=FONT_BOLD, text_color=t_color); l.pack(pady=25); b.label = l; return b

    def salvar_planilhas_outlook(self):
        banner = self.mostrar_banner("ANALISANDO E-MAILS...", "#D5F5E3", "#145A32")
        try:
            pythoncom.CoInitialize()
            acc_mail = EMAIL_PLANILHAS_THAILANY if self.usuario_atual == "thailany" else EMAIL_CONTA_PADRAO_VICTOR
            dest = PASTA_PLANILHAS_THAILANY if self.usuario_atual == "thailany" else PASTA_PLANILHAS_VICTOR
            outlook = win32.Dispatch("Outlook.Application"); ns = outlook.GetNamespace("MAPI"); inbox = None
            for acc in ns.Accounts:
                if acc.SmtpAddress.lower() == acc_mail.lower(): inbox = acc.DeliveryStore.GetDefaultFolder(6); break
            if not inbox: raise Exception(f"Conta {acc_mail} não encontrada.")
            items = inbox.Items; items.Sort("[ReceivedTime]", True); os.makedirs(dest, exist_ok=True)
            rel = [i for i in items if "listagem de funcionarios" in normalize_key(i.Subject or "") and i.Attachments.Count > 0]
            total, baixar_list = 0, []
            ts = os.path.join(dest, "temp_scan"); os.makedirs(ts, exist_ok=True)
            for i in rel:
                for a in i.Attachments:
                    if a.FileName.lower().endswith(".zip"):
                        tp = os.path.join(ts, a.FileName); a.SaveAsFile(tp)
                        with zipfile.ZipFile(tp, 'r') as zr: total += len([f for f in zr.namelist() if f.lower().endswith(".xlsx")])
                        baixar_list.append((i, a)); os.remove(tp)
            shutil.rmtree(ts); self.root.after(0, lambda: banner.label.configure(text=f"PLANILHAS SALVAS: 0 DE {total}"))
            count = 0
            for itm, att in baixar_list:
                nb = re.sub(r'[<>:"/\\|?*]', '', itm.Subject).strip(); cz = os.path.join(dest, att.FileName); att.SaveAsFile(cz)
                with zipfile.ZipFile(cz, 'r') as zr:
                    for zi in [f for f in zr.namelist() if f.lower().endswith(".xlsx")]:
                        cf = os.path.join(dest, f"{nb}.xlsx"); c = 1
                        while os.path.exists(cf): cf = os.path.join(dest, f"{nb}_{c}.xlsx"); c += 1
                        with zr.open(zi) as src, open(cf, 'wb') as tgt: shutil.copyfileobj(src, tgt)
                        count += 1; self.root.after(0, lambda c=count: banner.label.configure(text=f"PLANILHAS SALVAS: {c} DE {total}"))
                os.remove(cz)
            with open(TXT_ULTIMO_ENVIO, "w") as f: f.write(datetime.now().strftime("%d/%m/%Y"))
            hide_file(TXT_ULTIMO_ENVIO); messagebox.showinfo("Sucesso", "Concluído!")
        except Exception as e: messagebox.showerror("Erro", str(e))
        finally: banner.destroy(); pythoncom.CoUninitialize()

    def ajustar_planilhas_thread(self):
        banner = self.mostrar_banner("IDENTIFICANDO PLANILHAS...", "#FAD7A0", "#7E5109")
        try:
            pythoncom.CoInitialize(); pasta = PASTA_PLANILHAS_THAILANY if self.usuario_atual == "thailany" else PASTA_PLANILHAS_VICTOR
            all_files = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]
            if not all_files: messagebox.showinfo("Aviso", "Vazio."); return
            ex = win32.Dispatch("Excel.Application"); ex.Visible, ex.DisplayAlerts = False, False; aj = []
            for f in all_files:
                try:
                    wb = ex.Workbooks.Open(os.path.join(pasta, f), ReadOnly=True, Password=""); wb.Close(); aj.append(f)
                except: continue
            ex.Quit(); total = len(aj); self.root.after(0, lambda: banner.label.configure(text=f"PLANILHAS AJUSTADAS: 0 DE {total}"))
            hj, cnt = datetime.now(), 0
            for f in aj:
                p = os.path.join(pasta, f); wb = openpyxl.load_workbook(p)
                for ws in wb.worksheets:
                    ws.delete_rows(2, 1); mrg = list(ws.merged_cells.ranges); delr = []
                    for r in range(2, ws.max_row+1):
                        if not any(r >= m.min_row and r <= m.max_row for m in mrg):
                            st, dt = ws.cell(r, 6).value, ws.cell(r, 7).value
                            if str(st).strip().lower() == "pendente":
                                try:
                                    if not isinstance(dt, datetime): dt = datetime.strptime(str(dt), "%d/%m/%Y")
                                    if dt.month == hj.month and dt.year == hj.year: delr.append(r)
                                except: pass
                    for r in sorted(delr, reverse=True): ws.delete_rows(r, 1)
                    thin = Side(border_style="thin", color="000000"); border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    for r in range(1, ws.max_row+1):
                        for c in range(1, ws.max_column+1):
                            if not any(r >= m.min_row and r <= m.max_row and c >= m.min_col and c <= m.max_col for m in mrg):
                                cell = ws.cell(r, c); cell.border = border; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                wb.save(p); cnt += 1; self.root.after(0, lambda c=cnt: banner.label.configure(text=f"PLANILHAS AJUSTADAS: {c} DE {total}"))
            ex = win32.Dispatch("Excel.Application"); ex.Visible, ex.DisplayAlerts = False, False
            for f in aj:
                try:
                    wb = ex.Workbooks.Open(os.path.join(pasta, f))
                    for ws in wb.Worksheets: ws.Columns.AutoFit(); ws.Rows.AutoFit(); ws.Protect("spoc")
                    wb.Protect("spoc"); wb.Save(); wb.Close()
                except: pass
            ex.Quit()
            with open(TXT_ULTIMO_AJUSTE, "w") as fl: fl.write(hj.strftime("%d/%m/%Y"))
            hide_file(TXT_ULTIMO_AJUSTE); messagebox.showinfo("Sucesso", "Concluído.")
        except Exception as e: messagebox.showerror("Erro", str(e))
        finally: banner.destroy(); pythoncom.CoUninitialize()

    def tela_envio_cobranca(self):
        win = ctk.CTkToplevel(self.root); win.title("E.A.I. - ENVIO / COBRANÇA"); win.geometry("650x500"); centralizar_janela(win, 650, 500)
        ctk.CTkLabel(win, text="ESCOLHA A AÇÃO", font=FONT_TITLE).pack(pady=50)
        ctk.CTkButton(win, text="ENVIO", font=FONT_BUTTON, width=450, height=65, command=self.tela_empresas_envio).pack(pady=15)
        ctk.CTkButton(win, text="COBRANÇA", font=FONT_BUTTON, width=450, height=65, fg_color="#8E44AD", command=self.tela_empresas_cobranca).pack(pady=15)
        ctk.CTkButton(win, text="VOLTAR", font=FONT_BUTTON, fg_color="#E74C3C", command=win.destroy).pack(pady=40)

    def tela_empresas_envio(self):
        win = ctk.CTkToplevel(self.root); win.title("Envio"); win.state('zoomed')
        top = ctk.CTkFrame(win); top.pack(fill="x", pady=20)
        ctk.CTkLabel(top, text="ENVIO DA LISTAGEM DE FUNCIONÁRIOS", font=FONT_TITLE).pack()
        regs, hj = carregar_txt(TXT_ENVIO), datetime.now()
        bl = [l for l in self.labels if not (l in regs and datetime.strptime(regs[l], "%d/%m/%Y").month == hj.month) and not self._has_ok_status(l)]
        grid = ctk.CTkScrollableFrame(win, orientation="horizontal"); grid.pack(fill="both", expand=True)
        def send_all():
            for b in [w for w in grid.winfo_children() if isinstance(w, ctk.CTkButton)]:
                self._exec_env(b.cget("text"), b); time.sleep(0.5)
        ctk.CTkButton(top, text="ENVIAR TODOS", font=FONT_BUTTON, command=lambda: threading.Thread(target=send_all).start()).pack()
        for i, lbl in enumerate(bl):
            b = ctk.CTkButton(grid, text=lbl, width=320, height=80, font=FONT_BUTTON)
            b.grid(row=i%5, column=i//5, padx=10, pady=10)
            b.configure(command=lambda l=lbl, bt=b: threading.Thread(target=self._exec_env, args=(l, bt)).start())

    def _exec_env(self, l, bt):
        s, e = enviar_email_label(l, self.df, self.usuario_atual)
        if s: self.root.after(0, bt.destroy)

    def tela_empresas_cobranca(self):
        win = ctk.CTkToplevel(self.root); win.title("Cobrança"); win.state('zoomed')
        top = ctk.CTkFrame(win); top.pack(fill="x", pady=20)
        ctk.CTkLabel(top, text="COBRANÇA DA LISTAGEM DE FUNCIONÁRIOS", font=FONT_TITLE).pack()
        regs, hj = carregar_txt(TXT_COBRANCA), datetime.now().date()
        bl = [l for l in self.labels if l not in EXCLUIR_NA_COBRANCA and not (l in regs and datetime.strptime(regs[l], "%d/%m/%Y").date() == hj) and not self._has_ok_status(l)]
        grid = ctk.CTkScrollableFrame(win, orientation="horizontal"); grid.pack(fill="both", expand=True)
        def cob_all():
            for b in [w for w in grid.winfo_children() if isinstance(w, ctk.CTkButton)]:
                self._exec_cob(b.cget("text"), b); time.sleep(0.5)
        ctk.CTkButton(top, text="COBRAR TODOS", font=FONT_BUTTON, command=lambda: threading.Thread(target=cob_all).start()).pack()
        for i, lbl in enumerate(bl):
            b = ctk.CTkButton(grid, text=lbl, width=320, height=80, font=FONT_BUTTON, fg_color="#8E44AD")
            b.grid(row=i%5, column=i//5, padx=10, pady=10)
            b.configure(command=lambda l=lbl, bt=b: threading.Thread(target=self._exec_cob, args=(l, bt)).start())

    def _exec_cob(self, l, bt):
        s, e = cobrar_email_label(l, self.df, self.usuario_atual)
        if s: self.root.after(0, bt.destroy)

    def tela_procuracoes(self):
        win = ctk.CTkToplevel(self.root); win.title("Procurações"); win.state('zoomed')
        top = ctk.CTkFrame(win); top.pack(fill="x", pady=20)
        ctk.CTkLabel(top, text="PROCURAÇÕES VENCIDAS", font=FONT_TITLE).pack()
        hj, regs = datetime.now(), carregar_txt(TXT_PROCURACOES); pl = []
        for _, r in self.df.iloc[2:].iterrows():
            n, v, st = str(r.iloc[0]), r.iloc[IDX_AC_VALIDADE], str(r.iloc[IDX_F_BLOQUEADA])
            if not pd.isna(n) and "BLOQUEADA" not in st.upper() and not (n in regs and regs[n] == hj.strftime("%d/%m/%Y")):
                try:
                    if not isinstance(v, datetime): v = datetime.strptime(str(v), "%d/%m/%Y")
                    if v < hj: pl.append(n)
                except: pass
        grid = ctk.CTkScrollableFrame(win, orientation="horizontal"); grid.pack(fill="both", expand=True)
        def proc_all():
            for b in [w for w in grid.winfo_children() if isinstance(w, ctk.CTkButton)]:
                self._exec_proc(b.cget("text"), b); time.sleep(0.5)
        ctk.CTkButton(top, text="ENVIAR TODOS", font=FONT_BUTTON, command=lambda: threading.Thread(target=proc_all).start()).pack()
        for i, n in enumerate(pl):
            b = ctk.CTkButton(grid, text=n, width=340, height=85, font=FONT_BUTTON, fg_color="#34495E")
            b.grid(row=i%5, column=i//5, padx=10, pady=10)
            b.configure(command=lambda nm=n, bt=b: threading.Thread(target=self._exec_proc, args=(nm, bt)).start())

    def _exec_proc(self, n, bt):
        s, e = enviar_email_procuracao(n, self.df, self.usuario_atual)
        if s: self.root.after(0, bt.destroy)

    def abrir_janela_palavra_seguranca(self, event=None):
        win = ctk.CTkToplevel(self.root); win.title("Recuperação"); win.geometry("600x500"); centralizar_janela(win, 650, 500); win.after(200, lambda: win.focus())
        container = ctk.CTkFrame(win, corner_radius=15); container.pack(expand=True, padx=40, pady=40, fill="both")
        ctk.CTkLabel(container, text="Recuperação de Senha", font=FONT_TITLE).pack(pady=20)
        u_ent = ctk.CTkEntry(container, placeholder_text="Usuário", width=400, height=50); u_ent.pack(pady=10)
        p_ent = ctk.CTkEntry(container, placeholder_text="Palavra de Segurança", show="*", width=400, height=50); p_ent.pack(pady=10)
        tries = [3]
        lbl_t = ctk.CTkLabel(container, text=f"TENTATIVAS: {tries[0]}", text_color="red"); lbl_t.pack()
        def check():
            user, key = u_ent.get().strip().lower(), p_ent.get().strip().lower()
            secrets = {"victor": "creatina", "thailany": "spoc"}
            if user in secrets and key == secrets[user]:
                messagebox.showinfo("Senha", f"Olá {user.upper()}, sua senha é: 1001"); win.destroy()
            else:
                tries[0] -= 1; lbl_t.configure(text=f"TENTATIVAS: {tries[0]}")
                if tries[0] <= 0:
                    messagebox.showerror("Erro", "Tentativas esgotadas."); win.destroy()
                    webbrowser.open("https://wa.me/551221239207")
                else: messagebox.showerror("Erro", "Incorreto.")
        ctk.CTkButton(container, text="RECUPERAR", font=FONT_BUTTON, width=400, height=60, command=check).pack(pady=30)

def mostrar_intro(root, callback):
    if getattr(sys, 'frozen', False): bp = sys._MEIPASS
    else: bp = os.path.dirname(__file__)
    cp = os.path.join(bp, "patinho.gif")
    if not os.path.isfile(cp): cp = CAMINHO_PATINHO_2
    intro = tk.Toplevel(root); intro.attributes("-fullscreen", True); intro.configure(bg="white")
    if os.path.isfile(cp):
        try:
            img = tk.PhotoImage(file=cp)
            lbl = tk.Label(intro, image=img, bg="white"); lbl.image = img; lbl.pack(expand=True)
        except Exception: tk.Label(intro, text="CARREGANDO...", font=FONT_TITLE, bg="white").pack(expand=True)
    else: tk.Label(intro, text="CARREGANDO...", font=FONT_TITLE, bg="white").pack(expand=True)
    intro.after(2000, lambda: (intro.destroy(), callback()))

if __name__ == "__main__":
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    root = ctk.CTk(); app = App(root); mostrar_intro(root, lambda: root.deiconify()); root.mainloop()
