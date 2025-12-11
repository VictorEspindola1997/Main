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
FONT_NORMAL = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10)
FONT_ENTRY = (FONT_FAMILY, 10)
FONT_LABEL = (FONT_FAMILY, 11)
FONT_BUTTON = (FONT_FAMILY, 10)
FONT_FORGOT_PASSWORD = (FONT_FAMILY, 11, "underline")
FONT_TITLE = (FONT_FAMILY, 14)

MAX_LINHAS_COLUNA = 6
IDX_AJ_LABEL = 35
IDX_AK_ANEXOS = 36
IDX_AL_ASSUNTO = 37
IDX_EMAIL_TO = 32
IDX_EMAIL_CC = 33
IDX_F_BLOQUEADA = 5  # Coluna F para verificar "BLOQUEADA"
IDX_AA_STATUS = 26  # Coluna AA para verificar "OK"
IDX_AC_VALIDADE = 28  # Coluna AC para validade da procuração
IDX_AE_PROCURACAO_TO = 30  # Coluna AE para email de procuração
IDX_AF_PROCURACAO_CC = 31  # Coluna AF para email de procuração em cópia
EXCLUIR_NA_COBRANCA = {"GRUPO OSWALDO CRUZ"}

# ================== FUNÇÃO GERAR PENDENTES ==================


def executar_gerar_planilha(app_usuario, soc_usuario, soc_senha, soc_id, app_instance, btn_widget, status_banner):
    if app_usuario == "thailany":
        download_dir = PASTA_DOWNLOAD_THAILANY
        # Extrai o caminho 'D:\Desktop' de 'D:\Desktop\Oswaldo Cruz Gestão
        # Ocupacional\E-Social'
        desktop_dir = os.path.dirname(os.path.dirname(BASE_DIR_THAILANY))
    else:
        download_dir = PASTA_DOWNLOAD_VICTOR
        # Extrai 'C:\Users\TI Medicina\Desktop' de 'C:\Users\TI
        # Medicina\Desktop\ESOCIAL - PYTHON'
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

    # Usa um perfil de usuário temporário e limpo para garantir consistência
    temp_profile_dir = tempfile.mkdtemp(prefix="chrome_profile_eai_")
    chrome_options.add_argument(f"--user-data-dir={temp_profile_dir}")
    print(f"✨ Usando perfil de usuário temporário: {temp_profile_dir}")

    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 20)

    try:
        pythoncom.CoInitialize()
        driver.get("https://sistema.soc.com.br/WebSoc/")

        # Espera o campo de usuário aparecer e preenche
        wait.until(EC.presence_of_element_located(
            (By.NAME, "usu"))).send_keys(soc_usuario)
        driver.find_element(By.NAME, "senha").send_keys(soc_senha)

        # Aguarda carregamento do teclado virtual
        teclado_div = wait.until(
            EC.presence_of_element_located((By.ID, "pteclado")))
        botoes = teclado_div.find_elements(By.TAG_NAME, "input")

        # Dicionário valor->elemento para clicar rapidamente
        botoes_map = {btn.get_attribute("value"): btn for btn in botoes}

        for digito in soc_id:
            if digito not in botoes_map:
                raise ValueError(
                    f"Dígito {digito} não encontrado no teclado virtual.")
            botoes_map[digito].click()
            time.sleep(0.25)  # Pausa reduzida para performance

        # Clicar no botão Entrar
        wait.until(EC.element_to_be_clickable((By.ID, "bt_entrar"))).click()

        # Agora começa processo de digitar 267 no campo 'cod_programa'
        campo_cod_prog = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "cod_programa")))
        campo_cod_prog.click()
        campo_cod_prog.clear()
        campo_cod_prog.send_keys("267")

        wait.until(EC.element_to_be_clickable((By.ID, "btn_programa"))).click()

        # Achar e clicar no ícone excel dentro do iframe (otimizado)
        excel_icon_found = False
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(iframe)
                # Usa uma espera curta (1s) para cada iframe, tornando a busca rápida
                WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//img[contains(@src, 'excel.png')]"))
                ).click()
                excel_icon_found = True
                driver.switch_to.default_content()
                break
            except Exception:
                driver.switch_to.default_content()
        if not excel_icon_found:
            raise Exception("Ícone do Excel para gerar pedido não encontrado.")

        # Clicar em Consultar Pedidos no iframe (otimizado)
        consultar_btn_found = False
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.frame(iframe)
                # Usa uma espera curta (1s) para cada iframe
                WebDriverWait(driver, 1).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a.botaoT.btn-consultar-pedido"))
                ).click()
                consultar_btn_found = True
                driver.switch_to.default_content()
                break
            except Exception:
                driver.switch_to.default_content()
        if not consultar_btn_found:
            raise Exception("Botão 'Consultar Pedidos' não encontrado.")

        start_time = time.time()
        timeout = 300  # 5 minutos
        poll_interval = 3  # segundos (Otimizado para verificação rápida)
        download_found = False
        print("🕒 Buscando o link de download (até 5 minutos)...")

        while time.time() - start_time < timeout:
            # 1. Tenta encontrar o link de download em todos os iframes
            for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.frame(iframe)
                    # Procura por um link de download com um timeout curto
                    WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "div.div-download"))
                    )

                    # Se encontrou, pega o mais recente e clica
                    maior_id = -1
                    div_maior = None
                    divs = driver.find_elements(
                        By.CSS_SELECTOR, "div.div-download")
                    for div in divs:
                        div_id = div.get_attribute("id")
                        match = re.search(r"(\d+)-download", div_id)
                        if match:
                            id_num = int(match.group(1))
                            if id_num > maior_id:
                                maior_id = id_num
                                div_maior = div

                    if div_maior:
                        div_maior.find_element(By.TAG_NAME, "a").click()
                        download_found = True

                    driver.switch_to.default_content()
                    if download_found:
                        break
                except Exception:
                    driver.switch_to.default_content()

            if download_found:
                break

            # 2. Se não encontrou, clica no botão "Procurar" para atualizar e espera
            print(
                f"🕒 Link não encontrado. Clicando em 'Procurar' e aguardando {poll_interval} segundos...")
            try:
                procurar_found = False
                for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                    try:
                        driver.switch_to.frame(iframe)
                        WebDriverWait(driver, 2).until(
                            EC.element_to_be_clickable(
                                (By.XPATH, "//a[@href=\"javascript:doAcao('browse');\"]"))
                        ).click()
                        procurar_found = True
                        driver.switch_to.default_content()
                        break  # Exit iframe loop once clicked
                    except Exception:
                        driver.switch_to.default_content()

                if not procurar_found:
                    print(
                        "⚠️ Ícone 'Procurar' não foi encontrado para atualizar a lista.")

            except Exception as e:
                print(f"⚠️ Erro ao clicar em 'Procurar': {e}")
                driver.switch_to.default_content()

            time.sleep(poll_interval)

        if not download_found:
            raise Exception(
                "Tempo esgotado. O link de download não foi encontrado.")

        print("🕒 Aguardando o arquivo .zip ser baixado...")
        latest_zip = None
        zips_before = set(glob.glob(os.path.join(download_dir, '*.zip')))

        for _ in range(120):  # Espera até 60 segundos pelo novo .zip
            zips_after = set(glob.glob(os.path.join(download_dir, '*.zip')))
            new_zips = zips_after - zips_before
            if new_zips:
                latest_zip = max(new_zips, key=os.path.getctime)
                break
            time.sleep(0.5)

        if not latest_zip:
            raise FileNotFoundError(
                "❌ Download do arquivo .zip falhou ou demorou demais.")

        print(f"📦 Extraindo de: {latest_zip}")
        with zipfile.ZipFile(latest_zip, 'r') as zip_ref:
            xls_files = [f for f in zip_ref.namelist() if f.lower().endswith('.xls')]
            if not xls_files:
                raise FileNotFoundError(
                    "❌ Nenhum arquivo .xls encontrado no .zip.")
            xls_file = xls_files[0]
            zip_ref.extract(xls_file, download_dir)

        xls_path = os.path.join(download_dir, xls_file)

        for _ in range(10):  # Espera o arquivo ser extraído
            if os.path.exists(xls_path):
                break
            time.sleep(0.5)
        if not os.path.exists(xls_path):
            raise FileNotFoundError(
                f"❌ Arquivo .xls extraído não foi encontrado: {xls_path}")

        # --- Processamento otimizado com uma única instância do Excel ---
        print("🚀 Otimizando a planilha...")
        excel = None
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False  # Mantém oculto para performance
            excel.DisplayAlerts = False

            # 1. Abre .xls e salva como .xlsx, depois abre o .xlsx
            wb = excel.Workbooks.Open(xls_path)
            xlsx_filename = "Planilha de pendentes.xlsx"
            dest_path = os.path.join(desktop_dir, xlsx_filename)
            counter = 1
            while os.path.exists(dest_path):
                xlsx_filename = f"Planilha de pendentes_{counter}.xlsx"
                dest_path = os.path.join(desktop_dir, xlsx_filename)
                counter += 1
            wb.SaveAs(dest_path, FileFormat=51)
            wb.Close(False)
            wb = excel.Workbooks.Open(dest_path)
            ws = wb.Worksheets(1)
            ws.Activate()

            # 2. Filtra as linhas de forma eficiente com whitelist dinâmica
            # GARANTIA: A whitelist é sempre criada dinamicamente a partir da Coluna A
            # da planilha 'Controle do E-Social.xlsx'. Apenas as empresas listadas
            # nesta coluna serão mantidas na planilha de pendentes gerada.
            df_controle = app_instance.df
            whitelist_empresas = set(
                df_controle.iloc[:, 0].dropna().astype(str).str.upper())
            print(
                f"✅ Whitelist dinâmica criada com {len(whitelist_empresas)} empresas.")

            max_row = ws.UsedRange.Rows.Count
            for i in range(max_row, 4, -1):  # Itera de baixo para cima
                empresa_cell = ws.Cells(i, 1)
                empresa = str(empresa_cell.Value).strip().upper(
                ) if empresa_cell.Value else ""
                row_range = ws.Rows(i)

                # Verifica se a empresa está na whitelist para decidir se mantém ou apaga
                if empresa not in whitelist_empresas:
                    row_range.Delete()

            # 3. Formatação e AutoFit
            used_range = ws.UsedRange
            used_range.Borders.LineStyle = 1
            used_range.Borders.Weight = 2
            used_range.HorizontalAlignment = -4108  # xlCenter
            used_range.VerticalAlignment = -4108  # xlCenter
            used_range.WrapText = True
            ws.Columns.AutoFit()
            ws.Rows.AutoFit()

            for r in range(1, ws.UsedRange.Rows.Count + 1):
                if ws.Rows(r).RowHeight < 23:
                    ws.Rows(r).RowHeight = 23

            ws.Range("A1").Select()
            wb.Save()

            # 4. Exibe a planilha para o usuário
            excel.Visible = True
            excel.DisplayAlerts = True
            excel.ActiveWindow.Zoom = 100
            excel.ActiveWindow.ScrollRow = 1
            excel.ActiveWindow.ScrollColumn = 1

            # 5. Traz para o primeiro plano e maximiza a janela do Excel (com tentativas)
            start_time_excel = time.time()
            excel_maximized = False
            while time.time() - start_time_excel < 10:  # Tenta por 10 segundos
                try:
                    # Conecta usando o PID do processo Excel recém-criado
                    app = Application(backend="uia").connect(
                        process=excel.ProcessID)
                    window = app.top_window()
                    if window.is_visible():
                        window.set_focus().set_foreground().maximize()
                        excel_maximized = True
                        print("✅ Janela do Excel maximizada com sucesso.")
                        break
                except Exception:
                    # Ignora erros de conexão temporários enquanto o Excel inicia
                    time.sleep(0.25)  # Verifica 4x por segundo para máxima responsividade

            if not excel_maximized:
                print(
                    "⚠️ Aviso: Não foi possível focar e maximizar a janela do Excel após 10 segundos.")

            # Minimiza a janela principal para dar foco total ao Excel
            app_instance.root.after(0, app_instance.root.iconify)

        except Exception as e:
            if excel:
                excel.Quit()
            messagebox.showerror("Erro", f"Falha ao processar a planilha: {e}")
            raise e

        # Registrar a geração no arquivo TXT_ULTIMO_GERAR
        with open(TXT_ULTIMO_GERAR, "w", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%d/%m/%Y"))
        hide_file(TXT_ULTIMO_GERAR)
        app_instance.root.after(0, btn_widget.pack_forget)

    except Exception as e:
        print(f"❌ Erro: {e}")
        messagebox.showerror("Erro", f"Falha ao gerar planilha: {e}")
    finally:
        print("🟢 Processo finalizado. Fechando navegador em segundo plano.")
        driver.quit()
        pythoncom.CoUninitialize()
        if status_banner:
            app_instance.root.after(0, status_banner.destroy)
        # Limpa o diretório de perfil temporário
        try:
            shutil.rmtree(temp_profile_dir)
            print(
                f"🧹 Perfil temporário '{temp_profile_dir}' limpo com sucesso.")
        except Exception as e:
            print(f"⚠️ Aviso: Falha ao limpar o perfil temporário: {e}")


class GerarPlanilhaWindow(tk.Toplevel):
    def __init__(self, parent, usuario_atual, app_instance, btn_widget):
        super().__init__(parent)
        self.parent = parent
        self.usuario_atual = usuario_atual
        self.app_instance = app_instance
        self.btn_widget = btn_widget
        app_instance.gerar_planilha_window = self


        self.title("E.A.I. - GERAÇÃO DA PLANILHA DE PENDENTES")
        centralizar_janela(self, 600, 500)
        self.resizable(False, False)

        self.senha_visible = tk.BooleanVar(value=False)
        self.id_visible = tk.BooleanVar(value=False)

        frame = tk.Frame(self)
        frame.pack(expand=True, pady=20, padx=20)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=0)
        frame.grid_columnconfigure(2, weight=2)
        frame.grid_columnconfigure(3, weight=0)
        frame.grid_columnconfigure(4, weight=1)

        tk.Label(frame, text="USUÁRIO SOC:", font=FONT_LABEL).grid(row=0, column=1, sticky="e", padx=5, pady=5)
        self.entry_usuario_frame = create_aligned_entry(frame, font=FONT_ENTRY)
        self.entry_usuario_frame.grid(row=0, column=2, sticky="we", padx=5, pady=5)
        self.entry_usuario = self.entry_usuario_frame.entry

        tk.Label(frame, text="SENHA SOC:", font=FONT_LABEL).grid(row=1, column=1, sticky="e", padx=5, pady=5)
        self.entry_senha_frame = create_aligned_entry(frame, font=FONT_ENTRY, show='*')
        self.entry_senha_frame.grid(row=1, column=2, sticky="we", padx=5, pady=5)
        self.entry_senha = self.entry_senha_frame.entry
        senha_icon = tk.Label(frame, text="🔒", font=FONT_NORMAL, cursor="hand2")
        senha_icon.grid(row=1, column=3, sticky="w", padx=5, pady=5)
        senha_icon.bind("<Button-1>", lambda e: self.toggle_visibility(self.entry_senha, self.senha_visible, senha_icon))

        vcmd = self.register(self.apenas_numeros)
        tk.Label(frame, text="ID SOC:", font=FONT_LABEL).grid(row=2, column=1, sticky="e", padx=5, pady=5)
        self.entry_id_frame = create_aligned_entry(frame, font=FONT_ENTRY, show='*', validate="key", validatecommand=(vcmd, '%S'))
        self.entry_id_frame.grid(row=2, column=2, sticky="we", padx=5, pady=5)
        self.entry_id = self.entry_id_frame.entry
        id_icon = tk.Label(frame, text="🔒", font=FONT_NORMAL, cursor="hand2")
        id_icon.grid(row=2, column=3, sticky="w", padx=5, pady=5)
        id_icon.bind("<Button-1>", lambda e: self.toggle_visibility(self.entry_id, self.id_visible, id_icon))

        self.caps_lock_label_gerar = tk.Label(frame, text="", font=FONT_BOLD)
        self.caps_lock_label_gerar.grid(row=3, column=1, columnspan=2, sticky="we", padx=5, pady=(5, 0))
        parent_bg_gerar = frame.cget("background")
        self.caps_lock_tuple = (self.caps_lock_label_gerar, parent_bg_gerar)
        self.app_instance.caps_lock_labels.append(self.caps_lock_tuple)

        self.btn_gerar = tk.Button(frame, text="GERAR PLANILHA DE PENDENTES", font=FONT_BUTTON, height=2, command=self.executar)
        self.btn_gerar.grid(row=4, column=1, columnspan=2, pady=15, sticky="we")

        self.entry_usuario.bind("<Return>", lambda e: self.entry_senha.focus_set())
        self.entry_senha.bind("<Return>", lambda e: self.entry_id.focus_set())
        self.entry_id.bind("<Return>", lambda e: self.executar())
        self.btn_gerar.bind("<Return>", lambda e: self.executar())
        self.entry_usuario.focus_set()

        rodape = tk.Frame(self)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        self.protocol("WM_DELETE_WINDOW", self.on_win_destroy)

        voltar_btn = tk.Button(rodape, text="VOLTAR", font=FONT_BUTTON, width=10, command=self.on_win_destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

    def toggle_visibility(self, entry_widget, var, label):
        is_currently_hidden = not var.get()
        if is_currently_hidden:
            entry_widget.config(show="")
            label.config(text="🔓")
        else:
            entry_widget.config(show="*")
            label.config(text="🔒")
        var.set(is_currently_hidden)

    def apenas_numeros(self, char):
        return char.isdigit()

    def executar(self, event=None):
        soc_usuario = self.entry_usuario.get().strip()
        soc_senha = self.entry_senha.get().strip()
        soc_id = self.entry_id.get().strip()

        if not soc_usuario or not soc_senha or not soc_id:
            messagebox.showwarning("E.A.I. - ATENÇÃO", "POR FAVOR, PREENCHA TODOS OS CAMPOS (USUÁRIO, SENHA E ID).", parent=self)
            return
        if not soc_id.isdigit():
            messagebox.showwarning("E.A.I. - ATENÇÃO", "O ID DEVE CONTER APENAS NÚMEROS.", parent=self)
            return

        self.destroy()
        status_banner = self.app_instance.mostrar_status_geracao()
        threading.Thread(target=executar_gerar_planilha, args=(
            self.usuario_atual, soc_usuario, soc_senha, soc_id, self.app_instance, self.btn_widget, status_banner), daemon=True).start()

    def on_win_destroy(self):
        self.app_instance.caps_lock_labels.remove(self.caps_lock_tuple)
        self.app_instance.gerar_planilha_window = None
        self.destroy()


def centralizar_janela(janela, largura=500, altura=500):
    janela.update_idletasks()
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela // 2) - (largura // 2)
    y = (altura_tela // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

# ================== FUNÇÃO SALVAR PLANILHAS ==================


# ================== CHECA SE BOTÃO DEVE APARECER ==================


# ================== UTILITÁRIOS ==================


def create_aligned_entry(parent, **kwargs):
    """
    Cria um widget tk.Entry com o texto alinhado verticalmente no centro,
    envolvendo-o em um tk.Frame.

    Retorna o tk.Frame container. O widget tk.Entry real está anexado
    como um atributo `.entry` no frame retornado.
    """
    # A moldura externa recebe a borda e atua como o campo de entrada visual
    container_frame = tk.Frame(
        parent, relief="sunken", borderwidth=1, bg="white")

    # O widget Entry real não tem borda
    entry = tk.Entry(container_frame, relief="flat", bg="white", **kwargs)

    # Empacota a entrada com preenchimento vertical para centralizá-la
    entry.pack(fill="x", padx=2, pady=1)

    # Anexa a entrada ao frame para fácil acesso
    container_frame.entry = entry

    return container_frame


def centralizar(janela):
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def normalize_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()


def safe_email_field(val) -> str:
    if val is None:
        return ""
    try:
        if isinstance(val, float) and pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() in {"nan", "none", "null", ""}:
        return ""
    return s


def limpar_caminho(p: str) -> str:
    if p is None:
        return ""
    s = str(p).strip().strip('"').strip("'")
    return s


def extrair_nome_empresa_do_arquivo(caminho: str) -> str:
    nome = os.path.basename(caminho)
    nome = re.sub(r"(?i)listagem de funcionarios\s*-\s*", "", nome)
    nome = re.sub(r"\.(xlsx|xlsm|xls)$", "", nome, flags=re.I)
    return nome.strip()

# ================== DADOS ==================


def carregar_controle(usuario):
    if usuario == "thailany":
        caminho_controle = CAMINHO_CONTROLE_THAILANY
    else:
        caminho_controle = CAMINHO_CONTROLE_VICTOR
    if not os.path.isfile(caminho_controle):
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_controle}")
    df = pd.read_excel(caminho_controle, sheet_name=0, engine="openpyxl")
    print(f"Shape do DataFrame: {df.shape}")
    print(f"Colunas disponíveis: {df.columns.tolist()}")
    if df.shape[1] <= IDX_AJ_LABEL:
        raise ValueError(
            f"A planilha tem {df.shape[1]} colunas, mas o código espera pelo menos {IDX_AJ_LABEL + 1} colunas.")
    return df


def carregar_txt(nome_arquivo):
    registros = {}
    if os.path.isfile(nome_arquivo):
        with open(nome_arquivo, "r", encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if "|" in linha:
                    empresa, data = linha.split("|", 1)
                    registros[empresa] = data
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

# ================== INTRO PATINHO ==================


def mostrar_intro(root, callback):
    import sys
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(__file__)

    caminho_patinho = None
    caminho_patinho_1 = os.path.join(base_path, "patinho.gif")
    if os.path.isfile(caminho_patinho_1):
        caminho_patinho = caminho_patinho_1
    elif os.path.isfile(CAMINHO_PATINHO_2):
        caminho_patinho = CAMINHO_PATINHO_2
    intro = tk.Toplevel(root)
    intro.attributes("-fullscreen", True)
    intro.configure(bg="white")
    if caminho_patinho:
        try:
            img = tk.PhotoImage(file=caminho_patinho)
            lbl_img = tk.Label(intro, image=img, bg="white")
            lbl_img.image = img
            lbl_img.pack(expand=True)
        except Exception as e:
            lbl_txt = tk.Label(
                intro,
                text=f"[Erro carregando imagem: {e}]",
                fg="white",
                bg="white",
                font=FONT_TITLE)
            lbl_txt.pack(expand=True)
    else:
        lbl_txt = tk.Label(
            intro,
            text="[Imagem patinho.gif não encontrada]",
            fg="white",
            bg="white",
            font=FONT_TITLE)
        lbl_txt.pack(expand=True)
    intro.after(2000, lambda: (intro.destroy(), callback()))

# ================== ENVIO ==================


def enviar_email_label(label, df, usuario):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session
        row = get_row_by_label(df, label)
        if row is None:
            raise Exception(
                f"Botão/label '{label}' não encontrado na planilha.")
        if IDX_EMAIL_TO >= len(row):
            email_to = ""
        else:
            email_to = safe_email_field(row.iloc[IDX_EMAIL_TO])
        if IDX_EMAIL_CC >= len(row):
            email_cc = ""
        else:
            email_cc = safe_email_field(row.iloc[IDX_EMAIL_CC])
        if not email_to:
            raise Exception(
                f"E-mail de destino não configurado para '{label}'.")
        anexos = get_anexos_from_row(row, usuario)
        if not anexos:
            raise Exception(f"Anexos não encontrados para '{label}'.")
        assunto = get_assunto_from_row(row)
        if not assunto:
            nomes_empresas = [
                extrair_nome_empresa_do_arquivo(a) for a in anexos]
            assunto = "Listagem de Funcionários - " + \
                " / ".join(nomes_empresas)
        mail = outlook.CreateItem(0)
        try:
            accounts = session.Accounts
            email_conta = EMAIL_CONTA_PADRAO_THAILANY if usuario == "thailany" else EMAIL_CONTA_PADRAO_VICTOR
            account_found = False
            for acc in accounts:
                if str(acc.SmtpAddress).lower() == email_conta.lower():
                    mail._oleobj_.Invoke(64209, 0, 8, 0, acc)
                    account_found = True
                    break
            if not account_found:
                raise Exception(
                    f"Conta de e-mail {email_conta} não encontrada no Outlook.")
        except Exception as e:
            raise Exception(f"Erro ao selecionar conta de envio: {e}")
        mail.To = email_to
        if email_cc:
            mail.CC = email_cc
        mail.Subject = assunto
        mail.Display(False)
        time.sleep(0.25)
        assinatura_html = mail.HTMLBody or ""
        corpo_html = (
            "<html><body style='font-family: Calibri; font-size: 11pt;'>"
            "<p>Bom dia,</p>"
            "<p>Segue listagem de funcionários que estão Ativos em nosso sistema. "
            "Por gentileza confirmar se a listagem está correta, conferindo Nome, Cargo e Situação (Ativo/Afastado/Pendente).</p>"
            "</body></html>"
        )
        mail.HTMLBody = corpo_html + assinatura_html
        for anexo in anexos:
            mail.Attachments.Add(os.path.abspath(anexo))
        mail.Send()
        registrar_txt(TXT_ENVIO, label)
        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()

# ================== COBRANÇA ==================


def cobrar_email_label(label, df, usuario):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        email_conta = EMAIL_CONTA_PADRAO_THAILANY if usuario == "thailany" else EMAIL_CONTA_PADRAO_VICTOR
        account_found = False
        for account in ns.Accounts:
            if str(account.SmtpAddress).lower() == email_conta.lower():
                sent_items = account.DeliveryStore.GetDefaultFolder(
                    5)  # 5 = olFolderSentMail
                account_found = True
                print(f"✅ Conta encontrada: {email_conta}")
                break
        if not account_found:
            raise Exception(
                f"Conta de e-mail {email_conta} não encontrada no Outlook.")
        items = sent_items.Items
        # Ordena por data de envio descendente (mais recente primeiro)
        items.Sort("[SentOn]", True)
        row = get_row_by_label(df, label)
        if row is None:
            raise Exception(
                f"Botão/label '{label}' não encontrado na planilha.")
        if IDX_EMAIL_TO >= len(row):
            email_to = ""
        else:
            email_to = safe_email_field(row.iloc[IDX_EMAIL_TO])
        if IDX_EMAIL_CC >= len(row):
            email_cc = ""
        else:
            email_cc = safe_email_field(row.iloc[IDX_EMAIL_CC])
        if not email_to:
            raise Exception(
                f"E-mail de destino não configurado para '{label}'.")
        anexos = get_anexos_from_row(row, usuario)
        if not anexos:
            raise Exception(f"Anexos não encontrados para '{label}'.")
        assunto = get_assunto_from_row(row)
        if not assunto:
            nomes_empresas = [extrair_nome_empresa_do_arquivo(a)
                              for a in anexos]
            assunto = "Listagem de Funcionários - " + " / ".join(nomes_empresas)
        assunto_normalizado = unicodedata.normalize(
            "NFD", assunto).encode('ascii', 'ignore').decode('ascii').lower()
        assunto_normalizado = re.sub(
            r'^(re|fwd|fw):\s*', '', assunto_normalizado, flags=re.I)
        assunto_normalizado = re.sub(
            r'\s+', ' ', assunto_normalizado.strip())
        print(
            f"🔍 Procurando e-mail com assunto normalizado: '{assunto_normalizado}'")
        hoje = datetime.now()
        last_mail = None
        for itm in items:
            try:
                sent_on = itm.SentOn
                if sent_on.month != hoje.month or sent_on.year != hoje.year:
                    print(
                        f"⏭️ E-mail ignorado: Não é do mês corrente (SentOn: {sent_on})")
                    continue
                subj = (itm.Subject or "").lower()
                subj = unicodedata.normalize(
                    "NFD", subj).encode('ascii', 'ignore').decode('ascii')
                subj = re.sub(r'^(re|fwd|fw):\s*', '', subj, flags=re.I)
                subj = re.sub(r'\s+', ' ', subj.strip())
                sender = (itm.SenderEmailAddress or "").lower()
                print(
                    f"📧 Verificando e-mail: Assunto='{subj}', Remetente='{sender}', SentOn='{sent_on}'")
                if email_conta.lower() in sender and assunto_normalizado in subj:
                    last_mail = itm
                    print(
                        f"✅ E-mail encontrado: Assunto='{itm.Subject}', SentOn='{sent_on}'")
                    break
            except Exception as e:
                print(f"⚠️ Erro ao verificar e-mail: {e}")
                continue
        if last_mail is None:
            raise Exception(
                f"Não foi possível localizar e-mail anterior do mês corrente com assunto contendo: '{assunto}'.")
        reply = last_mail.ReplyAll()
        reply.Subject = last_mail.Subject
        reply.To = email_to
        if email_cc:
            reply.CC = email_cc
        corpo_cobranca = (
            "<html><body style='font-family: Calibri; font-size: 11pt;'>"
            "<p>Prezados,</p><p>Algum retorno?</p>"
            "</body></html>"
        )
        reply.HTMLBody = corpo_cobranca + (reply.HTMLBody or "")
        for anexo in anexos:
            reply.Attachments.Add(os.path.abspath(anexo))
        try:
            accounts = ns.Accounts
            account_found = False
            for acc in accounts:
                if str(acc.SmtpAddress).lower() == email_conta.lower():
                    reply._oleobj_.Invoke(64209, 0, 8, 0, acc)
                    account_found = True
                    print(f"✅ Conta de envio selecionada: {email_conta}")
                    break
            if not account_found:
                raise Exception(
                    f"Conta de e-mail {email_conta} não encontrada para envio.")
        except Exception as e:
            raise Exception(
                f"Erro ao selecionar conta de envio para cobrança: {e}")
        reply.Send()
        registrar_txt(TXT_COBRANCA, label)
        print(f"✅ Cobrança enviada para '{label}'")
        return True, None
    except Exception as e:
        print(f"❌ Erro ao cobrar '{label}': {e}")
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()

# ================== OCULTAR/EXIBIR ARQUIVOS TXT ==================


def hide_file(file_path):
    try:
        if os.path.exists(file_path):
            win32api.SetFileAttributes(
                file_path, win32con.FILE_ATTRIBUTE_HIDDEN)
    except Exception as e:
        print(f"Erro ao ocultar {file_path}: {e}")


def unhide_file(file_path):
    try:
        if os.path.exists(file_path):
            win32api.SetFileAttributes(
                file_path, win32con.FILE_ATTRIBUTE_NORMAL)
    except Exception as e:
        print(f"Erro ao exibir {file_path}: {e}")

# ================== APLICAÇÃO (UI) ==================


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("E.A.I. - E-SOCIAL ARTIFICIAL INTELLIGENCE")
        self.root.geometry("600x450")  # Aumentado em 100px
        centralizar(self.root)
        for txt_file in [
                TXT_ULTIMO_ENVIO,
                TXT_ENVIO,
                TXT_COBRANCA,
                TXT_PROCURACOES,
                TXT_ULTIMO_AJUSTE,
                TXT_ULTIMO_GERAR]:
            unhide_file(txt_file)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.usuario_atual = None
        self.df = pd.DataFrame()
        self.labels = []
        self.enviados = 0
        self.cobrados = 0
        self.baixados = 0
        self.ajustadas = 0
        self.procuracoes_enviadas = 0
        self.frame_empresas_envio = None
        self.frame_empresas_cobranca = None
        self.frame_empresas_procuracoes = None
        self.password_visible = False
        self.show_envio_cobranca_button = True
        self.caps_lock_labels = []
        self.gerar_planilha_window = None


        self._build_login_screen()
        self.check_caps_lock()

    def check_caps_lock(self):
        is_caps_on = win32api.GetKeyState(win32con.VK_CAPITAL) & 1
        for label, parent_bg in self.caps_lock_labels:
            try:
                if is_caps_on:
                    label.config(text="CAPS LOCK ATIVADA", bg="lightcoral")
                else:
                    label.config(text="", bg=parent_bg)
            except tk.TclError:
                # Widget might be destroyed, do nothing.
                # The cleanup logic should remove it from the list soon.
                pass
        self.root.after(100, self.check_caps_lock)

    def _build_login_screen(self):
        # This method builds the login UI and can be called multiple times.
        self.overlay_login = tk.Frame(self.root)
        self.overlay_login.place(relx=0, rely=0, relwidth=1, relheight=1)

        login_frame = tk.Frame(self.overlay_login)
        login_frame.pack(expand=True)

        login_frame.grid_columnconfigure(0, weight=1)
        login_frame.grid_columnconfigure(1, weight=0)
        login_frame.grid_columnconfigure(2, weight=2)
        login_frame.grid_columnconfigure(3, weight=0)
        login_frame.grid_columnconfigure(4, weight=1)

        tk.Label(login_frame, text="USUÁRIO:", font=FONT_LABEL).grid(
            row=0, column=1, sticky="e", padx=5, pady=5)
        self.usuario_entry_frame = create_aligned_entry(
            login_frame, font=FONT_ENTRY)
        self.usuario_entry_frame.grid(
            row=0, column=2, sticky="we", padx=5, pady=5)
        self.usuario_entry = self.usuario_entry_frame.entry

        tk.Label(login_frame, text="SENHA:", font=FONT_LABEL).grid(
            row=1, column=1, sticky="e", padx=5, pady=5)
        self.senha_entry_frame = create_aligned_entry(
            login_frame, show="*", font=FONT_ENTRY)
        self.senha_entry_frame.grid(
            row=1, column=2, sticky="we", padx=5, pady=5)
        self.senha_entry = self.senha_entry_frame.entry

        self.eye_icon_label = tk.Label(
            login_frame, text="🔒", font=FONT_NORMAL, cursor="hand2")
        self.eye_icon_label.grid(
            row=1, column=3, sticky="w", padx=(5, 5), pady=5)
        self.eye_icon_label.bind("<Button-1>", self.toggle_password_visibility)

        caps_lock_label = tk.Label(login_frame, text="", font=FONT_BOLD)
        caps_lock_label.grid(
            row=2, column=1, columnspan=2, sticky="we", padx=5, pady=(5, 0))
        parent_bg = login_frame.cget("background")
        self.caps_lock_labels.append((caps_lock_label, parent_bg))

        forgot_btn = tk.Label(
            login_frame, text="ESQUECI A MINHA SENHA", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        forgot_btn.grid(row=3, column=2, sticky="n", pady=5, padx=5)
        forgot_btn.bind("<Button-1>", self.abrir_janela_palavra_seguranca)

        btn_entrar = tk.Button(
            login_frame,
            text="ENTRAR",
            font=FONT_BUTTON,
            command=self.validar_login)
        btn_entrar.grid(row=4, column=2, pady=(10, 10), sticky="we", padx=5)

        footer_frame = tk.Frame(self.overlay_login)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")

        btn_sair_total = tk.Button(
            footer_frame,
            text="SAIR",
            font=FONT_BUTTON,
            width=10,
            command=self.root.destroy)
        btn_sair_total.pack(side=tk.LEFT, padx=10, pady=(0, 5))

        help_link = tk.Label(
            footer_frame, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

        self.usuario_entry.bind(
            "<Return>", lambda e: self.senha_entry.focus_set())
        self.senha_entry.bind("<Return>", lambda e: self.validar_login(e))
        btn_entrar.bind("<Return>", lambda e: self.validar_login(e))
        self.usuario_entry.focus_set()

    def tela_login(self):
        # Clear the root window and rebuild the login screen
        for widget in self.root.winfo_children():
            widget.destroy()
        self.root.title("E.A.I. - E-SOCIAL ARTIFICIAL INTELLIGENCE")
        self.root.geometry("600x450")  # Aumentado em 100px
        centralizar(self.root)

        # Reset state variables
        self.usuario_atual = None
        self.df = pd.DataFrame()
        self.labels = []
        self.show_envio_cobranca_button = True

        self._build_login_screen()

    def toggle_password_visibility(self, event=None):
        if self.password_visible:
            self.senha_entry.config(show="*")
            self.eye_icon_label.config(text="🔒")
            self.password_visible = False
        else:
            self.senha_entry.config(show="")
            self.eye_icon_label.config(text="🔓")
            self.password_visible = True

    def on_close(self):
        for txt_file in [
                TXT_ULTIMO_ENVIO,
                TXT_ENVIO,
                TXT_COBRANCA,
                TXT_PROCURACOES,
                TXT_ULTIMO_AJUSTE,
                TXT_ULTIMO_GERAR]:
            hide_file(txt_file)
        self.root.destroy()

    def abrir_janela_palavra_seguranca(self, event=None):
        initial_usuario = self.usuario_entry.get().strip().lower()

        win = tk.Toplevel(self.root)
        win.title("E.A.I. - RECUPERAÇÃO DE SENHA")
        centralizar_janela(win, 600, 500)  # Aumentado em 100px
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        tentativas_restantes = 3

        frame = tk.Frame(win)
        frame.pack(expand=True, padx=20, pady=20)

        frame.grid_columnconfigure(0, weight=1)
        frame.grid_columnconfigure(1, weight=2)
        frame.grid_columnconfigure(2, weight=0)
        frame.grid_columnconfigure(3, weight=1)

        tk.Label(frame, text="USUÁRIO:", font=FONT_LABEL).grid(
            row=0, column=0, sticky="e", padx=5, pady=5)
        usuario_rec_entry_frame = create_aligned_entry(frame, font=FONT_ENTRY)
        usuario_rec_entry_frame.grid(
            row=0, column=1, sticky="we", padx=5, pady=5)
        usuario_rec_entry = usuario_rec_entry_frame.entry

        tk.Label(frame, text="PALAVRA DE SEGURANÇA:", font=FONT_LABEL).grid(
            row=1, column=0, sticky="e", padx=5, pady=5)
        palavra_entry_frame = create_aligned_entry(
            frame, font=FONT_ENTRY, show="*")
        palavra_entry_frame.grid(row=1, column=1, sticky="we", padx=5, pady=5)
        palavra_entry = palavra_entry_frame.entry

        if initial_usuario:
            usuario_rec_entry.insert(0, initial_usuario)
            usuario_rec_entry.config(state="readonly")
            palavra_entry.focus_set()
        else:
            usuario_rec_entry.focus_set()

        palavra_visible = tk.BooleanVar(value=False)
        palavra_icon = tk.Label(
            frame, text="🔒", font=FONT_NORMAL, cursor="hand2")
        palavra_icon.grid(row=1, column=2, sticky="w", padx=5)

        def toggle_palavra_visibility():
            is_hidden = not palavra_visible.get()
            if is_hidden:
                palavra_entry.config(show="")
                palavra_icon.config(text="🔓")
            else:
                palavra_entry.config(show="*")
                palavra_icon.config(text="🔒")
            palavra_visible.set(is_hidden)

        palavra_icon.bind("<Button-1>", lambda e: toggle_palavra_visibility())

        contador_label = tk.Label(
            frame, text=f"RESTAM {tentativas_restantes} TENTATIVAS.", font=FONT_BOLD, bg="lightcoral")
        contador_label.grid(row=2, column=0, columnspan=4, pady=5)

        def verificar():
            nonlocal tentativas_restantes
            usuario = usuario_rec_entry.get().strip().lower()
            palavra_digitada = palavra_entry.get().strip().lower()

            if not usuario:
                messagebox.showwarning(
                    "ATENÇÃO", "POR FAVOR, INFORME O USUÁRIO.", parent=win)
                return

            if not palavra_digitada:
                self._show_palavra_seguranca_dialog(win, usuario)
                palavra_entry.focus_set()
                return

            PALAVRAS_SECRETAS = {"victor": "creatina", "thailany": "spoc"}

            if usuario in PALAVRAS_SECRETAS and palavra_digitada == PALAVRAS_SECRETAS[usuario]:
                messagebox.showinfo(
                    "SENHA RECUPERADA", f"OLÁ, {usuario.upper()}, SUA SENHA DO SISTEMA E.A.I. É: 1001", parent=win)
                win.destroy()
            else:
                tentativas_restantes -= 1
                if tentativas_restantes > 0:
                    if tentativas_restantes == 1:
                        contador_label.config(
                            text="RESTA 1 TENTATIVA.", font=FONT_BOLD, bg="lightcoral")
                    else:
                        contador_label.config(
                            text=f"RESTAM {tentativas_restantes} TENTATIVAS.", font=FONT_BOLD, bg="lightcoral")
                    messagebox.showwarning(
                        "DADOS INCORRETOS", "USUÁRIO OU PALAVRA DE SEGURANÇA INCORRETA.", parent=win)
                else:
                    messagebox.showerror(
                        "TENTATIVAS ESGOTadas", "VOCÊ ESGOTOU SUAS TENTATIVAS. POR FAVOR, SOLICITE AUXÍLIO DO DEPARTAMENTO DE T.I.", parent=win)
                    webbrowser.open(
                        "https://wa.me/551221239207?text=N%C3%A3o%20lembro%20a%20minha%20senha%2Fpalavra%20de%20seguran%C3%A7a%20do%20sistema%20E.A.I.%20Pode%20me%20ajudar%3F")
                    win.destroy()

        btn_verificar = tk.Button(
            frame, text="RECUPERAR SENHA", font=FONT_BUTTON, command=verificar)
        btn_verificar.grid(row=3, column=0, columnspan=4, pady=20, padx=20)
        palavra_entry.bind("<Return>", lambda e: verificar())
        usuario_rec_entry.bind("<Return>", lambda e: palavra_entry.focus_set())

        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=win.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

    def _enviar_uma_procuracao_thread(self, empresa, btn_widget):
        # A lógica de envio de e-mail será adicionada na próxima etapa
        success, err = enviar_email_procuracao(
            empresa, self.df, self.usuario_atual)

        def ui_update():
            if success:
                try:
                    btn_widget.destroy()
                except tk.TclError:
                    pass  # Widget já pode ter sido destruído
                self.procuracoes_enviadas += 1
            else:
                messagebox.showerror(
                    "Erro de Envio", f"Falha ao enviar e-mail para '{empresa}': {err}")

            # Atualiza o contador
            if self.frame_empresas_procuracoes and self.frame_empresas_procuracoes.winfo_exists():
                remaining_buttons = len([w for w in self.frame_empresas_procuracoes.winfo_children() if isinstance(w, tk.Button)])
                self.contador_label_procuracoes.config(text=f"QUANTIDADE DE PROCURAÇÕES VENCIDAS / A VENCER: {remaining_buttons}")

        self.root.after(0, ui_update)

    def _enviar_todas_procuracoes_thread(self):
        if not self.frame_empresas_procuracoes:
            return

        widgets = [w for w in self.frame_empresas_procuracoes.winfo_children() if isinstance(
            w, tk.Button)]

        for widget in widgets:
            empresa = widget.cget("text")
            # Chama a função de envio para cada um, sequencialmente
            self._enviar_uma_procuracao_thread(empresa, widget)
            time.sleep(0.5)  # Pequena pausa para não sobrecarregar

    def tela_procuracoes(self):
        win = tk.Toplevel(self.root)
        win.title("E.A.I. - PROCURAÇÕES")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1100x700")

        frame_top = tk.Frame(win)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=10)

        tk.Label(frame_top, text="PROCURAÇÕES", font=FONT_TITLE).pack(pady=6)

        btn_todos = tk.Button(
            frame_top,
            text="ENVIAR TODOS",
            bg="#1f5fbf",
            fg="white",
            font=FONT_BUTTON,
            command=lambda: threading.Thread(
                target=self._enviar_todas_procuracoes_thread, daemon=True).start()
        )
        btn_todos.pack(pady=6)

        # A lógica de exibição dos botões e contagem virá na próxima etapa
        hoje = datetime.now()
        registros_procuracoes = carregar_txt(TXT_PROCURACOES)
        botoes_a_criar = []

        # Itera a partir da 4a linha (índice 2) até o final do dataframe
        for index, row in self.df.iloc[2:].iterrows():
            nome_empresa = str(row.iloc[0])
            validade_cell = row.iloc[IDX_AC_VALIDADE]
            status_cell = str(row.iloc[IDX_F_BLOQUEADA])

            # Pula se o nome da empresa for nulo/vazio
            if pd.isna(nome_empresa) or not nome_empresa.strip():
                continue

            # Pula se na coluna F contiver "BLOQUEADA"
            if "BLOQUEADA" in status_cell.upper():
                continue

            # Verifica se já foi enviado hoje
            if nome_empresa in registros_procuracoes and registros_procuracoes[nome_empresa] == hoje.strftime("%d/%m/%Y"):
                continue

            # Converte a validade para datetime, se possível
            if isinstance(validade_cell, datetime):
                validade_data = validade_cell
            elif isinstance(validade_cell, str):
                try:
                    validade_data = datetime.strptime(
                        validade_cell, "%d/%m/%Y")
                except ValueError:
                    continue  # Ignora formatos de data inválidos
            else:
                continue  # Ignora células que não são data nem string

            # Adiciona à lista se a data de validade for anterior a hoje
            if validade_data < hoje:
                botoes_a_criar.append(nome_empresa)

        botoes_a_criar.sort(key=lambda x: normalize_key(x))

        self.procuracoes_enviadas = 0
        self.contador_label_procuracoes = tk.Label(
            frame_top, text=f"QUANTIDADE DE PROCURAÇÕES VENCIDAS / A VENCER: {len(botoes_a_criar)}", font=FONT_LABEL)
        self.contador_label_procuracoes.pack(pady=2)

        container = tk.Frame(win)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        hbar = tk.Scrollbar(container, orient=tk.HORIZONTAL,
                            command=canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(xscrollcommand=hbar.set)

        frame_empresas = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_empresas, anchor="nw")
        self.frame_empresas_procuracoes = frame_empresas

        for i, nome_empresa in enumerate(botoes_a_criar):
            r = i % MAX_LINHAS_COLUNA
            c = i // MAX_LINHAS_COLUNA
            btn = tk.Button(
                frame_empresas,
                text=nome_empresa,
                font=FONT_BUTTON,
                width=35,
                height=4,
                wraplength=310
            )
            btn.configure(
                command=lambda emp=nome_empresa, b=btn: threading.Thread(
                    target=self._enviar_uma_procuracao_thread, args=(emp, b), daemon=True
                ).start()
            )
            btn.grid(row=r, column=c, padx=10, pady=10)

        def _atualiza_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_empresas.bind("<Configure>", _atualiza_scroll)
        self._bind_horizontal_scroll(canvas)

        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape, text="VOLTAR", font=FONT_BUTTON, width=10, command=win.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

        centralizar(win)

    def _has_envio_actions(self):
        registros_envio = carregar_txt(TXT_ENVIO)
        botoes = [lbl for lbl in self.labels if not self._enviado_no_mes(
            lbl, registros_envio) and not self._has_ok_status(lbl)]
        return bool(botoes)

    def _has_cobranca_actions(self):
        registros_cobranca = carregar_txt(TXT_COBRANCA)
        botoes = [
            lbl for lbl in self.labels if not self._cobrado_hoje(
                lbl,
                registros_cobranca) and lbl not in EXCLUIR_NA_COBRANCA and not self._has_ok_status(lbl)
        ]
        return bool(botoes)

    def validar_login(self, event=None):
        usuario = self.usuario_entry.get().strip().lower()
        senha = self.senha_entry.get().strip()

        if not usuario:
            messagebox.showwarning(
                "E.A.I. - ATENÇÃO", "USUÁRIO NÃO DIGITADO. VERIFIQUE, POR FAVOR.", parent=self.root)
            self.usuario_entry.focus_set()
            return

        if not senha:
            messagebox.showwarning(
                "E.A.I. - ATENÇÃO", "SENHA NÃO DIGITADA. VERIFIQUE, POR FAVOR.", parent=self.root)
            self.senha_entry.focus_set()
            return

        USUARIOS = {"victor": "1001", "thailany": "1001"}
        if usuario in USUARIOS and senha == USUARIOS[usuario]:
            self.usuario_atual = usuario
            try:
                self.df = carregar_controle(usuario)
                self.labels = listar_labels(self.df)
            except Exception as e:
                messagebox.showerror(
                    "E.A.I. - ERRO", f"FALHA AO CARREGAR PLANILHA: {e}")
                self.root.destroy()
                return
            self.overlay_login.destroy()
            messagebox.showinfo(
                "E.A.I. - BEM-VINDO",
                f"LOGIN REALIZADO COM SUCESSO, {usuario.upper()}!")
            self.enviados = 0
            self.cobrados = 0
            self.baixados = 0
            self.ajustadas = 0
            self.root.geometry("600x500")  # Aumentado em 100px
            centralizar(self.root)
            self.tela_inicial()
        else:
            self._show_login_error_dialog()

    def _show_login_error_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("E.A.I. - DADOS INCORRETOS")
        centralizar_janela(dialog, 600, 450)  # Aumentado em 100px
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = tk.Frame(dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        label = tk.Label(
            main_frame, text="USUÁRIO E/OU SENHA INCORRETOS.\nGOSTARIA DE TENTAR NOVAMENTE OU REDEFINIR A SENHA?", font=FONT_LABEL)
        label.pack(pady=10)

        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=10)

        def tentar_novamente():
            dialog.destroy()
            self.usuario_entry.delete(0, tk.END)
            self.senha_entry.delete(0, tk.END)
            self.usuario_entry.focus_set()

        def redefinir_senha():
            dialog.destroy()
            self.abrir_janela_palavra_seguranca()

        btn_tentar = tk.Button(
            button_frame, text="TENTAR NOVAMENTE", font=FONT_BUTTON, command=tentar_novamente)
        btn_tentar.pack(side=tk.LEFT, padx=10)

        btn_redefinir = tk.Button(
            button_frame, text="REDEFINIR SENHA", font=FONT_BUTTON, command=redefinir_senha)
        btn_redefinir.pack(side=tk.LEFT, padx=10)

        rodape = tk.Frame(dialog)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=dialog.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=10)

    def _show_palavra_seguranca_dialog(self, parent_win, usuario):
        dialog = tk.Toplevel(parent_win)
        dialog.title("E.A.I. - PALAVRA DE SEGURANÇA NÃO DIGITADA")
        centralizar_janela(dialog, 600, 500)  # Aumentado em 100px
        dialog.resizable(False, False)
        dialog.transient(parent_win)
        dialog.grab_set()

        main_frame = tk.Frame(dialog)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        label = tk.Label(
            main_frame, text="PALAVRA DE SEGURANÇA NÃO DIGITADA", font=FONT_LABEL, bg="lightcoral")
        label.pack(pady=(10, 20), ipady=5, ipadx=5)

        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=20)

        def esqueci_palavra():
            webbrowser.open(
                "https://wa.me/551221239207?text=N%C3%A3o%20lembro%20a%20minha%20senha%2Fpalavra%20de%20seguran%C3%A7a%20do%20sistema%20E.A.I.%20Pode%20me%20ajudar%3F")
            dialog.destroy()

        btn_digitar = tk.Button(
            button_frame, text="DIGITAR NOVAMENTE", font=FONT_BUTTON, command=dialog.destroy, width=35)
        btn_digitar.pack(pady=5)

        btn_esqueci = tk.Button(
            button_frame, text="ESQUECI A PALAVRA DE SEGURANÇA", font=FONT_BUTTON, command=esqueci_palavra, width=35)
        btn_esqueci.pack(pady=5)

        rodape = tk.Frame(dialog)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=dialog.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

        parent_win.wait_window(dialog)

    def mostrar_status_geracao(self):
        # Destruir qualquer banner de status existente para evitar duplicação
        for widget in self.status_frame.winfo_children():
            widget.destroy()

        # Criar o novo banner de status
        status_banner = tk.Frame(self.status_frame, bg="lightblue")
        status_banner.pack(fill="x", pady=5)

        status_label = tk.Label(
            status_banner,
            text="PROCESSO DE GERAÇÃO INICIADO...",
            font=FONT_BOLD,
            bg="lightblue",
            fg="black"
        )
        status_label.pack(pady=5)

        # Agendar a troca para o contador regressivo após 10 segundos
        self.root.after(
            10000, lambda: self._iniciar_countdown(status_banner, status_label, 180))
        return status_banner

    def _iniciar_countdown(self, banner, label, segundos):
        if segundos >= 0:
            label.config(text=f"AGUARDE {segundos} SEGUNDOS, POR FAVOR")
            self.root.after(
                1000, lambda: self._iniciar_countdown(banner, label, segundos - 1))
        else:
            # Quando o contador terminar, destruir o banner
            try:
                banner.destroy()
            except tk.TclError:
                # O widget pode já ter sido destruído se o usuário fechou a tela
                pass

    def salvar_planilhas_outlook(self, botao, usuario):
        try:
            botao.pack_forget()
            counter_frame = tk.Frame(self.status_frame, bg="lightblue")
            counter_frame.pack(pady=5)
            counter_label = tk.Label(
                counter_frame,
                text="ANALISANDO E-MAILS...",
                font=FONT_BOLD,
                bg="lightblue",
                fg="black"
            )
            counter_label.pack(pady=5)

            if usuario == "thailany":
                email_conta = EMAIL_PLANILHAS_THAILANY
                pasta_destino = PASTA_PLANILHAS_THAILANY
            else:
                email_conta = EMAIL_CONTA_PADRAO_VICTOR
                pasta_destino = PASTA_PLANILHAS_VICTOR

            print(
                f"📧 Iniciando salvamento de planilhas para o usuário: {usuario}, email: {email_conta}")
            print(f"📂 Pasta de destino: {pasta_destino}")

            pythoncom.CoInitialize()
            outlook = win32.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            account_found = False
            for account in ns.Accounts:
                if str(account.SmtpAddress).lower() == email_conta.lower():
                    inbox = account.DeliveryStore.GetDefaultFolder(6)
                    account_found = True
                    break
            if not account_found:
                raise Exception(
                    f"Conta de e-mail {email_conta} não encontrada no Outlook.")

            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            os.makedirs(pasta_destino, exist_ok=True)

            # --- PASS 1: Pre-scan to get accurate total ---
            total_planilhas = 0
            relevant_emails = []
            temp_dir = os.path.join(pasta_destino, "temp_zip_scan")
            os.makedirs(temp_dir, exist_ok=True)

            print(
                "🔍 Fase 1: Analisando e-mails para contagem total de planilhas...")
            for itm in items:
                try:
                    assunto = str(itm.Subject or "").strip()
                    assunto_normalizado = unicodedata.normalize(
                        "NFD", assunto).encode('ascii', 'ignore').decode('ascii').lower()
                    assunto_normalizado = re.sub(
                        r'\s+', ' ', assunto_normalizado.strip())
                    if re.search(r'listagem\s*de\s*funcionarios', assunto_normalizado, re.IGNORECASE) and itm.Attachments.Count > 0:
                        relevant_emails.append(itm)
                except Exception:
                    continue

            for itm in relevant_emails:
                for att in itm.Attachments:
                    if att.FileName.lower().endswith(".zip"):
                        temp_zip_path = os.path.join(temp_dir, att.FileName)
                        try:
                            att.SaveAsFile(temp_zip_path)
                            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                                xlsx_files_in_zip = [
                                    f for f in zip_ref.namelist() if f.lower().endswith(".xlsx")]
                                total_planilhas += len(xlsx_files_in_zip)
                        finally:
                            if os.path.exists(temp_zip_path):
                                os.remove(temp_zip_path)
            shutil.rmtree(temp_dir)

            print(
                f"📊 Contagem finalizada. Total de planilhas a salvar: {total_planilhas}")
            self.root.after(
                0, lambda: counter_label.config(text=f"PLANILHAS SALVAS: 0 DE {total_planilhas}"))

            # --- PASS 2: Process and save files ---
            print("🚀 Fase 2: Salvando planilhas...")
            baixados = 0
            for itm in relevant_emails:
                try:
                    assunto = str(itm.Subject or "").strip()
                    nome_base = re.sub(r'[<>:"/\\|?*]', '', assunto).strip()

                    for att in itm.Attachments:
                        if att.FileName.lower().endswith(".zip"):
                            caminho_zip_temp = os.path.join(
                                pasta_destino, att.FileName)
                            att.SaveAsFile(caminho_zip_temp)
                            try:
                                with zipfile.ZipFile(caminho_zip_temp, 'r') as zip_ref:
                                    xlsx_files = [
                                        f for f in zip_ref.namelist() if f.lower().endswith(".xlsx")]
                                    for zip_info in xlsx_files:
                                        nome_arquivo = f"{nome_base}.xlsx"
                                        caminho_final = os.path.join(
                                            pasta_destino, nome_arquivo)
                                        contador = 1
                                        while os.path.exists(caminho_final):
                                            nome_arquivo = f"{nome_base}_{contador}.xlsx"
                                            caminho_final = os.path.join(
                                                pasta_destino, nome_arquivo)
                                            contador += 1

                                        # Use a more robust, single-step extraction to prevent files from disappearing
                                        with zip_ref.open(zip_info) as source, open(caminho_final, 'wb') as target:
                                            shutil.copyfileobj(source, target)

                                        baixados += 1
                                        self.baixados = baixados
                                        self.root.after(0, lambda b=baixados, t=total_planilhas: counter_label.config(
                                            text=f"PLANILHAS SALVAS: {b} DE {t}"))
                                        print(
                                            f"✅ Arquivo salvo: {caminho_final} ({baixados}/{total_planilhas})")
                            finally:
                                os.remove(caminho_zip_temp)
                except Exception as item_err:
                    print(
                        f"⚠️ Erro ao processar e-mail com assunto '{assunto}': {item_err}")
                    continue

            print(f"🏁 Total de arquivos baixados: {baixados}")
            with open(TXT_ULTIMO_ENVIO, "w", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%d/%m/%Y"))
            hide_file(TXT_ULTIMO_ENVIO)
            counter_frame.destroy()
            messagebox.showinfo(
                "E.A.I. - PLANILHAS SALVAS",
                "CONCLUIDO! TODAS AS PLANILHAS FORAM SALVAS COM SUCESSO NA PASTA.")
        except Exception as e:
            print(f"❌ Erro geral: {e}")
            if 'counter_frame' in locals():
                counter_frame.destroy()
            messagebox.showerror("ERRO", f"FALHA AO SALVAR PLANILHAS: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _ajustar_planilhas_thread(self, botao):
        botao.pack_forget()
        hoje = datetime.now()
        SENHA_PROTECAO = "spoc"
        if self.usuario_atual == "thailany":
            pasta = PASTA_PLANILHAS_THAILANY
        else:
            pasta = PASTA_PLANILHAS_VICTOR

        counter_frame = tk.Frame(self.status_frame, bg="lightblue")
        counter_frame.pack(pady=5)
        counter_label = tk.Label(
            counter_frame,
            text="OBTENDO A QUANTIDADE DE PLANILHAS A SEREM ALTERADAS...:",
            font=FONT_BOLD,
            bg="lightblue",
            fg="black",
            wraplength=500
        )
        counter_label.pack(pady=5, padx=10)

        try:
            all_files = [f for f in os.listdir(
                pasta) if f.endswith(".xlsx")]
            if not all_files:
                counter_frame.destroy()
                messagebox.showinfo(
                    "AJUSTAR PLANILHAS", "NENHUMA PLANILHA ENCONTRADA NA PASTA.")
                return

            # --- FASE 0: Pré-verificação com Excel para contagem precisa ---
            print("🔍 Fase 0: Verificando planilhas que precisam de ajuste...")
            arquivos_para_ajustar = []
            pythoncom.CoInitialize()
            excel_prescan = win32.Dispatch("Excel.Application")
            excel_prescan.Visible = False
            excel_prescan.DisplayAlerts = False
            try:
                for arquivo in all_files:
                    caminho_completo = os.path.join(pasta, arquivo)
                    try:
                        # Tenta abrir. Se estiver protegido por senha, vai falhar.
                        wb = excel_prescan.Workbooks.Open(
                            caminho_completo, ReadOnly=True, Password="")
                        wb.Close()
                        arquivos_para_ajustar.append(arquivo)
                    except Exception as e:
                        print(
                            f"ℹ️ Planilha '{arquivo}' já está protegida. Ignorando. Erro: {e}")
                        continue
            finally:
                excel_prescan.Quit()
                pythoncom.CoUninitialize()

            total_a_ajustar = len(arquivos_para_ajustar)
            print(f"📊 Total de planilhas a ajustar: {total_a_ajustar}")
            self.root.after(
                0, lambda: counter_label.config(text=f"PLANILHAS AJUSTADAS: 0 DE {total_a_ajustar}"))

            if total_a_ajustar == 0:
                counter_frame.destroy()
                messagebox.showinfo(
                    "AJUSTAR PLANILHAS", "Nenhuma planilha nova para ajustar.")
                return

            # --- FASE 1: Processamento com openpyxl ---
            print("🚀 Fase 1: Ajustando conteúdo das planilhas...")
            ajustadas = 0
            for arquivo in arquivos_para_ajustar:
                caminho = os.path.join(pasta, arquivo)
                # ... lógica do openpyxl ...
                wb = openpyxl.load_workbook(caminho)
                for ws in wb.worksheets:
                    ws.delete_rows(2, 1)
                    merged_cells = []
                    for merged_range in ws.merged_cells.ranges:
                        merged_cells.append(merged_range)
                    linhas_para_excluir = []
                    for r in range(2, ws.max_row + 1):
                        is_merged = False
                        for merged_range in merged_cells:
                            if r >= merged_range.min_row and r <= merged_range.max_row:
                                is_merged = True
                                break
                        if not is_merged:
                            try:
                                status = ws.cell(row=r, column=6).value
                                data_cell = ws.cell(row=r, column=7).value
                                if status and str(status).strip().lower() == "pendente":
                                    if isinstance(data_cell, datetime):
                                        data_val = data_cell
                                    else:
                                        try:
                                            data_val = datetime.strptime(
                                                str(data_cell), "%d/%m/%Y")
                                        except Exception:
                                            continue
                                    if data_val.month == hoje.month and data_val.year == hoje.year:
                                        linhas_para_excluir.append(r)
                            except BaseException:
                                continue
                    for r in sorted(linhas_para_excluir, reverse=True):
                        try:
                            ws.delete_rows(r, 1)
                        except BaseException:
                            continue
                    thin = Side(border_style="thin", color="000000")
                    border = Border(
                        top=thin, left=thin, right=thin, bottom=thin)
                    for row_idx in range(1, ws.max_row + 1):
                        for col_idx in range(1, ws.max_column + 1):
                            try:
                                is_in_merged = False
                                for merged_range in ws.merged_cells.ranges:
                                    if (row_idx >= merged_range.min_row and row_idx <= merged_range.max_row and
                                            col_idx >= merged_range.min_col and col_idx <= merged_range.max_col):
                                        is_in_merged = True
                                        break
                                if not is_in_merged:
                                    cell = ws.cell(
                                        row=row_idx, column=col_idx)
                                    cell.border = border
                                    cell.alignment = Alignment(
                                        horizontal="center", vertical="center", wrap_text=True)
                                    if col_idx == 7:
                                        cell.number_format = 'DD/MM/YYYY'
                                    else:
                                        cell.number_format = '@'
                            except BaseException:
                                continue
                wb.save(caminho)
                ajustadas += 1
                self.ajustadas = ajustadas
                self.root.after(0, lambda a=ajustadas: counter_label.config(
                    text=f"PLANILHAS AJUSTADAS: {a} DE {total_a_ajustar}"))
                print(
                    f"✅ Planilha ajustada: {caminho} ({ajustadas}/{total_a_ajustar})")

            # --- FASE 2: Proteção e Autofit com Excel ---
            print("🚀 Fase 2: Aplicando autofit e proteção...")
            self.root.after(
                0, lambda: counter_label.config(text=f"Aplicando senhas: 0 DE {total_a_ajustar}"))

            pythoncom.CoInitialize()
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            senhas_definidas = 0
            tentativas = 0

            for arquivo in arquivos_para_ajustar:
                tentativas += 1
                try:
                    caminho_completo = os.path.join(pasta, arquivo)
                    wb_excel = excel.Workbooks.Open(caminho_completo)
                    for ws in wb_excel.Worksheets:
                        ws.Cells.Select()
                        ws.Cells.Columns.AutoFit()
                        ws.Cells.Rows.AutoFit()
                        ws.Range("A1").Select()
                        ws.Protect(SENHA_PROTECAO)
                    wb_excel.Protect(SENHA_PROTECAO)
                    wb_excel.Save()
                    wb_excel.Close()
                    senhas_definidas += 1
                except Exception as e:
                    print(
                        f"⚠️ Erro ao aplicar autofit e senha em '{arquivo}'. Pulando. Erro: {e}")
                    try:
                        wb_excel.Close(SaveChanges=False)
                    except Exception:
                        pass
                finally:
                    self.root.after(0, lambda t=tentativas: counter_label.config(
                        text=f"Aplicando senhas: {t} DE {total_a_ajustar}"))

            excel.Quit()
            pythoncom.CoUninitialize()

            if senhas_definidas > 0:
                with open(TXT_ULTIMO_AJUSTE, "w", encoding="utf-8") as f:
                    f.write(datetime.now().strftime("%d/%m/%Y"))
                hide_file(TXT_ULTIMO_AJUSTE)

            counter_frame.destroy()
            messagebox.showinfo(
                "E.A.I. - PLANILHAS AJUSTADAS",
                f"CONCLUÍDO! TODAS AS PLANILHAS FORAM AJUSTADAS COM SUCESSO. SENHA: {SENHA_PROTECAO.lower()}")
        except Exception as e:
            print(f"❌ Erro geral: {e}")
            if 'counter_frame' in locals():
                counter_frame.destroy()
            messagebox.showerror("ERRO", f"FALHA AO AJUSTAR PLANILHAS: {e}")

    def chekar_botao_salvar(self, frame, usuario):
        if os.path.isfile(TXT_ULTIMO_ENVIO):
            with open(TXT_ULTIMO_ENVIO, "r", encoding="utf-8") as f:
                data_txt = f.read().strip()
            try:
                ultima_data = datetime.strptime(data_txt, "%d/%m/%Y")
                hoje = datetime.now()
                if ultima_data.month == hoje.month and ultima_data.year == hoje.year:
                    return None
            except Exception:
                pass
        btn = tk.Button(frame, text="SALVAR PLANILHAS",
                        width=30, font=FONT_BUTTON)
        btn.pack(pady=6)
        print(f"🔧 Configurando botão Salvar Planilhas para usuário: {usuario}")
        btn.configure(
            command=lambda b=btn: threading.Thread(
                target=self.salvar_planilhas_outlook, args=(b, usuario), daemon=True).start())
        return btn

    def _chekar_botao_gerar_pendentes(self, frame, usuario):
        if os.path.isfile(TXT_ULTIMO_GERAR):
            with open(TXT_ULTIMO_GERAR, "r", encoding="utf-8") as f:
                data_txt = f.read().strip()
            try:
                ultima_data = datetime.strptime(data_txt, "%d/%m/%Y")
                hoje = datetime.now()
                if ultima_data.month == hoje.month and ultima_data.year == hoje.year:
                    return None
            except Exception:
                pass
        btn = tk.Button(frame, text="GERAR PENDENTES",
                        width=30, font=FONT_BUTTON)
        btn.pack(pady=6)
        btn.configure(
            command=lambda b=btn: GerarPlanilhaWindow(
                self.root,
                usuario,
                self,
                b))
        return btn

    def chekar_botao_ajustar(self, frame):
        if os.path.isfile(TXT_ULTIMO_AJUSTE):
            with open(TXT_ULTIMO_AJUSTE, "r", encoding="utf-8") as f:
                data_txt = f.read().strip()
            try:
                ultima_data = datetime.strptime(data_txt, "%d/%m/%Y")
                hoje = datetime.now()
                if ultima_data.month == hoje.month and ultima_data.year == hoje.year:
                    return None
            except Exception:
                pass
        btn = tk.Button(
            frame,
            text="AJUSTAR PLANILHAS",
            width=30,
            font=FONT_BUTTON)
        btn.pack(pady=6)
        btn.configure(
            command=lambda b=btn: threading.Thread(
                target=self._ajustar_planilhas_thread, args=(b,), daemon=True).start())
        return btn

    def tela_inicial(self):
        for w in self.root.winfo_children():
            w.destroy()

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        self.root.title("E.A.I. - MENU DE OPÇÕES")

        tk.Label(
            main_frame,
            text="SELECIONE UMA OPÇÃO:",
            font=FONT_LABEL).pack(pady=(10, 20))

        # Buttons Frame
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(pady=10)

        self.btn_gerar_pendentes = self._chekar_botao_gerar_pendentes(
            frame=buttons_frame, usuario=self.usuario_atual)
        self.btn_salvar_planilhas = self.chekar_botao_salvar(
            buttons_frame, self.usuario_atual)
        self.btn_ajustar_planilhas = self.chekar_botao_ajustar(buttons_frame)

        if self.show_envio_cobranca_button:
            self.btn_envio_cobranca = tk.Button(
                buttons_frame,
                text="ENVIO / COBRANÇA",
                width=30,
                font=FONT_BUTTON,
                command=self.tela_envio_cobranca)
            self.btn_envio_cobranca.pack(pady=6)

        self.btn_procuracoes = tk.Button(
            buttons_frame,
            text="PROCURAÇÕES",
            width=30,
            font=FONT_BUTTON,
            command=self.tela_procuracoes)
        self.btn_procuracoes.pack(pady=6)

        # Status Frame - this is where the counters will go
        self.status_frame = tk.Frame(main_frame)
        self.status_frame.pack(pady=20, fill="x", expand=True)

        # Footer Frame
        rodape = tk.Frame(main_frame)
        rodape.pack(side=tk.BOTTOM, fill=tk.X)
        sair_btn = tk.Button(
            rodape,
            text="SAIR",
            font=FONT_BUTTON,
            width=10,
            command=self.confirmar_saida)
        sair_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

    def confirmar_saida(self):
        usuario_nome = self.usuario_atual.upper()
        if messagebox.askyesno("E.A.I. - CONFIRMAÇÃO DE SAÍDA", f"{usuario_nome}, TEM CERTEZA QUE DESEJA SAIR DO SISTEMA?"):
            messagebox.showinfo(
                "E.A.I. - AGRADECIMENTO", f"{usuario_nome}, OBRIGADO POR UTILIZAR O SISTEMA E.A.I. - E-SOCIAL ARTIFICIAL INTELLIGENCE.")
            self.tela_login()

    def tela_envio_cobranca(self):
        has_envio = self._has_envio_actions()
        has_cobranca = self._has_cobranca_actions()

        if not has_envio and not has_cobranca:
            messagebox.showinfo(
                "E.A.I. - AVISO", "NÃO HÁ AÇÕES DE ENVIO OU COBRANÇA DISPONÍVEIS NO MOMENTO.")
            # Find the button on the main screen and hide it.
            # This is a bit of a workaround, but necessary since we don't have a direct reference.
            for widget in self.root.winfo_children():
                if isinstance(widget, tk.Frame):
                    for btn_container in widget.winfo_children():
                        if isinstance(btn_container, tk.Frame):
                            for btn in btn_container.winfo_children():
                                if isinstance(btn, tk.Button) and "ENVIO / COBRANÇA" in btn.cget("text"):
                                    btn.pack_forget()
                                    return
            return

        win = tk.Toplevel(self.root)
        win.title("E.A.I. - LISTAGEM DE FUNCIONÁRIOS")
        centralizar_janela(win, 600, 500)
        tk.Label(win, text="ESCOLHA A AÇÃO:", font=FONT_LABEL).pack(pady=20)

        if has_envio:
            tk.Button(
                win,
                text="ENVIO",
                width=22,
                font=FONT_BUTTON,
                command=self.tela_empresas_envio).pack(pady=10)

        if has_cobranca:
            tk.Button(
                win,
                text="COBRANÇA",
                width=22,
                font=FONT_BUTTON,
                command=self.tela_empresas_cobranca).pack(pady=6)

        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X)
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=win.destroy
        )
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape,
            text="ENCONTROU UM ERRO? CLIQUE AQUI",
            font=FONT_FORGOT_PASSWORD,
            fg="blue",
            cursor="hand2"
        )
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

    def _bind_horizontal_scroll(self, widget_canvas):
        widget_canvas.bind(
            "<Enter>",
            lambda e: self._enable_mousewheel(widget_canvas))
        widget_canvas.bind(
            "<Leave>",
            lambda e: self._disable_mousewheel(widget_canvas))

    def _enable_mousewheel(self, canvas_widget):
        if sys.platform.startswith("win") or sys.platform == "darwin":
            canvas_widget.bind_all(
                "<MouseWheel>",
                lambda e: self._on_mousewheel_horizontal(
                    e,
                    canvas_widget))
        else:
            canvas_widget.bind_all(
                "<Button-4>",
                lambda e: self._on_mousewheel_horizontal(
                    e,
                    canvas_widget))
            canvas_widget.bind_all(
                "<Button-5>",
                lambda e: self._on_mousewheel_horizontal(
                    e,
                    canvas_widget))

    def _disable_mousewheel(self, canvas_widget):
        if sys.platform.startswith("win") or sys.platform == "darwin":
            canvas_widget.unbind_all("<MouseWheel>")
        else:
            canvas_widget.unbind_all("<Button-4>")
            canvas_widget.unbind_all("<Button-5>")

    def _on_mousewheel_horizontal(self, event, canvas_widget):
        try:
            if hasattr(event, "delta"):
                move = int(-1 * (event.delta / 120))
                canvas_widget.xview_scroll(move, "units")
            else:
                if event.num == 4:
                    canvas_widget.xview_scroll(-1, "units")
                elif event.num == 5:
                    canvas_widget.xview_scroll(1, "units")
        except Exception:
            pass

    def tela_empresas_envio(self):
        win = tk.Toplevel(self.root)
        win.title("E.A.I. - ENVIO DA LISTAGEM DE FUNCIONÁRIOS")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1100x700")
        frame_top = tk.Frame(win)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=10)

        tk.Label(
            frame_top,
            text="ENVIO DA LISTAGEM DE FUNCIONÁRIOS",
            font=FONT_TITLE).pack(
            pady=6)
        btn_todos = tk.Button(
            frame_top,
            text="ENVIAR TODOS",
            bg="#1f5fbf",
            fg="white",
            font=FONT_BUTTON,
            command=lambda: threading.Thread(
                target=self._enviar_todos_thread,
                daemon=True).start())
        btn_todos.pack(pady=6)
        registros_envio = carregar_txt(TXT_ENVIO)
        botoes = [lbl for lbl in self.labels if not self._enviado_no_mes(
            lbl, registros_envio) and not self._has_ok_status(lbl)]
        botoes.sort(key=lambda x: normalize_key(x))
        self.contador_label_envio = tk.Label(
            frame_top, text=f"QUANTIDADE DE LISTAGENS A SEREM ENVIADAS NESTE MÊS: {len(botoes)}", font=FONT_LABEL)
        self.contador_label_envio.pack(pady=2)
        container = tk.Frame(win)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        hbar = tk.Scrollbar(
            container,
            orient=tk.HORIZONTAL,
            command=canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(xscrollcommand=hbar.set)
        frame_empresas = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_empresas, anchor="nw")
        self.frame_empresas_envio = frame_empresas
        for i, label in enumerate(botoes):
            r = i % MAX_LINHAS_COLUNA
            c = i // MAX_LINHAS_COLUNA
            btn = tk.Button(
                frame_empresas,
                text=label,
                font=FONT_BUTTON,
                width=35,
                height=4,
                wraplength=310)
            btn.configure(
                command=lambda l=label, b=btn: threading.Thread(
                    target=self._enviar_uma_label_thread, args=(
                        l, b), daemon=True).start())
            btn.grid(row=r, column=c, padx=10, pady=10)

        def _atualiza_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_empresas.bind("<Configure>", _atualiza_scroll)
        self._bind_horizontal_scroll(canvas)

        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=win.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

        centralizar(win)

    def _enviado_no_mes(self, label, registros_envio=None):
        if registros_envio is None:
            registros_envio = carregar_txt(TXT_ENVIO)
        if label not in registros_envio:
            return False
        try:
            dt = datetime.strptime(registros_envio[label], "%d/%m/%Y")
            hoje = datetime.now()
            return dt.month == hoje.month and dt.year == hoje.year
        except Exception:
            return False

    def _enviar_uma_label_thread(self, label, btn_widget):
        success, err = enviar_email_label(label, self.df, self.usuario_atual)

        def ui_update():
            if success:
                try:
                    btn_widget.destroy()
                except Exception:
                    pass
                self.enviados += 1
            else:
                messagebox.showerror(
                    "ERRO", f"FALHA AO ENVIAR '{label}': {err}")

            if self.frame_empresas_envio and self.frame_empresas_envio.winfo_exists():
                remaining_buttons = len([w for w in self.frame_empresas_envio.winfo_children() if isinstance(w, tk.Button)])
                self.contador_label_envio.config(text=f"QUANTIDADE DE LISTAGENS A SEREM ENVIADAS NESTE MÊS: {remaining_buttons}")

        self.root.after(0, ui_update)

    def _enviar_todos_thread(self):
        if not self.frame_empresas_envio:
            return
        widgets = [
            w for w in self.frame_empresas_envio.winfo_children() if isinstance(
                w, tk.Button)]
        for w in list(widgets):
            label = w.cget("text")
            if self._enviado_no_mes(label) or self._has_ok_status(label):
                def remove_ui(widget=w):
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                self.root.after(0, remove_ui)
                continue
            success, err = enviar_email_label(
                label, self.df, self.usuario_atual)

            def ui_update_after_one(s=success, e=err, widget=w, lbl=label):
                if s:
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                    self.enviados += 1
                else:
                    messagebox.showerror(
                        "ERRO", f"FALHA AO ENVIAR '{lbl}': {e}")

                if self.frame_empresas_envio and self.frame_empresas_envio.winfo_exists():
                    remaining_buttons = len([w for w in self.frame_empresas_envio.winfo_children() if isinstance(w, tk.Button)])
                    self.contador_label_envio.config(text=f"QUANTIDADE DE LISTAGENS A SEREM ENVIADAS NESTE MÊS: {remaining_buttons}")
            self.root.after(0, ui_update_after_one)
            time.sleep(0.4)

    def _has_ok_status(self, label):
        row = get_row_by_label(self.df, label)
        if row is None:
            return False
        if IDX_AA_STATUS >= len(row):
            return False
        status = safe_email_field(row.iloc[IDX_AA_STATUS])
        return status.strip().upper() == "OK"

    def tela_empresas_cobranca(self):
        win = tk.Toplevel(self.root)
        win.title("E.A.I. - COBRANÇA DA LISTAGEM DE FUNCIONÁRIOS")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1100x700")
        frame_top = tk.Frame(win)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=10)
        tk.Label(
            frame_top,
            text="COBRANÇA DA LISTAGEM DE FUNCIONÁRIOS",
            font=FONT_TITLE).pack(
            pady=6)
        btn_todos = tk.Button(
            frame_top,
            text="COBRAR TODOS",
            bg="#1f5fbf",
            fg="white",
            font=FONT_BUTTON,
            command=lambda: threading.Thread(
                target=self._cobrar_todos_thread,
                daemon=True).start())
        btn_todos.pack(pady=6)
        registros_cobranca = carregar_txt(TXT_COBRANCA)
        botoes = [
            lbl for lbl in self.labels if not self._cobrado_hoje(
                lbl,
                registros_cobranca) and lbl not in EXCLUIR_NA_COBRANCA and not self._has_ok_status(lbl)]
        botoes.sort(key=lambda x: normalize_key(x))
        self.contador_label_cobranca = tk.Label(
            frame_top, text=f"QUANTIDADE DE LISTAGENS PENDENTES DE RETORNO HOJE: {len(botoes)}", font=FONT_LABEL)
        self.contador_label_cobranca.pack(pady=2)
        container = tk.Frame(win)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        hbar = tk.Scrollbar(
            container,
            orient=tk.HORIZONTAL,
            command=canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(xscrollcommand=hbar.set)
        frame_empresas = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_empresas, anchor="nw")
        self.frame_empresas_cobranca = frame_empresas
        for i, label in enumerate(botoes):
            r = i % MAX_LINHAS_COLUNA
            c = i // MAX_LINHAS_COLUNA
            btn = tk.Button(
                frame_empresas,
                text=label,
                font=FONT_BUTTON,
                width=35,
                height=4,
                wraplength=310)
            btn.configure(
                command=lambda l=label, b=btn: threading.Thread(
                    target=self._cobrar_uma_label_thread, args=(
                        l, b), daemon=True).start())
            btn.grid(row=r, column=c, padx=10, pady=10)

        def _atualiza_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_empresas.bind("<Configure>", _atualiza_scroll)
        self._bind_horizontal_scroll(canvas)

        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X, anchor="sw")
        voltar_btn = tk.Button(
            rodape,
            text="VOLTAR",
            font=FONT_BUTTON,
            width=10,
            command=win.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=(0, 5))

        help_link = tk.Label(
            rodape, text="ENCONTROU UM ERRO? CLIQUE AQUI", font=FONT_FORGOT_PASSWORD, fg="blue", cursor="hand2")
        help_link.pack(side=tk.RIGHT, anchor="se", padx=10, pady=(0, 5))
        help_link.bind("<Button-1>", abrir_link_ajuda)

        centralizar(win)

    def _cobrado_hoje(self, label, registros_cobranca=None):
        if registros_cobranca is None:
            registros_cobranca = carregar_txt(TXT_COBRANCA)
        if label not in registros_cobranca:
            return False
        try:
            dt = datetime.strptime(registros_cobranca[label], "%d/%m/%Y")
            hoje = datetime.now()
            return dt.date() == hoje.date()
        except Exception:
            return False

    def _cobrar_uma_label_thread(self, label, btn_widget):
        success, err = cobrar_email_label(label, self.df, self.usuario_atual)

        def ui_update():
            if success:
                try:
                    btn_widget.destroy()
                except Exception:
                    pass
                self.cobrados += 1
            else:
                messagebox.showerror(
                    "Erro", f"Falha ao cobrar '{label}': {err}")

            if self.frame_empresas_cobranca and self.frame_empresas_cobranca.winfo_exists():
                remaining_buttons = len([w for w in self.frame_empresas_cobranca.winfo_children() if isinstance(w, tk.Button)])
                self.contador_label_cobranca.config(text=f"QUANTIDADE DE LISTAGENS PENDENTES DE RETORNO HOJE: {remaining_buttons}")

        self.root.after(0, ui_update)

    def _cobrar_todos_thread(self):
        if not self.frame_empresas_cobranca:
            return
        widgets = [
            w for w in self.frame_empresas_cobranca.winfo_children() if isinstance(
                w, tk.Button)]
        for w in list(widgets):
            label = w.cget("text")
            if self._cobrado_hoje(label) or self._has_ok_status(label):
                def remove_ui(widget=w):
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                self.root.after(0, remove_ui)
                continue
            success, err = cobrar_email_label(
                label, self.df, self.usuario_atual)

            def ui_update_after_one(s=success, e=err, widget=w, lbl=label):
                if s:
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                    self.cobrados += 1
                else:
                    messagebox.showerror(
                        "Erro", f"Falha ao cobrar '{lbl}': {e}")
                if self.frame_empresas_cobranca and self.frame_empresas_cobranca.winfo_exists():
                    remaining_buttons = len([w for w in self.frame_empresas_cobranca.winfo_children() if isinstance(w, tk.Button)])
                    self.contador_label_cobranca.config(text=f"QUANTIDADE DE LISTAGENS PENDENTES DE RETORNO HOJE: {remaining_buttons}")
            self.root.after(0, ui_update_after_one)
            time.sleep(0.4)


def get_row_by_label(df, label):
    if IDX_AJ_LABEL >= df.shape[1]:
        raise ValueError(
            f"Índice {IDX_AJ_LABEL} fora dos limites. Colunas disponíveis: {df.shape[1]}")
    mask = df.iloc[:, IDX_AJ_LABEL].astype(str) == label
    if not mask.any():
        print(f"Label '{label}' não encontrado na coluna {IDX_AJ_LABEL}")
        return None
    return df[mask].iloc[0]


def get_row_by_name(df, name):
    # Procura pelo nome da empresa na primeira coluna (A)
    mask = df.iloc[:, 0].astype(str).str.upper() == name.upper()
    if not mask.any():
        print(f"Empresa '{name}' não encontrada na coluna A.")
        return None
    return df[mask].iloc[0]


def enviar_email_procuracao(empresa, df, usuario):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        session = outlook.Session

        row = get_row_by_name(df, empresa)
        if row is None:
            raise Exception(
                f"Empresa '{empresa}' não encontrada na planilha de controle.")

        email_to = safe_email_field(row.iloc[IDX_AE_PROCURACAO_TO])
        email_cc = safe_email_field(row.iloc[IDX_AF_PROCURACAO_CC])

        if not email_to:
            raise Exception(
                f"E-mail de destino (Coluna AE) não configurado para '{empresa}'.")

        assunto = f"Procuração eletrônica - {empresa}"

        mail = outlook.CreateItem(0)

        # Define a conta de envio
        try:
            accounts = session.Accounts
            email_conta = EMAIL_CONTA_PADRAO_THAILANY if usuario == "thailany" else EMAIL_CONTA_PADRAO_VICTOR
            account_found = False
            for acc in accounts:
                if str(acc.SmtpAddress).lower() == email_conta.lower():
                    mail._oleobj_.Invoke(64209, 0, 8, 0, acc)
                    account_found = True
                    break
            if not account_found:
                raise Exception(
                    f"Conta de e-mail {email_conta} não encontrada no Outlook.")
        except Exception as e:
            raise Exception(f"Erro ao selecionar conta de envio: {e}")

        mail.To = email_to
        if email_cc:
            mail.CC = email_cc
        mail.Subject = assunto

        # Força a captura da assinatura padrão
        mail.Display(False)
        time.sleep(0.25)
        assinatura_html = mail.HTMLBody or ""

        # Monta o corpo do e-mail com a formatação especificada
        corpo_html = f"""
<html>
<body style='font-family: Calibri; font-size: 11pt;'>
<p>Prezados,</p>
<p>A procuração eletrônica para acesso ao E-Social está vencida. Solicitamos, por gentileza, a renovação para que possamos continuar efetuando os envios ao sistema.</p>
<p>Essa renovação é feita no site do <a href="https://cav.receita.fazenda.gov.br/autenticacao/login">E-CAC</a>, seguem instruções:</p>
<ol>
    <li>Acessar <a href="https://cav.receita.fazenda.gov.br/autenticacao/login">E-CAC</a> e logar;</li>
    <li>Clicar no link <b>Procuração Eletrônica – Cadastra Procuração;</b></li>
    <li>Informar CPF/CNPJ em Dados do Procurador – <b>18.802.823/0001-04;</b></li>
    <li>Selecionar as procurações desejadas do <b>E-Social</b> e clicar em <b>Cadastrar a Procuração (SST);</b></li>
    <li>Clicar em <b>assinar documento.</b></li>
</ol>
<p><b>Vigência de 2 ou mais anos.</b></p>
<p><span style='background-color: yellow;'><b>CNPJ OSWALDO CRUZ: 18.802.823/0001-04</b></span></p>
</body>
</html>
"""

        mail.HTMLBody = corpo_html + assinatura_html
        mail.Send()

        registrar_txt(TXT_PROCURACOES, empresa)
        return True, None

    except Exception as e:
        return False, str(e)
    finally:
        pythoncom.CoUninitialize()


def get_anexos_from_row(row, usuario):
    try:
        if IDX_AK_ANEXOS >= len(row):
            return []
        raw = str(row.iloc[IDX_AK_ANEXOS])
    except Exception:
        raw = ""
    if not raw or raw.lower() in {"nan", "none"}:
        return []
    parts = re.split(r"[;,]", raw)
    anexos = []
    for p in parts:
        p_clean = limpar_caminho(p)
        if p_clean and os.path.isfile(p_clean):
            anexos.append(p_clean)
    return anexos


def get_assunto_from_row(row):
    try:
        if IDX_AL_ASSUNTO >= len(row):
            return ""
        raw = str(row.iloc[IDX_AL_ASSUNTO])
    except Exception:
        raw = ""
    return safe_email_field(raw)


def listar_labels(df):
    if IDX_AJ_LABEL >= df.shape[1]:
        raise ValueError(
            f"Índice {IDX_AJ_LABEL} fora dos limites. Colunas disponíveis: {df.shape[1]}")
    labels = list(df.iloc[2:, IDX_AJ_LABEL].dropna().astype(str))
    labels.sort(key=lambda x: normalize_key(x))
    return labels


def abrir_link_ajuda(event=None):
    webbrowser.open(
        "https://wa.me/551221239207?text=Encontrei%20um%20erro%20no%20sistema%20E.A.I.%20Pode%20me%20ajudar%3F")


# ================== EXECUÇÃO ==================
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    app = App(root)

    def abrir_login():
        app.overlay_login.lift()
        root.deiconify()
        app.usuario_entry.focus_set()
    mostrar_intro(root, abrir_login)
    root.mainloop()
