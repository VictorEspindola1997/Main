import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
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

FONT_PADRAO = ("Arial", 11, "bold")
MAX_LINHAS_COLUNA = 6
IDX_AK_LABEL = 36
IDX_AL_ANEXOS = 37
IDX_AM_ASSUNTO = 38
IDX_EMAIL_TO = 32
IDX_EMAIL_CC = 33
IDX_AA_STATUS = 26  # Coluna AA para verificar "OK"
EXCLUIR_NA_COBRANCA = {"GRUPO OSWALDO CRUZ"}

# --- Lista de empresas permitidas (whitelist) ---
whitelist_empresas = set(map(str.upper, [
    "A J O GIBELLO LTDA",
    "ABRAH AUTOMAÇÃO LTDA",
    "ADAVEN HOTEIS E TURISMO LTDA",
    "ADEMIR FRESCA CONSTRUTORA E INCORPORADORA LTDA",
    "ADFX CONSTRUTORA E INCORPORADORA SPE LTDA",
    "AGUIAR & COSTA COMERCIAL ELETRICO E ELETRONICO LTDA",
    "APLOS EDUCACIONAL LTDA EPP",
    "ASSOCIAÇAO DE EDUCAÇAO SAO VICENTE DE PAULO",
    "ASSOCIAÇÃO FRANCISCANA DE ASSISTENCIA SOCIAL CORAÇÃO DE MARIA",
    "ASSOCIAÇÃO SOCIOASSISTENCIAL SAO VICENTE DE PAULO (CASA DO ANCIÃO)",
    "ASSOCIAÇÃO VALEPARAIBANA DE ASSISTÊNCIA MÉDICA POLICIAL",
    "ASTRAL INDUSTRIA E COMERCIO DE MARMORES E GRANITOS LTDA",
    "ATRAN SG TRANSPORTADORA LTDA",
    "AUTO POSTO ÁGUIA DE TAUBATÉ LTDA",
    "AUTO POSTO CONFIANÇA PINDAMONHANGABA LTDA",
    "BCF SUPERMERCADO LTDA",
    "BLUE VIEW CONSTRUTORA E INCORPORADORA SPE LTDA",
    "C. RIBEIRO - MANUTENÇÃO AUTOMOTIVA - ME",
    "CAMPOS INSTALAÇÃO DE EQUIPAMENTOS PARA AQUECIMENTO EIRELI",
    "CASA SÃO FRANCISCO DE IDOSOS DE TAUBATÉ",
    "CEMITÉRIO PARQUE COLINA DA PAZ LTDA.",
    "CENOURÃO DA TERRA ALIMENTOS LTDA - ME",
    "CENTRO DE DIAGNÓSTICO ANDRADE",
    "CERÂMICA KATO LTDA. EPP",
    "CESTA DE ALIMENTOS MARUSTE LTDA",
    "CLINICA JACQUES FELIX LTDA.",
    "CLINICA NEPOMUCENO LTDA",
    "COMDIESEL PEÇAS E SERVIÇOS LTDA. - EPP",
    "COMÉRCIO DE VIDROS FARIA LIMA LTDA. ME",
    "CONGREGAÇÃO DOS PADRES DO SAGRADO CORAÇÃO DE JESUS - CPSCJ",
    "CONGREGACAO DOS PADRES DO SAGRADO CORACAO DE JESUS (SEMINARIO DEHONIANO)",
    "CONSTRUTORA ARAUJO SIMAO LTDA",
    "CSV TRANSPORTES LTDA. ME",
    "D. M. G. RIGHI - ME",
    "DARE EDUCAÇÃO LTDA - ME",
    "DEISE CIOLFI BARRADAS",
    "DISTRIBUIDORA DE ALIMENTOS SRM LTDA - ME",
    "DONABELLA EXPRESS PANIFICAÇÃO LTDA.",
    "ECONLIFE CARTOES DE BENEFICIOS LTDA",
    "EDIFICIO TRANCOSO ARAUJO SIMAO",
    "EMPRESA DE ÁGUAS MINERAIS PASSA TRES LTDA.",
    "ESCOLA DAMASCO LTDA (FILIAL)",
    "ESCOLA DAMASCO LTDA (MATRIZ)",
    "F&L COMERCIAL, IMPORTADORA E EXPORTADORA EIRELI",
    "FAM MONTAGENS INDUSTRIAIS LTDA. ME",
    "FREITAS E FRESCA CONSTRUTORA LTDA",
    "FUNDACAO DE APOIO A CIENCIA E NATUREZA - FUNAT",
    "FUNERÁRIA E VELÓRIO TREMEMBÉ LTDA",
    "FUNERÁRIA TAUBATÉ LTDA",
    "G E L INSTALL COMERCIO E SERVICOS DE SISTEMAS ELETRO ELETRONICOS LTDA",
    "G. A. FERRARI PIZZARIA E CIA LTDA EPP",
    "GLOBE MAQUINAS LTDA",
    "HOTEL TORIBA LTDA",
    "HOTEL TORIBA LTDA (BONANZA GRILL)",
    "HOTEL TORIBA LTDA (CASA BAMBUI)",
    "IMUNIVALE DEDETIZADORA LTDA",
    "IPSUN - ENGENHARIA E INSTALAÇÕES INDUSTRIAIS LTDA",
    "JK EDUCACIONAL LTDA",
    "LABORATÓRIO OSWALDO CRUZ - CAÇAPAVA",
    "LABORATÓRIO OSWALDO CRUZ - CAMPOS DO JORDÃO",
    "LABORATÓRIO OSWALDO CRUZ - CARAGUATATUBA",
    "LABORATÓRIO OSWALDO CRUZ - CARAGUATATUBA (SHOPPING SERRAMAR)",
    "LABORATÓRIO OSWALDO CRUZ - CRUZEIRO",
    "LABORATÓRIO OSWALDO CRUZ - GUARATINGUETÁ",
    "LABORATÓRIO OSWALDO CRUZ - GUARATINGUETÁ (SHOPPING BURITI)",
    "LABORATÓRIO OSWALDO CRUZ - JACAREÍ",
    "LABORATÓRIO OSWALDO CRUZ - JACAREÍ (SHOPPING)",
    "LABORATÓRIO OSWALDO CRUZ - PINDAMONHANGABA",
    "LABORATÓRIO OSWALDO CRUZ - PINDAMONHANGABA (SHOPPING PATIO PINDA)",
    "LABORATÓRIO OSWALDO CRUZ - SÃO JOSÉ DOS CAMPOS (AQUARIUS CARREFOUR)",
    "LABORATÓRIO OSWALDO CRUZ - SÃO JOSÉ DOS CAMPOS (ESPLANADA)",
    "LABORATÓRIO OSWALDO CRUZ - SÃO JOSÉ DOS CAMPOS (SHOPPING ORIENTE)",
    "LABORATÓRIO OSWALDO CRUZ - SÃO PAULO (VILA GUILHERMINA)",
    "LABORATÓRIO OSWALDO CRUZ - TAUBATÉ (INDEPENDENCIA)",
    "LABORATÓRIO OSWALDO CRUZ - TAUBATÉ (MATRIZ)",
    "LBJ EDUCACAÇÃO SOCIEDADE LTDA (Taubaté)",
    "LICIO LINS BARRADAS JUNIOR",
    "LIMA NETO INSTALACOES LTDA",
    "LOVE DRIVE-IN LTDA. EPP",
    "MAISNOVE MEDICINA DIAGNOSTICA LTDA",
    "MARIA MÉRCIA AGOSTINHO - ME",
    "MIAMI MOTEL VALE DO PARAÍBA LTDA. ME",
    "MINERAÇÃO CORRÊA LTDA.",
    "MMH MAGAZINE SÃO JOSE LTDA",
    "MOLNAR MENDES RESTAURANTE LTDA (FORNALHA GRILL - SHOPPING TAUBATÉ)",
    "MOVAP - MOTEL VALE DO PARAIBA LTDA - EPP.",
    "MRBAKER CONFEITARIA E PANIFICAÇÃO LTDA - ME.",
    "MUK-LOK COMÉRCIO DE FERRO E AÇO, MONTAGENS E LOCAÇÃO DE EQUIPAMENTOS LTDA. ME",
    "O. M. MENDES JUNIOR LTDA",
    "O. MOLNAR MENDES EIRELI ME  (FORNALHA GRILL - SHOPPING CARAGUATATUBA)",
    "OFTALMO CENTRO MILLENIUM LTDA - EPP",
    "OR HOTELARIA E ORGANIZACAO DE EVENTOS LTDA (Ort Hotel)",
    "ORGANIZAÇÃO ASSISTENCIAL DE LUTO SAGRADA FAMÍLIA LTDA",
    "ORGANIZAÇÃO EDUCACIONAL CESAR LATTES LTDA.",
    "OSWALDO CRUZ GESTÃO OCUPACIONAL LTDA",
    "PANIBRASIL MASSAS CONGELADAS LTDA.",
    "PARQUE BAMBUI TURISMO LTDA",
    "PAULO SERGIO PEREIRA ASSAF",
    "PEDRO F. RIBEIRO & CIA LTDA",
    "PENSE MATERIAIS DIDATICOS EIRELI",
    "PLASTCLIN - CLÍNICA DE CIRURGIA PLÁSTICA LTDA. EPP",
    "PNEUBREQ CENTRO AUTOMOTIVO LTDA - INDEPENDENCIA (MATRIZ)",
    "POTENZA CELANO FERRAMENTAS LTDA (FILIAL)",
    "POTENZA CELANO FERRAMENTAS LTDA (MATRIZ)",
    "POTENZA COMERCIO, LOCACAO E SERVICOS LTDA (FILIAL)",
    "POTENZA COMERCIO, LOCACAO E SERVICOS LTDA (MATRIZ)",
    "PROGRESSAO EDUCACIONAL EIRELI",
    "R. P. SILVA RESTAURANTE - ME.",
    "REINALDO ROCHA CARNEIRO BASTOS",
    "RIBEIRO FARIA MANUTENCAO AUTOMOTIVA LTDA",
    "RICARDO RIGHI DE CARVALHO E OUTRA",
    "RODÃO AUTO POSTO - LTDA.",
    "TEMCO TERRAPLANAGEM LTDA",
    "TRP TRANSPORTES E LOGISTICA",
    "UNIODONTO TAUBATÉ - COOPERATIVA DE TRABALHO ODONTOLÓGICO",
    "V. GRALHEIRA AUGUSTO GIBELLO E CIA LTDA",
    "VALLE SUPRIMENTOS E INFORMÁTICA EIRELI ME",
    "VALENTE & VALENTE LTDA",
    "VENEZIA MOTEL VALE DO PARAÍBA LTDA. ME",
    "VITROMORTH ESTRUTURAS METALICAS LTDA (FILIAL)",
    "VITROMORTH ESTRUTURAS METÁLICAS LTDA. EPP (MATRIZ)"
]))

# ================== FUNÇÃO GERAR PENDENTES ==================


def executar_gerar_planilha(usuario, senha, id_digitado):
    chromedriver_path = r"C:\Users\TI Medicina\Documents\Webdriver\chromedriver-win64\chromedriver.exe"
    if usuario == "thailany":
        download_dir = PASTA_DOWNLOAD_THAILANY
        desktop_dir = BASE_DIR_THAILANY
    else:
        download_dir = PASTA_DOWNLOAD_VICTOR
        desktop_dir = r"C:\Users\TI Medicina\Desktop"

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.popups": 2,
        "download.prompt_for_download": False
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    wait = WebDriverWait(driver, 2)

    try:
        driver.get("https://sistema.soc.com.br/WebSoc/")
        driver.maximize_window()
        time.sleep(1)

        driver.find_element(By.NAME, "usu").send_keys(usuario)
        driver.find_element(By.NAME, "senha").send_keys(senha)

        # Aguarda carregamento do teclado virtual
        time.sleep(2)
        # Digitar o ID solicitado pelo usuário (ex: 3230), clicando nos botões
        # conforme valor
        teclado_div = driver.find_element(By.ID, "pteclado")
        botoes = teclado_div.find_elements(By.TAG_NAME, "input")

        # Dicionário valor->elemento para clicar rapidamente
        botoes_map = {btn.get_attribute("value"): btn for btn in botoes}

        for digito in id_digitado:
            if digito not in botoes_map:
                raise ValueError(
                    f"Dígito {digito} não encontrado no teclado virtual.")
            botoes_map[digito].click()
            time.sleep(1)  # intervalo de 1 segundo entre cliques

        # Clicar no botão Entrar
        btn_entrar = driver.find_element(By.ID, "bt_entrar")
        btn_entrar.click()

        time.sleep(5)  # aguardar login e página carregar

        # Agora começa processo de digitar 267 no campo 'cod_programa'
        campo_cod_prog = wait.until(
            EC.element_to_be_clickable(
                (By.ID, "cod_programa")))
        campo_cod_prog.click()
        time.sleep(0.5)
        campo_cod_prog.clear()
        campo_cod_prog.send_keys("267")
        time.sleep(0.5)
        driver.find_element(By.ID, "btn_programa").click()

        # Achar e clicar no ícone excel dentro do iframe
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.frame(iframe)
            try:
                wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//img[contains(@src, 'excel.png')]"))).click()
                driver.switch_to.default_content()
                break
            except BaseException:
                driver.switch_to.default_content()
                continue

        # Clicar em Consultar Pedidos no iframe
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.frame(iframe)
            try:
                wait.until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a.botaoT.btn-consultar-pedido"))).click()
                driver.switch_to.default_content()
                break
            except BaseException:
                driver.switch_to.default_content()
                continue

        print("🕒 Aguardando 40 segundos para geração do arquivo...")
        time.sleep(120)

        maior_id = -1
        div_maior = None
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            driver.switch_to.frame(iframe)
            try:
                wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div.div-download")))
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
                    driver.switch_to.default_content()
                    break
            except BaseException:
                driver.switch_to.default_content()

        print("🕒 Aguardando 3 segundos para download...")
        time.sleep(3)

        zip_files = glob.glob(os.path.join(download_dir, '*.zip'))
        if not zip_files:
            raise FileNotFoundError("❌ Nenhum arquivo .zip encontrado.")

        latest_zip = max(zip_files, key=os.path.getctime)
        with zipfile.ZipFile(latest_zip, 'r') as zip_ref:
            xls_files = [
                f for f in zip_ref.namelist() if f.lower().endswith('.xls')]
            if not xls_files:
                raise FileNotFoundError(
                    "❌ Nenhum arquivo .xls encontrado no .zip.")
            xls_file = xls_files[0]  # Assume o primeiro .xls é o desejado
            zip_ref.extract(xls_file, download_dir)

        time.sleep(2)
        xls_path = os.path.join(download_dir, xls_file)
        if not os.path.exists(xls_path):
            raise FileNotFoundError(
                f"❌ Arquivo .xls extraído não encontrado: {xls_path}")

        # Converter .xls para .xlsx
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb_xls = excel.Workbooks.Open(xls_path)
            xlsx_filename = "Planilha de pendentes.xlsx"
            dest_path = os.path.join(desktop_dir, xlsx_filename)
            counter = 1
            while os.path.exists(dest_path):
                xlsx_filename = f"Planilha de pendentes_{counter}.xlsx"
                dest_path = os.path.join(desktop_dir, xlsx_filename)
                counter += 1
            wb_xls.SaveAs(dest_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
            wb_xls.Close()
        finally:
            excel.Quit()

        # Fechar Excel para evitar conflitos
        os.system('taskkill /f /im excel.exe')
        time.sleep(2)

        # Aplicar formatação na planilha .xlsx
        wb = openpyxl.load_workbook(dest_path)
        ws = wb.active

        # Determinar última linha e coluna com conteúdo
        max_row = 1
        max_col = 1
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and not isinstance(
                        cell, openpyxl.cell.cell.MergedCell):
                    max_row = max(max_row, cell.row)
                    max_col = max(max_col, cell.column)

        # Aplicar bordas e centralizar conteúdo
        thin = Side(border_style="thin", color="000000")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        for row in ws.iter_rows(
                min_row=1,
                max_row=max_row,
                min_col=1,
                max_col=max_col):
            for cell in row:
                if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                    cell.border = border
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center", wrap_text=True)

        # Autosize para colunas
        for col_idx in range(1, max_col + 1):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            for row in ws.iter_rows(
                    min_row=1,
                    max_row=max_row,
                    min_col=col_idx,
                    max_col=col_idx):
                cell = row[0]
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except BaseException:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column_letter].width = min(adjusted_width, 50)

        # Autosize para linhas
        for row_idx in range(1, max_row + 1):
            max_height = 23
            for col in ws.iter_cols(
                    min_row=row_idx,
                    max_row=row_idx,
                    min_col=1,
                    max_col=max_col):
                cell = col[0]
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    continue
                try:
                    if cell.value:
                        lines = str(cell.value).count("\n") + 1
                        max_height = max(max_height, lines * 15)
                except BaseException:
                    pass
            ws.row_dimensions[row_idx].height = max_height

        # Filtrar linhas conforme whitelist, preservando a linha 2
        linhas_para_remover = []
        for i, row in enumerate(
                ws.iter_rows(min_row=5), start=5):  # Começar da linha 5
            empresa = str(row[0].value).strip() if row[0].value else ""
            fundo_pintado = any(
                cell.fill and cell.fill.fill_type for cell in row if not isinstance(
                    cell, openpyxl.cell.cell.MergedCell))
            vazia = all(
                cell.value is None or str(
                    cell.value).strip() == "" for cell in row if not isinstance(
                    cell, openpyxl.cell.cell.MergedCell))

            if fundo_pintado or vazia:
                if empresa in whitelist_empresas:
                    continue
                linhas_para_remover.append(i)
            elif empresa not in whitelist_empresas:
                linhas_para_remover.append(i)

        for linha in reversed(linhas_para_remover):
            ws.delete_rows(linha, 1)

        wb.save(dest_path)

        # Reabrir a planilha e focar na linha 1
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        try:
            wb_excel = excel.Workbooks.Open(dest_path)
            ws_excel = wb_excel.ActiveSheet
            ws_excel.Range("A1").Select()  # Foca na célula A1
            excel.ActiveWindow.Zoom = 100  # Zoom padrão para legibilidade
            excel.ActiveWindow.ScrollRow = 1  # Garante que a linha 1 está visível
            excel.ActiveWindow.ScrollColumn = 1  # Garante que a coluna A está visível
        except Exception as e:
            print(f"⚠️ Erro ao abrir a planilha: {e}")
            messagebox.showerror("Erro", f"Falha ao abrir a planilha: {e}")
        finally:
            excel.DisplayAlerts = True

        # Maximizar janela do Excel
        try:
            app = Application(backend="uia").connect(title_re=".*Excel.*")
            app.window(title_re=".*Excel.*").set_focus().maximize()
        except Exception as e:
            print(f"⚠️ Não foi possível maximizar o Excel: {e}")

        # Registrar a geração no arquivo TXT_ULTIMO_GERAR
        with open(TXT_ULTIMO_GERAR, "w", encoding="utf-8") as f:
            f.write(datetime.now().strftime("%d/%m/%Y"))
        hide_file(TXT_ULTIMO_GERAR)

    except Exception as e:
        print(f"❌ Erro: {e}")
        messagebox.showerror("Erro", f"Falha ao gerar planilha: {e}")
    finally:
        print("🟢 Finalizado. Navegador permanece aberto.")
        # Não fechar o navegador automaticamente
        # driver.quit()


def on_gerar_planilha(parent, usuario_atual):
    parent.unbind("<Return>")
    def apenas_numeros(char):
        return char.isdigit()

    win = tk.Toplevel(parent)
    win.title("ENTRAR NO SISTEMA SOC")
    centralizar_janela(win, 500, 500)
    win.resizable(False, False)

    frame = tk.Frame(win)
    frame.pack(expand=True, pady=20)

    tk.Label(frame, text="Usuário SOC:", font=("Arial", 12)).grid(row=0, column=0, sticky="e", padx=10, pady=5)
    entry_usuario = tk.Entry(frame, font=("Arial", 12), width=15)
    entry_usuario.grid(row=0, column=1, pady=5)

    tk.Label(frame, text="Senha SOC:", font=("Arial", 12)).grid(row=1, column=0, sticky="e", padx=10, pady=5)
    entry_senha = tk.Entry(frame, font=("Arial", 12), width=15, show='*')
    entry_senha.grid(row=1, column=1, pady=5)

    # Botão de "cadeado" para senha (usando emojis)
    senha_visivel = False
    def toggle_senha():
        nonlocal senha_visivel
        senha_visivel = not senha_visivel
        entry_senha.config(show='' if senha_visivel else '*')
        btn_senha.config(text="🔓" if senha_visivel else "🔒")
        print(f"🔗 Senha {'mostrada' if senha_visivel else 'oculta'}")

    btn_senha = tk.Button(frame, text="🔒", font=("Arial", 12), width=2, borderwidth=0, command=toggle_senha)
    btn_senha.grid(row=1, column=2, padx=5)

    vcmd = win.register(apenas_numeros)
    tk.Label(frame, text="ID SOC:", font=("Arial", 12)).grid(row=2, column=0, sticky="e", padx=10, pady=5)
    entry_id = tk.Entry(frame, font=("Arial", 12), width=15, show='*', validate="key", validatecommand=(vcmd, '%S'))
    entry_id.grid(row=2, column=1, pady=5)

    # Botão de "cadeado" para ID (usando emojis)
    id_visivel = False
    def toggle_id():
        nonlocal id_visivel
        id_visivel = not id_visivel
        entry_id.config(show='' if id_visivel else '*')
        btn_id.config(text="🔓" if id_visivel else "🔒")
        print(f"🔗 ID {'mostrado' if id_visivel else 'oculto'}")

    btn_id = tk.Button(frame, text="🔒", font=("Arial", 12), width=2, borderwidth=0, command=toggle_id)
    btn_id.grid(row=2, column=2, padx=5)

    btn_gerar = tk.Button(frame, text="Gerar Planilha de pendentes", font=("Arial", 12), width=30, height=2)
    btn_gerar.grid(row=3, column=0, columnspan=2, pady=15)

    # Variável de controle para mensagem de credenciais
    mensagem_credenciais_mostrada = False

    def executar():
        nonlocal mensagem_credenciais_mostrada
        print("🔗 Enter pressed in Gerar Planilha")
        usuario = entry_usuario.get().strip()
        senha = entry_senha.get().strip()
        id_digitado = entry_id.get().strip()

        # Verificar se algum campo está vazio
        campos_vazios = not usuario or not senha or not id_digitado

        if campos_vazios:
            # Mostrar mensagem apenas uma vez
            if not mensagem_credenciais_mostrada:
                messagebox.showwarning("Atenção", "Digite todas as credenciais do SOC.")
                mensagem_credenciais_mostrada = True
            # Sempre definir foco no primeiro campo vazio
            if not usuario:
                entry_usuario.delete(0, tk.END)
                entry_usuario.focus_set()
            elif not senha:
                entry_senha.delete(0, tk.END)
                entry_senha.focus_set()
            elif not id_digitado:
                entry_id.delete(0, tk.END)
                entry_id.focus_set()
            return

        # Resetar flag se todos os campos foram preenchidos
        mensagem_credenciais_mostrada = False

        if not id_digitado.isdigit():
            messagebox.showwarning("Atenção", "O ID deve conter apenas números.")
            entry_id.delete(0, tk.END)
            entry_id.focus_set()
            return

        status_label = tk.Label(frame, text="Gerando planilha, aguarde...", font=("Arial", 12), fg="blue")
        status_label.grid(row=4, column=0, columnspan=2, pady=10)
        win.update()

        win.destroy()
        threading.Thread(target=executar_gerar_planilha, args=(usuario, senha, id_digitado), daemon=True).start()

    btn_gerar.configure(command=executar)
    entry_usuario.focus_set()
    win.update()

    entry_usuario.bind("<Return>", lambda e: executar())
    entry_senha.bind("<Return>", lambda e: executar())
    entry_id.bind("<Return>", lambda e: executar())
    btn_gerar.bind("<Return>", lambda e: executar())
    win.bind("<Return>", lambda e: [print("🔗 Enter pressed in Gerar Planilha window"), executar()])

    print("🔗 Enter key bound for Gerar Planilha inputs and button")

    # Re-vincular Enter ao fechar
    win.protocol("WM_DELETE_WINDOW", lambda: [win.destroy(), parent.bind("<Return>", lambda e: parent.focus_get().invoke() if isinstance(parent.focus_get(), tk.Button) else None)])


def centralizar_janela(janela, largura=500, altura=500):
    janela.update_idletasks()
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    x = (largura_tela // 2) - (largura // 2)
    y = (altura_tela // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")

# ================== FUNÇÃO SALVAR PLANILHAS ==================


def hide_file(file_path):
    try:
        if os.path.exists(file_path):
            import win32api
            import win32con
            win32api.SetFileAttributes(
                file_path, win32con.FILE_ATTRIBUTE_HIDDEN)
    except Exception as e:
        print(f"Erro ao ocultar {file_path}: {e}")


# ================== CHECA SE BOTÃO DEVE APARECER ==================


def chekar_botao_gerar_pendentes(self, frame, usuario):
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
    btn = tk.Button(frame, text="Gerar Pendentes", width=30, font=FONT_PADRAO)
    btn.pack(pady=6)
    btn.configure(
        command=lambda b=btn: [
            b.pack_forget(),
            on_gerar_planilha(
                self.root,
                usuario)])


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
        text="Ajustar Planilhas",
        width=30,
        font=FONT_PADRAO)
    btn.pack(pady=6)
    btn.configure(
        command=lambda b=btn: threading.Thread(
            target=self._ajustar_planilhas_thread, args=(
                b,), daemon=True).start())

# ================== UTILITÁRIOS ==================


def centralizar(janela):
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f"{largura}x{altura}+{x}+{y}")


def normalize_key(x):
    """
    Normaliza uma string para ordenação:
    - Converte para minúsculas
    - Remove acentos e espaços extras
    - Elimina caracteres especiais
    """
    x = str(x).strip().lower()
    x = ''.join(
        c for c in unicodedata.normalize('NFD', x)
        if unicodedata.category(c) != 'Mn'
    )
    x = re.sub(r'[^a-z0-9 ]', '', x)
    return x


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
    if df.shape[1] <= IDX_AK_LABEL:
        raise ValueError(
            f"A planilha tem {
                df.shape[1]} colunas, mas o código espera pelo menos {
                IDX_AK_LABEL +
                1} colunas.")
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
                font=("Arial", 18))
            lbl_txt.pack(expand=True)
    else:
        lbl_txt = tk.Label(
            intro,
            text="[Imagem patinho.gif não encontrada]",
            fg="white",
            bg="white",
            font=("Arial", 18))
        lbl_txt.pack(expand=True)
    intro.after(5000, lambda: (intro.destroy(), callback()))

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
            "Bom dia,<br><br>"
            "Segue listagem de funcionários que estão Ativos em nosso sistema. "
            "Por gentileza confirmar se a listagem está correta, conferindo Nome, Cargo e Situação (Ativo/Afastado/Pendente).<br><br>"
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
                sent_items = account.DeliveryStore.GetDefaultFolder(5)  # 5 = olFolderSentMail
                account_found = True
                print(f"✅ Conta encontrada: {email_conta}")
                break
        if not account_found:
            raise Exception(f"Conta de e-mail {email_conta} não encontrada no Outlook.")
        items = sent_items.Items
        items.Sort("[SentOn]", True)  # Ordena por data de envio descendente (mais recente primeiro)
        row = get_row_by_label(df, label)
        if row is None:
            raise Exception(f"Botão/label '{label}' não encontrado na planilha.")
        if IDX_EMAIL_TO >= len(row):
            email_to = ""
        else:
            email_to = safe_email_field(row.iloc[IDX_EMAIL_TO])
        if IDX_EMAIL_CC >= len(row):
            email_cc = ""
        else:
            email_cc = safe_email_field(row.iloc[IDX_EMAIL_CC])
        if not email_to:
            raise Exception(f"E-mail de destino não configurado para '{label}'.")
        anexos = get_anexos_from_row(row, usuario)
        if not anexos:
            raise Exception(f"Anexos não encontrados para '{label}'.")
        assunto = get_assunto_from_row(row)
        if not assunto:
            nomes_empresas = [extrair_nome_empresa_do_arquivo(a) for a in anexos]
            assunto = "Listagem de Funcionários - " + " / ".join(nomes_empresas)
        assunto_normalizado = unicodedata.normalize("NFD", assunto).encode('ascii', 'ignore').decode('ascii').lower()
        assunto_normalizado = re.sub(r'^(re|fwd|fw):\s*', '', assunto_normalizado, flags=re.I)
        assunto_normalizado = re.sub(r'\s+', ' ', assunto_normalizado.strip())
        print(f"🔍 Procurando e-mail com assunto normalizado: '{assunto_normalizado}'")
        hoje = datetime.now()
        last_mail = None
        for itm in items:
            try:
                sent_on = itm.SentOn
                if sent_on.month != hoje.month or sent_on.year != hoje.year:
                    print(f"⏭️ E-mail ignorado: Não é do mês corrente (SentOn: {sent_on})")
                    continue
                subj = (itm.Subject or "").lower()
                subj = unicodedata.normalize("NFD", subj).encode('ascii', 'ignore').decode('ascii')
                subj = re.sub(r'^(re|fwd|fw):\s*', '', subj, flags=re.I)
                subj = re.sub(r'\s+', ' ', subj.strip())
                sender = (itm.SenderEmailAddress or "").lower()
                print(f"📧 Verificando e-mail: Assunto='{subj}', Remetente='{sender}', SentOn='{sent_on}'")
                if email_conta.lower() in sender and assunto_normalizado in subj:
                    last_mail = itm
                    print(f"✅ E-mail encontrado: Assunto='{itm.Subject}', SentOn='{sent_on}'")
                    break
            except Exception as e:
                print(f"⚠️ Erro ao verificar e-mail: {e}")
                continue
        if last_mail is None:
            raise Exception(f"Não foi possível localizar e-mail anterior do mês corrente com assunto contendo: '{assunto}'.")
        reply = last_mail.ReplyAll()
        reply.Subject = last_mail.Subject
        reply.To = email_to
        if email_cc:
            reply.CC = email_cc
        reply.HTMLBody = "Prezados,<br><br>Algum retorno?<br><br>" + (reply.HTMLBody or "")
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
                raise Exception(f"Conta de e-mail {email_conta} não encontrada para envio.")
        except Exception as e:
            raise Exception(f"Erro ao selecionar conta de envio para cobrança: {e}")
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
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import os
import pythoncom
import win32com.client as win32
import zipfile
import unicodedata
import re
import threading
import time
from datetime import datetime
import openpyxl
from openpyxl.styles import Side, Border, Alignment
from openpyxl.utils import get_column_letter

class App:
    def __init__(self, root):
        # Validate required constants
        required_constants = [
            'TXT_ULTIMO_ENVIO', 'TXT_ENVIO', 'TXT_COBRANCA', 'TXT_ULTIMO_AJUSTE',
            'TXT_ULTIMO_GERAR', 'FONT_PADRAO', 'EMAIL_PLANILHAS_THAILANY',
            'PASTA_PLANILHAS_THAILANY', 'EMAIL_CONTA_PADRAO_VICTOR',
            'PASTA_PLANILHAS_VICTOR', 'IDX_AK_LABEL', 'IDX_AL_ANEXOS',
            'IDX_AM_ASSUNTO', 'MAX_LINHAS_COLUNA', 'EXCLUIR_NA_COBRANCA', 'IDX_AA_STATUS'
        ]
        for const in required_constants:
            if not globals().get(const):
                raise NameError(f"Constant {const} is not defined")

        self.root = root
        self.root.title("E.A.I. - E-SOCIAL ARTIFICIAL INTELLIGENCE")
        self.root.geometry("500x500")
        centralizar(self.root)
        for txt_file in [TXT_ULTIMO_ENVIO, TXT_ENVIO, TXT_COBRANCA, TXT_ULTIMO_AJUSTE, TXT_ULTIMO_GERAR]:
            unhide_file(txt_file)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.overlay_login = tk.Frame(root, bg="lightgray")
        self.overlay_login.place(relx=0, rely=0, relwidth=1, relheight=1)
        tk.Label(self.overlay_login, text="Usuário:", font=FONT_PADRAO).pack(pady=5)
        self.usuario_entry = tk.Entry(self.overlay_login, font=FONT_PADRAO)
        self.usuario_entry.pack(pady=5)
        self.usuario_entry.focus_set()
        tk.Label(self.overlay_login, text="Senha:", font=FONT_PADRAO).pack(pady=5)
        self.senha_entry = tk.Entry(self.overlay_login, show="*", font=FONT_PADRAO)
        self.senha_entry.pack(pady=5)
        btn_entrar = tk.Button(self.overlay_login, text="Entrar", font=FONT_PADRAO, width=15, takefocus=True, command=self.validar_login)
        btn_entrar.pack(pady=10)
        self.usuario_entry.bind("<Return>", lambda e: self.validar_login())
        self.senha_entry.bind("<Return>", lambda e: self.validar_login())
        btn_entrar.bind("<Return>", lambda e: self.validar_login())
        print("🔗 Login screen Enter bindings set")
        self.usuario_atual = None
        self.df = pd.DataFrame()
        self.labels = []
        self.enviados = 0
        self.cobrados = 0
        self.baixados = 0
        self.ajustadas = 0
        self.frame_empresas_envio = None
        self.frame_empresas_cobranca = None
        self.gerar_pendentes_btn = None
        self.salvar_btn = None
        self.ajustar_btn = None
        self.envio_cobranca_btn = None

    def _on_enter_press(self, event, parent):
        focused = parent.focus_get()
        if isinstance(focused, tk.Button):
            print(f"🔗 Enter pressed: Invoking button '{focused.cget('text')}' in {parent}")
            focused.invoke()
        return "break"

    def on_close(self):
        self.root.unbind("<Return>")
        print("🔗 Enter key unbound on close")
        for txt_file in [TXT_ULTIMO_ENVIO, TXT_ENVIO, TXT_COBRANCA, TXT_ULTIMO_AJUSTE, TXT_ULTIMO_GERAR]:
            hide_file(txt_file)
        self.root.destroy()

    def validar_login(self):
        print("🔍 Iniciando validação de login")
        usuario = self.usuario_entry.get().strip().lower()
        senha = self.senha_entry.get().strip()
        USUARIOS = {"victor": "1001", "thailany": "1001"}
        if usuario in USUARIOS and senha == USUARIOS[usuario]:
            self.usuario_atual = usuario
            try:
                self.df = carregar_controle(usuario)
                self.labels = listar_labels(self.df)
                print(f"✅ Login bem-sucedido para {usuario}")
            except Exception as e:
                print(f"❌ Erro ao carregar planilha: {e}")
                messagebox.showerror("Erro", f"Falha ao carregar planilha: {e}")
                self.root.destroy()
                return
            self.overlay_login.destroy()
            messagebox.showinfo("Bem-vindo", f"Login realizado com sucesso, {usuario.capitalize()}!")
            print("📺 Chamando tela_inicial")
            self.tela_inicial()
            self.root.update()
        else:
            print("❌ Credenciais inválidas")
            messagebox.showerror("Erro", "Usuário ou senha incorretos!")

    def salvar_planilhas_outlook(self, botao, usuario):
        try:
            botao.pack_forget()
            counter_frame = tk.Frame(self.root, bg="lightblue")
            counter_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
            counter_label = tk.Label(counter_frame, text="Arquivos baixados: 0 de 0", font=("Arial", 12, "bold"), bg="lightblue", fg="black")
            counter_label.pack(pady=5)

            pasta_destino = filedialog.askdirectory(title="Selecione a pasta para salvar as planilhas")
            if not pasta_destino:
                messagebox.showwarning("Aviso", "Nenhuma pasta selecionada. A operação foi cancelada.")
                counter_frame.destroy()
                return

            if usuario == "thailany":
                email_conta = EMAIL_PLANILHAS_THAILANY
            else:
                email_conta = EMAIL_CONTA_PADRAO_VICTOR

            print(f"📧 Iniciando salvamento de planilhas para o usuário: {usuario}, email: {email_conta}")
            print(f"📂 Pasta de destino: {pasta_destino}")

            pythoncom.CoInitialize()
            outlook = win32.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            account_found = False
            for account in ns.Accounts:
                print(f"🔍 Verificando conta: {account.SmtpAddress}")
                if str(account.SmtpAddress).lower() == email_conta.lower():
                    inbox = account.DeliveryStore.GetDefaultFolder(6)
                    account_found = True
                    print(f"✅ Conta encontrada: {email_conta}")
                    break
            if not account_found:
                raise Exception(f"Conta de e-mail {email_conta} não encontrada no Outlook.")

            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            os.makedirs(pasta_destino, exist_ok=True)
            baixados = 0
            total_emails = 0
            for itm in items:
                try:
                    assunto = str(itm.Subject or "").strip()
                    if not assunto:
                        print("⏭️ E-mail ignorado: assunto vazio")
                        continue
                    assunto_normalizado = unicodedata.normalize("NFD", assunto).encode('ascii', 'ignore').decode('ascii').lower()
                    assunto_normalizado = re.sub(r'\s+', ' ', assunto_normalizado.strip())
                    if re.search(r'listagem\s*de\s*funcionarios', assunto_normalizado, re.IGNORECASE) and itm.Attachments.Count:
                        total_emails += 1
                except Exception:
                    continue
            print(f"📥 Total de e-mails relevantes: {total_emails}")
            counter_label.config(text=f"Arquivos baixados: 0 de {total_emails}")

            for itm in items:
                try:
                    assunto = str(itm.Subject or "").strip()
                    print(f"📄 Verificando e-mail com assunto: '{assunto}'")
                    if not assunto:
                        print("⏭️ E-mail ignorado: assunto vazio")
                        continue
                    assunto_normalizado = unicodedata.normalize("NFD", assunto).encode('ascii', 'ignore').decode('ascii').lower()
                    assunto_normalizado = re.sub(r'\s+', ' ', assunto_normalizado.strip())
                    if not re.search(r'listagem\s*de\s*funcionarios', assunto_normalizado, re.IGNORECASE):
                        print(f"⏭️ E-mail ignorado: assunto não contém 'listagem de funcionarios'")
                        continue

                    nome_base = re.sub(r'[<>:"/\\|?*]', '', assunto).strip()
                    print(f"📋 Nome base gerado: {nome_base}")

                    if not itm.Attachments.Count:
                        print(f"⏭️ E-mail ignorado: sem anexos")
                        continue

                    for att in itm.Attachments:
                        nome_att = str(att.FileName)
                        print(f"📎 Verificando anexo: {nome_att}")
                        if nome_att.lower().endswith(".zip"):
                            caminho_zip_temp = os.path.join(pasta_destino, nome_att)
                            print(f"💾 Salvando anexo ZIP: {caminho_zip_temp}")
                            att.SaveAsFile(caminho_zip_temp)
                            try:
                                with zipfile.ZipFile(caminho_zip_temp, 'r') as zip_ref:
                                    xlsx_files = [f for f in zip_ref.namelist() if f.lower().endswith(".xlsx")]
                                    if not xlsx_files:
                                        print(f"⚠️ ZIP ignorado: não contém arquivos .xlsx")
                                        os.remove(caminho_zip_temp)
                                        continue
                                    for zip_info in xlsx_files:
                                        nome_arquivo = f"{nome_base}.xlsx"
                                        caminho_final = os.path.join(pasta_destino, nome_arquivo)
                                        contador = 1
                                        while os.path.exists(caminho_final):
                                            nome_arquivo = f"{nome_base}_{contador}.xlsx"
                                            caminho_final = os.path.join(pasta_destino, nome_arquivo)
                                            contador += 1
                                        print(f"📤 Extraindo arquivo XLSX: {zip_info} para {caminho_final}")
                                        zip_ref.extract(zip_info, pasta_destino)
                                        caminho_extracao = os.path.join(pasta_destino, zip_info)
                                        os.rename(caminho_extracao, caminho_final)
                                        baixados += 1
                                        self.baixados = baixados
                                        print(f"🔢 Contador atualizado: Arquivos baixados: {baixados} de {total_emails}")
                                        self.root.after(0, lambda: counter_label.config(
                                            text=f"Arquivos baixados: {baixados} de {total_emails}"))
                                        print(f"✅ Arquivo extraído: {caminho_final}")
                                os.remove(caminho_zip_temp)
                                print(f"🗑️ Arquivo ZIP temporário removido: {caminho_zip_temp}")
                            except Exception as zip_err:
                                print(f"⚠️ Erro ao processar ZIP {caminho_zip_temp}: {zip_err}")
                                continue
                        else:
                            print(f"⏭️ Anexo ignorado: {nome_att} (não é .zip)")
                except Exception as item_err:
                    print(f"⚠️ Erro ao processar e-mail com assunto '{assunto}': {item_err}")
                    continue

            print(f"🏁 Total de arquivos baixados: {baixados}")
            with open(TXT_ULTIMO_ENVIO, "w", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%d/%m/%Y"))
            hide_file(TXT_ULTIMO_ENVIO)
            counter_frame.destroy()
            messagebox.showinfo("Salvar Planilhas", f"Concluído! {baixados} arquivos salvos na pasta com sucesso.")
        except Exception as e:
            print(f"❌ Erro geral: {e}")
            if 'counter_frame' in locals():
                counter_frame.destroy()
            messagebox.showerror("Erro", f"Falha ao salvar planilhas: {e}")
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
        try:
            counter_frame = tk.Frame(self.root, bg="lightblue")
            counter_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
            counter_label = tk.Label(counter_frame, text="Planilhas ajustadas: 0 de 0", font=("Arial", 12, "bold"), bg="lightblue", fg="black")
            counter_label.pack(pady=5)

            arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]
            total_arquivos = len(arquivos)
            print(f"📂 Total de planilhas a ajustar: {total_arquivos}")
            counter_label.config(text=f"Planilhas ajustadas: 0 de {total_arquivos}")
            if not arquivos:
                counter_frame.destroy()
                messagebox.showinfo("Ajustar Planilhas", "Nenhuma planilha encontrada na pasta.")
                return
            ajustadas = 0
            for arquivo in arquivos:
                caminho = os.path.join(pasta, arquivo)
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
                                            data_val = datetime.strptime(str(data_cell), "%d/%m/%Y")
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
                    for r in range(1, ws.max_row + 1):
                        max_height = 23
                        for c in range(1, ws.max_column + 1):
                            try:
                                cell = ws.cell(row=r, column=c)
                                is_in_merged = False
                                for merged_range in merged_cells:
                                    if (r >= merged_range.min_row and r <= merged_range.max_row and
                                        c >= merged_range.min_col and c <= merged_range.max_col):
                                        is_in_merged = True
                                        break
                                if not is_in_merged and cell.value:
                                    linhas = str(cell.value).count("\n") + 1
                                    max_height = max(max_height, linhas * 15)
                            except BaseException:
                                continue
                        ws.row_dimensions[r].height = max_height
                    for col_idx in range(1, ws.max_column + 1):
                        max_length = 0
                        column_letter = get_column_letter(col_idx)
                        for row_idx in range(1, ws.max_row + 1):
                            try:
                                is_in_merged = False
                                for merged_range in ws.merged_cells.ranges:
                                    if (row_idx >= merged_range.min_row and row_idx <= merged_range.max_row and
                                        col_idx >= merged_range.min_col and col_idx <= merged_range.max_col):
                                        is_in_merged = True
                                        break
                                if not is_in_merged:
                                    cell = ws.cell(row=row_idx, column=col_idx)
                                    if cell.value:
                                        if '\n' in str(cell.value):
                                            lines = str(cell.value).split('\n')
                                            line_length = max(len(line) for line in lines)
                                        else:
                                            line_length = len(str(cell.value))
                                        if line_length > max_length:
                                            max_length = line_length
                            except BaseException:
                                continue
                        if max_length > 0:
                            adjusted_width = (max_length + 2) * 1.2
                            ws.column_dimensions[column_letter].width = min(adjusted_width, 50)
                    thin = Side(border_style="thin", color="000000")
                    border = Border(top=thin, left=thin, right=thin, bottom=thin)
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
                                    cell = ws.cell(row=row_idx, column=col_idx)
                                    cell.border = border
                                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                    if col_idx == 7:
                                        cell.number_format = 'DD/MM/YYYY'
                                    else:
                                        cell.number_format = '@'
                            except BaseException:
                                continue
                wb.save(caminho)
                ajustadas += 1
                self.ajustadas = ajustadas
                print(f"🔢 Contador atualizado: Planilhas ajustadas: {ajustadas} de {total_arquivos}")
                self.root.after(0, lambda: counter_label.config(text=f"Planilhas ajustadas: {ajustadas} de {total_arquivos}"))
                print(f"✅ Planilha ajustada: {caminho}")
            pythoncom.CoInitialize()
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            try:
                for arquivo in arquivos:
                    caminho_completo = os.path.join(pasta, arquivo)
                    wb_excel = excel.Workbooks.Open(caminho_completo)
                    for ws in wb_excel.Worksheets:
                        ws.Protect(SENHA_PROTECAO)
                    wb_excel.Protect(SENHA_PROTECAO)
                    wb_excel.Save()
                    wb_excel.Close()
            except Exception as e:
                counter_frame.destroy()
                messagebox.showerror("Erro", f"Erro ao aplicar senha: {e}")
                return
            finally:
                excel.Quit()
                pythoncom.CoUninitialize()
            with open(TXT_ULTIMO_AJUSTE, "w", encoding="utf-8") as f:
                f.write(datetime.now().strftime("%d/%m/%Y"))
            hide_file(TXT_ULTIMO_AJUSTE)
            counter_frame.destroy()
            messagebox.showinfo("Ajustar Planilhas", f"{ajustadas} planilhas ajustadas e protegidas com sucesso!\nSenha: {SENHA_PROTECAO}")
        except Exception as e:
            print(f"❌ Erro geral: {e}")
            if 'counter_frame' in locals():
                counter_frame.destroy()
            messagebox.showerror("Erro", f"Falha ao ajustar planilhas: {e}")

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
        btn = tk.Button(frame, text="Salvar Planilhas", width=30, font=FONT_PADRAO, takefocus=True)
        btn.pack(pady=6)
        print(f"🔧 Configurando botão Salvar Planilhas para usuário: {usuario}")
        btn.configure(command=lambda b=btn: threading.Thread(target=self.salvar_planilhas_outlook, args=(b, usuario), daemon=True).start())
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
        btn = tk.Button(frame, text="Ajustar Planilhas", width=30, font=FONT_PADRAO, takefocus=True)
        btn.pack(pady=6)
        btn.configure(command=lambda b=btn: threading.Thread(target=self._ajustar_planilhas_thread, args=(b,), daemon=True).start())
        return btn

    def tela_inicial(self):
        self.root.unbind("<Return>")
        print("📺 Renderizando tela_inicial")
        for w in self.root.winfo_children():
            w.destroy()
        tk.Label(self.root, text="Escolha uma opção", font=FONT_PADRAO).pack(pady=20)
        self.gerar_pendentes_btn = chekar_botao_gerar_pendentes(self, frame=self.root, usuario=self.usuario_atual)
        if self.gerar_pendentes_btn:
            self.gerar_pendentes_btn.focus_set()
            print("🔗 Initial focus set to Gerar Pendentes button")
        self.salvar_btn = self.chekar_botao_salvar(self.root, self.usuario_atual)
        self.ajustar_btn = self.chekar_botao_ajustar(self.root)
        self.envio_cobranca_btn = tk.Button(self.root, text="Envio / Cobrança", width=30, font=FONT_PADRAO, takefocus=True, command=self.tela_envio_cobranca)
        self.envio_cobranca_btn.pack(pady=6)
        rodape = tk.Frame(self.root)
        rodape.pack(side=tk.BOTTOM, fill=tk.X)
        sair_btn = tk.Button(rodape, text="Sair", font=FONT_PADRAO, width=10, takefocus=True, command=self.root.destroy)
        sair_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=10)
        self.root.bind("<Return>", lambda e: self._on_enter_press(e, self.root))
        print("🔗 Enter key bound for tela_inicial buttons")
        self.root.update()

    def tela_envio_cobranca(self):
        self.root.unbind("<Return>")
        print("📺 Renderizando tela_envio_cobranca")
        win = tk.Toplevel(self.root)
        win.title("Envio ou Cobrança")
        win.geometry("420x220")
        centralizar(win)
        tk.Label(win, text="Escolha a ação", font=FONT_PADRAO).pack(pady=20)
        envio_btn = tk.Button(win, text="Envio", width=22, font=FONT_PADRAO, takefocus=True, command=self.tela_empresas_envio)
        envio_btn.pack(pady=10)
        cobranca_btn = tk.Button(win, text="Cobrança", width=22, font=FONT_PADRAO, takefocus=True, command=self.tela_empresas_cobranca)
        cobranca_btn.pack(pady=6)
        rodape = tk.Frame(win)
        rodape.pack(side=tk.BOTTOM, fill=tk.X)
        voltar_btn = tk.Button(rodape, text="Voltar", font=FONT_PADRAO, width=10, takefocus=True, command=win.destroy)
        voltar_btn.pack(side=tk.LEFT, anchor="sw", padx=10, pady=10)
        win.bind("<Return>", lambda e: self._on_enter_press(e, win))
        print("🔗 Enter key bound for tela_envio_cobranca buttons")
        win.focus_set()
        envio_btn.focus_set()

    def _bind_horizontal_scroll(self, widget_canvas):
        widget_canvas.bind("<Enter>", lambda e: self._enable_mousewheel(widget_canvas))
        widget_canvas.bind("<Leave>", lambda e: self._disable_mousewheel(widget_canvas))

    def _enable_mousewheel(self, canvas_widget):
        import sys
        if sys.platform.startswith("win") or sys.platform == "darwin":
            canvas_widget.bind_all("<MouseWheel>", lambda e: self._on_mousewheel_horizontal(e, canvas_widget))
        else:
            canvas_widget.bind_all("<Button-4>", lambda e: self._on_mousewheel_horizontal(e, canvas_widget))
            canvas_widget.bind_all("<Button-5>", lambda e: self._on_mousewheel_horizontal(e, canvas_widget))

    def _disable_mousewheel(self, canvas_widget):
        import sys
        if sys.platform.startswith("win") or sys.platform == "darwin":
            canvas_widget.unbind_all("<MouseWheel>")
        else:
            canvas_widget.unbind_all("<Button-4>")
            canvas_widget.unbind_all("<Button-5>")

    def _on_mousewheel_horizontal(self, event, canvas_widget):
        try:
            if hasattr(event, "delta") and event.delta is not None:
                move = int(-1 * (event.delta / 120))
                canvas_widget.xview_scroll(move, "units")
            else:
                if hasattr(event, "num"):
                    if event.num == 4:
                        canvas_widget.xview_scroll(-1, "units")
                    elif event.num == 5:
                        canvas_widget.xview_scroll(1, "units")
        except Exception as e:
            print(f"⚠️ Error in _on_mousewheel_horizontal: {e}")

    def tela_empresas_envio(self):
        self.root.unbind("<Return>")
        print("📺 Renderizando tela_empresas_envio")
        win = tk.Toplevel(self.root)
        win.title("Envio da Listagem de Funcionários")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1100x700")
        frame_top = tk.Frame(win)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=10)
        voltar_btn = tk.Button(frame_top, text="Voltar", font=FONT_PADRAO, width=10, takefocus=True, command=lambda: (win.destroy(), self.tela_envio_cobranca()))
        voltar_btn.pack(side=tk.LEFT, anchor="nw", padx=10, pady=10)
        tk.Label(frame_top, text="ENVIO DA LISTAGEM DE FUNCIONÁRIOS", font=("Arial", 16, "bold")).pack(pady=6)
        btn_todos = tk.Button(frame_top, text="ENVIAR TODOS", bg="#1f5fbf", fg="white", font=FONT_PADRAO, takefocus=True, command=lambda: threading.Thread(target=self._enviar_todos_thread, daemon=True).start())
        btn_todos.pack(pady=6)
        registros_envio = carregar_txt(TXT_ENVIO)
        botoes = [lbl for lbl in self.labels if not self._enviado_no_mes(lbl, registros_envio) and not self._has_ok_status(lbl)]
        botoes.sort(key=normalize_key, reverse=True)  # Updated for descending order
        self.contador_label_envio = tk.Label(frame_top, text=f"{self.enviados}/{len(botoes)}", font=FONT_PADRAO)
        self.contador_label_envio.pack(pady=2)
        container = tk.Frame(win)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        hbar = tk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(xscrollcommand=hbar.set)
        frame_empresas = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_empresas, anchor="nw")
        self.frame_empresas_envio = frame_empresas
        for i, label in enumerate(botoes):
            r = i % MAX_LINHAS_COLUNA
            c = i // MAX_LINHAS_COLUNA
            btn = tk.Button(frame_empresas, text=label, font=FONT_PADRAO, width=36, height=5, wraplength=320, takefocus=True)
            btn.configure(command=lambda l=label, b=btn: threading.Thread(target=self._enviar_uma_label_thread, args=(l, b), daemon=True).start())
            btn.grid(row=r, column=c, padx=10, pady=10)
        def _atualiza_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_empresas.bind("<Configure>", _atualiza_scroll)
        self._bind_horizontal_scroll(canvas)
        win.bind("<Return>", lambda e: self._on_enter_press(e, win))
        print("🔗 Enter key bound for tela_empresas_envio buttons")
        centralizar(win)
        win.focus_set()
        voltar_btn.focus_set()

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
                messagebox.showerror("Erro", f"Falha ao enviar '{label}': {err}")
            remaining = len([w for w in self.frame_empresas_envio.winfo_children() if isinstance(w, tk.Button)])
            self.contador_label_envio.config(text=f"{self.enviados}/{remaining}")
        self.root.after(0, ui_update)

    def _enviar_todos_thread(self):
        if not self.frame_empresas_envio:
            return
        widgets = [w for w in self.frame_empresas_envio.winfo_children() if isinstance(w, tk.Button)]
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
            success, err = enviar_email_label(label, self.df, self.usuario_atual)
            def ui_update_after_one(s=success, e=err, widget=w, lbl=label):
                if s:
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                    self.enviados += 1
                else:
                    messagebox.showerror("Erro", f"Falha ao enviar '{lbl}': {e}")
                remaining = len([w2 for w2 in self.frame_empresas_envio.winfo_children() if isinstance(w2, tk.Button)])
                self.contador_label_envio.config(text=f"{self.enviados}/{remaining}")
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
        self.root.unbind("<Return>")
        print("📺 Renderizando tela_empresas_cobranca")
        win = tk.Toplevel(self.root)
        win.title("COBRANÇA DA LISTAGEM DE FUNCIONÁRIOS")
        try:
            win.state('zoomed')
        except Exception:
            win.geometry("1100x700")
        frame_top = tk.Frame(win)
        frame_top.pack(side=tk.TOP, fill=tk.X, pady=10)
        voltar_btn = tk.Button(frame_top, text="Voltar", font=FONT_PADRAO, width=10, takefocus=True, command=lambda: (win.destroy(), self.tela_envio_cobranca()))
        voltar_btn.pack(side=tk.LEFT, anchor="nw", padx=10, pady=10)
        tk.Label(frame_top, text="COBRANÇA DA LISTAGEM DE FUNCIONÁRIOS", font=("Arial", 16, "bold")).pack(pady=6)
        btn_todos = tk.Button(frame_top, text="COBRAR TODOS", bg="#1f5fbf", fg="white", font=FONT_PADRAO, takefocus=True, command=lambda: threading.Thread(target=self._cobrar_todos_thread, daemon=True).start())
        btn_todos.pack(pady=6)
        registros_cobranca = carregar_txt(TXT_COBRANCA)
        botoes = [lbl for lbl in self.labels if not self._cobrado_hoje(lbl, registros_cobranca) and lbl not in EXCLUIR_NA_COBRANCA and not self._has_ok_status(lbl)]
        botoes.sort(key=normalize_key, reverse=True)  # Updated for descending order
        self.contador_label_cobranca = tk.Label(frame_top, text=f"{self.cobrados}/{len(botoes)}", font=FONT_PADRAO)
        self.contador_label_cobranca.pack(pady=2)
        container = tk.Frame(win)
        container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(container)
        canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        hbar = tk.Scrollbar(container, orient=tk.HORIZONTAL, command=canvas.xview)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.configure(xscrollcommand=hbar.set)
        frame_empresas = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_empresas, anchor="nw")
        self.frame_empresas_cobranca = frame_empresas
        for i, label in enumerate(botoes):
            r = i % MAX_LINHAS_COLUNA
            c = i // MAX_LINHAS_COLUNA
            btn = tk.Button(frame_empresas, text=label, font=FONT_PADRAO, width=36, height=5, wraplength=320, takefocus=True)
            btn.configure(command=lambda l=label, b=btn: threading.Thread(target=self._cobrar_uma_label_thread, args=(l, b), daemon=True).start())
            btn.grid(row=r, column=c, padx=10, pady=10)
        def _atualiza_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        frame_empresas.bind("<Configure>", _atualiza_scroll)
        self._bind_horizontal_scroll(canvas)
        win.bind("<Return>", lambda e: self._on_enter_press(e, win))
        print("🔗 Enter key bound for tela_empresas_cobranca buttons")
        self.contador_label_cobranca.config(text=f"{self.cobrados}/{len(botoes)}")
        centralizar(win)
        win.focus_set()
        voltar_btn.focus_set()

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
                messagebox.showerror("Erro", f"Falha ao cobrar '{label}': {err}")
            remaining = len([w for w in self.frame_empresas_cobranca.winfo_children() if isinstance(w, tk.Button)])
            self.contador_label_cobranca.config(text=f"{self.cobrados}/{remaining}")
        self.root.after(0, ui_update)

    def _cobrar_todos_thread(self):
        if not self.frame_empresas_cobranca:
            return
        widgets = [w for w in self.frame_empresas_cobranca.winfo_children() if isinstance(w, tk.Button)]
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
            success, err = cobrar_email_label(label, self.df, self.usuario_atual)
            def ui_update_after_one(s=success, e=err, widget=w, lbl=label):
                if s:
                    try:
                        widget.destroy()
                    except Exception:
                        pass
                    self.cobrados += 1
                else:
                    messagebox.showerror("Erro", f"Falha ao cobrar '{lbl}': {e}")
                remaining = len([w2 for w2 in self.frame_empresas_cobranca.winfo_children() if isinstance(w2, tk.Button)])
                self.contador_label_cobranca.config(text=f"{self.cobrados}/{remaining}")
            self.root.after(0, ui_update_after_one)
            time.sleep(0.4)

def listar_labels(df):
    """Retorna a lista de labels da planilha (coluna IDX_AK_LABEL), removendo nulos e ordenando de forma decrescente."""
    if IDX_AK_LABEL >= df.shape[1]:
        raise ValueError(f"Índice {IDX_AK_LABEL} fora dos limites. Colunas disponíveis: {df.shape[1]}")
    labels = list(df.iloc[2:, IDX_AK_LABEL].dropna().astype(str))
    labels.sort(key=normalize_key, reverse=True)
    return labels

def get_row_by_label(df, label):
    if IDX_AK_LABEL >= df.shape[1]:
        raise ValueError(f"Índice {IDX_AK_LABEL} fora dos limites. Colunas disponíveis: {df.shape[1]}")
    mask = df.iloc[:, IDX_AK_LABEL].astype(str) == label
    if not mask.any():
        print(f"Label '{label}' não encontrado na coluna {IDX_AK_LABEL}")
        return None
    return df[mask].iloc[0]

def get_anexos_from_row(row, usuario):
    try:
        if IDX_AL_ANEXOS >= len(row):
            return []
        raw = str(row.iloc[IDX_AL_ANEXOS])
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
        if IDX_AM_ASSUNTO >= len(row):
            return ""
        raw = str(row.iloc[IDX_AM_ASSUNTO])
    except Exception:
        raw = ""
    return safe_email_field(raw)

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

quero que ao inves de ter fixo o repositorio de salvamento das planilhas quando clico em salvar planilhas, quero que seja aberta uma caixa de seleção de repositorio para que eu escolha o local. Ou seja, remova os caminhos fixos nesta função. atenção, não mexa em mais nada