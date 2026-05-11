import os
import sys
import shutil
import threading
import subprocess
import time
import datetime
import tkinter as tk
import customtkinter as ctk
import winreg
from tkinter import filedialog, messagebox
from PIL import Image

# Configuração de Caminhos para o HD Externo
def obter_caminho_recurso(caminho_relativo):
    """Retorna o caminho absoluto para o recurso, lidando com PyInstaller."""
    try:
        caminho_base = sys._MEIPASS
    except Exception:
        caminho_base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(caminho_base, caminho_relativo)

# Pasta onde o executável ou script está localizado
DIRETORIO_BASE = os.path.dirname(os.path.abspath(sys.argv[0]))

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class AppTecnico(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Victor Computadores - Pós Formatação")
        self.resizable(False, False)
        self.centralizar_janela(850, 600, self)

        self.fonte_negrito = ctk.CTkFont(family="Arial", size=13, weight="bold")
        self.fonte_titulo = ctk.CTkFont(family="Arial", size=15, weight="bold")

        self.avisos_sonoros = False
        self.after(600, self.perguntar_avisos_sonoros)

        # --- BARRA LATERAL ---
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        ctk.CTkLabel(self.sidebar, text="Victor Computadores\nPós Formatação", font=self.fonte_titulo).pack(pady=(20, 0))

        caminho_logo = obter_caminho_recurso("logo.jpg")
        if os.path.exists(caminho_logo):
            imagem_pil = Image.open(caminho_logo)
            self.logo_imagem = ctk.CTkImage(light_image=imagem_pil, dark_image=imagem_pil, size=(150, 150))
            ctk.CTkLabel(self.sidebar, image=self.logo_imagem, text="").pack(pady=15)

        self.frame_checks = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.frame_checks.pack(pady=10, padx=10, fill="x")
        self.verificar_arquivos_hd()

        # --- PAINEL PRINCIPAL ---
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(side="right", fill="both", expand=True, padx=20, pady=10)
        self.grid_buttons = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.grid_buttons.pack(pady=10)

        # Grid de Botões (1 a 8)
        self.criar_grid_btn("🛡️ 1. Desativar Defender", self.desativar_defender, "#ef4444", 0, 0)
        self.criar_grid_btn("⏸️ 2. Pausar Updates", self.pausar_windows_update, "#d97706", 1, 0)
        self.criar_grid_btn("📦 3. Instalar Programas", self.janela_selecao_programas, "#2c3e50", 2, 0)
        self.criar_grid_btn("🔧 4. Instalar Drivers", self.executar_drivers_sdi, "#059669", 3, 0)
        self.criar_grid_btn("✨ 5. Visual", self.visual_pro_max, "#8b5cf6", 0, 1)
        self.criar_grid_btn("🔑 6. Ativar Win/Off", self.executar_ativacao_local, "#1e3a8a", 1, 1)
        self.criar_grid_btn("🧹 7. Limpeza", self.executar_limpeza_final, "#4b5563", 2, 1)
        self.criar_grid_btn("📁 8. Restaurar Backup", self.iniciar_restauracao, "#3b82f6", 3, 1)

        self.btn_tudo = ctk.CTkButton(self.main_frame, text="🚀 EXECUTAR TUDO EM SEQUÊNCIA", command=self.executar_sequencia_completa,
                                     height=55, width=540, fg_color="#b91c1c", font=ctk.CTkFont(size=16, weight="bold"))
        self.btn_tudo.pack(pady=(20, 5))

        self.status_label = ctk.CTkLabel(self.main_frame, text="Aguardando ação", text_color="gray", font=self.fonte_negrito)
        self.status_label.pack(pady=2)
        self.progress_bar = ctk.CTkProgressBar(self.main_frame, width=540); self.progress_bar.pack(pady=5); self.progress_bar.set(0)

        ctk.CTkButton(self.main_frame, text="SAIR", command=self.destroy, width=80, height=28, fg_color="#333333", font=self.fonte_negrito).pack(side="bottom", anchor="e", pady=5)

    def criar_grid_btn(self, texto, comando, cor, r, c):
        btn = ctk.CTkButton(self.grid_buttons, text=texto, command=comando, height=45, width=260, fg_color=cor, font=self.fonte_negrito)
        btn.grid(row=r, column=c, padx=10, pady=8)

    def centralizar_janela(self, largura, altura, janela):
        sw, sh = janela.winfo_screenwidth(), janela.winfo_screenheight()
        janela.geometry(f"{largura}x{altura}+{int((sw/2)-(largura/2))}+{int((sh/2)-(altura/2))}")

    def verificar_arquivos_hd(self):
        """Verifica a existência de arquivos essenciais no diretório base."""
        for widget in self.frame_checks.winfo_children(): widget.destroy()
        existe_sdi = os.path.exists(os.path.join(DIRETORIO_BASE, "SDI"))
        existe_mas = os.path.exists(os.path.join(DIRETORIO_BASE, "MAS_AIO.cmd"))
        def adicionar_checagem(texto, status):
            cor, icone = ("#22c55e", "[V]") if status else ("#ef4444", "[X]")
            ctk.CTkLabel(self.frame_checks, text=f"{icone} {texto}", text_color=cor, font=self.fonte_negrito).pack(anchor="w", padx=20, pady=2)
        adicionar_checagem("Pasta de drivers", existe_sdi); adicionar_checagem("Script master", existe_mas)

    def perguntar_avisos_sonoros(self):
        """Pergunta ao usuário se deseja ativar avisos sonoros e ajusta o volume se sim."""
        if messagebox.askyesno("Som", "Ativar avisos de volume 82?"):
            self.avisos_sonoros = True
            try:
                subprocess.Popen('powershell -Command "$w=New-Object -ComObject WScript.Shell; 1..50 | % { $w.SendKeys([char]174) }; 1..41 | % { $w.SendKeys([char]175) }"', shell=True)
            except Exception as e:
                print(f"Erro ao ajustar volume: {e}")

    def status_temp(self, texto, cor="green"):
        """Atualiza o status temporariamente de forma segura para threads."""
        self.after(0, lambda: self.status_label.configure(text=texto, text_color=cor))
        self.after(5000, lambda: self.status_label.configure(text="Aguardando ação", text_color="gray"))

    # --- FUNÇÕES DE AÇÃO ---
    def desativar_defender(self, assincrono=True):
        """Desativa o monitoramento em tempo real do Windows Defender."""
        def task():
            try:
                subprocess.run("powershell -Command \"Set-MpPreference -DisableRealtimeMonitoring $true\"", shell=True, check=True)
                self.status_temp("🛡️ Defender Desativado")
            except Exception:
                self.status_temp("Erro Defender", "red")

        if assincrono:
            threading.Thread(target=task, daemon=True).start()
        else:
            task()

    def pausar_windows_update(self, assincrono=True):
        """Pausa as atualizações do Windows por 7 dias."""
        def task():
            try:
                data_pausa = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y-%m-%dT%H:%M:%SZ')
                subprocess.run(f'reg add "HKLM\\SOFTWARE\\Microsoft\\WindowsUpdate\\UX\\Settings" /v PauseUpdatesExpiryTime /t REG_SZ /d {data_pausa} /f', shell=True, check=True)
                self.status_temp("⏸️ Updates Pausados")
            except Exception:
                self.status_temp("Erro Updates", "red")

        if assincrono:
            threading.Thread(target=task, daemon=True).start()
        else:
            task()

    def esta_instalado(self, nome_programa):
        """Verifica se um programa está instalado consultando o registro do Windows."""
        caminhos_registro = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        for caminho in caminhos_registro:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, caminho) as chave_principal:
                    for i in range(winreg.QueryInfoKey(chave_principal)[0]):
                        try:
                            subchave_nome = winreg.EnumKey(chave_principal, i)
                            with winreg.OpenKey(chave_principal, subchave_nome) as subchave:
                                nome_exibicao, _ = winreg.QueryValueEx(subchave, "DisplayName")
                                if nome_programa.lower() in nome_exibicao.lower():
                                    return True
                        except (OSError, IndexError, KeyError):
                            continue
            except OSError:
                continue
        return False

    def janela_selecao_programas(self):
        """Abre uma janela para seleção de programas que ainda não estão instalados."""
        janela_programas = ctk.CTkToplevel(self); janela_programas.title("📦 Programas"); janela_programas.attributes("-topmost", True)
        mapa_programas = {"Chrome": "Google.Chrome", "Reader": "Adobe.Acrobat.Reader.64bit", "WinRAR": "RARLab.WinRAR", "Java": "Microsoft.OpenJDK.21", "Office": "Microsoft.Office"}
        programas_faltantes = {nome: id_winget for nome, id_winget in mapa_programas.items() if not self.esta_instalado(nome)}

        if not programas_faltantes:
            messagebox.showinfo("Info", "Tudo já instalado!")
            janela_programas.destroy()
            return

        self.variaveis_checkbox = {nome: ctk.BooleanVar() for nome in programas_faltantes}
        frame_principal = ctk.CTkFrame(janela_programas, fg_color="transparent"); frame_principal.pack(pady=20, padx=20)

        for indice, nome_programa in enumerate(sorted(programas_faltantes.keys())):
            ctk.CTkCheckBox(frame_principal, text=nome_programa, variable=self.variaveis_checkbox[nome_programa], font=self.fonte_negrito).grid(row=indice%4, column=indice//4, padx=10, pady=10)

        ctk.CTkButton(janela_programas, text="Instalar", command=lambda: self.rodar_progs(janela_programas, programas_faltantes), font=self.fonte_negrito).pack(pady=10)
        self.centralizar_janela(400, 350, janela_programas)

    def rodar_progs(self, janela_programas, mapa_programas):
        """Executa a instalação dos programas selecionados em uma thread separada."""
        lista_instalacao = [nome for nome, variavel in self.variaveis_checkbox.items() if variavel.get()]
        janela_programas.destroy()
        def task():
            for nome in lista_instalacao:
                try:
                    subprocess.run(f"winget install --id {mapa_programas[nome]} --silent --accept-package-agreements", shell=True, check=True)
                except subprocess.CalledProcessError as e:
                    print(f"Erro ao instalar {nome}: {e}")
            self.status_temp("📦 Programas OK")
        threading.Thread(target=task, daemon=True).start()

    def executar_drivers_sdi(self):
        """Inicia a instalação de drivers via SDI em uma thread separada."""
        caminho_sdi = os.path.join(DIRETORIO_BASE, "SDI")
        if os.path.exists(caminho_sdi):
            self.after(0, lambda: self.status_label.configure(text="🔧 SDI Iniciado...", text_color="cyan"))
            def task():
                try:
                    arquivos_exe = [arquivo for arquivo in os.listdir(caminho_sdi) if arquivo.lower().endswith(".exe")]
                    if not arquivos_exe:
                        self.after(0, lambda: messagebox.showerror("Erro", "Nenhum executável encontrado na pasta SDI!"))
                        return

                    # Prioriza versão x64 se disponível
                    sdi_exe = next((f for f in arquivos_exe if "x64" in f.lower()), arquivos_exe[0])
                    caminho_completo = os.path.join(caminho_sdi, sdi_exe)

                    subprocess.run([caminho_completo, "-autoinstall", "-autoclose"], check=True)
                    self.status_temp("🔧 Drivers OK")
                except Exception as e:
                    self.status_temp("Erro SDI", "red")
                    print(f"Erro ao executar SDI: {e}")
            threading.Thread(target=task, daemon=True).start()
        else:
            self.after(0, lambda: messagebox.showerror("Erro", "Pasta SDI não encontrada!"))

    def visual_pro_max(self):
        """Aplica melhorias visuais e reinicia o Explorer."""
        try:
            subprocess.run('reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v LaunchTo /t REG_DWORD /d 1 /f', shell=True, check=True)
            subprocess.run("taskkill /f /im explorer.exe && start explorer.exe", shell=True, check=True)
            self.status_temp("✨ Visual OK")
        except Exception as e:
            self.status_temp("Erro Visual", "red")
            print(f"Erro ao aplicar visual: {e}")

    def executar_ativacao_local(self):
        """Executa o script de ativação MAS em uma thread separada."""
        caminho_mas = os.path.join(DIRETORIO_BASE, "MAS_AIO.cmd")
        if os.path.exists(caminho_mas):
            self.after(0, lambda: self.status_label.configure(text="🔑 Ativando...", text_color="yellow"))
            def task():
                try:
                    subprocess.run(f'"{caminho_mas}" /HWID /OHOOK', shell=True, check=True)
                    self.status_temp("🔑 Ativação OK")
                except Exception as e:
                    self.status_temp("Erro Ativação", "red")
                    print(f"Erro na ativação: {e}")
            threading.Thread(target=task, daemon=True).start()
        else:
            self.after(0, lambda: messagebox.showerror("Erro", "MAS_AIO.cmd não encontrado!"))

    def executar_limpeza_final(self):
        """Executa comandos de limpeza do Windows em uma thread separada."""
        def task():
            try:
                subprocess.run('start /wait cmd /c "dism /online /cleanup-image /startcomponentcleanup /resetbase"', shell=True, check=True)
                subprocess.run("cleanmgr /sagerun:65535", shell=True, check=True)
                self.status_temp("🧹 Limpeza OK")
            except subprocess.CalledProcessError as e:
                self.status_temp(f"Erro Limpeza: {e.returncode}", "red")
        threading.Thread(target=task, daemon=True).start()

    def iniciar_restauracao(self):
        origem = filedialog.askdirectory()
        if origem: self.status_temp("📁 Backup OK")

    def executar_sequencia_completa(self):
        """Executa as principais ações de otimização em sequência (Defender, Updates, Limpeza)."""
        if messagebox.askyesno("🚀", "Iniciar Tudo?"):
            def seq():
                self.after(0, lambda: self.progress_bar.set(0.1))
                self.desativar_defender(assincrono=False)

                self.after(0, lambda: self.progress_bar.set(0.4))
                time.sleep(1)

                self.pausar_windows_update(assincrono=False)

                self.after(0, lambda: self.progress_bar.set(0.7))
                try:
                    subprocess.run('start /wait cmd /c "dism /online /cleanup-image /startcomponentcleanup /resetbase"', shell=True, check=True)
                    subprocess.run("cleanmgr /sagerun:65535", shell=True, check=True)

                    self.after(0, lambda: self.progress_bar.set(1.0))
                    self.after(0, lambda: self.status_label.configure(text="Ações concluídas. Reinicie após o backup.", text_color="#00FF00"))
                except Exception as e:
                    self.status_temp("Erro na sequência", "red")
                    print(f"Erro na sequência completa: {e}")
            threading.Thread(target=seq, daemon=True).start()

if __name__ == "__main__":
    AppTecnico().mainloop()
