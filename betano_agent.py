import time
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# E.A.I. - Betano Monitoring Agent GUI (Agente de Monitoramento Betano com Interface)
# Focado em perfil conservador, consistente e banca baixa.

class BetanoAgent:
    def __init__(self, bankroll=100.0, risk_percent=0.02):
        self.bankroll = bankroll
        self.risk_percent = risk_percent
        self.url = "https://www.betano.bet.br/sport/futebol/"
        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless")
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        self.chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    def get_matches(self, mock=False):
        if mock:
            time.sleep(1) # Simula delay
            return self._get_mock_data()

        driver = None
        try:
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.get(self.url)
            wait = WebDriverWait(driver, 20)
            elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='event-row']")))

            matches = []
            for el in elements:
                try:
                    name = el.find_element(By.CSS_SELECTOR, "[data-testid='event-title']").text
                    odds_elements = el.find_elements(By.CSS_SELECTOR, "[data-testid='odds-button']")
                    if len(odds_elements) >= 3:
                        odds = {
                            '1 (Casa)': float(odds_elements[0].text.replace(',', '.')),
                            'X (Empate)': float(odds_elements[1].text.replace(',', '.')),
                            '2 (Fora)': float(odds_elements[2].text.replace(',', '.'))
                        }
                        matches.append({'name': name, 'odds': odds})
                except Exception:
                    continue
            return matches
        except Exception as e:
            raise e
        finally:
            if driver:
                driver.quit()

    def _get_mock_data(self):
        return [
            {'name': 'Flamengo vs Palmeiras', 'odds': {'1 (Casa)': 2.10, 'X (Empate)': 3.40, '2 (Fora)': 3.60}},
            {'name': 'Real Madrid vs Getafe', 'odds': {'1 (Casa)': 1.25, 'X (Empate)': 5.50, '2 (Fora)': 12.00}},
            {'name': 'Man City vs Everton', 'odds': {'1 (Casa)': 1.15, 'X (Empate)': 7.00, '2 (Fora)': 18.00}},
            {'name': 'Bayern vs Bochum', 'odds': {'1 (Casa)': 1.12, 'X (Empate)': 9.00, '2 (Fora)': 25.00}},
            {'name': 'Juventus vs Lecce', 'odds': {'1 (Casa)': 1.45, 'X (Empate)': 4.20, '2 (Fora)': 8.50}},
            {'name': 'Arsenal vs Wolves', 'odds': {'1 (Casa)': 1.35, 'X (Empate)': 4.80, '2 (Fora)': 9.00}},
            {'name': 'PSG vs Metz', 'odds': {'1 (Casa)': 1.18, 'X (Empate)': 6.50, '2 (Fora)': 15.00}},
            {'name': 'Inter vs Verona', 'odds': {'1 (Casa)': 1.22, 'X (Empate)': 5.80, '2 (Fora)': 13.00}},
        ]

    def analisar_oportunidades(self, mock=False):
        matches = self.get_matches(mock=mock)
        simples = []
        super_seguros = []

        for m in matches:
            for mercado, odd in m['odds'].items():
                if 1.20 <= odd <= 1.50:
                    simples.append({'name': m['name'], 'mercado': mercado, 'odd': odd})
                elif 1.05 <= odd < 1.20:
                    super_seguros.append({'name': m['name'], 'mercado': mercado, 'odd': odd})

        duplas = []
        if len(super_seguros) >= 2:
            s1 = super_seguros[0]
            s2 = super_seguros[1]
            total_odd = s1['odd'] * s2['odd']
            duplas.append({
                'match1': s1['name'], 'mercado1': s1['mercado'], 'odd1': s1['odd'],
                'match2': s2['name'], 'mercado2': s2['mercado'], 'odd2': s2['odd'],
                'total_odd': total_odd
            })

        return simples, duplas

class BetanoGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("E.A.I. - Betano AI Agent")
        self.root.geometry("600x750")
        self.agent = BetanoAgent()
        self.is_monitoring = tk.BooleanVar(value=False)
        self.use_mock = tk.BooleanVar(value=True) # Padrão para teste, usuário pode desmarcar

        # Layout
        tk.Label(root, text="E.A.I. - MONITORAMENTO BETANO", font=("Arial", 16, "bold")).pack(pady=10)

        frame_inputs = tk.Frame(root)
        frame_inputs.pack(pady=10)

        tk.Label(frame_inputs, text="Banca Atual (R$):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.entry_banca = tk.Entry(frame_inputs)
        self.entry_banca.insert(0, "100.00")
        self.entry_banca.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_inputs, text="Meta de Lucro Hoje (R$):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.entry_meta = tk.Entry(frame_inputs)
        self.entry_meta.insert(0, "5.00")
        self.entry_meta.grid(row=1, column=1, padx=5, pady=5)

        frame_opts = tk.Frame(root)
        frame_opts.pack(pady=5)

        tk.Checkbutton(frame_opts, text="Monitoramento Contínuo (5 min)", variable=self.is_monitoring, command=self.toggle_monitoring).pack(side="left", padx=10)
        tk.Checkbutton(frame_opts, text="Usar Dados de Teste (Mock)", variable=self.use_mock).pack(side="left", padx=10)

        self.btn_analisar = tk.Button(root, text="ANALISAR OPORTUNIDADES AGORA", command=self.executar_analise_thread, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), height=2)
        self.btn_analisar.pack(pady=10, fill="x", padx=50)

        self.status_label = tk.Label(root, text="Pronto.", fg="blue")
        self.status_label.pack()

        self.text_area = scrolledtext.ScrolledText(root, width=70, height=25, font=("Consolas", 10))
        self.text_area.pack(pady=10, padx=10)

        tk.Label(root, text="Aviso: Aposte com responsabilidade. | DEPARTAMENTO DE T.I.", fg="gray").pack(side="bottom", pady=5)

    def toggle_monitoring(self):
        if self.is_monitoring.get():
            self.status_label.config(text="Monitoramento contínuo ativado.")
            self.auto_refresh()
        else:
            self.status_label.config(text="Monitoramento contínuo desativado.")

    def auto_refresh(self):
        if self.is_monitoring.get():
            self.executar_analise_thread()
            # Agenda próxima execução para daqui a 5 minutos (300.000 ms)
            self.root.after(300000, self.auto_refresh)

    def executar_analise_thread(self):
        # Roda em uma thread separada para não travar a interface
        thread = threading.Thread(target=self.executar_analise, daemon=True)
        thread.start()

    def executar_analise(self):
        try:
            banca = float(self.entry_banca.get().replace(',', '.'))
            meta = float(self.entry_meta.get().replace(',', '.'))
        except ValueError:
            self.root.after(0, lambda: messagebox.showerror("Erro", "Insira valores válidos."))
            return

        self.root.after(0, lambda: self.status_label.config(text="🔍 Analisando Betano... Aguarde."))
        self.root.after(0, lambda: self.text_area.delete(1.0, tk.END))
        self.root.after(0, lambda: self.text_area.insert(tk.END, "⏳ Coletando dados em tempo real...\n"))

        try:
            self.agent.bankroll = banca
            simples, duplas = self.agent.analisar_oportunidades(mock=self.use_mock.get())

            self.root.after(0, lambda: self._exibir_resultados(banca, meta, simples, duplas))
            self.root.after(0, lambda: self.status_label.config(text=f"✅ Última atualização: {time.strftime('%H:%M:%S')}"))
        except Exception as e:
            err_msg = f"❌ Erro no scraper: {str(e)}"
            self.root.after(0, lambda: self.text_area.insert(tk.END, f"\n{err_msg}\nVerifique se o Chrome/ChromeDriver está instalado."))
            self.root.after(0, lambda: self.status_label.config(text="❌ Erro na análise."))

    def _exibir_resultados(self, banca, meta, simples, duplas):
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, "="*50 + "\n")
        self.text_area.insert(tk.END, f"RELATÓRIO DE INTELIGÊNCIA - {time.strftime('%d/%m/%Y %H:%M')}\n")
        self.text_area.insert(tk.END, f"BANCA: R$ {banca:.2f} | META: R$ {meta:.2f}\n")
        self.text_area.insert(tk.END, "="*50 + "\n\n")

        # Gestão de Banca
        stake = banca * self.agent.risk_percent
        self.text_area.insert(tk.END, f"💰 GESTÃO SUGERIDA:\n- Unidade de Aposta ({self.agent.risk_percent*100:.0f}%): R$ {stake:.2f}\n\n")

        # Plano para Meta
        if simples:
            odd_media = sum(s['odd'] for s in simples) / len(simples)
            lucro_por_aposta = (stake * odd_media) - stake
            vits = int(meta / lucro_por_aposta) + 1 if lucro_por_aposta > 0 else "N/A"
            self.text_area.insert(tk.END, f"📊 RUMO À META:\n- Você precisa de aprox. {vits} acertos para bater sua meta hoje.\n\n")

        # Exibe Simples
        self.text_area.insert(tk.END, "🛡️ APOSTAS SIMPLES (MAIOR SEGURANÇA):\n")
        if not simples:
            self.text_area.insert(tk.END, "Nenhuma oportunidade simples encontrada.\n")
        for s in simples:
            self.text_area.insert(tk.END, f"• {s['name']}\n  -> Vitória: {s['mercado']} | Odd: {s['odd']:.2f}\n")

        # Exibe Duplas
        self.text_area.insert(tk.END, "\n🔗 DUPLAS RECOMENDADAS:\n")
        if not duplas:
            self.text_area.insert(tk.END, "Sem duplas seguras no momento.\n")
        for d in duplas:
            self.text_area.insert(tk.END, f"• DUPLA SEGURA (Odd Total: {d['total_odd']:.2f})\n")
            self.text_area.insert(tk.END, f"  1. {d['match1']} ({d['mercado1']})\n")
            self.text_area.insert(tk.END, f"  2. {d['match2']} ({d['mercado2']})\n")

        self.text_area.insert(tk.END, "\n" + "="*50 + "\n")
        self.text_area.insert(tk.END, "Dica: Perfil conservador evita múltiplas de 3+ jogos.\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = BetanoGUI(root)
    root.mainloop()
