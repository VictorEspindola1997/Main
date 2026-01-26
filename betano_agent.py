import time
import sys
import threading
import json
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# E.A.I. - Betano Monitoring Agent GUI v3
# Focado em perfil conservador, consistente e banca baixa.
# Desenvolvido para o Departamento de T.I.

class BetanoAgent:
    def __init__(self, bankroll=100.0, risk_percent=0.02, driver_path=r"C:\Users\TI Medicina\Documents\Webdriver\chromedriver-win64\chromedriver.exe"):
        self.bankroll = bankroll
        self.risk_percent = risk_percent
        self.driver_path = driver_path
        self.performance_file = "performance.json"
        self.learning_data = self._load_learning_data()

        # URLs de monitoramento por esporte
        self.sports_urls = {
            "FUTEBOL": "https://www.betano.bet.br/sport/futebol/",
            "BASQUETE": "https://www.betano.bet.br/sport/basquete/",
            "TÊNIS": "https://www.betano.bet.br/sport/tenis/"
        }

        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless") # Análise em segundo plano
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        self.chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    def _load_learning_data(self):
        if os.path.exists(self.performance_file):
            try:
                with open(self.performance_file, "r") as f:
                    return json.load(f)
            except:
                return {"greens": 0, "reds": 0, "leagues": {}}
        return {"greens": 0, "reds": 0, "leagues": {}}

    def save_result(self, is_green, league="Geral"):
        if is_green: self.learning_data["greens"] += 1
        else: self.learning_data["reds"] += 1

        if league not in self.learning_data["leagues"]:
            self.learning_data["leagues"][league] = {"greens": 0, "reds": 0}

        if is_green: self.learning_data["leagues"][league]["greens"] += 1
        else: self.learning_data["leagues"][league]["reds"] += 1

        with open(self.performance_file, "w") as f:
            json.dump(self.learning_data, f)

    def get_all_matches(self):
        """Varre múltiplos esportes em busca de oportunidades."""
        all_matches = []
        driver = None
        try:
            # Tenta usar o path configurado. Se falhar (ex: Linux), tenta o padrão do sistema.
            try:
                service = Service(executable_path=self.driver_path)
                driver = webdriver.Chrome(service=service, options=self.chrome_options)
            except Exception:
                driver = webdriver.Chrome(options=self.chrome_options)

            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            for sport_name, url in self.sports_urls.items():
                print(f"🔍 Analisando {sport_name}...")
                driver.get(url)
                try:
                    wait = WebDriverWait(driver, 15)
                    # Localiza eventos e cabeçalhos de liga/categoria
                    elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='event-row'], .events-list-group__header")))

                    current_league = "Geral"
                    for el in elements:
                        try:
                            # Detecta troca de liga
                            if "events-list-group__header" in el.get_attribute("class"):
                                current_league = el.text.strip()
                                continue

                            # Extrai dados da partida
                            name = el.find_element(By.CSS_SELECTOR, "[data-testid='event-title']").text
                            odds_elements = el.find_elements(By.CSS_SELECTOR, "[data-testid='odds-button']")

                            match_info = {'name': name, 'sport': sport_name, 'league': current_league, 'odds': {}, 'goals': {}}

                            if len(odds_elements) >= 2:
                                # Vencedor (1X2 para futebol, Vencedor para Tênis/Basquete)
                                if sport_name == "FUTEBOL" and len(odds_elements) >= 3:
                                    match_info['odds'] = {
                                        '1 (Casa)': float(odds_elements[0].text.replace(',', '.')),
                                        'X (Empate)': float(odds_elements[1].text.replace(',', '.')),
                                        '2 (Fora)': float(odds_elements[2].text.replace(',', '.'))
                                    }
                                else:
                                    # Para Tênis e Basquete, foca no Vencedor do Jogo
                                    match_info['odds'] = {
                                        '1 (Casa/Jogador 1)': float(odds_elements[0].text.replace(',', '.')),
                                        '2 (Fora/Jogador 2)': float(odds_elements[1].text.replace(',', '.'))
                                    }

                            # Tenta capturar gols se for Futebol
                            if sport_name == "FUTEBOL":
                                for odd_btn in odds_elements:
                                    txt = (odd_btn.get_attribute("aria-label") or "").lower()
                                    if "mais de 1.5" in txt or "over 1.5" in txt:
                                        match_info['goals']['+1.5 Gols'] = float(odd_btn.text.replace(',', '.'))
                                    elif "mais de 2.5" in txt or "over 2.5" in txt:
                                        match_info['goals']['+2.5 Gols'] = float(odd_btn.text.replace(',', '.'))

                            all_matches.append(match_info)
                        except:
                            continue
                except:
                    continue # Pula esporte se der erro

            return all_matches
        finally:
            if driver: driver.quit()

    def calcular_stake(self):
        """Calcula um valor de aposta realista e conservador."""
        # Sugere o maior entre R$ 2,00 ou 2% da banca
        return max(2.0, self.bankroll * self.risk_percent)

    def analisar_oportunidades(self):
        matches = self.get_all_matches()
        simples = []
        super_seguros = []

        for m in matches:
            stats = self.learning_data["leagues"].get(m['league'], {"greens": 0, "reds": 0})
            win_rate = stats["greens"] / (stats["greens"] + stats["reds"]) if (stats["greens"] + stats["reds"]) > 0 else 0

            # Analisa Vencedor
            for mercado, odd in m['odds'].items():
                if 1.20 <= odd <= 1.50:
                    simples.append({'name': m['name'], 'sport': m['sport'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Vencedor', 'win_rate': win_rate})
                elif 1.05 <= odd < 1.20:
                    super_seguros.append({'name': m['name'], 'sport': m['sport'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Vencedor', 'win_rate': win_rate})

            # Analisa Gols (Apenas Futebol)
            for mercado, odd in m['goals'].items():
                if 1.25 <= odd <= 1.45:
                    simples.append({'name': m['name'], 'sport': m['sport'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Gols', 'win_rate': win_rate})
                elif 1.05 <= odd < 1.25:
                    super_seguros.append({'name': m['name'], 'sport': m['sport'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Gols', 'win_rate': win_rate})

        duplas = []
        if len(super_seguros) >= 2:
            s1 = super_seguros[0]; s2 = super_seguros[1]
            duplas.append({
                'match1': f"[{s1['sport']}] {s1['name']}", 'mercado1': s1['mercado'], 'odd1': s1['odd'],
                'match2': f"[{s2['sport']}] {s2['name']}", 'mercado2': s2['mercado'], 'odd2': s2['odd'],
                'total_odd': s1['odd'] * s2['odd']
            })

        return simples, duplas

class BetanoGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("E.A.I. - Betano AI Agent v3")
        self.root.geometry("700x850")
        self.agent = BetanoAgent()
        self.is_monitoring = tk.BooleanVar(value=False)
        self.last_results = []

        # UI Layout
        tk.Label(root, text="🤖 E.A.I. - MONITORAMENTO MULTIESPORTIVO", font=("Arial", 16, "bold")).pack(pady=10)

        frame_inputs = tk.Frame(root)
        frame_inputs.pack(pady=10)

        tk.Label(frame_inputs, text="Banca Atual (R$):").grid(row=0, column=0, padx=5, pady=5)
        self.entry_banca = tk.Entry(frame_inputs, width=15)
        self.entry_banca.insert(0, "100.00")
        self.entry_banca.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_inputs, text="Meta de Lucro (R$):").grid(row=0, column=2, padx=5, pady=5)
        self.entry_meta = tk.Entry(frame_inputs, width=15)
        self.entry_meta.insert(0, "5.00")
        self.entry_meta.grid(row=0, column=3, padx=5, pady=5)

        tk.Checkbutton(root, text="Monitoramento Contínuo em Segundo Plano (5 min)", variable=self.is_monitoring, command=self.toggle_monitoring).pack(pady=5)

        self.btn_analisar = tk.Button(root, text="🚀 INICIAR ANÁLISE REAL DO DIA", command=self.executar_analise_thread, bg="#1E88E5", fg="white", font=("Arial", 11, "bold"), width=40, height=2)
        self.btn_analisar.pack(pady=10)

        frame_learning = tk.Frame(root, relief="groove", borderwidth=1)
        frame_learning.pack(fill="x", padx=15, pady=5)
        self.label_assertividade = tk.Label(frame_learning, text=self._get_learning_status(), font=("Arial", 9, "italic"))
        self.label_assertividade.pack(side="left", padx=5)

        tk.Button(frame_learning, text="Green ✅", command=lambda: self.registrar_resultado(True), bg="#e8f5e9", font=("Arial", 8)).pack(side="right", padx=2)
        tk.Button(frame_learning, text="Red ❌", command=lambda: self.registrar_resultado(False), bg="#ffeef0", font=("Arial", 8)).pack(side="right", padx=2)

        self.status_label = tk.Label(root, text="Pronto para analisar partidas reais.", fg="gray")
        self.status_label.pack()

        self.text_area = scrolledtext.ScrolledText(root, width=85, height=35, font=("Consolas", 10))
        self.text_area.pack(pady=10, padx=10)

    def _get_learning_status(self):
        total = self.agent.learning_data["greens"] + self.agent.learning_data["reds"]
        if total == 0: return "Agente estudando mercados..."
        pct = (self.agent.learning_data["greens"] / total) * 100
        return f"Assertividade Histórica: {pct:.1f}% ({total} jogos analisados)."

    def registrar_resultado(self, is_green):
        if not self.last_results:
            messagebox.showwarning("Aviso", "Analise oportunidades primeiro.")
            return
        res = self.last_results[0]
        self.agent.save_result(is_green, res.get('league', 'Geral'))
        self.label_assertividade.config(text=self._get_learning_status())
        messagebox.showinfo("Sucesso", f"Obrigado pelo feedback! O agente agora conhece melhor a liga {res.get('league')}.")

    def toggle_monitoring(self):
        if self.is_monitoring.get(): self.auto_refresh()

    def auto_refresh(self):
        if self.is_monitoring.get():
            self.executar_analise_thread()
            self.root.after(300000, self.auto_refresh)

    def executar_analise_thread(self):
        threading.Thread(target=self.executar_analise, daemon=True).start()

    def executar_analise(self):
        try:
            banca = float(self.entry_banca.get().replace(',', '.'))
            meta = float(self.entry_meta.get().replace(',', '.'))
            self.agent.bankroll = banca
        except:
            return

        self.root.after(0, lambda: self.status_label.config(text="🔍 Escaneando Futebol, Basquete e Tênis na Betano..."))
        self.root.after(0, lambda: self.text_area.delete(1.0, tk.END))
        self.root.after(0, lambda: self.text_area.insert(tk.END, "⏳ Estudando partidas reais do dia. Aguarde...\n"))

        try:
            simples, duplas = self.agent.analisar_oportunidades()
            self.last_results = simples
            self.root.after(0, lambda: self._exibir_resultados(banca, meta, simples, duplas))
            self.root.after(0, lambda: self.status_label.config(text=f"✅ Atualizado em: {time.strftime('%H:%M:%S')}"))
        except Exception as e:
            self.root.after(0, lambda: self.status_label.config(text="❌ Erro: Verifique o ChromeDriver."))
            self.root.after(0, lambda: self.text_area.insert(tk.END, f"\nErro técnico: {str(e)}\nCaminho esperado: {self.agent.driver_path}"))

    def _exibir_resultados(self, banca, meta, simples, duplas):
        self.text_area.insert(tk.END, "="*65 + "\n")
        self.text_area.insert(tk.END, f"   RELATÓRIO E.A.I. - ANÁLISE MULTIESPORTIVA REAL\n")
        self.text_area.insert(tk.END, "="*65 + "\n\n")

        stake = self.agent.calcular_stake()
        self.text_area.insert(tk.END, f"💰 VALOR SUGERIDO POR APOSTA: R$ {stake:.2f}\n")
        self.text_area.insert(tk.END, f"🎯 META DE HOJE: R$ {meta:.2f}\n\n")

        self.text_area.insert(tk.END, "🛡️ OPORTUNIDADES IDENTIFICADAS (CONSERVADORAS):\n")
        if not simples:
            self.text_area.insert(tk.END, "Nenhuma janela de segurança encontrada nos mercados principais.\n")

        simples.sort(key=lambda x: x['win_rate'], reverse=True)
        for s in simples:
            badge = " [⭐]" if s['win_rate'] >= 0.75 else ""
            self.text_area.insert(tk.END, f"• [{s['sport']}] {s['name']} ({s['league']}){badge}\n  -> {s['mercado']} ({s['tipo']}) | Odd: {s['odd']:.2f}\n")

        if duplas:
            self.text_area.insert(tk.END, "\n🔗 DUPLA DE SEGURANÇA RECOMENDADA:\n")
            for d in duplas:
                self.text_area.insert(tk.END, f"• ODD TOTAL: {d['total_odd']:.2f}\n")
                self.text_area.insert(tk.END, f"  1. {d['match1']} ({d['mercado1']})\n")
                self.text_area.insert(tk.END, f"  2. {d['match2']} ({d['mercado2']})\n")

        self.text_area.insert(tk.END, "\n" + "="*65 + "\n")
        self.text_area.insert(tk.END, "Análise concluída. O Agente priorizou odds entre 1.20 e 1.50.\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = BetanoGUI(root)
    root.mainloop()
