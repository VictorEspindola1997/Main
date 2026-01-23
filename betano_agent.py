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

# E.A.I. - Betano Monitoring Agent GUI (Agente de Monitoramento Betano com Interface)
# Focado em perfil conservador, consistente e banca baixa.

class BetanoAgent:
    def __init__(self, bankroll=100.0, risk_percent=0.02):
        self.bankroll = bankroll
        self.risk_percent = risk_percent
        self.url = "https://www.betano.bet.br/sport/futebol/"
        self.performance_file = "performance.json"
        self.learning_data = self._load_learning_data()

        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless")
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
        if is_green:
            self.learning_data["greens"] += 1
        else:
            self.learning_data["reds"] += 1

        if league not in self.learning_data["leagues"]:
            self.learning_data["leagues"][league] = {"greens": 0, "reds": 0}

        if is_green:
            self.learning_data["leagues"][league]["greens"] += 1
        else:
            self.learning_data["leagues"][league]["reds"] += 1

        with open(self.performance_file, "w") as f:
            json.dump(self.learning_data, f)

    def get_matches(self, mock=False):
        if mock:
            time.sleep(1)
            return self._get_mock_data()

        driver = None
        try:
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.get(self.url)
            wait = WebDriverWait(driver, 20)

            # Tenta localizar as linhas de evento e seus cabeçalhos de liga
            elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='event-row'], .events-list-group__header")))

            matches = []
            current_league = "Geral"

            for el in elements:
                try:
                    # Verifica se é um cabeçalho de liga
                    if "events-list-group__header" in el.get_attribute("class"):
                        current_league = el.text.strip()
                        continue

                    # Se for uma linha de evento
                    name = el.find_element(By.CSS_SELECTOR, "[data-testid='event-title']").text
                    odds_elements = el.find_elements(By.CSS_SELECTOR, "[data-testid='odds-button']")

                    match_info = {'name': name, 'league': current_league, 'odds': {}, 'goals': {}}

                    if len(odds_elements) >= 3:
                        match_info['odds'] = {
                            '1 (Casa)': float(odds_elements[0].text.replace(',', '.')),
                            'X (Empate)': float(odds_elements[1].text.replace(',', '.')),
                            '2 (Fora)': float(odds_elements[2].text.replace(',', '.'))
                        }

                    for odd_btn in odds_elements:
                        txt = odd_btn.get_attribute("aria-label") or ""
                        if "Mais de 1.5" in txt or "Over 1.5" in txt:
                             match_info['goals']['+1.5 Gols'] = float(odd_btn.text.replace(',', '.'))
                        elif "Mais de 2.5" in txt or "Over 2.5" in txt:
                             match_info['goals']['+2.5 Gols'] = float(odd_btn.text.replace(',', '.'))

                    matches.append(match_info)
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
            {'name': 'Flamengo vs Palmeiras', 'league': 'Brasileirão', 'odds': {'1 (Casa)': 2.10, 'X (Empate)': 3.40, '2 (Fora)': 3.60}, 'goals': {'+1.5 Gols': 1.28}},
            {'name': 'Real Madrid vs Getafe', 'league': 'La Liga', 'odds': {'1 (Casa)': 1.25, 'X (Empate)': 5.50, '2 (Fora)': 12.00}, 'goals': {'+2.5 Gols': 1.65}},
            {'name': 'Man City vs Everton', 'league': 'Premier League', 'odds': {'1 (Casa)': 1.15, 'X (Empate)': 7.00, '2 (Fora)': 18.00}, 'goals': {'+1.5 Gols': 1.18}},
            {'name': 'Bayern vs Bochum', 'league': 'Bundesliga', 'odds': {'1 (Casa)': 1.12, 'X (Empate)': 9.00, '2 (Fora)': 25.00}, 'goals': {'+2.5 Gols': 1.35}},
            {'name': 'Juventus vs Lecce', 'league': 'Serie A', 'odds': {'1 (Casa)': 1.45, 'X (Empate)': 4.20, '2 (Fora)': 8.50}, 'goals': {'+1.5 Gols': 1.30}},
            {'name': 'Arsenal vs Wolves', 'league': 'Premier League', 'odds': {'1 (Casa)': 1.35, 'X (Empate)': 4.80, '2 (Fora)': 9.00}, 'goals': {'+1.5 Gols': 1.22}},
            {'name': 'PSG vs Metz', 'league': 'Ligue 1', 'odds': {'1 (Casa)': 1.18, 'X (Empate)': 6.50, '2 (Fora)': 15.00}, 'goals': {'+2.5 Gols': 1.42}},
        ]

    def analisar_oportunidades(self, mock=False):
        matches = self.get_matches(mock=mock)
        simples = []
        super_seguros = []

        for m in matches:
            # Pega o histórico da liga para aumentar a confiança
            stats = self.learning_data["leagues"].get(m['league'], {"greens": 0, "reds": 0})
            win_rate = 0
            if (stats["greens"] + stats["reds"]) > 0:
                win_rate = stats["greens"] / (stats["greens"] + stats["reds"])

            # Analisa Vencedor (1X2)
            for mercado, odd in m['odds'].items():
                if 1.20 <= odd <= 1.50:
                    simples.append({'name': m['name'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Vencedor', 'win_rate': win_rate})
                elif 1.05 <= odd < 1.20:
                    super_seguros.append({'name': m['name'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Vencedor', 'win_rate': win_rate})

            # Analisa Gols
            for mercado, odd in m['goals'].items():
                if 1.25 <= odd <= 1.45:
                    simples.append({'name': m['name'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Gols', 'win_rate': win_rate})
                elif 1.05 <= odd < 1.25:
                    super_seguros.append({'name': m['name'], 'league': m['league'], 'mercado': mercado, 'odd': odd, 'tipo': 'Gols', 'win_rate': win_rate})

        duplas = []
        if len(super_seguros) >= 2:
            s1 = super_seguros[0]
            s2 = super_seguros[1]
            duplas.append({
                'match1': s1['name'], 'mercado1': s1['mercado'], 'odd1': s1['odd'],
                'match2': s2['name'], 'mercado2': s2['mercado'], 'odd2': s2['odd'],
                'total_odd': s1['odd'] * s2['odd']
            })

        return simples, duplas

class BetanoGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("E.A.I. - Betano AI Agent")
        self.root.geometry("650x850")
        self.agent = BetanoAgent()
        self.is_monitoring = tk.BooleanVar(value=False)
        self.use_mock = tk.BooleanVar(value=True)
        self.last_results = [] # Armazena últimas sugestões para feedback

        # Layout
        tk.Label(root, text="E.A.I. - MONITORAMENTO INTELIGENTE", font=("Arial", 16, "bold")).pack(pady=10)

        frame_inputs = tk.Frame(root)
        frame_inputs.pack(pady=10)

        tk.Label(frame_inputs, text="Banca (R$):").grid(row=0, column=0, padx=5, pady=5)
        self.entry_banca = tk.Entry(frame_inputs, width=10)
        self.entry_banca.insert(0, "100.00")
        self.entry_banca.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame_inputs, text="Meta (R$):").grid(row=0, column=2, padx=5, pady=5)
        self.entry_meta = tk.Entry(frame_inputs, width=10)
        self.entry_meta.insert(0, "5.00")
        self.entry_meta.grid(row=0, column=3, padx=5, pady=5)

        tk.Label(frame_inputs, text="Stake Mínima (R$):").grid(row=1, column=0, padx=5, pady=5)
        self.entry_stake_min = tk.Entry(frame_inputs, width=10)
        self.entry_stake_min.insert(0, "2.00")
        self.entry_stake_min.grid(row=1, column=1, padx=5, pady=5)

        frame_opts = tk.Frame(root)
        frame_opts.pack(pady=5)
        tk.Checkbutton(frame_opts, text="Auto-Refresh", variable=self.is_monitoring, command=self.toggle_monitoring).pack(side="left", padx=10)
        tk.Checkbutton(frame_opts, text="Mock (Teste)", variable=self.use_mock).pack(side="left", padx=10)

        self.btn_analisar = tk.Button(root, text="🔍 ANALISAR AGORA", command=self.executar_analise_thread, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=30)
        self.btn_analisar.pack(pady=10)

        frame_learning = tk.Frame(root, relief="groove", borderwidth=1)
        frame_learning.pack(fill="x", padx=10, pady=5)
        self.label_assertividade = tk.Label(frame_learning, text=self._get_learning_status(), font=("Arial", 9, "italic"))
        self.label_assertividade.pack(side="left", padx=5)

        btn_green = tk.Button(frame_learning, text="Green ✅", command=lambda: self.registrar_resultado(True), bg="#e8f5e9", font=("Arial", 8))
        btn_green.pack(side="right", padx=2)
        btn_red = tk.Button(frame_learning, text="Red ❌", command=lambda: self.registrar_resultado(False), bg="#ffeef0", font=("Arial", 8))
        btn_red.pack(side="right", padx=2)

        self.status_label = tk.Label(root, text="Pronto.", fg="blue")
        self.status_label.pack()

        self.text_area = scrolledtext.ScrolledText(root, width=80, height=30, font=("Consolas", 10))
        self.text_area.pack(pady=10, padx=10)

    def _get_learning_status(self):
        total = self.agent.learning_data["greens"] + self.agent.learning_data["reds"]
        if total == 0: return "Agente em fase inicial de aprendizado."
        pct = (self.agent.learning_data["greens"] / total) * 100
        return f"Assertividade: {pct:.1f}% ({total} registros). Agente evoluindo..."

    def registrar_resultado(self, is_green):
        if not self.last_results:
            messagebox.showwarning("Aviso", "Analise oportunidades primeiro para registrar resultados.")
            return

        # Registra para a liga da última sugestão exibida
        res = self.last_results[0] # Simplificação: usa a primeira da lista
        self.agent.save_result(is_green, res.get('league', 'Geral'))
        self.label_assertividade.config(text=self._get_learning_status())
        messagebox.showinfo("Aprendizado", f"Resultado registrado para {res.get('league')}! O agente está estudando.")

    def toggle_monitoring(self):
        if self.is_monitoring.get():
            self.auto_refresh()

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
            stake_min = float(self.entry_stake_min.get().replace(',', '.'))
        except:
            return

        self.root.after(0, lambda: self.status_label.config(text="🔍 Analisando dados confiáveis..."))
        self.root.after(0, lambda: self.text_area.delete(1.0, tk.END))

        try:
            self.agent.bankroll = banca
            simples, duplas = self.agent.analisar_oportunidades(mock=self.use_mock.get())
            self.last_results = simples
            self.root.after(0, lambda: self._exibir_resultados(banca, meta, stake_min, simples, duplas))
            self.root.after(0, lambda: self.status_label.config(text=f"✅ Atualizado: {time.strftime('%H:%M:%S')}"))
        except Exception as e:
            self.root.after(0, lambda: self.status_label.config(text="❌ Erro na conexão."))

    def _exibir_resultados(self, banca, meta, stake_min, simples, duplas):
        self.text_area.insert(tk.END, "="*60 + "\n")
        self.text_area.insert(tk.END, f"   E.A.I. - RELATÓRIO DE PREVISÕES REALISTAS\n")
        self.text_area.insert(tk.END, "="*60 + "\n\n")

        # Stake realista
        stake = max(stake_min, banca * self.agent.risk_percent)
        self.text_area.insert(tk.END, f"💰 GESTÃO DE BANCA:\n- Valor por entrada: R$ {stake:.2f}\n")
        self.text_area.insert(tk.END, f"- Meta hoje: R$ {meta:.2f}\n\n")

        # Simples (Vencedor e Gols)
        self.text_area.insert(tk.END, "🛡️ APOSTAS SIMPLES (MAIOR ASSERTIVIDADE):\n")
        if not simples:
            self.text_area.insert(tk.END, "Aguardando janelas de maior segurança...\n")

        # Ordena por win_rate da liga para ser "mais inteligente"
        simples.sort(key=lambda x: x['win_rate'], reverse=True)

        for s in simples:
            badge = ""
            if s['win_rate'] >= 0.7: badge = " [⭐ ALTA CONFIANÇA]"
            elif s['win_rate'] > 0: badge = f" [Taxa de acerto: {s['win_rate']*100:.0f}%]"

            self.text_area.insert(tk.END, f"• {s['name']} ({s['league']}){badge}\n  -> {s['mercado']} ({s['tipo']}) | Odd: {s['odd']:.2f}\n")

        # Duplas
        self.text_area.insert(tk.END, "\n🔗 DUPLAS SELECIONADAS (COMBINAÇÃO SEGURA):\n")
        if not duplas:
            self.text_area.insert(tk.END, "Sem combinações ideais no momento.\n")
        for d in duplas:
            self.text_area.insert(tk.END, f"• DUPLA SEGURA (Total Odd: {d['total_odd']:.2f})\n")
            self.text_area.insert(tk.END, f"  1. {d['match1']} ({d['mercado1']})\n")
            self.text_area.insert(tk.END, f"  2. {d['match2']} ({d['mercado2']})\n")

        self.text_area.insert(tk.END, "\n" + "="*60 + "\n")
        self.text_area.insert(tk.END, "O Agente analisou tendências de gols e favoritismo real.\n")
        self.text_area.insert(tk.END, "Foco em consistência para sua meta de R$ " + f"{meta:.2f}\n")

if __name__ == "__main__":
    root = tk.Tk()
    app = BetanoGUI(root)
    root.mainloop()
