import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# E.A.I. - Betano Monitoring Agent (Agente de Monitoramento Betano)
# Desenvolvido para usuários com perfil conservador e banca baixa.

class BetanoAgent:
    def __init__(self, bankroll=100.0, risk_percent=0.01):
        """
        Inicializa o agente com configurações de banca e risco.
        :param bankroll: Valor total da banca (ex: 100.0)
        :param risk_percent: Porcentagem da banca a ser apostada (0.01 = 1%)
        """
        self.bankroll = bankroll
        self.risk_percent = risk_percent
        self.url = "https://www.betano.bet.br/sport/futebol/"

        # Configurações do Selenium para rodar em diversos ambientes
        self.chrome_options = Options()
        self.chrome_options.add_argument("--headless")  # Roda sem abrir a janela do navegador
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        # User-agent e headers para evitar detecção básica
        self.chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")

    def get_matches(self, mock=False):
        """
        Extrai as partidas e odds do site da Betano.
        """
        if mock:
            return self._get_mock_data()

        print("🔍 Acessando Betano para buscar odds atualizadas...")
        driver = None
        try:
            # Nota: Em um ambiente real, o ChromeDriver deve estar no PATH
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.get(self.url)

            # Aguarda os elementos das partidas carregarem
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
            print(f"❌ Erro ao acessar o site: {e}")
            return []
        finally:
            if driver:
                driver.quit()

    def _get_mock_data(self):
        """
        Dados simulados para teste.
        """
        return [
            {'name': 'Flamengo vs Palmeiras', 'odds': {'1 (Casa)': 2.10, 'X (Empate)': 3.40, '2 (Fora)': 3.60}},
            {'name': 'Real Madrid vs Getafe', 'odds': {'1 (Casa)': 1.25, 'X (Empate)': 5.50, '2 (Fora)': 12.00}},
            {'name': 'Man City vs Everton', 'odds': {'1 (Casa)': 1.15, 'X (Empate)': 7.00, '2 (Fora)': 18.00}},
            {'name': 'Benfica vs Porto', 'odds': {'1 (Casa)': 2.40, 'X (Empate)': 3.10, '2 (Fora)': 2.80}},
            {'name': 'Bayern vs Bochum', 'odds': {'1 (Casa)': 1.12, 'X (Empate)': 9.00, '2 (Fora)': 25.00}},
            {'name': 'Juventus vs Lecce', 'odds': {'1 (Casa)': 1.45, 'X (Empate)': 4.20, '2 (Fora)': 8.50}},
            {'name': 'Arsenal vs Wolves', 'odds': {'1 (Casa)': 1.35, 'X (Empate)': 4.80, '2 (Fora)': 9.00}},
        ]

    def monitorar(self, mock=False):
        """
        Analisa as partidas e notifica sobre oportunidades conservadoras.
        """
        matches = self.get_matches(mock=mock)

        timestamp = time.strftime("%H:%M:%S")
        print(f"\n[{timestamp}] --- Verificação em andamento ---")

        encontradas = 0
        for m in matches:
            for mercado, odd in m['odds'].items():
                # Estratégia Conservadora: Odds entre 1.20 e 1.50
                if 1.20 <= odd <= 1.50:
                    stake = self.bankroll * self.risk_percent
                    print(f"✨ OPORTUNIDADE: {m['name']} | {mercado} | Odd: {odd:.2f} | Sugestão: R$ {stake:.2f}")
                    encontradas += 1

        if encontradas == 0:
            print("📭 Nenhuma oportunidade encontrada.")
        return encontradas

    def iniciar_loop(self, intervalo=300, mock=True):
        """
        Executa o monitoramento continuamente.
        :param intervalo: Tempo entre verificações em segundos (padrão 5 min)
        """
        print("="*50)
        print("        E.A.I. - AGENTE BETANO ATIVO")
        print("="*50)
        print(f"Banca: R$ {self.bankroll:.2f} | Risco: {self.risk_percent*100:.1f}%")
        print(f"Monitorando a cada {intervalo} segundos...")
        print("Pressione Ctrl+C para parar.\n")

        try:
            while True:
                self.monitorar(mock=mock)
                time.sleep(intervalo)
        except KeyboardInterrupt:
            print("\n🛑 Monitoramento encerrado pelo usuário.")

if __name__ == "__main__":
    # Configuração Inicial
    banca_inicial = 100.00
    percentual_risco = 0.01 # 1%

    agente = BetanoAgent(bankroll=banca_inicial, risk_percent=percentual_risco)

    # Para demonstração, usamos mock=True.
    # Para uso real, alterar para mock=False e garantir que o ChromeDriver esteja instalado.
    agente.iniciar_loop(intervalo=10, mock=True) # Loop rápido para teste
