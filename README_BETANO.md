# Agente de Inteligência Betano (E.A.I.)

Este script foi desenvolvido para auxiliar apostadores iniciantes com perfil **conservador** e **banca baixa** a monitorar oportunidades na casa de apostas Betano.

## Funcionalidades
- **Interface Gráfica:** Fácil de usar, permitindo inserir sua banca e meta diária.
- **Mercados Ampliados:** Agora analisa mercados de **Vencedor (1X2)** e **Gols (Mais de 1.5/2.5)**.
- **Duplas Seguras:** Combina jogos de favoritismo extremo para criar bilhetes consistentes.
- **Gestão Inteligente:** O agente calcula automaticamente um valor de aposta (stake) realista (mínimo de R$ 2,00) baseado na sua banca.
- **Sistema de Aprendizado:** O agente possui um módulo de estudo. Ao marcar "Green" ou "Red", ele registra a assertividade e "aprende" quais padrões estão funcionando melhor para você.
- **Monitoramento Contínuo:** Opção de atualizar automaticamente os jogos a cada 5 minutos.

## Como Instalar e Rodar

### 1. Pré-requisitos
Você precisa ter o **Python** instalado em seu computador.
Além disso, instale a biblioteca Selenium:
```bash
pip install selenium
```

### 2. Navegador
O script utiliza o Google Chrome. Certifique-se de ter o Chrome instalado.
Em alguns casos, será necessário baixar o `chromedriver` compatível com sua versão do Chrome e colocá-lo na pasta do script (ou no PATH do sistema).

### 3. Rodando o Script
Execute o arquivo:
```bash
python betano_agent.py
```

## Como usar
1. Insira o valor total que você tem na banca (ex: 100.00).
2. Insira quanto você quer ganhar hoje (Meta).
3. Clique em **Iniciar Análise Real**.
5. Após fazer sua aposta e o jogo acabar, use os botões **Green ✅** ou **Red ❌** para o agente estudar os resultados.
4. O Agente mostrará os melhores jogos e mercados (geralmente Vitória do Favorito).
5. Se quiser que o robô fique vigiando o site sozinho, marque a caixa **Monitoramento Contínuo**.

## Dicas de Segurança
- Priorize as **Apostas Simples**. Elas são o caminho mais consistente para o lucro a longo prazo.
- Use as **Duplas** apenas quando houver jogos de favoritismo extremo.
- **Nunca** aposte um dinheiro que você não pode perder.

---
*Desenvolvido pelo Departamento de T.I. para auxílio estratégico.*
