# Agente de Inteligência Betano (E.A.I.)

Este script foi desenvolvido para auxiliar apostadores iniciantes com perfil **conservador** e **banca baixa** a monitorar oportunidades na casa de apostas Betano.

## Funcionalidades
- **Interface Gráfica:** Fácil de usar, permitindo inserir sua banca e meta diária.
- **Análise Inteligente:** Filtra jogos com odds seguras (entre 1.20 e 1.50) para apostas simples.
- **Duplas Seguras:** Combina jogos de odds muito baixas (ex: 1.10) para criar bilhetes com maior retorno e baixo risco.
- **Gestão de Banca:** Sugere automaticamente o valor da aposta (stake) baseado em 2% da sua banca.
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
3. Clique em **Analisar Oportunidades**.
4. O Agente mostrará os melhores jogos e mercados (geralmente Vitória do Favorito).
5. Se quiser que o robô fique vigiando o site sozinho, marque a caixa **Monitoramento Contínuo**.

## Dicas de Segurança
- Priorize as **Apostas Simples**. Elas são o caminho mais consistente para o lucro a longo prazo.
- Use as **Duplas** apenas quando houver jogos de favoritismo extremo.
- **Nunca** aposte um dinheiro que você não pode perder.

---
*Desenvolvido pelo Departamento de T.I. para auxílio estratégico.*
