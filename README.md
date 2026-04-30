# Finance Dashboard

Dashboard web de acompanhamento de investimentos pessoais, com cotações em tempo real da B3, histórico de dividendos e insights automáticos.

> Construído para controlar minha própria carteira — cobre renda variável (ações + FIIs), renda fixa (Selic) e cripto, com dados desde junho de 2024.

---

## Funcionalidades

- **Cotações em tempo real da B3** — ações, FIIs e ETFs via Yahoo Finance (`yfinance`), com cache em memória de 15 minutos e busca paralela (ThreadPoolExecutor)
- **Taxa Selic ao vivo** — consumida da API pública do Banco Central do Brasil e anualizada via `(1 + diária)^252 − 1`
- **Gráficos de desempenho mensal** — evolução do patrimônio, META vs REAL de rendimento, renda mensal detalhada (Selic + dividendos FIIs + dividendos ações)
- **Alocação da carteira** — gráfico de rosca e tabela com a composição atual do portfólio
- **Motor de insights automáticos** — 8 insights baseados em regras: status da reserva de emergência, reserva de oportunidade, meta de cripto (10%), tendência META vs REAL nos últimos 6 meses, melhor FII por yield acumulado e sugestões de aporte e reinvestimento
- **Barra de mínimo/máximo de 52 semanas** — posição visual do preço de cada ativo dentro do intervalo anual
- **Parser Excel dual-formato** — detecta automaticamente o novo formato organizado (`finance_data.xlsx`) ou usa o legado (`finance.xlsx`) como fallback

## Tecnologias

| Camada | Tecnologia |
|---|---|
| Backend | Python 3.13 + Flask 3.x |
| Leitura de dados | openpyxl |
| Dados de mercado | yfinance + API REST do Banco Central |
| Frontend | HTML/CSS/JS — Chart.js 4.x + Bootstrap 5 |
| Concorrência | ThreadPoolExecutor (buscas paralelas de cotações) |
| Cache | Dicionário em memória com TTL (15 min cotações, 1 h Selic) |

## Arquitetura

```
finance_data.xlsx  ←──── alimentado manualmente no Excel
       │
       ▼
 data_parser.py   ←──── lê e normaliza os dados
       │
       ▼
    app.py        ←──── API Flask + motor de insights
  /api/data            dados do portfólio parseados
  /api/market          cotações em tempo real + Selic
  /api/insights        recomendações automáticas
       │
       ▼
 templates/index.html  ←──── dashboard escuro com Chart.js
```

O frontend consome os três endpoints de forma assíncrona no carregamento. Os dados de mercado são carregados separadamente (após o dashboard principal renderizar) para não bloquear a interface enquanto o Yahoo Finance responde.

## Estrutura do Excel

A planilha `finance_data.xlsx` tem quatro abas:

- **Mensal** — uma linha por mês: valor do portfólio, total investido, ganho/perda, Selic, dividendos FIIs, dividendos ações, rendimento META/REAL
- **Dividendos** — um registro por evento de renda: ativo, tipo (FII/Ação/SELIC/ETF/Cripto), valor, data de pagamento
- **Carteira_Atual** — posições atuais com tickers do Yahoo Finance e % da carteira
- **Resumo** — totais consolidados: alocação variável/fixa/cripto, lucro total, aportes, reservas

> Os arquivos Excel estão excluídos deste repositório via `.gitignore` por conterem dados financeiros pessoais. Execute `python create_template.py` para gerar um template em branco com a estrutura correta.

## Como rodar

```bash
# Instalar dependências
pip install -r requirements.txt

# Gerar o template Excel (ou trazer seu próprio finance_data.xlsx)
python create_template.py

# Iniciar o servidor
python app.py
# → http://localhost:5000
```

O app relê os dados do Excel a cada requisição — basta salvar a planilha e dar F5 no navegador, sem precisar reiniciar o servidor.

## Habilidades demonstradas

- **Backend:** design de API REST com Flask, normalização de dados com tratamento de Unicode (`unicodedata.normalize`), parser multi-formato com detecção automática
- **Engenharia de dados:** pipeline ETL do Excel → dicionários Python → API JSON, remoção de acentos para matching robusto de labels, I/O concorrente com ThreadPoolExecutor
- **Frontend:** JS assíncrono com `async/await` e `Promise.all`, Chart.js 4.x (linha, barra, rosca, barra empilhada), UI dark responsiva com CSS custom properties
- **Integrações externas:** Yahoo Finance (yfinance), API REST do Banco Central do Brasil
- **Design de software:** cache em memória com TTL, estratégias de fallback, separação entre parsing de dados, lógica de negócio e apresentação
