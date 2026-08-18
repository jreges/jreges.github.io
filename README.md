# Painel financeiro

Dashboard pessoal de **fluxo de caixa** e **posição consolidada de investimentos**,
alimentada por um snapshot da sua planilha do Google Sheets. Um arquivo só,
sem dependências externas, sem backend — abre offline e nada sai do seu computador.

## O que ela mostra

- **Patrimônio total** com evolução mês a mês e variação no último mês.
- **Fluxo de caixa**: entradas × saídas por mês (com a linha de resultado), maiores
  categorias de gasto e fontes de entrada dos últimos 12 meses — considerando só
  movimentações **reais** (transferências internas, aportes e pagamentos de fatura
  ficam de fora).
- **Posição consolidada**: alocação por classe de ativo (CDB, FIIs, FI, Previdência,
  Renda fixa, Ações, Tesouro, Cripto, BDR, ETF) e distribuição por instituição.
- **Metas** rumo à independência financeira (10K → 2M).

## Como atualizar os dados

O snapshot é manual — você atualiza quando quiser, em 3 passos:

1. No Google Sheets: **Arquivo → Fazer download → Microsoft Excel (.xlsx)**.
2. Salve como **`Financeiro.xlsx`** nesta pasta (substituindo o anterior).
3. Rode o gerador:

   ```bash
   pip install openpyxl        # só na primeira vez
   python build_data.py
   ```

   Isso regenera `data.json` (o snapshot) e `data.js` (mesmo conteúdo, usado para
   abrir a página com duplo clique).

4. Abra (ou recarregue) o **`index.html`**.

## Como abrir

- **Mais simples:** duplo clique em `index.html`. Funciona direto porque a página
  lê o `data.js`.
- **Alternativa (servindo o `data.json`):**
  ```bash
  python -m http.server 8000
  # abra http://localhost:8000
  ```

## Privacidade e hospedagem

Estes são dados financeiros pessoais, então vale reforçar o que já conversamos:

- **Repositório privado ≠ site privado.** No GitHub Pages, mesmo com o repositório
  privado, o site publicado fica acessível por qualquer um com a URL. Para dados
  sensíveis, isso **não** é privacidade de verdade.
- **Recomendado começar local** (é o que está aqui): a dashboard roda no seu
  computador e nada vai para a internet.
- **Se quiser acesso remoto depois**, hospede em **Cloudflare Pages** ou **Netlify**
  com autenticação (Access / senha) — continua de graça e aí sim fica protegido.
  O `index.html` é idêntico nos dois casos; só muda o "onde".

## Notas sobre os dados (o que foi encontrado ao montar)

Ao ler a planilha, algumas escolhas foram feitas para os números baterem com o seu
resumo:

- A posição de investimentos vem da aba **`Patrimonio_real`**, que é a que reconcilia
  exatamente com o `Dash` (R$ 204.203,45). As abas `Carteira de investimento` e
  `Investimentos` estavam **defasadas** (fechavam em ~R$ 135 mil e ~R$ 3,7 mil) e por
  isso **não** foram usadas — vale revisá-las na planilha quando puder.
- A natureza de cada lançamento (real × interno) segue o mapa da aba **`Config_cat`**,
  que tem prioridade sobre a coluna `Natureza` do fluxo. Isso corrige lançamentos de
  categoria "Investimento" que estavam marcados como reais e inflavam as saídas.
- **Duas datas de referência:** o patrimônio está atualizado até **ago/26**, mas os
  lançamentos de fluxo de caixa vão até **mai/26**. Importar os extratos mais recentes
  alinha os dois (a página avisa sobre isso no rodapé enquanto houver defasagem).

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | A dashboard. Abra este. |
| `build_data.py` | Lê `Financeiro.xlsx` e gera o snapshot. |
| `data.json` | Snapshot dos dados (canônico). |
| `data.js` | Mesmo snapshot, para abrir via duplo clique. |
| `Financeiro.xlsx` | *(você adiciona)* export da sua planilha. |
