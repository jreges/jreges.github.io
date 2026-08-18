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

A página funciona de dois jeitos, com o mesmo `index.html`:

- **Modo local (duplo clique):** se o `data.js` estiver na mesma pasta, é só abrir o
  `index.html`. Ele carrega os dados automaticamente.
- **Modo upload (GitHub Pages ou qualquer lugar sem os dados ao lado):** ao abrir sem
  o `data.js`, a página mostra um botão **"Escolher arquivo…"**. Você seleciona o
  `data.json` do seu computador e ele é lido **só no navegador** — nada é enviado para
  a internet. É esse modo que torna o GitHub Pages seguro (veja abaixo).

## Publicar no GitHub Pages (com segurança)

Dá para usar o GitHub Pages, **desde que os dados nunca sejam publicados.** O ponto
que exige cuidado: no GitHub Pages, o site publicado é **público** mesmo quando o
repositório é privado — só o Enterprise Cloud (pago) permite site com acesso
restrito. Um repositório privado protege o *código*, não o *site*.

A solução é separar o app dos dados: **só o `index.html` vai para o GitHub**; o
`data.json` fica no seu computador e é carregado pelo botão "Escolher arquivo…".

Passo a passo:

1. Crie um repositório no GitHub (pode ser privado; a proteção real vem do passo 3).
2. Suba **apenas** o `index.html` (e, se quiser, o `README.md` e o `build_data.py`).
   O `.gitignore` desta pasta já bloqueia `data.js`, `data.json` e `*.xlsx` para você
   não subir os dados por engano.
3. No repositório: **Settings → Pages → Build and deployment → Source: Deploy from a
   branch → `main` / root → Save**. Em ~1 min o site fica no ar em
   `https://SEU-USUARIO.github.io/NOME-DO-REPO/`.
4. Abra a URL, clique em **"Escolher arquivo…"** e selecione o `data.json` gerado
   pelo `build_data.py`. Os dados são lidos no navegador e **não** vão para o GitHub.

> **Teste de sanidade:** abra a URL do Pages numa janela anônima (deslogado). Você deve
> ver a tela "Carregar dados", **nunca** seus números. Se aparecerem valores, algum
> `data.*` foi commitado sem querer — remova do repositório.

### Alternativa: acesso remoto com os dados já carregados

Se um dia quiser abrir de qualquer lugar **sem** ter o arquivo em mãos, aí sim os
dados precisam estar hospedados — e isso exige autenticação de verdade. Nesse caso,
**Cloudflare Pages** ou **Netlify** com senha/login (Access) são o caminho, continuam
de graça, e o `index.html` é o mesmo. Só aí faz sentido publicar o `data.js` junto,
porque o acesso fica protegido por login.

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

| Arquivo | O que é | Vai pro GitHub? |
|---|---|---|
| `index.html` | A dashboard. Abra este. | **Sim** (é só a interface) |
| `build_data.py` | Lê `Financeiro.xlsx` e gera o snapshot. | Opcional |
| `README.md` | Este guia. | Opcional |
| `.gitignore` | Impede subir os dados por engano. | Sim |
| `data.json` | Snapshot dos dados (canônico). | **Não** — dados pessoais |
| `data.js` | Mesmo snapshot, para o modo duplo clique. | **Não** — dados pessoais |
| `Financeiro.xlsx` | *(você adiciona)* export da sua planilha. | **Não** — dados pessoais |
