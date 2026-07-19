# Conciliador Contábil Inteligente

Sistema em Python que lê o razão contábil de uma conta de passivo/provisão
(exportado em Excel), concilia automaticamente os lançamentos a crédito
(entradas/provisões) com os lançamentos a débito (saídas/reversões/baixas) e
devolve um novo Excel com o detalhe da conciliação, um resumo por período e
uma ponte de diferenças contra o balancete.

Construído e testado em cima do arquivo real `Smart - Provisão de Bônus`
(conta 2.1.1.05.0005, período 06/2026) — os números de exemplo abaixo vêm
desse arquivo.

## Instalação

```bash
pip install -r requirements.txt
```

## Painel visual (arrastar o arquivo e rodar no navegador)

Se você prefere não mexer em linha de comando nem em `config.yaml`, existe um
painel visual em Streamlit: arraste o Excel, confira os parâmetros na barra
lateral (já vêm com valores padrão) e clique em "Rodar conciliação".

```bash
pip install -r requirements.txt
streamlit run app_revisao.py
```

Isso abre uma aba no seu navegador (`http://localhost:8501`) com:

- Cartões com o total de lançamentos, % conciliado, saldo em aberto e a
  diferença contra o balancete
- Gráfico de saldo em aberto por período e de lançamentos por regra aplicada
- Tabelas de resumo por período, itens em aberto e detalhe completo (com as
  mesmas cores da Etapa 6: verde = conciliado, âmbar/coral = em aberto)
- Um botão para baixar o Excel conciliado, pronto, sem precisar rodar nada na
  linha de comando

## Uso por linha de comando (alternativa ao painel)

```bash
python main.py \
  --arquivo "caminho/para/seu_razao.xlsx" \
  --aba "01.Razão (2)" \
  --conta "2.1.1.05.0005" \
  --saida "razao_conciliado.xlsx"
```

Ou usando um arquivo de configuração (veja `config/config.exemplo.yaml`):

```bash
cp config/config.exemplo.yaml config/config.yaml   # e ajuste os valores
python main.py --config config/config.yaml
```

Parâmetros passados na linha de comando sempre sobrepõem o `config.yaml`.

### Principais parâmetros

| Parâmetro | Padrão | O que faz |
|---|---|---|
| `--arquivo` | *(obrigatório)* | Excel de origem com o razão contábil |
| `--aba` | `01.Razão (2)` | Nome da aba a ser lida |
| `--linha-cabecalho` | `6` | Linha onde está o cabeçalho (Período, Conta, Data...) |
| `--conta` | — | Código da conta (usado na Ponte_Balancete e para ler o saldo automaticamente) |
| `--tolerancia` | `0.01` | Tolerância de arredondamento (R$) para considerar dois valores "iguais" |
| `--max-grupo` | `6` | Tamanho máximo de grupo testado na Etapa 3 (busca combinatória) |
| `--similaridade-min` | `80` | % mínimo de similaridade de texto exigido na Etapa 2 |
| `--data-corte` | data mais recente do razão | Data de referência para calcular o aging dos itens em aberto |
| `--saldo-balancete` | *(lido automaticamente se possível)* | Informe manualmente se preferir não depender da leitura automática |

## O que o sistema faz (a cascata de 6 regras)

Cada lançamento é classificado pelo **sinal** de `Valor = Débito − Crédito`
(positivo = saída/débito, negativo = entrada/crédito — nunca pelo texto do
histórico sozinho, que na prática pode repetir a mesma palavra tanto em
provisões quanto em baixas). A partir daí, roda em cascata:

1. **Match exato 1:1** — um crédito e um débito de mesmo valor absoluto
2. **Match por referência/texto** — valor próximo + histórico parecido (`rapidfuzz`)
3. **Match agrupado (N:1 / 1:N)** — um grupo de lançamentos de um lado soma
   exatamente o valor de um lançamento do outro lado (busca combinatória
   limitada a `--max-grupo` itens, em centavos, para nunca sofrer erro de
   arredondamento de ponto flutuante)
4. **Netting por período** — se o que sobrou em um ano/período fecha em zero
   como um todo, mesmo sem uma combinação específica ter sido encontrada na
   Etapa 3, o período inteiro é dado como conciliado em bloco (rede de
   segurança para grupos maiores que `--max-grupo`)
5. **FIFO global cronológico** — o que ainda sobrar é compensado na ordem das
   datas (o crédito mais antigo é baixado primeiro), inclusive atravessando
   anos, com baixa parcial quando um lado não cobre o outro exatamente
6. **Itens em aberto** — o que não fechou em nenhuma etapa anterior, com
   aging (dias desde o lançamento) calculado

Cada match grava **qual regra o encontrou** e **com quais linhas exatamente**
(`id_match`, `regra_aplicada`, `contraparte`) — nada é conciliado "no escuro".

### O que o sistema encontrou no arquivo de exemplo

- **2021**: 12 créditos mensais (entre R$ 22.500 e R$ 40.500) fecham
  exatamente contra 1 débito de R$ 324.000,00 em 31/12/2021. Como esse grupo
  tem 13 itens (mais que o `--max-grupo` padrão de 6), quem resolve é a
  **Etapa 4** — o período fecha em zero como um todo.
- **2022**: depois de isolar um par de R$ 8.000,00 (Etapa 1), o restante se
  decompõe em **dois** grupos N:1 que a **Etapa 3** encontra sozinha: 4
  provisões mensais de R$ 59.713,20 + 1 crédito residual de R$ 8.000,00
  somam exatamente o débito de 30/11 (R$ 246.852,80); as 6 provisões mensais
  restantes somam exatamente o débito de 30/12 (R$ 358.279,20).
- **2023 a 2026**: nenhum desses anos fecha sozinho. A **Etapa 5 (FIFO
  global)** consome os créditos mais antigos primeiro — o resultado é que,
  por essa lógica cronológica, tudo até 2024 acaba coberto pelos débitos
  registrados depois, e o que **realmente permanece em aberto hoje** está
  concentrado em 2025 (R$ 195.120,86) e 2026 (R$ 230.810,98), somando os
  R$ 425.931,84 que batem exatamente com o saldo da conta no balancete em
  30/06/2026.

> A suposição de FIFO (o mais antigo é baixado primeiro) é um critério
> razoável quando não há uma referência explícita ligando um débito a um
> crédito específico — mas é uma suposição, não um fato contábil. Revise as
> linhas com `regra_aplicada = FIFO` antes de assinar o fechamento do mês.

## Saída gerada

Um novo arquivo `.xlsx` (o original nunca é sobrescrito) com 4 abas:

- **Detalhe_Conciliacao** — cada lançamento + status, regra aplicada,
  contraparte(s), valor residual, aging — com formatação condicional
  (verde = conciliado, âmbar = em aberto ≤ 90 dias, coral = em aberto > 90 dias)
- **Resumo_Periodo** — o mesmo cálculo que hoje é feito manualmente na aba
  `03.Conciliação`, automatizado: provisionado x revertido x saldo x % conciliado
- **Itens_Em_Aberto** — recorte só do que ainda não fechou, ordenado por aging
- **Ponte_Balancete** — reproduz o formato da aba `Capa`: saldo no balancete
  (1) × saldo no relatório auxiliar pós-conciliação (2) × diferença (3)

Todas as somas na saída são fórmulas do Excel (`SUBTOTAL`/`SUM`), não valores
fixos — a planilha recalcula se uma linha for editada depois.

## Estrutura do projeto

```
conciliador-contabil/
├── src/
│   ├── excel_io.py           # leitura do razão + leitor opcional de saldo do balancete
│   ├── classificador.py       # tipo textual (Provisão/Reversão/...) + direção (crédito/débito)
│   ├── motor_conciliacao.py   # classe ConciliadorContabil - as 6 etapas da cascata
│   └── relatorios.py          # monta o Excel de saída (4 abas + formatação)
├── tests/
│   └── test_motor_conciliacao.py
├── config/
│   └── config.exemplo.yaml
├── main.py                    # CLI
├── app_revisao.py             # painel visual (Streamlit) - arraste o arquivo e rode no navegador
├── requirements.txt
└── README.md
```

## Rodando os testes

```bash
pytest tests/ -v
```

Os 4 primeiros testes usam dados sintéticos e cobrem os 3 padrões reais
encontrados no arquivo de origem (grupo N:1 que fecha em bloco, par exato
1:1, saldo genuinamente em aberto). Os últimos 5 são de integração e rodam a
cascata inteira contra o arquivo real, conferindo que:

- nada é perdido ou duplicado (soma dos resíduos = soma original do razão);
- a Ponte_Balancete fecha em zero;
- 2021 e 2022 ficam 100% conciliados.

## Limitações conhecidas e próximos passos

- **Revisão humana continua sendo obrigatória.** Trate a saída como uma
  *proposta* de conciliação (especialmente as linhas resolvidas via Etapa 5 -
  FIFO), não como o fechamento contábil em si.
- **Lançamentos de reclassificação** (`RECLAS.VLR`) hoje entram no mesmo
  motor de match que provisão/reversão. Se no seu processo eles devem ser
  tratados como simples transferência (sem "casar" com uma provisão), separe-
  os antes de rodar a cascata.
- **A Etapa 2 (referência/texto)** não encontrou nenhum par no arquivo de
  exemplo — o `similaridade_min` (80%) pode precisar de ajuste dependendo do
  quão padronizado é o texto do histórico na sua empresa.
- **Camada de IA (Claude/GPT) para explicações automáticas dos itens em
  aberto** ainda não está implementada nesta versão — é o próximo passo
  natural (ver o prompt `prompt_conciliador_contabil_inteligente.md` anexo,
  seção 6, para o desenho dessa camada).
- **O painel Streamlit (`app_revisao.py`) roda a conciliação, mas ainda não
  permite aprovar/rejeitar matches individualmente** antes de gerar o Excel
  final — hoje ele mostra o resultado da cascata inteira e deixa baixar. Um
  próximo passo natural é adicionar essa aprovação linha a linha, principalmente
  para os matches da Etapa 5 (FIFO).
