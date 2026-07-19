# Prompt mestre: Conciliador Contábil Inteligente (Python + Excel)

> **Como usar:** copie este documento inteiro e cole em uma ferramenta de codificação com IA (Claude Code, Cursor, Copilot Chat, ChatGPT etc.), anexando também o arquivo Excel do razão contábil a ser conciliado. Os valores entre colchetes `[ ]` são parâmetros que você ajusta a cada execução (conta, período, planilha).
>
> Este prompt já foi testado ponta a ponta: existe uma implementação de referência funcionando (pasta `conciliador-contabil/`, entregue junto com este documento), validada linha a linha contra um razão real de 114 lançamentos. Os números e exemplos abaixo não são hipotéticos — são o que o sistema realmente encontrou.

---

## 1. Papel e objetivo

Você é um(a) engenheiro(a) de software sênior especializado em automação de processos contábeis e financeiros, com domínio de:
- Python (`pandas`, `openpyxl`, `numpy`)
- Lógica contábil de partidas dobradas (débito/crédito, contas de ativo/passivo/resultado, provisão e reversão)
- Algoritmos de conciliação e "matching" — o mesmo problema resolvido por ferramentas como BlackLine, Trintech, ou a transação de compensação de partidas do SAP FI (F-03/F.13)

Construa, em Python, o **Conciliador Contábil Inteligente**: um sistema que lê um razão contábil em Excel, classifica cada lançamento como entrada (crédito) ou saída (débito), concilia os valores positivos com os negativos através de uma cascata de regras determinísticas e auditáveis, e devolve um Excel com o detalhe da conciliação, um resumo por período e uma ponte de diferenças contra o balancete.

**Restrição não negociável:** a entrada e a saída do sistema são sempre arquivos `.xlsx`. O Python é o motor de processamento — mas quem usa o resultado é um analista contábil dentro do Excel, não em um banco de dados ou uma tela web. Cada execução tem que poder ser conferida célula a célula em uma planilha comum.

## 2. Contexto de negócio

Trabalho como analista contábil e concilio mensalmente contas de passivo do tipo "provisão" (neste caso, Provisão de Bônus a Pagar). A dinâmica contábil é sempre a mesma:

- Todo mês, uma **provisão** é lançada a **crédito** da conta (aumenta o passivo) — reconhece uma despesa/obrigação antes do pagamento efetivo.
- Em algum momento posterior — no mesmo mês, meses depois, ou até em outro ano — essa provisão é **revertida/baixada** com um lançamento a **débito** (reduz o passivo): pode ser pagamento efetivo, ajuste, estorno ou reclassificação.
- Conciliar a conta significa **casar cada crédito (provisão) com o(s) débito(s) que o liquidam**, sobrando no final só o saldo genuinamente em aberto.

É o mesmo problema de "compensação de partidas em aberto" (*open item clearing*) que ERPs resolvem para contas correntes de fornecedores/clientes — aplicado aqui a uma conta de provisão, sem um módulo pronto do ERP para isso.

### O que já existe na planilha usada como referência

O arquivo real usado para construir e testar este sistema (`Smart - Provisão de Bônus`, período 06/2026, empresa AMBAR TECH) tem estas abas:

| Aba | Conteúdo |
|---|---|
| `Parametros` | Parâmetros do relatório extraído do ERP (Totvs Protheus) — conta, datas, filial |
| **`01.Razão (2)`** | **A aba principal a ser conciliada** — razão analítico da conta 2.1.1.05.0005 (Provisão Bônus a Pagar), com todos os lançamentos de 2021 a 06/2026 |
| `01.Razão` | Uma segunda ordenação do mesmo razão (mesmos 114 lançamentos, mesma soma total — só a ordem das linhas difere; não use como fonte adicional, é o mesmo dado) |
| `00.Balancete` | Balancete de verificação de toda a empresa — usado só para validar o saldo final da conta |
| `02.Suporte` | Memória de cálculo de "quanto deveria ser provisionado" mês a mês — não é o razão real |
| `03.Conciliação` | Conciliação manual atual, feita "na unha": soma Provisão do ano × Reversão do ano, por ano |
| `Capa` | Modelo de "ponte" de diferenças (papel de trabalho): Balancete (1) × Relatório Auxiliar (2) = Diferença (3); Diferenças identificadas na contabilidade (4) e no relatório (5); Diferença a identificar (3-4-5) |

**O objetivo central é substituir a aba `03.Conciliação` (hoje manual, só no nível de ano) por um motor automático que concilia lançamento a lançamento sempre que possível, só recorrendo ao saldo do período quando não há match individual — e que alimenta a `Capa` automaticamente.**

**Achado relevante durante a análise:** a aba `Capa` deste arquivo referencia a conta `2.1.5.10.050`, enquanto o razão e o balancete usam `2.1.1.05.0005` em todas as outras abas. Isso é sintoma comum de planilha-modelo reaproveitada de outra conta e não totalmente atualizada — exatamente o tipo de erro manual que este sistema deve eliminar. **Por isso a conta contábil é sempre um parâmetro explícito de entrada, nunca algo deduzido lendo uma aba solta.**

## 3. Estrutura de dados de entrada (schema real da aba `01.Razão (2)`)

- Cabeçalho na **linha 6**, dados a partir da **linha 7** (linhas 1-5 têm células soltas de totais/fórmulas — ignore ao ler; a leitura correta é filtrar por `Data` não nula).
- Existe uma fórmula de conferência perto do topo (`=SUBTOTAL(9,M7:M120)`), mas **não confie no valor em cache dela**: no arquivo de referência, essa fórmula tinha um cache desatualizado (R$ -329.237,80) enquanto o total real recalculado é R$ -425.931,84 — o Excel só recalcula ao reabrir/editar, e um arquivo gerado por script e nunca reaberto fica com esse cache velho para sempre. **Sempre some a coluna Valor em Python; use o cache só como alerta de divergência, nunca como fonte de verdade.**

| Coluna | Nome | Tipo | Descrição | Exemplo real |
|---|---|---|---|---|
| B | `Período` | int (ano) | Ano do lançamento | `2023` |
| C | `Conta` | texto | Código sintético (12 primeiros chars da coluna D — hoje é fórmula `=MID(D7,1,12)`; recalcule em Python, não dependa da fórmula) | `2.1.1.05.000` |
| D | `Conta + Descrição` | texto | Código completo + nome da conta | `2.1.1.05.0005 - PROVISAO BONUS A PAGAR` |
| E | `Data` | data | Data do lançamento contábil | `2023-08-31` |
| F | `LOTE/SUB/DOC/LINHA` | texto | Identificador único do lançamento no ERP | `000001001000003002` |
| G | `HISTORICO` | texto | Descrição livre digitada pelo contador | `VLR PROVISAO BONUS` |
| H | `C/PARTIDA` | texto | Conta contrapartida da partida dobrada | `4.1.1.01.0026` |
| I | `FL` | texto | Filial | `0301` |
| J | `C CUSTO` | texto | Centro de custo | `0.3.7.000007` |
| K | `DEBITO` | número | Valor a débito (sempre ≥ 0) | `116611.00` |
| L | `CREDITO` | número | Valor a crédito (sempre ≥ 0) | `0.00` |
| M | `Valor` | número | **DEBITO − CREDITO** — positivo = débito (saída/baixa), negativo = crédito (entrada/provisão) | `116611.00` |
| N | `Obs.` | texto | Observação manual — hoje contém principalmente a marca `"Efeito zero"` em lançamentos que o contador já sabe que se compensam | `Efeito zero` |

### Casos reais extraídos do arquivo (usados como teste de aceite do sistema — seção 11)

**Caso 1 — grupo N:1 grande, resolvido pela Etapa 4 (netting de período), não pela Etapa 3:**
Em 2021, 12 créditos de provisão (entre R$ 22.500,00 e R$ 40.500,00, somando R$ 324.000,00) fecham exatamente contra 1 débito de 31/12/2021 (`"EST PROV BONUS"`, R$ 324.000,00). Como esse grupo tem 13 lançamentos — mais do que o tamanho máximo de grupo testado por padrão na Etapa 3 (6) — quem efetivamente resolve esse caso é a **Etapa 4**: o ano inteiro fecha em zero como bloco. Isso é o comportamento correto e esperado, não uma falha: a Etapa 4 existe justamente como rede de segurança para grupos maiores do que o motor de busca combinatória testa por padrão.

**Caso 2 — dois grupos N:1 menores, resolvidos pela Etapa 3:**
Em 2022, depois de isolar um par de R$ 8.000,00 por match exato (crédito em 01/03 × débito em 04/03, mesma pessoa — "EDUARDO DE SETTIALVES"), sobram 11 créditos e 2 débitos. Longe de ser um bloco só, isso se decompõe em **dois matches N:1 independentes**, ambos dentro do limite padrão de 6 itens por grupo: (a) 4 provisões mensais de R$ 59.713,20 + o crédito residual de R$ 8.000,00 somam exatamente R$ 246.852,80, batendo com o débito de 30/11/2022; (b) as 6 provisões mensais restantes (R$ 59.713,20 × 6) somam exatamente R$ 358.279,20, batendo com o débito de 30/12/2022. Um motor que só testa 1:1 nunca encontraria isso — precisa de busca combinatória sobre subconjuntos.

**Caso 3 — heterogeneidade de texto: o histórico sozinho não confiável para direção:**
Em 2022, o texto `"PROV BONUS"` aparece tanto nos créditos mensais (provisão nova) quanto nos dois débitos de fechamento de ano (`"PROV BONUS"` e `"PROV EST BONUS"` — que na verdade são reversões). Em 2026, créditos de provisão (contrapartida `4.1.1.01.0026`) convivem com lançamentos de reversão/reclassificação via folha de pagamento (`"REV.VLR FOPG 313 - PREMIO"`, `"RECLAS.VLR FOPG 402 - DESCONTO DIVERSOS"`) usando contrapartidas diferentes (`3.1.1.04.0020`, `4.1.1.01.0020`, `4.4.1.01.0011`). **Conclusão prática: classifique a direção sempre pelo sinal do Valor, nunca pelo texto do histórico — o texto só ajuda como sinal auxiliar de agrupamento.**

**Caso 4 — o que sobra depois de tudo isso precisa de FIFO cronológico, mesmo atravessando anos:**
2023, 2024, 2025 e 2026 (parcial) não fecham sozinhos ano a ano (saldos líquidos de -342.223,00 / -280.060,00 / +281.614,67 / -85.263,51, respectivamente). Aplicando compensação cronológica global (o crédito mais antigo é baixado primeiro pelo débito mais antigo disponível, inclusive atravessando anos), tudo o que foi provisionado até 2024 acaba coberto pelos débitos lançados depois — e o saldo que **realmente** permanece em aberto na conta hoje está concentrado em 2025 (R$ 195.120,86) e 2026 (R$ 230.810,98), somando exatamente **R$ 425.931,84**, que bate com o saldo da conta no balancete em 30/06/2026. Trate essa suposição de "mais antigo primeiro" como um critério razoável na ausência de referência explícita — não como fato contábil certo; por isso ela fica marcada com uma regra própria (`FIFO`), separada das etapas de match exato/agrupado, para o analista revisar antes de fechar o mês.

## 4. Regra central: conciliar valor positivo com valor negativo

A tarefa fundamental, em uma frase: **para cada lançamento a crédito (Valor negativo = entrada/provisão), encontrar o(s) lançamento(s) a débito (Valor positivo = saída/reversão) que o compensam, dentro de uma tolerância de arredondamento, e marcar ambos os lados como conciliados — registrando sempre qual regra fez o match.**

Formalmente, é uma variação do **problema da soma de subconjuntos (subset sum)** aplicado a duas listas (créditos e débitos) — NP-difícil no caso geral, mas tratável aqui porque: os valores têm só 2 casas decimais (trabalhe em **centavos, como inteiro**, nunca em float puro, para nunca sofrer erro de arredondamento tipo `0.1 + 0.2 != 0.3`); o volume de lançamentos por conta/período é pequeno (dezenas, não milhões); e o tamanho dos grupos testados pode ser limitado (ex.: no máximo 6 lançamentos por grupo) sem perder a maioria dos casos reais.

## 5. Algoritmo de conciliação — cascata de regras (o "cérebro" do sistema)

Implemente as regras **em ordem de prioridade**, como passes sucessivos: a cada passe, remova do "pool" os lançamentos já conciliados antes de rodar o próximo. Cada regra grava, para cada match, **qual regra encontrou o match e os IDs exatos das linhas envolvidas** — isso é o que torna o sistema auditável.

1. **Match exato 1:1.** Para cada crédito, procure um débito de mesmo valor absoluto (tolerância configurável, padrão R$ 0,01), priorizando mesma `C/PARTIDA` e data mais próxima. Marque como `Conciliado - Exato`.
2. **Match por referência/texto.** Para o que sobrar, procure pares com valor próximo (tolerância um pouco mais larga) **e** histórico textualmente parecido (`rapidfuzz.fuzz.token_sort_ratio`, limiar mínimo configurável — padrão 80%). Marque como `Conciliado - Referência`.
3. **Match agrupado (N:1 ou 1:N), por período.** Teste se um grupo de N créditos soma (dentro da tolerância) a 1 débito, ou vice-versa, **escopado dentro do mesmo período** (ano, no exemplo — pode ser mês, dependendo do seu processo) para evitar combinações artificiais entre lançamentos sem relação nenhuma. Use `itertools.combinations` sobre os valores em centavos, testando tamanhos crescentes (2, 3, 4...) até um limite configurável (padrão 6). Marque como `Conciliado - Agrupado`, registrando os IDs de todos os lançamentos do grupo.
4. **Netting por período.** Para o que ainda sobrar dentro do mesmo período, some todos os créditos e débitos remanescentes daquele período; se a diferença for zero (dentro da tolerância), concilie o período inteiro como um bloco. Essa etapa é a rede de segurança para grupos maiores do que o limite testado na Etapa 3 (foi o que resolveu o Caso 1 da seção 3).
5. **FIFO global cronológico.** Para o que sobrar depois disso — agora **sem** escopo de período, já que uma provisão de um ano pode legitimamente só ser baixada em anos seguintes — ordene todos os créditos e débitos remanescentes por data crescente e compense o mais antigo de cada lado primeiro, permitindo baixa parcial. Isso não zera o saldo; reduz ao mínimo o que fica de fato em aberto e deixa explícito quais lançamentos específicos ainda restam (e com qual regra cada um foi tratado).
6. **Itens em aberto + aging.** Tudo que sobrar é marcado `Em Aberto` (ou `Em Aberto (parcial)`, se recebeu alguma baixa parcial na Etapa 5). Calcule o aging (dias entre a data do lançamento e a data de corte) e classifique em faixas (0-30, 31-60, 61-90, 91-180, 181-365, 365+ dias).

> Comece com essas 6 etapas como regras determinísticas e explicáveis — é o que um auditor consegue conferir manualmente. Só evolua para uma formulação de otimização (problema de atribuição via `PuLP`/`OR-Tools`, minimizando o saldo residual) se o volume de lançamentos crescer a ponto de o `itertools.combinations` da Etapa 3 ficar lento — não comece por aí.

## 6. Camada "inteligente" (o que diferencia de um script de conciliação comum)

1. **Classificação por texto, não só por sinal.** Extraia um "tipo de lançamento" do `HISTORICO` via palavras-chave (`PROV`/`PROVISÃO` → Provisão; `REV`/`REVERSÃO`/`ESTORNO` → Reversão; `RECLAS` → Reclassificação; `PG`/`PAGAMENTO` → Pagamento) e use isso como sinal **auxiliar** — nunca como único critério de direção (ver Caso 3 da seção 3).
2. **Correspondência difusa (`rapidfuzz`)** do texto do histórico, para achar pares que um match exato de valor não pegaria.
3. **Camada opcional de LLM (API da Anthropic) para os itens que sobrarem sem match determinístico:** envie ao modelo os lançamentos em aberto de um período (histórico, valor, contrapartida, data) e peça (a) uma sugestão de pareamento com justificativa textual, e (b) uma nota explicativa no mesmo estilo da coluna `Obs.` que já existe (ex.: `"Saldo em aberto referente a Provisão de Bônus - 2024"`), automatizando a redação do papel de trabalho da aba `Capa`. **Trate toda sugestão de LLM como rascunho para revisão humana, nunca como match automático definitivo** — grave-a em colunas separadas (`Sugestao_IA`, `Justificativa_IA`), nunca misturada com os status `Conciliado - *` das etapas determinísticas.
4. **Detecção de anomalias:** compare o valor de cada provisão mensal com a média móvel dos meses anteriores da mesma natureza; sinalize (não bloqueie) lançamentos fora do padrão para revisão prioritária.

## 7. Arquitetura técnica

- **Linguagem:** Python 3.11+
- **Dados:** `pandas` (regras e transformação), `openpyxl` (ler/escrever `.xlsx` preservando fórmulas e formatação)
- **Matching combinatório:** `itertools` para a Etapa 3; opcional `networkx` para modelar créditos × débitos como grafo bipartido (útil para depurar visualmente os matches); opcional `PuLP`/`OR-Tools` para a versão de otimização mencionada na seção 5
- **Texto difuso:** `rapidfuzz`
- **Camada de IA (opcional):** SDK `anthropic`
- **Interface:** CLI (`argparse`) como ponto de entrada padrão — `python main.py --arquivo caminho.xlsx --aba "01.Razão (2)" --conta 2.1.1.05.0005`. Como evolução, um dashboard em `streamlit` para o analista aprovar/rejeitar matches sugeridos antes de gerar o Excel final.
- **Logging e trilha de auditoria:** `logging` padrão, com data/hora, parâmetros usados e cada decisão de match. Em contabilidade isso não é opcional.
- **Testes:** `pytest`, com casos sintéticos cobrindo os 4 padrões da seção 3 **e** um teste de integração rodando contra o arquivo real, conferindo os invariantes financeiros (nada pode ser perdido ou duplicado; a ponte com o balancete fecha em zero).

### Estrutura de projeto usada na implementação de referência

```
conciliador-contabil/
├── src/
│   ├── excel_io.py           # leitura do razão + leitor de saldo do balancete
│   ├── classificador.py       # tipo textual + direção (crédito/débito)
│   ├── motor_conciliacao.py   # classe ConciliadorContabil - as 6 etapas
│   └── relatorios.py          # monta o Excel de saída (4 abas + formatação)
├── tests/
│   └── test_motor_conciliacao.py
├── config/
│   └── config.exemplo.yaml
├── main.py                    # CLI
├── requirements.txt
└── README.md
```

## 8. Contrato de entrada e saída

**Entrada (sempre parâmetros explícitos, nunca hardcoded no meio do código):** `arquivo_origem`, `nome_aba` (padrão `"01.Razão (2)"`), `linha_cabecalho` (padrão 6), `conta_contabil` (nunca deduzida de outra aba — ver o achado da seção 2), `tolerancia` (padrão R$ 0,01), `tamanho_max_grupo` (padrão 6), `similaridade_min` (padrão 80%), `data_corte` (para aging).

**Saída — um novo `.xlsx`** (nunca sobrescreva o original) com estas abas:

- **`Detalhe_Conciliacao`** — cada linha do razão original + `Status`, `ID_Match`, `Regra_Aplicada`, `Contraparte(s)`, `Valor_Residual`, `Aging_Dias`, `Faixa_Aging` — com formatação condicional (verde = conciliado, âmbar = em aberto ≤ 90 dias, coral = em aberto > 90 dias) e linha de totais com fórmula `SUBTOTAL` (não valor fixo)
- **`Resumo_Periodo`** — automatiza o que a aba `03.Conciliação` faz manualmente hoje: por período, total provisionado, total revertido/baixado, saldo líquido, saldo em aberto, % conciliado
- **`Itens_Em_Aberto`** — recorte só do que não fechou, ordenado por aging decrescente
- **`Ponte_Balancete`** — reproduz o formato da aba `Capa`: (1) saldo no balancete, (2) saldo no relatório auxiliar (soma dos itens em aberto pós-conciliação), (3) diferença — no arquivo de referência, essa diferença fecha em **R$ 0,00**

## 9. Casos de borda a tratar

- **Linhas de cabeçalho/rodapé não são dados** — filtre por `Data` não nula, nunca por posição fixa de linha além do cabeçalho
- **Nome de aba com parênteses** (`"01.Razão (2)"`) — cuidado ao referenciar em fórmulas entre aspas simples
- **Cache de fórmula desatualizado** — nunca confie em um valor de `SUBTOTAL` ou similar sem recalcular a partir dos dados brutos (ver seção 3)
- **Marca manual `Obs. = "Efeito zero"` já existente** — trate como sinal de bootstrap: se um grupo com essa marca já soma zero, é uma confirmação a mais (não a única fonte); se a marca existir mas o grupo **não** somar zero, isso é uma inconsistência de dados a reportar, não a ignorar
- **Múltiplos lançamentos com mesmo valor e mesma data** — ambíguo por natureza; defina um critério de desempate (ex.: mesma `C/PARTIDA`, depois ordem de leitura) e documente esse critério no relatório de saída, nunca escolha silenciosamente
- **Reversão em período fiscal diferente da provisão** — o match (Etapa 5, especialmente) tem que poder atravessar anos, como no Caso 4 da seção 3
- **Lançamentos de reclassificação (`RECLAS.VLR`)** — avalie se devem ser tratados como "transferência" (fora do cálculo de provisão × reversão) em vez de forçados a casar como se fossem baixa real; confirme essa regra com o usuário antes de assumir
- **A conta contábil é sempre um parâmetro explícito** — nunca deduza de uma aba de "capa" ou de outro arquivo (ver o achado da seção 2 — planilhas-modelo reaproveitadas de outras contas são comuns e é exatamente esse tipo de erro manual que o sistema deve eliminar)

## 10. Entregáveis esperados

1. `config.yaml` de exemplo já preenchido com os parâmetros do arquivo de referência
2. `excel_io.py` — leitura (schema da seção 3) e leitor auxiliar de saldo do balancete
3. `classificador.py` — tipo textual + direção
4. `motor_conciliacao.py` — classe `ConciliadorContabil` com as 6 etapas, cada uma retornando também a regra/motivo do match
5. `relatorios.py` — as 4 abas de saída, com formatação condicional e fórmulas nativas
6. `tests/test_motor_conciliacao.py` — cobrindo os 4 casos reais da seção 3 + integração
7. `main.py` — CLI ponta a ponta
8. `README.md` — como rodar, o que cada parâmetro faz, limitações conhecidas
9. (Evolução natural, não obrigatória na primeira versão) `app_revisao.py` em Streamlit para aprovar/rejeitar sugestões antes do Excel final; camada de IA da seção 6.3

## 11. Critérios de aceite (testados e confirmados na implementação de referência)

- [x] O total conciliado + o total em aberto, somados, bate com a soma da coluna Valor recalculada em Python (**R$ -425.931,84** no arquivo de referência) — nunca com o cache da fórmula do Excel
- [x] 2021 fecha 100% (12 créditos = 1 débito de R$ 324.000,00), via Etapa 4
- [x] 2022 fecha 100% (par de R$ 8.000,00 via Etapa 1 + dois grupos N:1 via Etapa 3)
- [x] A Ponte_Balancete fecha em **R$ 0,00** de diferença
- [x] Nenhuma linha fica sem status final (nem "Não processado" residual)
- [x] Nenhuma fórmula quebrada (`#NAME?`, `#REF!`, `#VALUE!`) no arquivo de saída

## 12. Boas práticas obrigatórias (não negociáveis)

- **Nunca sobrescreva o arquivo original** — sempre grave em um novo arquivo (ex.: `<nome_original>_conciliado_<AAAAMMDD>.xlsx`)
- **Toda conciliação automática precisa ser auditável** — se um auditor perguntar "por que essas duas linhas foram consideradas a mesma coisa?", o sistema responde com a regra exata e os IDs das linhas, nunca "a IA decidiu"
- **Revisão humana antes de fechar o mês** — trate a saída como uma proposta de conciliação, não como o fechamento contábil em si; isso vale especialmente para os matches da Etapa 5 (FIFO), que são uma suposição razoável, não um fato
- **Idempotência** — rodar o sistema duas vezes sobre o mesmo arquivo de entrada tem que produzir exatamente o mesmo resultado

---

*Este prompt foi construído e validado a partir da estrutura real de uma planilha de conciliação de Provisão de Bônus (conta de passivo), mas nomes de colunas, contas e regras de classificação por palavra-chave devem ser tratados como configuráveis — o objetivo final é reutilizar este mesmo motor em qualquer conta de provisão/passivo conciliada mês a mês, não só nesta.*
