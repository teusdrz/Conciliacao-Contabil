"""
motor_conciliacao.py
O motor do Conciliador Contabil Inteligente: uma cascata de 6 regras que
concilia lancamentos a credito (entrada/provisao, Valor<0) com lancamentos a
debito (saida/reversao, Valor>0), na ordem:

    1. Match exato 1:1               (mesmo valor absoluto, um credito x um debito,
                                       em toda a base, sem restricao de periodo/ano)
    2. Match por referencia/texto     (valor proximo + historico parecido, tambem
                                       em toda a base, sem restricao de periodo/ano)
    3. Match agrupado N:1 / 1:N / N:M (soma de subconjuntos que fecha em zero,
                                       em toda a base - credito(s) x debito(s)
                                       podem estar em anos/periodos diferentes)
    4. Netting por periodo            (fecha o residuo do periodo inteiro em bloco)
    5. FIFO global                    (compensa cronologicamente o que sobrar,
                                       inclusive atravessando anos, com baixa parcial)
    6. Itens em aberto + aging        (o que nao fechou em nenhuma etapa anterior)

Tudo trabalha em CENTAVOS (inteiros) para nunca sofrer erro de arredondamento
de ponto flutuante. Cada match grava a regra que o encontrou e os IDs das
linhas envolvidas - nenhum match "silencioso".
"""
from __future__ import annotations

import logging
from dataclasses import dataclass, field

import pandas as pd
from rapidfuzz import fuzz

logger = logging.getLogger("conciliador.motor")

FAIXAS_AGING = [
    (0, 30, "0-30 dias"),
    (31, 60, "31-60 dias"),
    (61, 90, "61-90 dias"),
    (91, 180, "91-180 dias"),
    (181, 365, "181-365 dias"),
    (366, 10_000, "> 365 dias"),
]


def _faixa_aging(dias: int) -> str:
    for lo, hi, rotulo in FAIXAS_AGING:
        if lo <= dias <= hi:
            return rotulo
    return "> 365 dias"


@dataclass
class EventoAuditoria:
    etapa: str
    regra: str
    id_grupo: str
    ids_lancamentos: list[str]
    valor_total_centavos: int
    detalhe: str = ""


class ConciliadorContabil:
    def __init__(
        self,
        df: pd.DataFrame,
        tolerancia: float = 0.01,
        max_grupo: int = 6,
        similaridade_min: int = 80,
        data_corte: pd.Timestamp | None = None,
    ):
        self.tol_cent = max(1, round(tolerancia * 100))
        self.max_grupo = max_grupo
        self.similaridade_min = similaridade_min

        self.df = df.copy().reset_index(drop=True)
        self.data_corte = pd.Timestamp(data_corte) if data_corte is not None else self.df["data"].max()

        self.df["valor_centavos"] = (self.df["valor"] * 100).round().astype(int)
        self.df["residual_centavos"] = self.df["valor_centavos"]
        self.df["status"] = "Não processado"
        self.df["id_match"] = pd.array([None] * len(self.df), dtype="string")
        self.df["regra_aplicada"] = ""
        self.df["contraparte"] = ""
        # a coluna Obs. do arquivo de origem pode trazer texto velho/manual
        # (ex.: planilhas reais as vezes tem "Efeito zero" digitado numa linha
        # que na verdade ainda esta em aberto, ou lixo de outra analise). O
        # motor e a UNICA fonte da verdade para o Obs de saida: sempre comeca
        # limpo e so grava algo quando a propria cascata concilia a linha.
        self.df["obs"] = ""

        self._contador_grupo = 0
        self.trilha_auditoria: list[EventoAuditoria] = []

    # ------------------------------------------------------------------ util
    def _novo_id_grupo(self) -> str:
        self._contador_grupo += 1
        return f"M{self._contador_grupo:04d}"

    def _pool_aberto(self) -> pd.DataFrame:
        return self.df[self.df["residual_centavos"] != 0]

    def _marcar_grupo(
        self, indices: list[int], regra: str, etapa: str, detalhe: str = "",
        tolerancia_cent: int | None = None,
    ) -> str:
        id_grupo = self._novo_id_grupo()
        ids_lote = self.df.loc[indices, "id_lancamento"].tolist()
        valor_total = int(self.df.loc[indices, "residual_centavos"].sum())
        # "efeito zero" so pode ser gravado no Obs se a soma do VALOR BRUTO
        # original das linhas deste grupo realmente fecha em zero (dentro da
        # tolerancia) - nao basta cada linha ter chegado a residual zero
        # individualmente. Isso importa na Etapa 5 (FIFO): um componente
        # cronologico pode zerar o residual de uma linha usando valor "emprestado"
        # de uma contraparte que ficou so parcialmente baixada (fora deste grupo);
        # nesse caso a soma bruta do subconjunto marcado aqui nao fecha em zero,
        # entao o Obs nao deve dizer "efeito zero".
        tol = self.tol_cent if tolerancia_cent is None else tolerancia_cent
        soma_valor_bruto = int(self.df.loc[indices, "valor_centavos"].sum())
        fecha_em_zero = abs(soma_valor_bruto) <= tol
        for idx in indices:
            meu_id = self.df.loc[idx, "id_lancamento"]
            contrapartes = [i for i in ids_lote if i != meu_id]
            self.df.loc[idx, "status"] = f"Conciliado - {regra}"
            self.df.loc[idx, "id_match"] = id_grupo
            self.df.loc[idx, "regra_aplicada"] = regra
            self.df.loc[idx, "contraparte"] = ", ".join(contrapartes)
            self.df.loc[idx, "residual_centavos"] = 0
            self.df.loc[idx, "obs"] = (
                "efeito zero" if fecha_em_zero else "baixa parcial (FIFO) - ver contraparte"
            )
        self.trilha_auditoria.append(
            EventoAuditoria(etapa, regra, id_grupo, ids_lote, valor_total, detalhe)
        )
        return id_grupo

    # ------------------------------------------------------------- etapa 1
    def etapa1_match_exato(self) -> "ConciliadorContabil":
        pool = self._pool_aberto()
        creditos = pool[pool["residual_centavos"] < 0].sort_values("data")
        usados: set[int] = set()
        for idx_c, c in creditos.iterrows():
            if idx_c in usados:
                continue
            alvo = -c["residual_centavos"]
            pool_atual = self._pool_aberto()
            debitos = pool_atual[
                (pool_atual["residual_centavos"] > 0)
                & (~pool_atual.index.isin(usados))
                & ((pool_atual["residual_centavos"] - alvo).abs() <= self.tol_cent)
            ].copy()
            if debitos.empty:
                continue
            debitos["mesma_contrapartida"] = debitos["c_partida"] == c["c_partida"]
            debitos["dist_dias"] = (debitos["data"] - c["data"]).abs().dt.days
            debitos = debitos.sort_values(["mesma_contrapartida", "dist_dias"], ascending=[False, True])
            idx_d = debitos.index[0]
            self._marcar_grupo([idx_c, idx_d], "Exato", "1")
            usados.add(idx_c)
            usados.add(idx_d)
        logger.info("Etapa 1 (match exato 1:1): %d par(es) conciliado(s).", len(usados) // 2)
        return self

    # ------------------------------------------------------------- etapa 2
    def etapa2_match_referencia(self) -> "ConciliadorContabil":
        pool = self._pool_aberto()
        creditos = pool[pool["residual_centavos"] < 0].sort_values("data")
        n_matches = 0
        usados: set[int] = set()
        tol_larga = self.tol_cent * 5  # tolerancia um pouco mais folgada que a Etapa 1
        for idx_c, c in creditos.iterrows():
            if idx_c in usados:
                continue
            alvo = -c["residual_centavos"]
            pool_atual = self._pool_aberto()
            candidatos = pool_atual[
                (pool_atual["residual_centavos"] > 0)
                & (~pool_atual.index.isin(usados))
                & ((pool_atual["residual_centavos"] - alvo).abs() <= tol_larga)
            ].copy()
            if candidatos.empty:
                continue
            candidatos["similaridade"] = candidatos["historico"].map(
                lambda h: fuzz.token_sort_ratio(str(h), str(c["historico"]))
            )
            candidatos = candidatos[candidatos["similaridade"] >= self.similaridade_min]
            if candidatos.empty:
                continue
            candidatos = candidatos.sort_values("similaridade", ascending=False)
            idx_d = candidatos.index[0]
            self._marcar_grupo(
                [idx_c, idx_d], "Referência", "2",
                detalhe=f"similaridade texto={candidatos.loc[idx_d, 'similaridade']:.0f}%",
                tolerancia_cent=tol_larga,
            )
            usados.add(idx_c)
            usados.add(idx_d)
            n_matches += 1
        logger.info("Etapa 2 (match por referência/texto): %d par(es) conciliado(s).", n_matches)
        return self

    # ------------------------------------------------------------- etapa 3
    _LIMITE_ESTADOS_SUBCONJUNTO = 4_000  # trava de seguranca contra explosao combinatoria

    @classmethod
    def _tabela_somas_alcancaveis(
        cls, itens: list[tuple[int, int]], limite_superior: int, tam_max: int
    ) -> list[dict[int, list[int]]]:
        """itens: [(indice_df, valor_centavos), ...], todos com o mesmo sinal
        (ja normalizados como positivos por quem chama). Constroi, em uma
        unica passada, todas as somas alcancaveis usando de 1 a 'tam_max'
        itens sem ultrapassar 'limite_superior'.

        Implementado como programacao dinamica podada pelo valor-alvo (nao
        como forca bruta de itertools.combinations): como todos os itens tem
        o mesmo sinal, qualquer soma parcial que ja ultrapasse o limite
        nunca pode virar um match valido e e descartada na hora. A tabela e
        montada UMA VEZ e reaproveitada para todos os alvos do lote (em vez
        de refazer a busca do zero a cada alvo), o que evita a explosao
        combinatoria de C(n, tam_max) quando a base de busca e grande (ex.:
        toda a base em vez de so um periodo). 'limite_estados' funciona como
        trava de seguranca: se um nivel crescer alem do limite, paramos de
        expandi-lo ali (protege contra travamentos em cenarios sem solucao
        real, sem afetar o resultado nos casos normais).
        """
        limite_estados = cls._LIMITE_ESTADOS_SUBCONJUNTO
        estados: list[dict[int, list[int]]] = [dict() for _ in range(tam_max + 1)]
        estados[0][0] = []
        for idx, valor in itens:
            if valor > limite_superior:
                continue  # item sozinho ja estoura o alvo, nunca entra em nenhuma combinacao
            for tam in range(tam_max - 1, -1, -1):
                origem = estados[tam]
                if not origem:
                    continue
                destino = estados[tam + 1]
                if len(destino) >= limite_estados:
                    continue
                for soma_atual, combo_atual in origem.items():
                    nova_soma = soma_atual + valor
                    if nova_soma > limite_superior:
                        continue
                    if nova_soma not in destino:
                        destino[nova_soma] = combo_atual + [idx]
                        if len(destino) >= limite_estados:
                            break
        return estados

    @staticmethod
    def _buscar_na_tabela(
        estados: list[dict[int, list[int]]], alvo: int, tol: int, usados: set[int] = frozenset()
    ) -> list[int] | None:
        """Procura, na tabela ja construida, a menor combinacao (tamanho
        crescente) cuja soma bate com 'alvo' dentro da tolerancia, ignorando
        combinacoes que usem algum indice ja consumido por outro match neste
        mesmo lote ('usados'). Como a tolerancia e pequena, basta checar as
        poucas chaves vizinhas de 'alvo' em vez de varrer a tabela inteira -
        assim a busca por alvo continua rapida mesmo com a tabela cheia."""
        for tam in range(2, len(estados)):
            nivel = estados[tam]
            if not nivel:
                continue
            for delta in range(-tol, tol + 1):
                combo = nivel.get(alvo + delta)
                if combo is not None and usados.isdisjoint(combo):
                    return combo
        return None

    def _etapa3_um_lado(self, direcao_alvo: str) -> int:
        """direcao_alvo='debito' -> procura 1 debito == soma de N creditos (N:1)
        direcao_alvo='credito' -> procura 1 credito == soma de N debitos (1:N)

        Busca em toda a base (sem filtrar por periodo/ano): o credito e os
        debitos que somam com ele podem estar em anos diferentes - o que
        importa e o conjunto zerar (positivo compensando negativo) dentro da
        tolerancia configurada.
        """
        n_matches = 0
        progresso = True
        while progresso:
            progresso = False
            pool = self._pool_aberto()
            if direcao_alvo == "debito":
                alvos = pool[pool["residual_centavos"] > 0].sort_values(
                    "residual_centavos", ascending=False
                )
                fonte = pool[pool["residual_centavos"] < 0]
            else:
                alvos = pool[pool["residual_centavos"] < 0].sort_values(
                    "residual_centavos", ascending=True
                )
                fonte = pool[pool["residual_centavos"] > 0]
            if alvos.empty or len(fonte) < 2:
                break
            itens_fonte = [(idx, -v if direcao_alvo == "debito" else v)
                           for idx, v in fonte["residual_centavos"].items()]
            tam_max = min(self.max_grupo, len(itens_fonte))
            if tam_max < 2:
                break
            limite_superior = int(alvos["residual_centavos"].abs().max()) + self.tol_cent
            tabela = self._tabela_somas_alcancaveis(itens_fonte, limite_superior, tam_max)

            # Reaproveita a MESMA tabela para varios alvos do lote (em vez de
            # reconstrui-la a cada match individual) - assim que um alvo e um
            # combo sao consumidos, seus indices entram em 'usados' para que
            # nenhum outro alvo do mesmo lote tente reutiliza-los.
            usados: set[int] = set()
            for idx_alvo, valor_alvo in alvos["residual_centavos"].items():
                if idx_alvo in usados:
                    continue
                alvo_abs = abs(valor_alvo)
                combo = self._buscar_na_tabela(tabela, alvo_abs, self.tol_cent, usados)
                if combo:
                    regra = "Agrupado (N:1)" if direcao_alvo == "debito" else "Agrupado (1:N)"
                    self._marcar_grupo(
                        combo + [idx_alvo], regra, "3",
                        detalhe=f"{len(combo)} lançamento(s) vs. 1 (toda a base, todos os períodos)",
                    )
                    usados.update(combo)
                    usados.add(idx_alvo)
                    n_matches += 1
                    progresso = True
            # Reconstroi a tabela apenas quando o lote inteiro ja foi
            # percorrido (progresso=True refaz o pool para pegar o que ainda
            # sobrou), em vez de a cada match individual.
        return n_matches

    def etapa3_match_agrupado(self) -> "ConciliadorContabil":
        n1 = self._etapa3_um_lado("debito")   # N creditos : 1 debito
        n2 = self._etapa3_um_lado("credito")  # 1 credito : N debitos
        logger.info(
            "Etapa 3 (match agrupado N:1 / 1:N em toda a base): %d + %d grupo(s) conciliado(s).",
            n1, n2,
        )
        return self

    # ------------------------------------------------------------- etapa 4
    def etapa4_netting_periodo(self) -> "ConciliadorContabil":
        n_periodos = 0
        for periodo in sorted(self.df["periodo"].dropna().unique()):
            pool = self._pool_aberto()
            pool = pool[pool["periodo"] == periodo]
            if pool.empty:
                continue
            residuo = int(pool["residual_centavos"].sum())
            if abs(residuo) <= self.tol_cent:
                self._marcar_grupo(
                    list(pool.index), "Saldo Período", "4",
                    detalhe=f"período {periodo} fecha em bloco (resíduo {residuo/100:+.2f})",
                )
                n_periodos += 1
        logger.info("Etapa 4 (netting por período): %d período(s) fechado(s) em bloco.", n_periodos)
        return self

    # ------------------------------------------------------------- etapa 5
    def etapa5_fifo_global(self) -> "ConciliadorContabil":
        pool = self._pool_aberto().sort_values("data")
        creditos = [[idx, -v] for idx, v in pool[pool["residual_centavos"] < 0]["residual_centavos"].items()]
        debitos = [[idx, v] for idx, v in pool[pool["residual_centavos"] > 0]["residual_centavos"].items()]

        i, j = 0, 0
        trocas: list[tuple[int, int, int]] = []  # (idx_credito, idx_debito, valor_centavos)
        while i < len(creditos) and j < len(debitos):
            idx_c, res_c = creditos[i]
            idx_d, res_d = debitos[j]
            montante = min(res_c, res_d)
            trocas.append((idx_c, idx_d, montante))
            creditos[i][1] -= montante
            debitos[j][1] -= montante
            self.df.loc[idx_c, "residual_centavos"] = -creditos[i][1]
            self.df.loc[idx_d, "residual_centavos"] = debitos[j][1]
            if creditos[i][1] <= self.tol_cent:
                i += 1
            if debitos[j][1] <= self.tol_cent:
                j += 1

        # agrupa trocas em componentes conexos (uma linha pode ter sido parcialmente
        # baixada por mais de uma contraparte ao longo do processo)
        grafo: dict[int, set[int]] = {}
        for idx_c, idx_d, _ in trocas:
            grafo.setdefault(idx_c, set()).add(idx_d)
            grafo.setdefault(idx_d, set()).add(idx_c)

        visitados: set[int] = set()
        n_grupos_fifo = 0
        n_linhas_zeradas = 0
        for no in grafo:
            if no in visitados:
                continue
            componente, fila = set(), [no]
            while fila:
                atual = fila.pop()
                if atual in componente:
                    continue
                componente.add(atual)
                fila.extend(grafo.get(atual, set()) - componente)
            visitados |= componente

            zerados = [idx for idx in componente if self.df.loc[idx, "residual_centavos"] == 0]
            parciais = [idx for idx in componente if self.df.loc[idx, "residual_centavos"] != 0]
            if zerados:
                id_grupo = self._marcar_grupo(zerados, "FIFO", "5", detalhe="baixa cronológica")
                n_grupos_fifo += 1
                n_linhas_zeradas += len(zerados)
                # linhas parciais dentro do mesmo componente recebem o mesmo id_match
                # de referência, mas continuam com residual != 0 (tratadas na etapa 6)
                for idx in parciais:
                    self.df.loc[idx, "id_match"] = id_grupo
                    self.df.loc[idx, "contraparte"] = ", ".join(
                        self.df.loc[zerados, "id_lancamento"].tolist()
                    )
        logger.info(
            "Etapa 5 (FIFO global cronológico): %d linha(s) totalmente baixada(s) em %d grupo(s).",
            n_linhas_zeradas, n_grupos_fifo,
        )
        return self

    # ------------------------------------------------------------- etapa 6
    def etapa6_itens_em_aberto(self) -> "ConciliadorContabil":
        aberto = self.df["residual_centavos"] != 0
        parcialmente_baixado = self.df["id_match"].notna() & aberto
        self.df.loc[aberto & parcialmente_baixado, "status"] = "Em Aberto (parcial)"
        self.df.loc[aberto & ~parcialmente_baixado, "status"] = "Em Aberto"
        self.df.loc[aberto, "regra_aplicada"] = self.df.loc[aberto, "regra_aplicada"].where(
            self.df.loc[aberto, "regra_aplicada"] != "", "Nenhuma - saldo em aberto"
        )

        self.df["aging_dias"] = (self.data_corte - self.df["data"]).dt.days
        self.df["faixa_aging"] = self.df["aging_dias"].map(_faixa_aging)
        self.df.loc[~aberto, ["aging_dias", "faixa_aging"]] = None

        logger.info("Etapa 6 (itens em aberto): %d linha(s) permanecem em aberto.", int(aberto.sum()))
        return self

    # --------------------------------------------------------------- runner
    def rodar_cascata(self) -> pd.DataFrame:
        (
            self.etapa1_match_exato()
            .etapa2_match_referencia()
            .etapa3_match_agrupado()
            .etapa4_netting_periodo()
            .etapa5_fifo_global()
            .etapa6_itens_em_aberto()
        )
        self.df["valor_residual"] = self.df["residual_centavos"] / 100
        self.df["valor_conciliado"] = self.df["valor"] - self.df["valor_residual"]
        return self.df

    # ---------------------------------------------------------- resumo/ponte
    def resumo_por_periodo(self) -> pd.DataFrame:
        df = self.df
        linhas = []
        for periodo in sorted(df["periodo"].dropna().unique()):
            grupo = df[df["periodo"] == periodo]
            provisao = -grupo.loc[grupo["valor"] < 0, "valor"].sum()  # positivo p/ leitura
            reversao = grupo.loc[grupo["valor"] > 0, "valor"].sum()
            saldo = grupo["valor"].sum()
            saldo_aberto = grupo["residual_centavos"].sum() / 100
            pct_conciliado = 0.0 if grupo["valor"].abs().sum() == 0 else (
                1 - grupo["residual_centavos"].abs().sum() / (abs(grupo["valor_centavos"]).sum() or 1)
            )
            linhas.append(
                {
                    "periodo": int(periodo),
                    "total_provisionado": round(provisao, 2),
                    "total_revertido_baixado": round(reversao, 2),
                    "saldo_liquido_periodo": round(saldo, 2),
                    "saldo_em_aberto_no_periodo": round(saldo_aberto, 2),
                    "pct_conciliado": round(pct_conciliado * 100, 1),
                }
            )
        return pd.DataFrame(linhas)

    def ponte_balancete(self, saldo_balancete: float | None) -> pd.DataFrame:
        relatorio_auxiliar = round(float(self.df["residual_centavos"].sum()) / 100, 2)
        linhas = [
            {"item": "(1) Saldo no Balancete", "valor": saldo_balancete},
            {"item": "(2) Saldo no Relatório Auxiliar (itens em aberto pós-conciliação)", "valor": relatorio_auxiliar},
        ]
        if saldo_balancete is not None:
            diferenca = round(saldo_balancete - relatorio_auxiliar, 2)
            linhas.append({"item": "(3) Diferença (1) - (2)", "valor": diferenca})
        return pd.DataFrame(linhas)
