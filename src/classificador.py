"""
classificador.py
Classifica cada lancamento por DIRECAO (sinal do Valor) e por TIPO TEXTUAL
(palavra-chave no HISTORICO). Os dois sinais sao independentes e devem ser
usados em conjunto pelo motor de conciliacao - nenhum dos dois sozinho basta.

Por que os dois separados: no arquivo que motivou este projeto, o texto
"PROV BONUS" aparece tanto em creditos (provisao nova) quanto em debitos de
fechamento de ano (estorno/baixa). Ou seja, o HISTORICO nao identifica com
seguranca a direcao contabil - so o sinal (Valor = Debito - Credito) e
confiavel para isso. O texto ainda ajuda a agrupar lancamentos que "conversam"
entre si (ex.: mesma referencia de folha de pagamento).
"""
from __future__ import annotations

import re

import pandas as pd

# Ordem importa: o primeiro padrao que casar vence. RECLAS antes de REV porque
# historicos como "RECLAS.VLR FOPG 313 - PREMIO" nao devem cair em Reversao.
_PADROES_TIPO = [
    ("Reclassificacao", re.compile(r"\bRECLAS\b", re.IGNORECASE)),
    ("Reversao/Estorno", re.compile(r"\b(REV|REVERS|ESTORNO|EST)\b", re.IGNORECASE)),
    ("Pagamento", re.compile(r"\b(PG|PAGO|PAGTO|PAGAMENTO)\b", re.IGNORECASE)),
    ("Ajuste", re.compile(r"\bAJUSTE\b", re.IGNORECASE)),
    ("Provisao", re.compile(r"\bPROV", re.IGNORECASE)),
]


def classificar_tipo_textual(historico: str) -> str:
    """Devolve um rotulo textual (Provisao / Reversao-Estorno / Reclassificacao /
    Pagamento / Ajuste / Outro) a partir de palavras-chave no HISTORICO.
    Isto e um sinal AUXILIAR - nunca decide sozinho a direcao contabil."""
    texto = historico or ""
    for rotulo, padrao in _PADROES_TIPO:
        if padrao.search(texto):
            return rotulo
    return "Outro"


def classificar_direcao(valor: float) -> str:
    """A UNICA fonte confiavel de direcao: o sinal do Valor (Debito-Credito)."""
    if valor > 0:
        return "Saida (Debito)"
    if valor < 0:
        return "Entrada (Credito)"
    return "Neutro (zero)"


def aplicar_classificacao(df: pd.DataFrame) -> pd.DataFrame:
    """Adiciona as colunas tipo_textual e direcao ao DataFrame do razao."""
    df = df.copy()
    df["tipo_textual"] = df["historico"].map(classificar_tipo_textual)
    df["direcao"] = df["valor"].map(classificar_direcao)
    return df
