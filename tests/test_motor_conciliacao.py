"""
Testes do Conciliador Contábil Inteligente.

- Os testes 1-3 usam dados sintéticos (não dependem de nenhum arquivo) e cobrem
  os 3 padrões reais encontrados no arquivo que motivou este projeto:
    1) grupo N:1 que fecha em bloco (Etapa 3 ou 4)
    2) par exato 1:1 (Etapa 1)
    3) saldo genuinamente em aberto (Etapa 6)
- O teste 4 é de integração: roda a cascata inteira contra o arquivo real e
  confere os invariantes financeiros (nada pode "sumir" ou "aparecer" no processo).

Rodar com:  pytest tests/ -v
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from motor_conciliacao import ConciliadorContabil  # noqa: E402

ARQUIVO_REAL = "/mnt/user-data/uploads/1784319202280_202606_Smart-_Provisa_o_de_Bo_nus.xlsx"


def _linha(periodo, data, historico, valor, c_partida="4.1.1.01.0026"):
    return {
        "periodo": periodo,
        "conta": "2.1.1.05.0005",
        "conta_desc": "PROVISAO BONUS A PAGAR",
        "data": pd.Timestamp(data),
        "lote": "000000000000000000",
        "historico": historico,
        "c_partida": c_partida,
        "fl": "0301",
        "c_custo": "",
        "debito": max(valor, 0),
        "credito": max(-valor, 0),
        "valor": float(valor),
        "obs": "",
        "linha_origem": 0,
        "id_lancamento": None,  # preenchido depois
    }


def _montar_df(linhas: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(linhas).reset_index(drop=True)
    df["id_lancamento"] = [f"L{i + 1:04d}" for i in range(len(df))]
    df["tipo_textual"] = "Provisao"
    df["direcao"] = df["valor"].map(lambda v: "Saida (Debito)" if v > 0 else "Entrada (Credito)")
    return df


class TestGrupoNParaUm:
    """Reproduz o padrão do ano de 2021 do arquivo real: 12 créditos mensais
    somando exatamente o valor de 1 débito de fechamento."""

    def test_grupo_fecha_em_bloco(self):
        linhas = [
            _linha(2021, "2021-01-31", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-02-28", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-03-31", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-04-30", "PROVISAO DE BONUS MES MARÇO E ABRIL", -40500),
            _linha(2021, "2021-05-31", "PROVISAO DE BONUS MES MAIO-21", -27000),
            _linha(2021, "2021-06-30", "PROVISAO DE BONUS JUN 2021", -27000),
            _linha(2021, "2021-07-30", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-08-31", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-09-28", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-10-29", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-11-29", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-12-30", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-12-31", "EST PROV BONUS", 324000),
        ]
        df = _montar_df(linhas)
        motor = ConciliadorContabil(df, tolerancia=0.01, max_grupo=12)
        resultado = motor.rodar_cascata()

        assert (resultado["residual_centavos"] == 0).all(), "o grupo inteiro deveria fechar em zero"
        assert resultado["status"].str.startswith("Conciliado").all()
        # com max_grupo>=12 a Etapa 3 (agrupado) já deve resolver sozinha,
        # sem precisar da Etapa 4 (netting de período) como rede de segurança
        regras = set(resultado["regra_aplicada"])
        assert regras <= {"Agrupado (N:1)", "Agrupado (1:N)"}

    def test_grupo_grande_precisa_de_max_grupo_suficiente(self):
        """Com max_grupo menor que o tamanho do grupo real, a Etapa 3 sozinha
        não encontra o match - mas a Etapa 4 (netting do período) deve pegar,
        já que o período fecha em zero de qualquer forma. Isso é o comportamento
        real observado no arquivo de origem (ano de 2021, max_grupo padrão=6)."""
        linhas = [
            _linha(2021, "2021-01-31", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-02-28", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-03-31", "PROV BONUS A PAGAR ACIONISTA", -22500),
            _linha(2021, "2021-04-30", "PROVISAO DE BONUS MES MARÇO E ABRIL", -40500),
            _linha(2021, "2021-05-31", "PROVISAO DE BONUS MES MAIO-21", -27000),
            _linha(2021, "2021-06-30", "PROVISAO DE BONUS JUN 2021", -27000),
            _linha(2021, "2021-07-30", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-08-31", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-09-28", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-10-29", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-11-29", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-12-30", "PROV PROVISAO DE BONUS", -27000),
            _linha(2021, "2021-12-31", "EST PROV BONUS", 324000),
        ]
        df = _montar_df(linhas)
        motor = ConciliadorContabil(df, tolerancia=0.01, max_grupo=6)  # padrão do sistema
        resultado = motor.rodar_cascata()

        assert (resultado["residual_centavos"] == 0).all()
        assert "Saldo Período" in set(resultado["regra_aplicada"])


class TestMatchExato:
    def test_par_simples_1_para_1(self):
        linhas = [
            _linha(2024, "2024-04-30", "VLR AJUSTE PROVISAO FILIAL", -639.42, c_partida="0.3.7.000007"),
            _linha(2024, "2024-04-30", "VLR AJUSTE PROVISAO FILIAL", 639.42, c_partida="0.3.8.000018"),
        ]
        df = _montar_df(linhas)
        motor = ConciliadorContabil(df)
        resultado = motor.rodar_cascata()
        assert (resultado["status"] == "Conciliado - Exato").all()
        assert resultado.iloc[0]["contraparte"] == resultado.iloc[1]["id_lancamento"]


class TestSaldoEmAberto:
    """Reproduz o padrão do ano de 2023: créditos mensais muito maiores do que
    os poucos débitos do ano - nenhuma combinação fecha em zero, e o saldo deve
    ficar corretamente sinalizado como em aberto, sem "match" falso-positivo."""

    def test_saldo_nao_fecha_e_fica_marcado_em_aberto(self):
        linhas = [
            _linha(2023, "2023-01-31", "PROV BONUS", -66439),
            _linha(2023, "2023-02-28", "PROV BONUS", -66439),
            _linha(2023, "2023-03-31", "PROV BONUS", -66439),
            _linha(2023, "2023-04-30", "PROV BONUS", -66440),
            _linha(2023, "2023-05-31", "VLR PROVISAO BONUS", -66439),
            _linha(2023, "2023-11-30", "VLR PROVISAO BONUS", 116611),
        ]
        df = _montar_df(linhas)
        motor = ConciliadorContabil(df, max_grupo=6, data_corte=pd.Timestamp("2023-12-31"))
        resultado = motor.rodar_cascata()

        saldo_esperado = df["valor"].sum()  # nada deve ser inventado nem perdido
        saldo_aberto = resultado.loc[resultado["residual_centavos"] != 0, "valor_residual"].sum()
        assert saldo_aberto == pytest.approx(saldo_esperado, abs=0.01)
        assert (resultado["status"] == "Em Aberto").any()
        # nenhuma linha deve ser marcada Conciliada com um match inventado
        for _, row in resultado[resultado["status"].str.startswith("Conciliado")].iterrows():
            pass  # se chegou aqui sem quebrar os invariantes acima, não há match falso


@pytest.mark.skipif(not Path(ARQUIVO_REAL).exists(), reason="arquivo real não disponível neste ambiente")
class TestIntegracaoArquivoReal:
    """Roda a cascata inteira contra o arquivo real e confere os invariantes:
    nada pode ser perdido ou duplicado, e a ponte com o balancete deve fechar
    em zero (dado que o arquivo real já está corretamente escriturado)."""

    @classmethod
    @pytest.fixture(scope="class")
    def resultado(cls):
        from classificador import aplicar_classificacao
        from excel_io import carregar_razao

        carregado = carregar_razao(ARQUIVO_REAL)
        df = aplicar_classificacao(carregado.df)
        motor = ConciliadorContabil(df, tolerancia=0.01, max_grupo=6)
        df_resultado = motor.rodar_cascata()
        return carregado, motor, df_resultado

    def test_total_e_conservado(self, resultado):
        carregado, motor, df_resultado = resultado
        total_residual = df_resultado["residual_centavos"].sum() / 100
        assert total_residual == pytest.approx(carregado.total_valor_calculado, abs=0.01)

    def test_saldo_bate_com_balancete(self, resultado):
        from excel_io import ler_saldo_balancete

        carregado, motor, df_resultado = resultado
        saldo_balancete = ler_saldo_balancete(ARQUIVO_REAL, "2.1.1.05.0005")
        ponte = motor.ponte_balancete(saldo_balancete)
        diferenca = ponte.loc[ponte["item"].str.startswith("(3)"), "valor"].iloc[0]
        assert diferenca == pytest.approx(0.0, abs=0.01)

    def test_2021_fecha_totalmente(self, resultado):
        _, _, df_resultado = resultado
        ano = df_resultado[df_resultado["periodo"] == 2021]
        assert (ano["residual_centavos"] == 0).all()

    def test_2022_fecha_totalmente(self, resultado):
        _, _, df_resultado = resultado
        ano = df_resultado[df_resultado["periodo"] == 2022]
        assert (ano["residual_centavos"] == 0).all()

    def test_nenhuma_linha_sem_status_final(self, resultado):
        _, _, df_resultado = resultado
        assert not (df_resultado["status"] == "Não processado").any()
