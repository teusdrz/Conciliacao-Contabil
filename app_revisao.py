"""
app_revisao.py
Painel visual do Conciliador Contabil Inteligente: arraste o Excel do razao,
confira os parametros (ja vem com valores padrao sensatos) e rode a
conciliacao direto no navegador - sem precisar editar config.yaml nem
lidar com caminho de arquivo na linha de comando.

Rodar com:
    streamlit run app_revisao.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent / "src"))

from classificador import aplicar_classificacao  # noqa: E402
from excel_io import carregar_razao, ler_saldo_balancete  # noqa: E402
from motor_conciliacao import ConciliadorContabil  # noqa: E402
from relatorios import gerar_excel_saida  # noqa: E402

st.set_page_config(page_title="Conciliador Contábil Inteligente", layout="wide")

CORES_STATUS = {
    "Conciliado - Exato": "#1D9E75",
    "Conciliado - Referência": "#5DCAA5",
    "Conciliado - Agrupado (N:1)": "#0F6E56",
    "Conciliado - Agrupado (1:N)": "#0F6E56",
    "Conciliado - Saldo Período": "#085041",
    "Conciliado - FIFO": "#639922",
    "Em Aberto (parcial)": "#EF9F27",
    "Em Aberto": "#D85A30",
}


def _cor_linha(status: str) -> str:
    cor = CORES_STATUS.get(status, "#B4B2A9")
    return f"background-color: {cor}22"  # transparência leve, so funciona bem em fundo claro ou escuro


def main() -> None:
    st.title("Conciliador Contábil Inteligente")
    st.caption(
        "Arraste o Excel do razão contábil, confira os parâmetros na barra lateral e rode a "
        "conciliação. O resultado aparece aqui e também fica disponível para download."
    )

    with st.sidebar:
        st.header("Parâmetros")
        aba = st.text_input("Aba do razão", value="01.Razão (2)")
        linha_cabecalho = st.number_input("Linha do cabeçalho", min_value=1, value=6)
        conta = st.text_input("Código da conta (opcional, para a Ponte com o balancete)", value="")
        st.divider()
        tolerancia = st.number_input("Tolerância (R$)", min_value=0.0, value=0.01, step=0.01, format="%.2f")
        max_grupo = st.slider("Tamanho máximo de grupo (Etapa 3)", min_value=2, max_value=15, value=6)
        similaridade_min = st.slider("Similaridade mínima de texto % (Etapa 2)", min_value=50, max_value=100, value=80)
        st.divider()
        usar_balancete = st.checkbox("Ler saldo do balancete automaticamente (aba 00.Balancete do mesmo arquivo)", value=True)
        aba_balancete = st.text_input("Aba do balancete", value="00.Balancete", disabled=not usar_balancete)
        saldo_manual = st.number_input(
            "Ou informe o saldo do balancete manualmente (R$)",
            value=0.0, step=0.01, format="%.2f",
            help="Deixe 0,00 e marque a opção acima se quiser que o sistema tente ler sozinho.",
        )

    arquivo_up = st.file_uploader("Arraste ou selecione o Excel do razão contábil (.xlsx)", type=["xlsx"])

    if arquivo_up is None:
        st.info("Nenhum arquivo carregado ainda. Assim que você soltar o `.xlsx` aqui, o botão de rodar aparece.")
        return

    tmp_dir = Path(tempfile.mkdtemp(prefix="conciliador_"))
    caminho_tmp = tmp_dir / arquivo_up.name
    caminho_tmp.write_bytes(arquivo_up.getvalue())

    if not st.button("▶ Rodar conciliação", type="primary"):
        st.caption(f"Arquivo carregado: **{arquivo_up.name}** ({arquivo_up.size / 1024:.0f} KB). Clique no botão para rodar.")
        return

    with st.spinner("Lendo o razão, classificando e rodando a cascata de conciliação..."):
        try:
            carregado = carregar_razao(caminho_tmp, aba=aba, linha_cabecalho=int(linha_cabecalho))
        except Exception as exc:  # noqa: BLE001
            st.error(f"Não consegui ler a aba '{aba}' desse arquivo. Detalhe técnico: {exc}")
            return

        if carregado.diferenca_cache != 0:
            st.warning(
                f"Aviso: o total recalculado (R$ {carregado.total_valor_calculado:,.2f}) difere do valor em "
                f"cache de alguma fórmula de conferência da planilha original "
                f"(R$ {carregado.total_valor_cache_formula:,.2f}). O sistema sempre usa o valor recalculado."
            )

        df = aplicar_classificacao(carregado.df)
        motor = ConciliadorContabil(
            df, tolerancia=tolerancia, max_grupo=int(max_grupo), similaridade_min=int(similaridade_min)
        )
        resultado = motor.rodar_cascata()
        resumo = motor.resumo_por_periodo()

        saldo_balancete = None
        if usar_balancete and conta:
            saldo_balancete = ler_saldo_balancete(caminho_tmp, conta, aba=aba_balancete)
        if saldo_balancete is None and saldo_manual != 0:
            saldo_balancete = saldo_manual
        ponte = motor.ponte_balancete(saldo_balancete)

        caminho_saida = tmp_dir / f"{Path(arquivo_up.name).stem}_conciliado.xlsx"
        gerar_excel_saida(resultado, resumo, ponte, caminho_saida, conta=conta, periodo_referencia="")

    st.success("Conciliação concluída.")

    # ---------------------------------------------------------------- KPIs
    total_lancamentos = len(resultado)
    total_conciliado = int((resultado["residual_centavos"] == 0).sum())
    pct_conciliado = total_conciliado / total_lancamentos if total_lancamentos else 0
    saldo_aberto = resultado["residual_centavos"].sum() / 100
    diferenca_ponte = None
    if saldo_balancete is not None:
        diferenca_ponte = round(saldo_balancete - saldo_aberto, 2)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Lançamentos", total_lancamentos)
    c2.metric("% conciliado", f"{pct_conciliado:.0%}", f"{total_conciliado}/{total_lancamentos}")
    c3.metric("Saldo em aberto", f"R$ {saldo_aberto:,.2f}")
    c4.metric(
        "Diferença vs. balancete",
        f"R$ {diferenca_ponte:,.2f}" if diferenca_ponte is not None else "—",
        help="(1) Saldo no balancete − (2) saldo em aberto pós-conciliação. Deveria fechar em 0,00.",
    )

    st.divider()

    # ------------------------------------------------------------- gráficos
    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Saldo em aberto por período")
        fig = px.bar(
            resumo, x="periodo", y="saldo_em_aberto_no_periodo",
            labels={"periodo": "Período", "saldo_em_aberto_no_periodo": "Saldo em aberto (R$)"},
            color="saldo_em_aberto_no_periodo", color_continuous_scale=["#1D9E75", "#D85A30"],
        )
        fig.update_layout(showlegend=False, coloraxis_showscale=False, margin=dict(t=10))
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        st.subheader("Lançamentos por regra aplicada")
        contagem = resultado["status"].value_counts().reset_index()
        contagem.columns = ["status", "quantidade"]
        fig2 = px.bar(
            contagem, x="quantidade", y="status", orientation="h",
            color="status", color_discrete_map=CORES_STATUS,
        )
        fig2.update_layout(showlegend=False, margin=dict(t=10), yaxis_title="", xaxis_title="Lançamentos")
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    # --------------------------------------------------------------- tabelas
    st.subheader("Resumo por período")
    st.dataframe(
        resumo.style.format({
            "total_provisionado": "R$ {:,.2f}", "total_revertido_baixado": "R$ {:,.2f}",
            "saldo_liquido_periodo": "R$ {:,.2f}", "saldo_em_aberto_no_periodo": "R$ {:,.2f}",
            "pct_conciliado": "{:.1f}%",
        }),
        use_container_width=True, hide_index=True,
    )

    abertos = resultado[resultado["residual_centavos"] != 0].sort_values("aging_dias", ascending=False)
    with st.expander(f"Itens em aberto ({len(abertos)})", expanded=len(abertos) > 0):
        if abertos.empty:
            st.success("Nenhum item em aberto — conta 100% conciliada.")
        else:
            colunas = ["id_lancamento", "periodo", "data", "historico", "valor", "valor_residual", "status", "aging_dias", "faixa_aging"]
            st.dataframe(
                abertos[colunas].style.format({"valor": "R$ {:,.2f}", "valor_residual": "R$ {:,.2f}"}),
                use_container_width=True, hide_index=True,
            )

    with st.expander("Detalhe completo de todos os lançamentos"):
        colunas_det = ["id_lancamento", "periodo", "data", "historico", "valor", "status", "regra_aplicada", "contraparte"]
        st.dataframe(
            resultado[colunas_det].sort_values(["periodo", "data"])
            .style.format({"valor": "R$ {:,.2f}"})
            .map(_cor_linha, subset=["status"]),
            use_container_width=True, hide_index=True, height=420,
        )

    st.divider()
    st.download_button(
        "⬇ Baixar Excel conciliado",
        data=caminho_saida.read_bytes(),
        file_name=caminho_saida.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    main()
