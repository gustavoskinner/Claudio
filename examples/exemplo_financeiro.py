"""
Exemplo: Apresentação Financeira - Análise de Investimentos

Tema: Financial
Uso: Reuniões de investimento, análise de portfólio, wealth management.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="financial", titulo="Análise de Portfólio Q1 2026")

    # 1. Capa
    prs.slide_capa(
        titulo="Análise de Portfólio\nQ1 2026",
        subtitulo="Estratégia de Investimentos e Alocação de Ativos",
        autor="Wealth Management Division",
        data="Abril 2026",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "Panorama Macroeconômico",
        "Performance do Portfólio",
        "Alocação de Ativos",
        "Análise de Risco",
        "Recomendações",
    ])

    # 3. Macro
    prs.slide_metricas(
        titulo="Panorama Macroeconômico",
        subtitulo="Indicadores chave do cenário atual",
        metricas=[
            {"valor": "12.25%", "label": "Taxa SELIC", "icone": "🏛️"},
            {"valor": "4.8%", "label": "IPCA 12m", "icone": "📊"},
            {"valor": "R$ 5.12", "label": "USD/BRL", "icone": "💱"},
            {"valor": "+2.3%", "label": "PIB 2026E", "icone": "📈"},
        ],
    )

    # 4. Performance
    prs.slide_metricas(
        titulo="Performance do Portfólio",
        subtitulo="Rentabilidade acumulada no trimestre",
        metricas=[
            {"valor": "+4.8%", "label": "Portfólio Total", "icone": "💰"},
            {"valor": "+3.2%", "label": "CDI (benchmark)", "icone": "📉"},
            {"valor": "+1.6pp", "label": "Alpha Gerado", "icone": "⭐"},
            {"valor": "0.82", "label": "Sharpe Ratio", "icone": "📐"},
        ],
    )

    # 5. Alocação
    prs.slide_dois_colunas(
        titulo="Alocação Atual vs. Recomendada",
        titulo_esq="Alocação Atual",
        conteudo_esq=[
            "Renda Fixa: 45% (R$ 4.5M)",
            "Ações Brasil: 20% (R$ 2.0M)",
            "Ações Global: 15% (R$ 1.5M)",
            "Multimercado: 12% (R$ 1.2M)",
            "Alternativos: 8% (R$ 0.8M)",
        ],
        titulo_dir="Alocação Recomendada",
        conteudo_dir=[
            "Renda Fixa: 40% (-5pp)",
            "Ações Brasil: 22% (+2pp)",
            "Ações Global: 18% (+3pp)",
            "Multimercado: 12% (manter)",
            "Alternativos: 8% (manter)",
        ],
    )

    # 6. Análise de Risco
    prs.slide_conteudo(
        titulo="Análise de Risco do Portfólio",
        subtitulo="Métricas de risco e stress test",
        pontos=[
            "VaR (95%, 1 dia): R$ 82.4K — dentro do limite aprovado",
            "Maximum Drawdown 12m: -6.2% (meta: < -10%)",
            "Correlação média entre classes: 0.35 (diversificação adequada)",
            "Stress Test cenário adverso: -12.5% (recuperação em 8 meses)",
            "Liquidez D+0/D+1: 68% do portfólio (meta: > 60%)",
        ],
    )

    # 7. Recomendações
    prs.slide_comparacao(
        titulo="Cenários e Recomendações",
        opcao_a={
            "titulo": "Cenário Base (60%)",
            "pontos": [
                "SELIC cai para 11.5% até dez/2026",
                "Migrar 5% de pós-fixado para prefixado",
                "Aumentar exposição a ações globais",
                "Manter hedge cambial em 50%",
                "Retorno esperado: CDI + 2.5%",
            ],
        },
        opcao_b={
            "titulo": "Cenário Estresse (25%)",
            "cor": "#C62828",
            "pontos": [
                "SELIC sobe para 13.5%",
                "Aumentar caixa para 15%",
                "Reduzir ações Brasil para 15%",
                "Aumentar hedge para 80%",
                "Proteção via opções de índice",
            ],
        },
    )

    # 8. Citação
    prs.slide_citacao(
        citacao="O mercado pode permanecer irracional por mais tempo do que "
                "você pode permanecer solvente.",
        autor="John Maynard Keynes",
        cargo="Economista",
    )

    # 9. Encerramento
    prs.slide_encerramento(
        mensagem="Obrigado pela Confiança",
        contato="wealth@banco.com.br | (11) 3000-4000",
        subtexto="Este material é meramente informativo e não constitui recomendação de investimento",
    )

    return prs.salvar("exemplo_financeiro.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
