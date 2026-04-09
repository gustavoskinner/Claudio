"""
Exemplo: Apresentação Corporativa - Relatório Trimestral

Tema: Corporate Dark
Uso: Reuniões executivas, board meetings, relatórios financeiros.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="corporate_dark", titulo="Relatório Q1 2026")

    # 1. Capa
    prs.slide_capa(
        titulo="Relatório Trimestral\nQ1 2026",
        subtitulo="Resultados e Perspectivas Estratégicas",
        autor="Diretoria Executiva",
        data="Abril 2026",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "Destaques do Trimestre",
        "Resultados Financeiros",
        "Performance por Unidade de Negócio",
        "Indicadores Operacionais",
        "Plano Estratégico Q2",
        "Perguntas e Discussão",
    ])

    # 3. Divisor de seção
    prs.slide_divisor(
        titulo_secao="Destaques do Trimestre",
        numero="01",
        descricao="Principais conquistas e marcos alcançados",
    )

    # 4. Métricas principais
    prs.slide_metricas(
        titulo="Indicadores Chave de Performance",
        subtitulo="Comparativo com o trimestre anterior",
        metricas=[
            {"valor": "R$ 45M", "label": "Receita Bruta", "icone": "📊"},
            {"valor": "+23%", "label": "Crescimento YoY", "icone": "📈"},
            {"valor": "1.2K", "label": "Novos Clientes", "icone": "👥"},
            {"valor": "94%", "label": "Satisfação NPS", "icone": "⭐"},
        ],
    )

    # 5. Conteúdo - Destaques
    prs.slide_conteudo(
        titulo="Principais Conquistas",
        subtitulo="Q1 2026",
        pontos=[
            "Expansão para 3 novos mercados na América Latina",
            "Lançamento da plataforma 2.0 com aumento de 40% na retenção",
            "Parceria estratégica com 2 líderes globais do setor",
            "Redução de 18% no custo de aquisição de clientes (CAC)",
            "Certificação ISO 27001 conquistada em tempo recorde",
        ],
    )

    # 6. Divisor
    prs.slide_divisor(
        titulo_secao="Resultados Financeiros",
        numero="02",
        descricao="Análise detalhada de receita, custos e margens",
    )

    # 7. Duas colunas - Receita vs Custos
    prs.slide_dois_colunas(
        titulo="Receita vs. Custos Operacionais",
        titulo_esq="Receita",
        conteudo_esq=[
            "Receita bruta: R$ 45.2M (+23% YoY)",
            "Receita recorrente (MRR): R$ 3.8M",
            "Ticket médio: R$ 12.5K (+8%)",
            "Churn rate: 2.1% (-0.5pp)",
            "LTV médio: R$ 180K",
        ],
        titulo_dir="Custos",
        conteudo_dir=[
            "COGS: R$ 15.8M (35% da receita)",
            "Marketing & Vendas: R$ 8.2M",
            "P&D: R$ 6.1M (investimento crescente)",
            "G&A: R$ 4.3M (otimizado -12%)",
            "EBITDA: R$ 10.8M (margem 24%)",
        ],
    )

    # 8. Timeline
    prs.slide_timeline(
        titulo="Marcos Estratégicos 2026",
        eventos=[
            {"ano": "Jan", "titulo": "Lançamento v2.0", "descricao": "Nova plataforma"},
            {"ano": "Fev", "titulo": "Expansão LATAM", "descricao": "México e Colômbia"},
            {"ano": "Mar", "titulo": "ISO 27001", "descricao": "Certificação obtida"},
            {"ano": "Abr", "titulo": "Série B", "descricao": "Rodada de investimento"},
            {"ano": "Jun", "titulo": "Meta 100K", "descricao": "100K usuários ativos"},
        ],
    )

    # 9. Comparação
    prs.slide_comparacao(
        titulo="Estratégia Q2: Duas Abordagens",
        opcao_a={
            "titulo": "Crescimento Agressivo",
            "pontos": [
                "Investimento pesado em marketing",
                "Expansão para 5 novos mercados",
                "Contratação de +50 colaboradores",
                "Foco em market share",
                "ROI esperado em 18 meses",
            ],
        },
        opcao_b={
            "titulo": "Crescimento Sustentável",
            "pontos": [
                "Otimização dos canais existentes",
                "Consolidação nos mercados atuais",
                "Contratação seletiva (+15 pessoas)",
                "Foco em rentabilidade",
                "ROI esperado em 8 meses",
            ],
        },
    )

    # 10. Citação
    prs.slide_citacao(
        citacao="O futuro pertence àqueles que enxergam as possibilidades "
                "antes que elas se tornem óbvias.",
        autor="John Sculley",
        cargo="Ex-CEO Apple",
    )

    # 11. Encerramento
    prs.slide_encerramento(
        mensagem="Obrigado!",
        contato="investidores@empresa.com.br",
        website="www.empresa.com.br",
        subtexto="Dúvidas e próximos passos",
    )

    return prs.salvar("exemplo_corporativo.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
