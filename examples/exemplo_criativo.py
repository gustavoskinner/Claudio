"""
Exemplo: Apresentação Criativa - Portfólio de Agência de Design

Tema: Creative Bold
Uso: Portfólios, propostas criativas, branding, pitch de agência.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="creative_bold", titulo="BOLD Studio - Portfólio")

    # 1. Capa
    prs.slide_capa(
        titulo="BOLD Studio",
        subtitulo="Design que transforma marcas em experiências memoráveis",
        autor="Equipe Criativa",
        data="2026",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "Quem Somos",
        "Nossos Serviços",
        "Cases de Sucesso",
        "Processo Criativo",
        "Resultados em Números",
        "Próximos Passos",
    ])

    # 3. Sobre a agência
    prs.slide_divisor(
        titulo_secao="Quem Somos",
        descricao="Criatividade com propósito. Design com resultado.",
    )

    prs.slide_conteudo(
        titulo="Uma Agência Diferente",
        texto="Somos um estúdio de design e branding que acredita no poder "
              "da criatividade estratégica para transformar negócios.",
        pontos=[
            "12 anos de mercado e +500 projetos entregues",
            "Equipe multidisciplinar de 35 criativos",
            "Premiações: Cannes Lions, D&AD, ABEDESIGN",
            "Clientes em 8 países da América Latina e Europa",
        ],
    )

    # 4. Serviços
    prs.slide_dois_colunas(
        titulo="Nossos Serviços",
        titulo_esq="Branding & Identidade",
        conteudo_esq=[
            "Construção de marca do zero",
            "Redesign e rebranding estratégico",
            "Manual de identidade visual",
            "Naming e arquitetura de marca",
            "Brand strategy e posicionamento",
        ],
        titulo_dir="Digital & Experiência",
        conteudo_dir=[
            "UX/UI Design para web e mobile",
            "Design de produto digital",
            "Motion design e animação",
            "Social media design",
            "Design de embalagem e PDV",
        ],
    )

    # 5. Métricas
    prs.slide_metricas(
        titulo="Resultados que Falam",
        metricas=[
            {"valor": "500+", "label": "Projetos Entregues", "icone": "🎨"},
            {"valor": "98%", "label": "Clientes Satisfeitos", "icone": "❤️"},
            {"valor": "12x", "label": "ROI Médio do Cliente", "icone": "📈"},
            {"valor": "35", "label": "Prêmios Internacionais", "icone": "🏆"},
        ],
    )

    # 6. Processo
    prs.slide_timeline(
        titulo="Nosso Processo Criativo",
        eventos=[
            {"ano": "Fase 1", "titulo": "Imersão", "descricao": "Pesquisa e entendimento profundo"},
            {"ano": "Fase 2", "titulo": "Ideação", "descricao": "Brainstorm e conceitos"},
            {"ano": "Fase 3", "titulo": "Criação", "descricao": "Design e prototipagem"},
            {"ano": "Fase 4", "titulo": "Validação", "descricao": "Testes e refinamento"},
            {"ano": "Fase 5", "titulo": "Entrega", "descricao": "Implementação e suporte"},
        ],
    )

    # 7. Case de sucesso
    prs.slide_conteudo(
        titulo="Case: Rebranding FreshMarket",
        subtitulo="De supermercado regional a marca amada nacionalmente",
        pontos=[
            "Desafio: marca datada com baixo reconhecimento entre jovens",
            "Solução: nova identidade visual vibrante e posicionamento 'comida de verdade'",
            "Resultado: +340% no engajamento nas redes sociais",
            "Resultado: +28% em vendas nos primeiros 6 meses pós-rebranding",
            "Prêmio ABEDESIGN 2025 — Melhor Rebranding do ano",
        ],
    )

    # 8. Citação
    prs.slide_citacao(
        citacao="Design não é apenas o que parece e como se sente. "
                "Design é como funciona.",
        autor="Steve Jobs",
    )

    # 9. Comparação de pacotes
    prs.slide_comparacao(
        titulo="Nossos Pacotes",
        opcao_a={
            "titulo": "Essencial",
            "pontos": [
                "Identidade visual completa",
                "Manual de marca (30 páginas)",
                "Papelaria básica",
                "3 rodadas de revisão",
                "Entrega em 4 semanas",
            ],
        },
        opcao_b={
            "titulo": "Premium",
            "pontos": [
                "Tudo do Essencial +",
                "UX/UI do site institucional",
                "Social media kit (20 templates)",
                "Motion logo + animações",
                "Suporte por 6 meses",
            ],
        },
    )

    # 10. Encerramento
    prs.slide_encerramento(
        mensagem="Vamos Criar Juntos?",
        contato="hello@boldstudio.com.br",
        website="www.boldstudio.com.br",
        subtexto="BOLD Studio — Criatividade sem limites",
    )

    return prs.salvar("exemplo_criativo.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
