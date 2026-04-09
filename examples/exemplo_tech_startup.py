"""
Exemplo: Pitch Deck de Startup de Tecnologia

Tema: Tech Startup
Uso: Pitch para investidores, demo day, apresentação de produto.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="tech_startup", titulo="NexusAI - Pitch Deck")

    # 1. Capa
    prs.slide_capa(
        titulo="NexusAI",
        subtitulo="Inteligência Artificial que democratiza a análise de dados",
        autor="Maria Silva, CEO & Co-founder",
        data="Abril 2026",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "O Problema",
        "Nossa Solução",
        "Tamanho de Mercado",
        "Modelo de Negócio",
        "Tração e Métricas",
        "Time",
        "Investimento",
    ])

    # 3. O Problema
    prs.slide_divisor(
        titulo_secao="O Problema",
        numero="01",
        descricao="87% das empresas não conseguem extrair valor dos seus dados",
    )

    prs.slide_conteudo(
        titulo="Empresas Afogadas em Dados",
        texto="Empresas coletam terabytes de dados diariamente, mas menos de "
              "15% conseguem transformar esses dados em decisões acionáveis.",
        pontos=[
            "Ferramentas de BI tradicionais exigem meses de implementação",
            "Cientistas de dados custam em média R$ 25K/mês e são escassos",
            "Dashboards estáticos não respondem perguntas em tempo real",
            "70% dos projetos de dados falham por complexidade técnica",
        ],
    )

    # 4. Nossa Solução
    prs.slide_divisor(
        titulo_secao="Nossa Solução",
        numero="02",
        descricao="IA conversacional para análise de dados empresariais",
    )

    prs.slide_dois_colunas(
        titulo="NexusAI: Dados ao Alcance de Todos",
        titulo_esq="Como Funciona",
        conteudo_esq=[
            "Conecte suas fontes de dados em 5 minutos",
            "Faça perguntas em linguagem natural",
            "Receba visualizações e insights automáticos",
            "Compartilhe análises com sua equipe",
        ],
        titulo_dir="Diferenciais",
        conteudo_dir=[
            "Zero código necessário",
            "Suporte a +50 fontes de dados",
            "Respostas em menos de 3 segundos",
            "Segurança enterprise (SOC 2 Type II)",
        ],
    )

    # 5. Mercado
    prs.slide_metricas(
        titulo="Tamanho de Mercado (TAM/SAM/SOM)",
        subtitulo="Mercado global de Business Intelligence",
        metricas=[
            {"valor": "$33B", "label": "TAM Global", "icone": "🌍"},
            {"valor": "$8.5B", "label": "SAM LATAM", "icone": "🎯"},
            {"valor": "$850M", "label": "SOM Brasil", "icone": "🇧🇷"},
        ],
    )

    # 6. Tração
    prs.slide_metricas(
        titulo="Tração e Métricas",
        subtitulo="Últimos 12 meses",
        metricas=[
            {"valor": "320+", "label": "Empresas Ativas", "icone": "🏢"},
            {"valor": "R$ 2.4M", "label": "ARR", "icone": "💰"},
            {"valor": "15%", "label": "Crescimento Mensal", "icone": "🚀"},
            {"valor": "145%", "label": "Net Revenue Retention", "icone": "📊"},
        ],
    )

    # 7. Timeline
    prs.slide_timeline(
        titulo="Roadmap de Produto",
        eventos=[
            {"ano": "2025 Q3", "titulo": "MVP", "descricao": "Lançamento beta"},
            {"ano": "2025 Q4", "titulo": "Product-Market Fit", "descricao": "100 clientes"},
            {"ano": "2026 Q1", "titulo": "Expansão", "descricao": "Enterprise tier"},
            {"ano": "2026 Q3", "titulo": "LATAM", "descricao": "México e Colômbia"},
            {"ano": "2027 Q1", "titulo": "Global", "descricao": "EUA e Europa"},
        ],
    )

    # 8. Modelo de negócio
    prs.slide_comparacao(
        titulo="Planos e Pricing",
        opcao_a={
            "titulo": "Starter (R$ 990/mês)",
            "pontos": [
                "Até 5 usuários",
                "3 fontes de dados",
                "1.000 consultas/mês",
                "Suporte via chat",
                "Dashboards ilimitados",
            ],
        },
        opcao_b={
            "titulo": "Enterprise (R$ 4.990/mês)",
            "pontos": [
                "Usuários ilimitados",
                "Fontes ilimitadas",
                "Consultas ilimitadas",
                "SLA 99.9% + Suporte dedicado",
                "SSO, RBAC e auditoria",
            ],
        },
    )

    # 9. Citação cliente
    prs.slide_citacao(
        citacao="Com a NexusAI, reduzimos o tempo de análise de dias para "
                "segundos. Nosso time de vendas agora toma decisões baseadas "
                "em dados em tempo real.",
        autor="Carlos Mendes",
        cargo="VP de Vendas, TechCorp",
    )

    # 10. Investimento
    prs.slide_conteudo(
        titulo="A Rodada",
        texto="Buscamos R$ 15M em Série A para acelerar crescimento e expansão.",
        pontos=[
            "40% - Produto e Engenharia (IA de próxima geração)",
            "30% - Vendas e Marketing (expansão comercial LATAM)",
            "20% - Operações e Infraestrutura (SOC 2, LGPD)",
            "10% - Reserva estratégica para M&A",
        ],
    )

    # 11. Encerramento
    prs.slide_encerramento(
        mensagem="Vamos Conversar?",
        contato="maria@nexusai.com.br",
        website="www.nexusai.com.br",
        subtexto="NexusAI - Dados que falam a sua língua",
    )

    return prs.salvar("exemplo_tech_startup.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
