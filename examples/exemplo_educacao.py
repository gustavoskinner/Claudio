"""
Exemplo: Apresentação Educacional - Aula Universitária

Tema: Education
Uso: Aulas, workshops, treinamentos, seminários acadêmicos.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="education", titulo="Introdução à IA")

    # 1. Capa
    prs.slide_capa(
        titulo="Inteligência Artificial\ne Machine Learning",
        subtitulo="Fundamentos, Aplicações e o Futuro da Tecnologia",
        autor="Prof. Dr. Ana Costa",
        data="Semestre 2026.1",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "O que é Inteligência Artificial?",
        "História e Evolução",
        "Tipos de Aprendizado de Máquina",
        "Aplicações no Mundo Real",
        "Ética e Desafios",
        "Exercícios Práticos",
    ])

    # 3. Definição
    prs.slide_conteudo(
        titulo="O que é Inteligência Artificial?",
        texto="Inteligência Artificial (IA) é o campo da ciência da computação "
              "dedicado a criar sistemas capazes de realizar tarefas que "
              "normalmente requerem inteligência humana.",
        pontos=[
            "Reconhecimento de padrões em grandes volumes de dados",
            "Tomada de decisão automatizada baseada em evidências",
            "Processamento de linguagem natural (texto e fala)",
            "Visão computacional e interpretação de imagens",
            "Aprendizado contínuo a partir de experiências",
        ],
    )

    # 4. Timeline histórica
    prs.slide_timeline(
        titulo="Evolução da Inteligência Artificial",
        eventos=[
            {"ano": "1950", "titulo": "Teste de Turing", "descricao": "Alan Turing propõe o teste"},
            {"ano": "1997", "titulo": "Deep Blue", "descricao": "IA vence campeão de xadrez"},
            {"ano": "2012", "titulo": "Deep Learning", "descricao": "Revolução em visão computacional"},
            {"ano": "2022", "titulo": "IA Generativa", "descricao": "ChatGPT e difusão"},
            {"ano": "2026", "titulo": "Agentes IA", "descricao": "IA autônoma em produção"},
        ],
    )

    # 5. Tipos de ML
    prs.slide_divisor(
        titulo_secao="Tipos de Aprendizado",
        numero="03",
        descricao="Supervisionado, Não-Supervisionado e por Reforço",
    )

    prs.slide_dois_colunas(
        titulo="Aprendizado Supervisionado vs. Não-Supervisionado",
        titulo_esq="Supervisionado",
        conteudo_esq=[
            "Dados rotulados (com respostas conhecidas)",
            "Classificação: spam vs. não-spam",
            "Regressão: prever preço de imóvel",
            "Exemplos: SVM, Random Forest, Redes Neurais",
        ],
        titulo_dir="Não-Supervisionado",
        conteudo_dir=[
            "Dados sem rótulos (descobre padrões)",
            "Clustering: segmentar clientes",
            "Redução dimensional: PCA, t-SNE",
            "Exemplos: K-Means, DBSCAN, Autoencoders",
        ],
    )

    # 6. Aplicações
    prs.slide_metricas(
        titulo="IA em Números",
        subtitulo="O impacto global da Inteligência Artificial",
        metricas=[
            {"valor": "$184B", "label": "Mercado Global IA", "icone": "💡"},
            {"valor": "77%", "label": "Empresas usando IA", "icone": "🏢"},
            {"valor": "97M", "label": "Empregos criados até 2030", "icone": "👨‍💻"},
            {"valor": "40%", "label": "Ganho de produtividade", "icone": "⚡"},
        ],
    )

    # 7. Comparação
    prs.slide_comparacao(
        titulo="IA Tradicional vs. Deep Learning",
        opcao_a={
            "titulo": "IA Tradicional (ML Clássico)",
            "pontos": [
                "Features engenhadas manualmente",
                "Funciona bem com poucos dados",
                "Mais interpretável e explicável",
                "Menor custo computacional",
                "Bom para dados estruturados",
            ],
        },
        opcao_b={
            "titulo": "Deep Learning",
            "pontos": [
                "Aprende features automaticamente",
                "Precisa de grandes volumes de dados",
                "Resultados superiores em visão e NLP",
                "Requer GPUs/TPUs para treino",
                "Estado da arte em tarefas complexas",
            ],
        },
    )

    # 8. Ética
    prs.slide_conteudo(
        titulo="Ética e Desafios da IA",
        subtitulo="Responsabilidade no desenvolvimento e uso",
        pontos=[
            "Viés algorítmico: modelos podem perpetuar discriminação",
            "Privacidade: uso de dados pessoais para treinamento",
            "Transparência: necessidade de IA explicável (XAI)",
            "Impacto no emprego: automação e requalificação",
            "Regulamentação: LGPD, EU AI Act e marcos legais",
            "Segurança: deepfakes, desinformação e uso malicioso",
        ],
    )

    # 9. Citação
    prs.slide_citacao(
        citacao="A inteligência artificial é a nova eletricidade. "
                "Assim como a eletricidade transformou quase tudo há 100 anos, "
                "a IA vai transformar todas as indústrias.",
        autor="Andrew Ng",
        cargo="Professor Stanford, Co-fundador Coursera",
    )

    # 10. Encerramento
    prs.slide_encerramento(
        mensagem="Dúvidas?",
        contato="prof.ana.costa@universidade.edu.br",
        subtexto="Próxima aula: Redes Neurais Convolucionais na prática",
    )

    return prs.salvar("exemplo_educacao.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
