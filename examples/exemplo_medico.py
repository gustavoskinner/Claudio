"""
Exemplo: Apresentação Médica/Saúde

Tema: Medical
Uso: Congressos médicos, apresentação de pesquisa clínica, hospitalar.
"""

from claudio import Presentation


def criar_apresentacao():
    prs = Presentation(theme="medical", titulo="Pesquisa Clínica - Cardio")

    # 1. Capa
    prs.slide_capa(
        titulo="Estudo CARDIO-PREV\n2024-2026",
        subtitulo="Resultados do ensaio clínico randomizado multicêntrico",
        autor="Dr. Roberto Lima et al.",
        data="Congresso Nacional de Cardiologia 2026",
    )

    # 2. Agenda
    prs.slide_agenda(topicos=[
        "Contexto e Justificativa",
        "Objetivos do Estudo",
        "Metodologia",
        "Resultados Primários",
        "Resultados Secundários",
        "Conclusões e Implicações",
    ])

    # 3. Contexto
    prs.slide_conteudo(
        titulo="Contexto e Justificativa",
        texto="Doenças cardiovasculares representam a principal causa de "
              "mortalidade no Brasil, com 400 mil óbitos anuais.",
        pontos=[
            "14 milhões de brasileiros vivem com insuficiência cardíaca",
            "Custo anual ao SUS: R$ 22.1 bilhões",
            "Terapias atuais apresentam limitações significativas",
            "Nova classe de inibidores mostra potencial em estudos pré-clínicos",
        ],
    )

    # 4. Métricas
    prs.slide_metricas(
        titulo="Desenho do Estudo",
        subtitulo="Ensaio clínico randomizado, duplo-cego, controlado por placebo",
        metricas=[
            {"valor": "2.400", "label": "Pacientes Recrutados", "icone": "👥"},
            {"valor": "36", "label": "Centros Participantes", "icone": "🏥"},
            {"valor": "24m", "label": "Seguimento Médio", "icone": "📅"},
        ],
    )

    # 5. Resultados
    prs.slide_comparacao(
        titulo="Resultados Primários: Desfecho Composto",
        opcao_a={
            "titulo": "Grupo Intervenção (n=1.200)",
            "pontos": [
                "Mortalidade CV: 8.2% (vs 12.1%)",
                "Internação por IC: 14.5% (vs 21.3%)",
                "Evento composto: 18.7%",
                "HR: 0.72 (IC 95%: 0.61-0.85)",
                "NNT: 12",
            ],
        },
        opcao_b={
            "titulo": "Grupo Placebo (n=1.200)",
            "cor": "#757575",
            "pontos": [
                "Mortalidade CV: 12.1%",
                "Internação por IC: 21.3%",
                "Evento composto: 27.4%",
                "Referência para comparação",
                "p < 0.001",
            ],
        },
    )

    # 6. Conclusões
    prs.slide_conteudo(
        titulo="Conclusões",
        pontos=[
            "Redução significativa de 28% no desfecho composto (p < 0.001)",
            "Perfil de segurança favorável com efeitos adversos leves",
            "Benefício consistente em todos os subgrupos pré-especificados",
            "Resultados sustentados ao longo de 24 meses de seguimento",
            "Potencial para redefinir o tratamento padrão da IC",
        ],
    )

    # 7. Citação
    prs.slide_citacao(
        citacao="A melhor maneira de prever o futuro da medicina é criá-lo "
                "através de evidências robustas e ciência rigorosa.",
        autor="Dr. Roberto Lima",
        cargo="Investigador Principal, CARDIO-PREV",
    )

    # 8. Encerramento
    prs.slide_encerramento(
        mensagem="Obrigado",
        contato="roberto.lima@hospital.edu.br",
        subtexto="Financiamento: CNPq, FAPESP | ClinicalTrials.gov: NCT04XXXXXX",
    )

    return prs.salvar("exemplo_medico.pptx")


if __name__ == "__main__":
    caminho = criar_apresentacao()
    print(f"Apresentação salva em: {caminho}")
