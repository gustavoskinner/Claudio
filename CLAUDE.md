# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Projeto

Claudio é um gerador de apresentações PowerPoint (.pptx) modernas e sofisticadas em Python. Usa `python-pptx` para criar slides com design profissional, suportando 12 temas visuais e múltiplos tipos de slides.

## Comandos

```bash
pip install -r requirements.txt     # instalar dependências
python examples/exemplo_corporativo.py  # gerar exemplo
python -c "from claudio import Presentation; p = Presentation(theme='tech_startup'); p.slide_capa(titulo='Teste'); p.salvar('teste.pptx')"
```

## Arquitetura

- `claudio/themes.py` — Define a dataclass `Theme` e 12 temas pré-configurados (`THEMES` dict). Cada tema contém paleta de cores (hex), tipografia e tamanhos de fonte.
- `claudio/utils.py` — Funções de baixo nível para criar formas (retângulos, círculos, linhas), cards com sombra, barras de progresso, listas com bullet points e manipulação de XML do pptx para transparência.
- `claudio/presentation.py` — Classe `Presentation` que é a API principal. Usa layout blank e cria cada tipo de slide (capa, agenda, conteúdo, métricas, timeline, divisor, comparação, citação, encerramento) programaticamente sobre ele.
- `examples/` — Scripts completos que demonstram uso para diferentes áreas (corporativo, startup, educação, criativo, médico, financeiro).

## Convenções

- Toda a API pública usa nomes em português (ex: `slide_capa`, `salvar`, `titulo`).
- Cores nos temas são armazenadas como strings hex (`#1A73E8`) e convertidas para `RGBColor` via `Theme.rgb()`.
- Slides usam dimensões widescreen 16:9 (13.33" x 7.5").
- Transparência é aplicada via manipulação direta do XML lxml (`p:spPr > a:solidFill > a:srgbClr > a:alpha`).
