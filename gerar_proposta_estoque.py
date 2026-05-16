"""
Gera a proposta de Estruturação Organizacional com Gestão de Estoques,
replicando exatamente o modelo visual da Secran Gestão e inserindo
4 novos slides para o módulo de Gestão de Estoques (posição 13–16).
"""
import zipfile
import shutil
import copy
import os
import re
from lxml import etree

INPUT = "/root/.claude/uploads/c4fc3ea4-0abf-4098-8f7e-8b5b16aa0365/4700d83a-PROPOSTA_DE_ESTRUTURA__O_ORGANIZACIONAL__CC_COOPERATIVA___1_1.pptx"
OUTPUT = "/home/user/Claudio/PROPOSTA_ESTRUTURACAO_ORGANIZACIONAL_COM_GESTAO_DE_ESTOQUES.pptx"

NS_A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG= "http://schemas.openxmlformats.org/package/2006/content-types"
REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"


# ---------------------------------------------------------------------------
# Conteúdo dos novos slides de Gestão de Estoques
# ---------------------------------------------------------------------------

# Slide 26 – divisor (baseado no slide13.xml)
DIVISOR_TEXT = "GESTÃO DE ESTOQUES COMO MÓDULO DO PROJETO"

# Slide 27 – DIAGNÓSTICO DE ESTOQUE (baseado no slide10.xml)
DIAG_TITLE_PARTS = ["DIAGNÓSTICO ", "DE ESTOQUE"]
DIAG_BANNERS = [
    # (número-no-título, título-do-banner, texto-de-conteúdo)
    ("1", "LEVANTAMENTO DO ESTOQUE",
     "Mapear todos os itens do inventário, quantidades, localização, "
     "estado de conservação e valor financeiro total do estoque atual."),
    ("2", "ANÁLISE DE GIRO E DEMANDA",
     "Identificar produtos de alta e baixa rotatividade para otimizar "
     "reposição e evitar obsolescência e capital imobilizado."),
    ("3", "DIAGNÓSTICO DE PERDAS",
     "Avaliar índices de desperdício, desvios, validade vencida e rupturas "
     "que impactam diretamente o resultado financeiro da empresa."),
    ("4", "AVALIAÇÃO DOS CONTROLES",
     "Analisar sistemas, planilhas e processos de controle de entradas "
     "e saídas de materiais, identificando falhas nos registros."),
]

# Slide 28 – ESTRUTURAÇÃO DE ESTOQUE (baseado no slide11.xml)
ESTRU_TITLE = "ESTRUTURAÇÃO DA GESTÃO DE ESTOQUES"
ESTRU_BANNERS = [
    ("1", "POLÍTICA DE ESTOQUES",
     "Definir níveis mínimos, máximos e ponto de pedido para cada "
     "categoria de produto, garantindo abastecimento sem excessos."),
    ("2", "CLASSIFICAÇÃO ABC",
     "Categorizar itens por valor e frequência de uso, priorizando "
     "os de maior impacto financeiro no negócio."),
    ("3", "PROCESSOS DE ENTRADA E SAÍDA",
     "Estruturar fluxos de recebimento, conferência, armazenamento "
     "e expedição de forma padronizada e eficiente."),
    ("4", "INDICADORES DE ESTOQUE (KPIs)",
     "Criar manual de KPIs: giro, cobertura, acuracidade, ruptura "
     "e custo de manutenção de estoque por categoria."),
    ("5", "LAYOUT E ARMAZENAGEM",
     "Definir organização física do almoxarifado para maximizar "
     "eficiência, segurança e rastreabilidade dos itens."),
    ("6", "MODELO DE REPOSIÇÃO",
     "Implementar sistema de reposição baseado em demanda histórica, "
     "sazonalidade e ponto de pedido definido pela política."),
]

# Slide 29 – IMPLANTAÇÃO DE ESTOQUE (baseado no slide12.xml)
IMPL_TITLE = "IMPLANTAÇÃO DA GESTÃO DE ESTOQUES"
IMPL_BANNERS = [
    ("1", "TREINAMENTO DA EQUIPE",
     "Capacitar colaboradores nos processos, ferramentas e rotinas "
     "de controle de estoque definidos na fase de estruturação."),
    ("2", "IMPLEMENTAÇÃO DO SISTEMA",
     "Configurar e ativar o sistema de gestão (WMS/ERP) com os "
     "parâmetros, categorias e fluxos de estoque definidos."),
    ("3", "AUDITORIAS DE INVENTÁRIO",
     "Realizar contagens cíclicas e inventários periódicos para "
     "garantir acuracidade dos registros e confiabilidade dos dados."),
    ("4", "RELATÓRIOS E MONITORAMENTO",
     "Implantar dashboard de indicadores de estoque para "
     "acompanhamento em tempo real pela gestão da empresa."),
]


# ---------------------------------------------------------------------------
# Helpers para manipular XML de texto
# ---------------------------------------------------------------------------

def get_sp_texts(slide_root):
    """Retorna lista de (sp_elem, sp_name, [texto_t_elements]) para sps com texto."""
    result = []
    for sp in slide_root.iter(f'{{{NS_P}}}sp'):
        cnvpr = sp.find(f'.//{{{NS_P}}}cNvPr')
        name = cnvpr.get('name', '?') if cnvpr is not None else '?'
        t_elems = [t for t in sp.findall(f'.//{{{NS_A}}}t')
                   if t.text and t.text.strip()]
        if t_elems:
            result.append((sp, name, t_elems))
    return result


def replace_sp_text(sp, new_texts):
    """
    Substitui o conteúdo textual de um sp.
    new_texts pode ser str ou lista de str (cada item = parágrafo).
    Preserva a formatação do primeiro run do primeiro parágrafo.
    """
    txBody = sp.find(f'.//{{{NS_P}}}txBody')
    if txBody is None:
        txBody = sp.find(f'.//{{{NS_A}}}txBody')
    if txBody is None:
        return

    if isinstance(new_texts, str):
        new_texts = [new_texts]

    # Coletar parágrafos existentes para reutilizar formatação
    existing_paras = txBody.findall(f'{{{NS_A}}}p')

    # Obter o run template (para preservar fonte/cor/tamanho)
    run_template = None
    for p in existing_paras:
        for r in p.findall(f'{{{NS_A}}}r'):
            run_template = r
            break
        if run_template is not None:
            break

    # Remover parágrafos existentes (mas manter bodyPr e lstStyle se existirem)
    for p in list(existing_paras):
        txBody.remove(p)

    for text in new_texts:
        new_p = etree.SubElement(txBody, f'{{{NS_A}}}p')
        if run_template is not None:
            new_r = copy.deepcopy(run_template)
            t_elem = new_r.find(f'{{{NS_A}}}t')
            if t_elem is None:
                t_elem = etree.SubElement(new_r, f'{{{NS_A}}}t')
            t_elem.text = text
            new_p.append(new_r)
        else:
            new_r = etree.SubElement(new_p, f'{{{NS_A}}}r')
            t_elem = etree.SubElement(new_r, f'{{{NS_A}}}t')
            t_elem.text = text


def set_single_t_text(t_elems, text):
    """Define o texto de uma lista de elementos t, consolidando em um só."""
    if not t_elems:
        return
    t_elems[0].text = text
    for t in t_elems[1:]:
        t.text = ""


# ---------------------------------------------------------------------------
# Modificar slide do DIVISOR (slide13)
# ---------------------------------------------------------------------------

def modify_divisor(xml_bytes, new_text):
    root = etree.fromstring(xml_bytes)
    for sp, name, t_elems in get_sp_texts(root):
        if name == "CaixaDeTexto 43":
            set_single_t_text(t_elems, new_text)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------------------
# Modificar slide do DIAGNÓSTICO (slide10)
# ---------------------------------------------------------------------------

def modify_diagnostico(xml_bytes, title_parts, banners):
    """
    title_parts: ['DIAGNÓSTICO', 'DE ESTOQUE']
    banners: list of (num_str, titulo_str, conteudo_str)  — 4 itens
    """
    root = etree.fromstring(xml_bytes)
    sps = get_sp_texts(root)

    # Título principal (CaixaDeTexto 57) – pode ter 2 runs/parágrafos
    for sp, name, t_elems in sps:
        if name == "CaixaDeTexto 57":
            if len(t_elems) >= 2:
                t_elems[0].text = title_parts[0]
                t_elems[1].text = title_parts[1]
            elif len(t_elems) == 1:
                t_elems[0].text = " ".join(title_parts)
            break

    # Banner titles (nos grupos grpSp)
    # CaixaDeTexto 12 → banner1, 22 → banner2, 46 → banner3, 54 → banner4
    title_map = {"CaixaDeTexto 12": 0, "CaixaDeTexto 22": 1,
                 "CaixaDeTexto 46": 2, "CaixaDeTexto 54": 3}
    for sp, name, t_elems in sps:
        if name in title_map:
            idx = title_map[name]
            set_single_t_text(t_elems, banners[idx][1])

    # Banner numbers (CaixaDeTexto 11, 21, 45, 53)
    num_map = {"CaixaDeTexto 11": 0, "CaixaDeTexto 21": 1,
               "CaixaDeTexto 45": 2, "CaixaDeTexto 53": 3}
    for sp, name, t_elems in sps:
        if name in num_map:
            idx = num_map[name]
            set_single_t_text(t_elems, banners[idx][0])

    # Conteúdo textual: a ordem no spTree é banner4, banner3, banner1, banner2
    # Identificados pelos nomes texto1 / texto2 em sequência
    content_order = [3, 2, 0, 1]  # índice do banner
    content_idx = 0
    for sp, name, t_elems in sps:
        if name in ("texto1", "texto2") and content_idx < 4:
            idx = content_order[content_idx]
            replace_sp_text(sp, banners[idx][2])
            content_idx += 1

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------------------
# Modificar slide da ESTRUTURAÇÃO (slide11) – 6 banners
# ---------------------------------------------------------------------------

def modify_estruturacao(xml_bytes, title, banners):
    root = etree.fromstring(xml_bytes)
    sps = get_sp_texts(root)

    # Título
    for sp, name, t_elems in sps:
        if name == "CaixaDeTexto 57":
            set_single_t_text(t_elems, title)
            break

    # Banner titles
    title_map = {
        "CaixaDeTexto 12": 0, "CaixaDeTexto 22": 1,
        "CaixaDeTexto 30": 2, "CaixaDeTexto 38": 3,
        "CaixaDeTexto 46": 4, "CaixaDeTexto 54": 5,
    }
    for sp, name, t_elems in sps:
        if name in title_map:
            set_single_t_text(t_elems, banners[title_map[name]][1])

    # Banner numbers
    num_map = {
        "CaixaDeTexto 11": 0, "CaixaDeTexto 21": 1,
        "CaixaDeTexto 29": 2, "CaixaDeTexto 37": 3,
        "CaixaDeTexto 45": 4, "CaixaDeTexto 53": 5,
    }
    for sp, name, t_elems in sps:
        if name in num_map:
            set_single_t_text(t_elems, banners[num_map[name]][0])

    # Conteúdo: ordem no spTree = banner6, banner5, banner4, banner3, banner1, banner2
    content_order = [5, 4, 3, 2, 0, 1]
    content_idx = 0
    for sp, name, t_elems in sps:
        if name in ("texto1", "texto2") and content_idx < 6:
            idx = content_order[content_idx]
            replace_sp_text(sp, banners[idx][2])
            content_idx += 1

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------------------
# Modificar slide da IMPLANTAÇÃO (slide12) – 4 banners
# ---------------------------------------------------------------------------

def modify_implantacao(xml_bytes, title, banners):
    root = etree.fromstring(xml_bytes)
    sps = get_sp_texts(root)

    # Título
    for sp, name, t_elems in sps:
        if name == "CaixaDeTexto 57":
            set_single_t_text(t_elems, title)
            break

    # Banner titles
    title_map = {
        "CaixaDeTexto 12": 0, "CaixaDeTexto 22": 1,
        "CaixaDeTexto 30": 2, "CaixaDeTexto 38": 3,
    }
    for sp, name, t_elems in sps:
        if name in title_map:
            set_single_t_text(t_elems, banners[title_map[name]][1])

    # Banner numbers
    num_map = {
        "CaixaDeTexto 11": 0, "CaixaDeTexto 21": 1,
        "CaixaDeTexto 29": 2, "CaixaDeTexto 37": 3,
    }
    for sp, name, t_elems in sps:
        if name in num_map:
            set_single_t_text(t_elems, banners[num_map[name]][0])

    # Conteúdo: ordem no spTree = banner4, banner3, banner1, banner2
    content_order = [3, 2, 0, 1]
    content_idx = 0
    for sp, name, t_elems in sps:
        if name in ("texto1", "texto2") and content_idx < 4:
            idx = content_order[content_idx]
            replace_sp_text(sp, banners[idx][2])
            content_idx += 1

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ---------------------------------------------------------------------------
# Criar arquivo .rels para novos slides
# ---------------------------------------------------------------------------

def make_slide_rels(layout_target):
    """Gera bytes do arquivo .rels de um slide com apenas a referência ao layout."""
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        f'Target="{layout_target}"/>'
        '</Relationships>'
    )
    return xml.encode('utf-8')


# ---------------------------------------------------------------------------
# Montagem do ZIP
# ---------------------------------------------------------------------------

def build_presentation():
    shutil.copy(INPUT, OUTPUT)

    # Ler todos os arquivos do ZIP original
    with zipfile.ZipFile(INPUT, 'r') as z_in:
        original_files = {name: z_in.read(name) for name in z_in.namelist()}

    # Templates de slides
    slide10_xml = original_files['ppt/slides/slide10.xml']
    slide11_xml = original_files['ppt/slides/slide11.xml']
    slide12_xml = original_files['ppt/slides/slide12.xml']
    slide13_xml = original_files['ppt/slides/slide13.xml']

    # Gerar XMLs dos novos slides
    new_slide26 = modify_divisor(slide13_xml, DIVISOR_TEXT)
    new_slide27 = modify_diagnostico(slide10_xml, DIAG_TITLE_PARTS, DIAG_BANNERS)
    new_slide28 = modify_estruturacao(slide11_xml, ESTRU_TITLE, ESTRU_BANNERS)
    new_slide29 = modify_implantacao(slide12_xml, IMPL_TITLE, IMPL_BANNERS)

    # Rels dos novos slides (layout = mesmo do slide10/11/12 para 27/28/29, slide13 para 26)
    rels_13_layout = '../slideLayouts/slideLayout187.xml'
    rels_10_layout = '../slideLayouts/slideLayout260.xml'

    new_rels26 = make_slide_rels(rels_13_layout)
    new_rels27 = make_slide_rels(rels_10_layout)
    new_rels28 = make_slide_rels(rels_10_layout)
    new_rels29 = make_slide_rels(rels_10_layout)

    # -----------------------------------------------------------------------
    # Atualizar presentation.xml – inserir novos slides depois de rId19 (slide12)
    # -----------------------------------------------------------------------
    prs_xml = original_files['ppt/presentation.xml']
    prs_root = etree.fromstring(prs_xml)
    ns = {
        'p': NS_P,
        'r': NS_R,
    }
    etree.register_namespace('p', NS_P)
    etree.register_namespace('r', NS_R)

    sldIdLst = prs_root.find(f'{{{NS_P}}}sldIdLst')

    # Encontrar a posição do slide12 (rId19) para inserir depois
    insert_after_idx = None
    for i, sld in enumerate(sldIdLst):
        r_id = sld.get(f'{{{NS_R}}}id')
        if r_id == 'rId19':
            insert_after_idx = i
            break

    if insert_after_idx is None:
        raise RuntimeError("Não encontrou rId19 no sldIdLst")

    # IDs únicos para os novos slides
    new_slide_entries = [
        ('rId33', '300001', 'slide26.xml'),
        ('rId34', '300002', 'slide27.xml'),
        ('rId35', '300003', 'slide28.xml'),
        ('rId36', '300004', 'slide29.xml'),
    ]

    for offset, (rid, sld_id, _) in enumerate(new_slide_entries):
        new_sld = etree.Element(f'{{{NS_P}}}sldId')
        new_sld.set('id', sld_id)
        new_sld.set(f'{{{NS_R}}}id', rid)
        sldIdLst.insert(insert_after_idx + 1 + offset, new_sld)

    new_prs_xml = etree.tostring(prs_root, xml_declaration=True,
                                  encoding='UTF-8', standalone=True)

    # -----------------------------------------------------------------------
    # Atualizar ppt/_rels/presentation.xml.rels
    # -----------------------------------------------------------------------
    prs_rels_xml = original_files['ppt/_rels/presentation.xml.rels']
    prs_rels_root = etree.fromstring(prs_rels_xml)
    REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

    for rid, _, fname in new_slide_entries:
        new_rel = etree.SubElement(prs_rels_root, 'Relationship')
        new_rel.set('Id', rid)
        new_rel.set('Type', REL_SLIDE)
        new_rel.set('Target', f'slides/{fname}')

    new_prs_rels_xml = etree.tostring(prs_rels_root, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)

    # -----------------------------------------------------------------------
    # Atualizar [Content_Types].xml
    # -----------------------------------------------------------------------
    ct_xml = original_files['[Content_Types].xml']
    ct_root = etree.fromstring(ct_xml)
    SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

    for _, _, fname in new_slide_entries:
        ov = etree.SubElement(ct_root, f'{{{NS_PKG}}}Override')
        ov.set('PartName', f'/ppt/slides/{fname}')
        ov.set('ContentType', SLIDE_CT)

    new_ct_xml = etree.tostring(ct_root, xml_declaration=True,
                                 encoding='UTF-8', standalone=True)

    # -----------------------------------------------------------------------
    # Escrever novo ZIP
    # -----------------------------------------------------------------------
    new_files = dict(original_files)
    new_files['ppt/presentation.xml'] = new_prs_xml
    new_files['ppt/_rels/presentation.xml.rels'] = new_prs_rels_xml
    new_files['[Content_Types].xml'] = new_ct_xml
    new_files['ppt/slides/slide26.xml'] = new_slide26
    new_files['ppt/slides/slide27.xml'] = new_slide27
    new_files['ppt/slides/slide28.xml'] = new_slide28
    new_files['ppt/slides/slide29.xml'] = new_slide29
    new_files['ppt/slides/_rels/slide26.xml.rels'] = new_rels26
    new_files['ppt/slides/_rels/slide27.xml.rels'] = new_rels27
    new_files['ppt/slides/_rels/slide28.xml.rels'] = new_rels28
    new_files['ppt/slides/_rels/slide29.xml.rels'] = new_rels29

    with zipfile.ZipFile(OUTPUT, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for name, data in new_files.items():
            z_out.writestr(name, data)

    print(f"✓ Arquivo gerado: {OUTPUT}")
    print(f"  Total de slides: 25 originais + 4 novos (Gestão de Estoques) = 29 slides")


if __name__ == '__main__':
    build_presentation()
