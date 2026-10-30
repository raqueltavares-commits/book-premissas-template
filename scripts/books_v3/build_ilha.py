"""Ilha do Campeche II Spot no layout v3 (template Campeche). Conteudo = book legado (pptx da Rachel) com as
regras do roteiro v3: tudo DEFINIDO, estilo Industrial, trocas de MDF/metal aprovadas pela Raquel em 15/10/2026
(Terrino -> Argento, Savana -> Solanum, metal preto fosco -> metal preto RAL 9005).

python build_ilha.py DADOS saida.pptx   (DADOS = trabalho/v3_dados/ilha)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3"))
from bookv3 import *  # noqa: F401,F403

D, OUT = sys.argv[1], sys.argv[2]
OLD = os.path.join(REPO, "books", "legado", "ILHA_DO_CAMPECHE_SPOT_II_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
sw = lambda n: os.path.join(D, "..", "sw_common", n)
O = lambda k: glob.glob(os.path.join(D, "old", k + "_*"))[0]
REC = lambda n: src(f"REC{n:02d}")
RF = lambda n: src(f"ROOF{n:02d}")
FE = lambda n: src(f"E{n:02d}")

# imagens grandes do legado (marcadores batem com elas)
FRENTE, TERREO = O("s04_11"), O("s08_11")
# amostras do legado (pic id = id do rotulo - 2)
PLACAS, BORDAS, RIPADO, ESQ, PILARES, GCVIDRO = O("s05_12"), O("s05_19"), O("s05_26"), O("s05_33"), O("s05_40"), O("s05_47")
ROOFTOP, CAIXA, RUSTICA, LETREIRO, FLOREIRA, PISCINA = O("s06_12"), O("s06_19"), O("s06_26"), O("s06_33"), O("s06_40"), O("s06_47")
CIMENTO, CONCRETO, CEMENTO, LINHO = O("s17_12"), O("s17_19"), O("s17_33"), O("s17_40")
ARGENTO, SOLANUM = sw("amostra_argento.png"), sw("amostra_solanum.png")
PRETO, GRAFITE = sw("amostra_metal_preto_ral9005.png"), sw("amostra_metal_grafite_ral7024.png")

B = Book(TPL, os.path.join(D, "cache"))

# ================= 1 · CAPA =================
S = B.dup(1)
S.P(16, FE(4), (0.55, 0.5))
S.P(17, FE(2), (0.5, 0.5))
S.P(18, REC(3), (0.4, 0.5))
S.P(19, RF(6), (0.6, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.sh(21).text_frame.paragraphs[0].runs[0].font.size = Pt(8.5)
S.T(22, "Ilha do Campeche II Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(32)
S.T(24, "Lançamento  ×  Pré-executivo  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Industrial")
S.T(25, "Avenida do Campeche · Campeche · Florianópolis/SC", "DOCUMENTO TÉCNICO INTERNO")

# ================= 2 · CONTEXTUALIZAÇÃO =================
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Ilha do Campeche II Spot a partir da "
        "leitura cruzada entre as imagens de lançamento (REV 02), o pré-executivo da AD Arquitetura e Engenharia, o "
        "Briefing de premissas e o projeto de interiores da Seazone: o que a imagem vendeu, o que o projeto traz, onde "
        "compatibilizar e o que a obra deve seguir na compra e na execução.")
S.P(37, FE(7), (0.5, 0.45))
S.P(40, RIPADO)
S.P(43, RF(4), (0.62, 0.5))
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], [
        "O que as imagens de lançamento mostram na fachada?",
        "Qual é o revestimento das paredes, das bordas e dos pilares?",
        "Onde o pré-executivo precisa ser compatibilizado à imagem?",
        "Como são o ripado, os guarda-corpos e as floreiras?",
        "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
        "O que não pode ser alterado sem validação?",
        "O que a obra precisa observar na execução?"]):
    S.T(sid, q)
S.T(62, "Regra do material: nenhuma informação foi inventada. Estilo de interiores: Industrial. Prevalecem a imagem de "
        "lançamento, o conceito Seazone e o Briefing; o pré-executivo e o memorial compatibilizam. Fontes (1.26 - [3851] "
        "Ilha do Campeche II Spot / 02 - Projetos): 08 - Projeto Arquitetônico (imagens 3D REV 02, pré-executivo AD R00, "
        "interiores), 09 - Memorial Descritivo e 00 - Premissas de Projeto. Interiores das unidades, índices "
        "urbanísticos, áreas e custos não fazem parte deste Book.")

# ================= 3 · COMO USAR =================
S = B.dup(3)
S.P(75, RIPADO)
S.T(77, "T03")
S.T(78, "Código", "Código da textura (T03) ou da indicação das Considerações (IND 07) para achar o item no projeto.")
S.P(87, FE(1), (0.5, 0.5), None, 0.8)
S.P(90, REC(4), (0.5, 0.5))
S.P(93, RF(6), (0.6, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas frontal e do térreo · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção e corredores · mini market · rooftop e piscina · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T24 para compra e aprovação de amostras, com MDFs e metais do estilo Industrial")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# ================= 4 · DIVISOR 01 =================
S = B.dup(4)
S.P(101, FE(7), (0.5, 0.45))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# ================= 5 · MOODBOARD FACHADA =================
S = B.dup(5)
S.P(118, FE(1), (0.5, 0.5)); label_w(S, 119, "ESQUINA")
S.P(120, FE(3), (0.5, 0.5)); label_w(S, 121, "FRENTE")
S.P(122, FE(4), (0.55, 0.5)); label_w(S, 123, "E04")
S.P(124, FE(8), (0.55, 0.5)); label_w(S, 125, "E08")
S.P(126, FE(2), (0.5, 0.5)); label_w(S, 127, "TÉRREO · NOITE")
S.P(128, FE(6), (0.5, 0.78)); label_w(S, 129, "PISCINA")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Módulos claros", "paredes em placas claras e bordas de laje com aspecto de pedra e spots embutidos"),
        ("Ripado amadeirado", "brises em alumínio em frente às janelas, entre pilares grafite"),
        ("Esquadrias pretas", "linha 32 gold com perfil preto e guarda-corpo de vidro"),
        ("Térreo", "textura cinza rústica, vitrines pretas, floreira em balanço e letreiro iluminado"),
        ("Coroamento", "rooftop com empenas escuras, floreira e piscina; caixa d'água com logo-casinha")]):
    S.T(sid, a, b)

# ================= 6 · PALETA =================
S = B.dup(6)
for pic, lab, txt, code, a, b, img in [
        (151, 152, 153, "T01", "Placas claras", "paredes dos módulos", PLACAS),
        (154, 155, 156, "T02", "Bordas com aspecto de pedra", "lajes com spots embutidos", BORDAS),
        (157, 158, 159, "T03", "Brise alumínio amadeirado", "frente das janelas", RIPADO),
        (160, 161, 162, "T05", "Pilares grafite", "faixas entre os módulos", PILARES),
        (163, 164, 165, "T04", "Esquadrias pretas", "linha 32 gold, perfil preto", ESQ),
        (166, 167, 168, "T09", "Textura cinza rústica", "pilares e paredes do térreo", RUSTICA)]:
    S.P(pic, img); S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = paredes em placas claras (T01) + bordas de laje com aspecto de pedra e spots (T02) + brise "
                     "em alumínio amadeirado (T03) + pilares grafite (T05) + esquadrias pretas (T04) + térreo em textura cinza "
                     "rústica (T09). Detalhe de cada material nas páginas de texturas T01–T12.")

# ================= 7 · LEITURA GERAL =================
S = B.dup(7)
win = S.P(181, FE(1))
label_w(S, 182, "LANÇAMENTO · ESQUINA")
# marcadores posicionados sobre o render E01 (legenda do slide 4 do legado)
place_markers(S, 181, win, list(range(183, 193)), [
    (0.60, 0.12), (0.36, 0.19), (0.30, 0.69), (0.585, 0.43), (0.645, 0.33),
    (0.343, 0.29), (0.69, 0.63), (0.36, 0.505), (0.20, 0.78), (0.66, 0.85)])
S.T(193, "A fachada é lida em módulos claros: paredes em placas claras e bordas de laje com aspecto de pedra e spots, "
         "separados por pilares grafite, com ripado amadeirado em frente às janelas, esquadrias pretas e guarda-corpo de "
         "vidro. No térreo, textura cinza rústica, vitrines, floreira em balanço e o letreiro iluminado; no topo, rooftop "
         "com empenas escuras e a caixa d'água com o logo-casinha.")
for sid, (a, b) in zip(range(196, 215, 2), [
        ("Caixa d'água com identidade", "logo-casinha e sigla \"AD\" iluminadas"),
        ("Rooftop", "empenas escuras, vidro e floreira"),
        ("Bordas das lajes", "faixa clara com aspecto de pedra e spots"),
        ("Paredes em placas claras", "paginação em placas nas faces dos módulos"),
        ("Ripado vertical amadeirado", "brise em alumínio em frente às janelas"),
        ("Pilares escuros", "faixas grafite entre os módulos"),
        ("Esquadrias pretas", "linha 32 gold, perfil preto"),
        ("Guarda-corpo das sacadas", "vidro com alumínio e montante aparente"),
        ("Floreira em balanço", "jardineira sobre o térreo"),
        ("Letreiro e vitrines", "\"Ilha do Campeche II Spot\" iluminado")]):
    S.T(sid, a, b)

# ================= 8-10 · TEXTURAS DA FACHADA =================
NOTE = ("Amostras extraídas das imagens de lançamento (REV 02): referência visual, não substitui amostra física. "
        "Especificação: Briefing (Premissas de Projeto), Memorial Descritivo e pré-executivo AD R00 (sem legenda de material).")
for title, cards in [
    ("Texturas da fachada · paredes e ripado (1/3)", [
        ("T01", "Placas claras nas paredes", "Faces dos módulos, em todos os pavimentos.",
         "Placas claras com aspecto de pedra/concreto e paginação em placas, como na imagem (memorial: revestimento conforme projeto).",
         "Pré-executivo desenha branco sem legenda: compatibilizar o executivo; aprovar amostra contra a imagem antes da compra.", PLACAS),
        ("T02", "Bordas das lajes", "Faixas horizontais de todos os pavimentos.",
         "Faixa clara com textura de pedra e spots embutidos, como na imagem.",
         "Executivo traz o acabamento no quadro de materiais; Luminotécnico: Projetos entrega antes do fechamento de fachada.", BORDAS),
        ("T03", "Ripado vertical amadeirado", "Frente das janelas, em todos os pavimentos.",
         "Brise em alumínio ripado com acabamento amadeirado, como na imagem.",
         "Memorial cita madeira: executar em alumínio, acabamento amadeirado no tom da imagem; aprovar amostra antes da compra.", RIPADO),
        ("T04", "Esquadrias pretas", "Janelas e portas-janela.",
         "Linha 32 gold, perfil preto (Briefing) · áreas comuns em preto fosco (Briefing).",
         "Converge com a imagem: caixilho preto com vidro incolor.", ESQ)]),
    ("Texturas da fachada · pilares e coroamento (2/3)", [
        ("T05", "Pilares escuros", "Faixas verticais entre os módulos.",
         "Pilares em tom grafite, como na imagem de lançamento.",
         "Pré-executivo mostra branco: compatibilizar o executivo; aprovar amostra contra a imagem antes da compra.", PILARES),
        ("T06", "Guarda-corpo das sacadas", "Sacadas privativas.",
         "Vidro com alumínio anodizado e montante aparente (Briefing), corrimão preto como na imagem.",
         "Memorial cita alumínio com vidro: compatibilizar para o modelo do Briefing antes do pedido.", GCVIDRO),
        ("T07", "Rooftop: empenas e floreira", "Último pavimento.",
         "Empenas escuras triangulares, vidro incolor com perfil \"U\" inox (Briefing) e floreira contínua, como na imagem.",
         "Detalhar no executivo; floreira com impermeabilização e irrigação no hidrossanitário.", ROOFTOP),
        ("T08", "Caixa d'água com identidade", "Volume superior.",
         "Fechamento ripado claro com logo-casinha e sigla \"AD\" iluminadas, como na imagem.",
         "Pré-executivo mostra volume branco: executar como na imagem; Projetos detalha fixação e ponto elétrico.", CAIXA)]),
    ("Texturas da fachada · térreo, letreiro e piscina (3/3)", [
        ("T09", "Textura rústica do térreo", "Pilares e paredes do térreo.",
         "Textura cinza rústica, como na imagem (Briefing: área externa com pintura seguindo o conceito da fachada).",
         "Executivo define o produto no quadro de materiais; aprovar painel-teste contra a imagem.", RUSTICA),
        ("T10", "Letreiro iluminado", "Fachadas do térreo, nas duas faces da esquina.",
         "Letras brancas iluminadas \"Ilha do Campeche II Spot\", como na imagem (memorial: nome do edifício na fachada).",
         "Como na imagem; Projetos detalha fixação e ponto elétrico antes do fechamento de fachada.", LETREIRO),
        ("T11", "Floreira em balanço", "Sobre as vitrines do térreo.",
         "Jardineira com vegetação pendente, como na imagem.",
         "Incluir impermeabilização e irrigação no hidrossanitário; paisagismo detalhado no executivo.", FLOREIRA),
        ("T12", "Piscina do rooftop", "Borda com vista para o mar.",
         "Revestimento esverdeado 20×20 (Briefing), como a pastilha verde da imagem · aquecimento solar (Briefing).",
         "Compatibilizar memorial (porcelanato para piscina); aprovar amostra contra a imagem antes do pedido.", PISCINA)])]:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, c in zip(TEX, cards):
        card(S, ids, *c)

# ================= 11 · VISTA FRONTAL (tpl 12, vertical) =================
S = B.dup(12)
S.T(354, "Fachada frontal")
win = S.P(358, FRENTE)
label_w(S, 359, "LANÇAMENTO · FRENTE")
pts = old_markers(OLD, 4)
place_markers(S, 358, win, list(range(360, 369)), pts[:3] + pts[4:])   # placas claras (4) vai no quadro de amostras
S.T(370, LEG)
vista_items(S, V12, [
    ("Caixa d'água", "Logo-casinha e sigla \"AD\" iluminadas, como na imagem"),
    ("Rooftop", "Empenas escuras, vidro com perfil \"U\" inox e floreira"),
    ("Bordas das lajes", "Faixa clara com aspecto de pedra e spots embutidos"),
    ("Ripado", "Brise em alumínio amadeirado em frente às janelas"),
    ("Pilares", "Tom grafite, como na imagem"),
    ("Esquadrias", "Linha 32 gold, perfil preto (Briefing)"),
    ("Guarda-corpo", "Vidro com alumínio anodizado e montante aparente"),
    ("Floreira", "Jardineira em balanço sobre o térreo"),
    ("Letreiro e vitrines", "Letras brancas iluminadas e caixilho preto")])
S.P(398, PLACAS); S.P(399, BORDAS); S.P(400, PILARES)
S.T(401, "T01 · T02 · T05")
S.T(403, "PRÉ-EXECUTIVO AD R00", "As elevações 009 e 010 desenham o volume todo branco, sem legenda de material. Seguir a "
         "imagem: placas claras, bordas com aspecto de pedra e pilares grafite; compatibilizar o executivo com quadro de materiais.")

# ================= 12 · VISTA TÉRREO (tpl 13) =================
S = B.dup(13)
S.T(410, "Térreo: vitrines, letreiro e floreira")
win = S.P(414, TERREO)
label_w(S, 415, "LANÇAMENTO · TÉRREO")
place_markers(S, 414, win, list(range(416, 426)), old_markers(OLD, 8))
S.T(427, "ATENÇÃO · IDENTIDADE DO TÉRREO", "Letreiro iluminado, vitrines pretas, floreira em balanço e textura rústica são "
         "a \"cara\" do térreo no lançamento: executar como na imagem. Projetos detalha fixação, elétrica e paisagismo antes "
         "do fechamento de fachada.")
S.T(429, LEG)
vista_items(S, V13, [
    ("Letreiro", "Letras brancas iluminadas nas duas fachadas"),
    ("Floreira", "Jardineira em balanço · impermeabilização e irrigação"),
    ("Vitrines", "Caixilho preto fosco, vidro incolor"),
    ("Textura", "Cinza rústica no térreo, como na imagem"),
    ("Spots", "Embutidos sob as lajes · luminotécnico de Projetos"),
    ("Ripado", "Alumínio amadeirado em frente às janelas, como na imagem"),
    ("Guarda-corpo", "Vidro com alumínio anodizado e montante aparente"),
    ("Jardim", "Palmeiras e forração · paisagismo no executivo")])

# ================= 13 · PONTOS DE ATENÇÃO FACHADA (tpl 24) =================
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, FRENTE, (0.5, 0.5))
label_w(S, 773, "LANÇAMENTO · FRENTE")
checklist(S, [
    ("Executivo de arquitetura", "Emitir liberado para obra, com quadro de materiais da fachada conforme este book."),
    ("Placas, bordas e pilares", "Seguir a imagem: placas claras, bordas com aspecto de pedra e pilares grafite; aprovar amostra antes da compra."),
    ("Ripado e brises", "Alumínio com acabamento amadeirado no tom da imagem; compatibilizar o memorial (madeira) e detalhar a fixação."),
    ("Esquadrias e guarda-corpos", "Linha 32 gold, perfil preto; sacadas em vidro com alumínio e montante aparente; rooftop com perfil \"U\" inox."),
    ("Identidade e luz", "Caixa d'água, letreiros e spots das bordas como na imagem: fixação, pontos e luminotécnico antes do fechamento."),
    ("Térreo e paisagismo", "Textura cinza rústica, floreiras impermeabilizadas e irrigadas; muro com pintura cor Seazone, logo e nome (Briefing).")])

# ================= 14 · DIVISOR 02 =================
S = B.dup(14)
S.P(465, REC(4), (0.45, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito Seazone × decisão  ·  recepção e corredores  ·  mini market  ·  rooftop e piscina  ·  pontos de atenção")

# ================= 15 · CONCEITO × DECISÃO (tpl 15) =================
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
concept_cards(S, [
    ("Recepção", REC(3), (0.35, 0.5),
     "Painel ripado amadeirado com letreiro, jardim vertical aramado, textura rústica, espelho orgânico, portas pretas",
     "Painel PVC ripado e jardim aramado (Briefing) · paredes em cimento queimado Campeche"),
    ("Corredores", REC(6), (0.5, 0.5),
     "Painel ripado, jardim vertical e portas pretas",
     "Cores do Briefing: Cinza Urbano + Calopsita (Suvinil)"),
    ("Mini market", RF(2), (0.6, 0.5),
     "Estante de madeira, vending e pendentes de fibra no corredor do rooftop",
     "Pontos da vending e iluminação da estante no elétrico"),
    ("Rooftop", RF(6), (0.6, 0.5),
     "Jardim vertical com letreiro, deck, parede cinza, ducha com painel ripado, piscina esverdeada",
     "Painel PVC ripado + textura cinza (Briefing) · deck em porcelanato madeira 120×20"),
    ("Executivo de interiores", REC(5), (0.4, 0.5),
     "Imagens e texturas da pasta 10 - Projeto de Interiores",
     "Seguir o conceito Seazone até a entrega do executivo")])
S.T(507, "Imagens de interiores Seazone (10 - Projeto de Interiores / 02- Imagens) × Briefing, memorial e texturas do projeto de interiores.")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Executivo de interiores: seguir o conceito Seazone (imagens e texturas da pasta "
         "\"01- Arquivos de projeto / Texturas\") até a entrega do executivo. Letreiros, jardins verticais, vending e pontos do "
         "mini market entram no elétrico e no hidrossanitário antes do fechamento.")

# ================= 16 · FICHA RECEPÇÃO (tpl 27) =================
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO E CORREDORES")
S.T(849, "Ficha: recepção e corredores")
S.P(853, REC(3), (0.4, 0.5)); label_w(S, 854, "CONCEITO · RECEPÇÃO")
S.P(855, REC(4), (0.5, 0.5)); label_w(S, 856, "JARDIM VERTICAL ARAMADO")
S.T(857, "Painel ripado com letreiro, jardim vertical aramado, textura rústica e espelho orgânico; porta preta marcada na IND 07 (Considerações).")
rows(S, ROWS27, [
    ("Paredes", "Textura cimento queimado Campeche (texturas de interiores), visual rústico da imagem, tom cinza médio (Briefing)"),
    ("Painel", "Painel PVC ripado amadeirado com letreiro ILHA DO CAMPECHE II SPOT e frase na parede (Briefing e imagem)"),
    ("Destaque", "Jardim vertical aramado em tela metálica preta RAL 9005 com plantas (Briefing e imagem)"),
    ("Piso", "Porcelanato 90×90 natural/acetinado (Briefing) · compatibilizar a pasta de texturas (Cemento Soft 84×84)"),
    ("Forro", "Gesso com pé-direito livre (Briefing) · compatibilizar a imagem (forro modular)"),
    ("Portas e rodapés", "Pretos nas áreas comuns, porta com bandeira até o forro (Briefing e IND 07)"),
    ("Mobiliário", "Poltrona clara, espelho orgânico com moldura preta e tapete redondo (imagem)"),
    ("Corredores", "Cinza Urbano + Calopsita Suvinil (Briefing) · painel ripado e portas pretas (imagem)"),
    ("Acesso", "Perfil metálico com vidro e reconhecimento facial (Briefing) · esquadrias comuns em preto fosco")], PILLS27)

# ================= 17 · MOODBOARD (tpl 31) =================
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · MOODBOARD")
S.T(978, "Recepção, corredores, mini market e rooftop")
for pic, img, foc in zip(range(982, 993),
                         [REC(3), REC(4), REC(5), REC(6), REC(1), REC(2), RF(1), RF(3), RF(6), RF(7), FE(6)],
                         [(.4, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.6, .5), (.6, .5), (.5, .78)]):
    S.P(pic, img, foc)
S.T(993, "CONCEITO", "Painel PVC ripado amadeirado com letreiro, jardim vertical aramado em tela preta, textura rústica, "
         "espelho orgânico e portas pretas.",
         "ROOFTOP", "Jardim vertical com letreiro, deck em porcelanato madeira, piscina esverdeada, ducha com painel ripado "
         "e mini market com vending.",
         "ESTILO", "Industrial: MDF Berneck Argento e Solanum com metal preto RAL 9005.")
for pic, lab, name, img in [(994, 995, "Cimento", CIMENTO), (996, 997, "Concreto", CONCRETO),
                            (998, 999, "Linho", LINHO), (1000, 1001, "Solanum", SOLANUM)]:
    S.P(pic, img); S.T(lab, name)

# ================= 18 · FICHA MINI MARKET (tpl 30) =================
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · MINI MARKET")
S.T(939, "Ficha de premissas: mini market e corredor do rooftop")
S.P(943, RF(2), (0.5, 0.5)); label_w(S, 944, "CONCEITO · MINI MARKET")
S.P(945, RF(3), (0.5, 0.5)); label_w(S, 946, "VENDING + ESTANTE")
S.T(947, "Mini market no corredor do rooftop: estante de madeira, vending machine e pendentes de fibra.")
rows(S, ROWS30, [
    ("Mini market", "Estante de madeira com produtos e vending machine no corredor (imagem)"),
    ("Marcenaria", "MDF Berneck Solanum e Argento com metal preto RAL 9005 (estilo Industrial)"),
    ("Paredes", "Pintura cinza médio + painel PVC ripado amadeirado + textura cinza (Briefing)"),
    ("Iluminação", "Pendentes de fibra no corredor (imagem)"),
    ("Pontos", "Vending e iluminação da estante: incluir no elétrico antes do fechamento"),
    ("Piso", "Porcelanato 90×90 natural/acetinado (Briefing)"),
    ("Forro", "Gesso com pé-direito livre (Briefing)"),
    ("Portas", "Pretas, com bandeira até o forro (Briefing)"),
    ("Rodapés", "Pretos, como as portas das áreas comuns (Briefing)")], PILLS30)

# ================= 19 · FICHA ROOFTOP (tpl 32) =================
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP")
S.T(1008, "Ficha de premissas: deck, jardim vertical e piscina")
S.P(1012, RF(6), (0.6, 0.5)); label_w(S, 1013, "CONCEITO · DECK + JARDIM")
S.P(1014, FE(6), (0.5, 0.78)); label_w(S, 1015, "LANÇAMENTO · PISCINA")
S.T(1016, "Deck amadeirado com vista para o mar, jardim vertical com letreiro e piscina esverdeada (IND 06 marca o deck).")
rows(S, ROWS32, [
    ("Destaque", "Jardim vertical com letreiro Ilha do Campeche II SPOT iluminado (imagem)"),
    ("Paredes", "Pintura cinza médio + painel PVC ripado + textura cinza (Briefing)"),
    ("Piso externo", "Porcelanato que imita madeira 120×20, rodapé embutido (Briefing e IND 06)"),
    ("Piscina", "Revestimento esverdeado 20×20 (Briefing), como a pastilha verde da imagem"),
    ("Aquecimento", "Solar para a piscina (Briefing)"),
    ("Guarda-corpo", "Vidro incolor com perfil \"U\" inox (Briefing)"),
    ("Ducha", "Ducha com painel ripado na parede (imagem) · incluir no hidrossanitário"),
    ("Jardins", "Paisagismo, irrigação e impermeabilização detalhados no executivo"),
    ("Pontos", "Letreiro do jardim vertical: Projetos detalha fixação e ponto elétrico")], PILLS32)

# ================= 20 · PONTOS DE ATENÇÃO INTERIORES (tpl 24) =================
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, REC(5), (0.25, 0.5))
label_w(S, 773, "CONCEITO · RECEPÇÃO")
checklist(S, [
    ("Executivo de interiores", "Seguir o conceito Seazone (imagens e texturas) até a entrega do executivo de interiores."),
    ("Letreiros e jardins verticais", "Recepção e rooftop: fixação, irrigação, impermeabilização e pontos elétricos antes do fechamento."),
    ("Mini market", "Pontos da vending e iluminação da estante no elétrico antes do fechamento."),
    ("Cores das áreas comuns", "Recepção em cimento queimado Campeche; corredores Cinza Urbano + Calopsita (Briefing); aprovar painel-teste."),
    ("Piso e forro", "Porcelanato 90×90 (Briefing; compatibilizar Cemento Soft 84×84) e forro de gesso com pé-direito livre."),
    ("Piscina e deck", "Esverdeado 20×20 com aquecimento solar; deck em porcelanato que imita madeira 120×20 (Briefing).")])

# ================= 21 · DIVISOR 03 =================
S = B.dup(33)
S.P(1048, RF(5), (0.55, 0.5))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T24: amostras tiradas das texturas e imagens Seazone e padrões do catálogo Industrial para compra e aprovação")

# ================= 22-23 · MATERIAIS =================
NOTE_M = "Texturas da pasta \"10 - Projeto de Interiores / 01- Arquivos de projeto / Texturas\" e recortes das imagens Seazone. Referência visual, não substitui amostra física."
for title, cards in [
    ("Materiais · paredes, piso e tecido (1/3)", [
        ("T13", "Cimento queimado Campeche", "Paredes da recepção e das áreas comuns.",
         "Textura \"Cimento queimado Campeche\" do projeto de interiores, visual rústico da imagem, tom cinza médio (Briefing).",
         "Aprovar painel-teste contra a imagem antes da compra.", CIMENTO),
        ("T14", "Concreto alisado", "Paredes de referência do projeto de interiores.",
         "Textura \"Concreto alisado\" do projeto de interiores (pasta de texturas).",
         "Locação detalhada no executivo de interiores, seguindo as imagens; aprovar painel-teste.", CONCRETO),
        ("T15", "Porcelanato 90×90", "Piso da recepção, corredores e mini market.",
         "Porcelanato 90×90 natural/acetinado (Briefing), tom do Cemento Soft Elizabeth da pasta de texturas.",
         "Compatibilizar a pasta de texturas (84×84) para 90×90; rodapé preto nas áreas comuns.", CEMENTO),
        ("T16", "Linho chumbo", "Estofados das áreas comuns.",
         "Tecido \"Linho chumbo\" do projeto de interiores.",
         "Referência de mobiliário; aprovar amostra de tecido antes da compra.", LINHO)]),
    ("Materiais · painel, jardins e deck (2/3)", [
        ("T17", "Painel PVC ripado", "Recepção (letreiro), corredores, mini market e ducha do rooftop.",
         "Painel PVC ripado amadeirado (Briefing e imagem).",
         "Aprovar amostra de cor contra a imagem; prever fixação e ponto do letreiro.", REC(3), (0.0, 0.75), 0.18),
        ("T18", "Jardim vertical aramado", "Recepção e corredores.",
         "Tela metálica preta RAL 9005 com plantas (Briefing e imagem).",
         "Irrigação, impermeabilização e fixação detalhadas no executivo.", REC(4), (0.1, 0.45), 0.45),
        ("T19", "Jardim vertical do rooftop", "Parede do rooftop, com letreiro Ilha do Campeche II SPOT.",
         "Jardim vertical com letreiro iluminado, como na imagem.",
         "Paisagismo, irrigação e ponto elétrico do letreiro detalhados no executivo.", RF(6), (0.75, 0.3), 0.3),
        ("T20", "Porcelanato madeira 120×20", "Deck do rooftop e borda da piscina.",
         "Porcelanato que imita madeira 120×20 com rodapé embutido (Briefing · IND 06).",
         "Aprovar amostra contra a imagem antes da compra.", RF(8), (0.58, 0.86), 0.12)])]:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title); S.T(1065, NOTE_M)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

# ================= 24 · MDFs E METAIS (tpl 37) =================
S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e metais (3/3)")
S.T(1146, "Padrões Berneck e cores de metal do catálogo Industrial Seazone (rev. 13/10/2026). Referência visual, não substitui amostra física.")
for ids, c in zip(MD, [
        ("T21", "MDF Argento", "Marcenaria das áreas comuns, em composição com o Solanum.",
         "MDF Berneck Argento (catálogo Industrial), cinza claro.",
         "Compatibilizar memorial de interiores (Terrino); aprovar amostra física antes da compra.", ARGENTO),
        ("T22", "MDF Solanum", "Marcenaria das áreas comuns e cabeceiras, com o Argento.",
         "MDF Berneck Solanum (catálogo Industrial), amadeirado.",
         "Compatibilizar memorial de interiores (Savana); aprovar amostra física antes da compra.", SOLANUM),
        ("T23", "Metal preto RAL 9005", "Jardim vertical aramado, moldura dos espelhos e metais da marcenaria.",
         "Metal pintado preto fosco, ref. RAL 9005 (catálogo Industrial).",
         "Compatibilizar memorial de interiores (metal preto fosco); mesmo acabamento em todas as peças metálicas.", PRETO),
        ("T24", "Metal grafite RAL 7024", "Segunda cor de metal do catálogo Industrial, para peças soltas.",
         "Metal pintado grafite, ref. RAL 7024 (catálogo Industrial).",
         "Padrão das áreas comuns é o preto RAL 9005 (T23); grafite só em peça solta, com amostra aprovada junto com os MDFs.", GRAFITE)]):
    card(S, ids, *c)

# ================= 25 · ENCERRAMENTO =================
S = B.dup(38)
S.P(1168, FE(7), (0.5, 0.45))
S.T(1170, "ILHA DO CAMPECHE II SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)
S.T(1172, "Lançamento × Pré-executivo  ·  Seazone")

finish(B, OUT, "ILHA DO CAMPECHE II SPOT")
