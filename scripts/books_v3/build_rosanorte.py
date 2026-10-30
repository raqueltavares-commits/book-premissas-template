"""Rosa Norte Spot no layout v3 (template Campeche). Conteudo = decisoes aprovadas na v2 (13/10/2026).
python build_rosanorte.py DADOS saida.pptx   (DADOS = trabalho/v3_dados/rosanorte)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3"))
from bookv3 import *  # noqa: F401,F403

D, OUT = sys.argv[1], sys.argv[2]
OLD = os.path.join(REPO, "books", "legado", "ROSA_NORTE_SPOT_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
sw = lambda n: os.path.join(D, "sw", n)
O = lambda k: glob.glob(os.path.join(D, "old", k + "_*"))[0]
REC = lambda n: src(f"REC{n:02d}")
RF = lambda n: src(f"ROOF{n:02d}")
F1, F2, L1, L4, JD = src("FRONTAL1"), src("FRONTAL2"), src("LATERAL1"), src("LATERAL4"), src("JARDIM")
# amostras do legado
PRATA, NANQUIM, BRISE, FORRO, GRADIL, GC = O("s05_12"), O("s05_19"), O("s05_26"), O("s05_33"), O("s05_40"), O("s05_47")
ESQ, PERG, FLOR, BOTANICA, PEDRA, LETREIRO = O("s06_12"), O("s06_19"), O("s06_26"), O("s06_33"), O("s06_40"), O("s06_47")
RIPADO, LEDCORR, PENDENTE, FORROROOF, PAINELSEU, JARDIM = O("s19_12"), O("s19_19"), O("s19_26"), O("s19_33"), O("s19_40"), O("s19_47")
CINZACRISTAL = O("s20_19")

B = Book(TPL, os.path.join(D, "cache"))

# 1 · CAPA
S = B.dup(1)
S.P(16, F2, (0.5, 0.5))
S.P(17, L1, (0.5, 0.5))
S.P(18, REC(2), (0.5, 0.5))
S.P(19, RF(8), (0.5, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.sh(21).text_frame.paragraphs[0].runs[0].font.size = Pt(8.5)
S.T(22, "Rosa Norte Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(34)
S.T(24, "Lançamento  ×  Executivo  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Industrial")
S.T(25, "Avenida Central da Praia do Rosa · Ibiraquera · Imbituba/SC", "DOCUMENTO TÉCNICO INTERNO")

# 2 · CONTEXTUALIZAÇÃO
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Rosa Norte Spot a partir da leitura cruzada "
        "entre as imagens de lançamento, o projeto executivo da MOMO Arquitetura e o projeto de interiores da Seazone: o que a "
        "imagem vendeu, o que o executivo especifica, onde divergem e como a obra deve executar.")
S.P(37, L1, (0.5, 0.5), 0.6)
S.P(40, BRISE)
S.P(43, F1, (0.5, 0.82), 0.25)
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], [
        "O que as imagens de lançamento mostram na fachada?",
        "Qual cor e revestimento o executivo define?",
        "Onde o executivo diverge da imagem vendida?",
        "O que a conferência pediu para corrigir?",
        "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
        "O que não pode ser alterado sem validação?",
        "O que a obra precisa observar na execução?"]):
    S.T(sid, q)
S.T(62, "Regra do material: nenhuma informação foi inventada. Estilo de interiores: Industrial. Prevalecem a imagem de lançamento, "
        "o conceito Seazone e o Briefing; o executivo compatibiliza. Fontes (1.19 - [3096] Rosa Norte Spot / 02 - Projetos): "
        "08 - Projeto Arquitetônico (imagens 3D, executivo, interiores), 09 - Memorial Descritivo, 00 - Premissas de Projeto e a "
        "Conferência de Projetos. Interiores das unidades, índices urbanísticos, áreas e custos não fazem parte deste Book.")

# 3 · COMO USAR
S = B.dup(3)
S.P(75, FORRO)
S.T(77, "T03")
S.T(78, "Código", "Código da textura/prancha (ex.: T03, prancha 54) para achar o item no executivo e no memorial.")
S.P(87, L1, (0.5, 0.5), None, 0.8)
S.P(90, REC(2), (0.5, 0.5))
S.P(93, RF(8), (0.5, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas frontal, jardim e lateral · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção e corredores · rooftop interno e externo · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T24 para compra e aprovação de amostras, com MDFs e metais do estilo Industrial")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# 4 · DIVISOR 01
S = B.dup(4)
S.P(101, JD, (0.5, 0.45))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# 5 · MOODBOARD
S = B.dup(5)
S.P(118, L1, (0.5, 0.5)); label_w(S, 119, "LATERAL 1 · ESQUINA")
S.P(120, F1, (0.5, 0.5)); label_w(S, 121, "FRONTAL 1")
S.P(122, JD, (0.5, 0.5)); label_w(S, 123, "JARDIM")
S.P(124, src("LATERAL3"), (0.4, 0.5)); label_w(S, 125, "LATERAL 3")
S.P(126, L4, (0.5, 0.5)); label_w(S, 127, "LATERAL 4")
S.P(128, src("LATERAL2"), (0.5, 0.5)); label_w(S, 129, "LATERAL 2")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Volumes Prata", "paredes claras em pintura Prata com textura rolada"),
        ("Vigas Nanquim", "vigas de bordo e lajes escuras marcando os pavimentos"),
        ("Brises amadeirados", "ripas verticais em alumínio amadeirado em frente às janelas"),
        ("Forro de madeira", "PVC cor madeira com spots sob as lajes e marquises"),
        ("Base em pedra", "hall com painel amadeirado e muro em pedra São Tomé")]):
    S.T(sid, a, b)

# 6 · PALETA
S = B.dup(6)
for pic, lab, txt, code, a, b, img in [
        (151, 152, 153, "T01", "Pintura Prata", "paredes, floreiras, caixa d'água", PRATA),
        (154, 155, 156, "T02", "Pintura Nanquim", "vigas de bordo e lajes", NANQUIM),
        (157, 158, 159, "T03", "Brise alumínio amadeirado", "frente de todas as janelas", BRISE),
        (160, 161, 162, "T04", "Forro PVC Madeira Brun 02", "lajes, marquises e rooftop", FORRO),
        (163, 164, 165, "T06", "Vidro + perfil \"U\" inox", "sacadas e ático (Briefing)", "GC"),
        (166, 167, 168, "T11", "Pedra São Tomé branca", "muro da base e escada", PEDRA)]:
    if img == "GC": S.P(pic, JD, (0.30, 0.56), 0.08)
    else: S.P(pic, img)
    S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = volumes claros em pintura Prata + vigas de bordo em Nanquim + brises verticais em alumínio amadeirado + "
                     "forro PVC cor madeira com spots + guarda-corpo de vidro + base em pedra São Tomé e painel amadeirado no hall. "
                     "Detalhe de cada material nas páginas de texturas T01–T12.")

# 7 · LEITURA GERAL
S = B.dup(7)
win = S.P(181, L1)
label_w(S, 182, "LANÇAMENTO · LATERAL 1 (ESQUINA)")
place_markers(S, 181, win, list(range(183, 193)), old_markers(OLD, 4))
S.T(193, "A fachada vendida combina volumes claros (Prata) marcados por vigas de bordo escuras (Nanquim), brises verticais "
         "amadeirados em frente às janelas e forro de madeira com spots sob as lajes. No topo, pergolado escuro e caixa d'água com "
         "o logo Seazone; na base, hall com painel amadeirado e pele de vidro sobre muro de pedra.")
for sid, (a, b) in zip(range(196, 215, 2), [
        ("Caixa d'água com logo", "Seazone iluminado no topo"),
        ("Pergolado do ático", "metálico, pintura Nanquim"),
        ("Guarda-corpo do ático", "vidro com perfil \"U\" inox (Briefing)"),
        ("Floreiras", "vegetação pendente nas bordas"),
        ("Brise vertical amadeirado", "ripas em alumínio em frente às janelas"),
        ("Forro e iluminação", "PVC cor madeira com spots embutidos"),
        ("Vigas de bordo e lajes", "pintura Nanquim + textura rolada"),
        ("Hall de acesso", "painel amadeirado com logo Seazone"),
        ("Pele de vidro", "fachada envidraçada da recepção"),
        ("Muro de pedra", "pedra São Tomé retalho branca")]):
    S.T(sid, a, b)

# 8-10 · TEXTURAS
NOTE = ("Amostras extraídas das imagens de lançamento: referência visual, não substitui amostra física. Especificação: "
        "elevações 39 a 42, detalhamentos 54/55 (EX-ARQ R01), Memorial Descritivo R01 e Conferência de Projetos.")
for title, cards in [
    ("Texturas da fachada · cores, brises e forro (1/3)", [
        ("T01", "Pintura Prata", "Paredes externas, floreiras e caixa d'água.",
         "Pintura acrílica cor Prata + textura rolada Suvinil (elevações e memorial 26.3.3).",
         "Converge com a imagem: painel-teste antes da compra.", PRATA),
        ("T02", "Pintura Nanquim", "Vigas de bordo, lajes aparentes e pergolado.",
         "Pintura acrílica cor Nanquim + textura rolada Suvinil · lajes impermeabilizadas aparentes (elevações).",
         "Mesmo tom no pergolado (pintura eletrostática Nanquim).", NANQUIM),
        ("T03", "Brises verticais", "Frente das janelas das unidades.",
         "Brise em alumínio com acabamento amadeirado, como na imagem (prancha 55 e memorial) · compatibilizar elevações.",
         "Executivo indica alumínio: manter, amadeirado no tom da imagem, em TODAS as janelas (pranchas 40 e 41); aprovar amostra.", BRISE),
        ("T04", "Forro de madeira", "Face inferior das lajes, marquises e rooftop.",
         "Forro PVC cor Madeira Brun 02, Plasbil (elevações, pranchas de forro e memorial 26.3.3).",
         "Aprovar amostra da cor contra a imagem.", FORRO)]),
    ("Texturas da fachada · gradil, vidro e esquadrias (2/3)", [
        ("T05", "Gradil das condensadoras", "Caixas abaixo das janelas das unidades.",
         "Gradil metálico cor próxima da madeira do forro, conforme detalhamento (elevações).",
         "Reposicionar as condensadoras sem tirar brises: brise em todas as janelas (conferência).", GRADIL),
        ("T06", "Guarda-corpo", "Sacadas e ático.",
         "Vidro temperado laminado com perfil \"U\" inox (Briefing).",
         "Executivo indica alumínio preto fosco: compatibilizar antes do pedido.", JD, (0.30, 0.56), 0.08),
        ("T07", "Esquadrias", "Janelas, porta-janelas e hall.",
         "Alumínio pintura eletrostática preto fosco, vidro temperado incolor · linha 32 (gold) (memorial e conferência).",
         "Indicar a linha 32 no quadro de esquadrias antes do pedido.", ESQ),
        ("T08", "Pergolado do ático", "Terraço do rooftop.",
         "Pergolado em aço/alumínio com pintura eletrostática cor Nanquim (prancha 54 e memorial).",
         "Converge com a imagem (pergolado escuro).", PERG)]),
    ("Texturas da fachada · floreiras, hall e base (3/3)", [
        ("T09", "Floreiras", "Bordas das sacadas e do ático.",
         "Floreiras com pintura acrílica Prata e textura rolada · memorial: jardineiras com vegetação.",
         "Seguir as floreiras das imagens: incluir as 3 que faltam na lateral direita (prancha 42).", FLOR),
        ("T10", "Painel do hall + logo", "Fachada do hall de acesso.",
         "Porcelanato Botanica Ambar MA 19,7×120 (Decortiles) com logo Seazone iluminado.",
         "Manter o porcelanato amadeirado, como na imagem; aprovar amostra antes da compra.", BOTANICA),
        ("T11", "Pedra São Tomé", "Muro da base e escada de acesso.",
         "Pedra São Tomé retalho cor branca (elevação frontal) · pingadeiras em granito Andorinha.",
         "Pedra clara como na imagem; recuperar a borda cinza do muro da imagem.", PEDRA),
        ("T12", "Comunicação visual", "Caixa d'água e fachada do hall.",
         "Logo Seazone iluminado na caixa d'água e letreiro ROSA NORTE SPOT, como na imagem.",
         "Incluir logo e letreiro nas quatro elevações, com pontos elétricos.", LETREIRO)])]:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, c in zip(TEX, cards):
        card(S, ids, *c)

# 11 · VISTA FRONTAL (tpl 13)
S = B.dup(13)
S.T(410, "Fachada frontal (Av. Central da Praia do Rosa)")
win = S.P(414, F1)
label_w(S, 415, "LANÇAMENTO · FRONTAL 1")
place_markers(S, 414, win, list(range(416, 426)), old_markers(OLD, 7))
S.T(427, "LEITURA DA FACHADA FRONTAL", "Frente assimétrica: brises amadeirados nas laterais, floreiras ao centro e vigas escuras "
         "marcando os pavimentos. O executivo reproduz cores, forro e pergolado; completar a comunicação visual, brises em alumínio amadeirado em "
         "todas as janelas e manter os revestimentos da base como na imagem.")
S.T(429, LEG)
vista_items(S, V13, [
    ("Caixa d'água", "Prata + textura · logo Seazone iluminado, como na imagem"),
    ("Pergolado", "Metálico, pintura eletrostática Nanquim"),
    ("Guarda-corpo do ático", "Vidro com perfil \"U\" inox (Briefing)"),
    ("Vigas de bordo", "Nanquim + textura rolada · forro PVC madeira"),
    ("Brises", "Alumínio amadeirado em todas as janelas"),
    ("Floreira", "Pintura Prata + textura rolada"),
    ("Sacada", "Guarda-corpo de vidro com perfil \"U\" inox"),
    ("Hall", "Botanica Ambar 19,7×120 + logo iluminado"),
    ("Letreiro", "ROSA NORTE SPOT iluminado, como na imagem"),
    ("Base e escada", "Pedra São Tomé retalho branca")])

# 12 · VISTA JARDIM (tpl 12, vertical)
S = B.dup(12)
S.T(354, "Fachada do jardim e sacadas")
win = S.P(358, JD)
label_w(S, 359, "LANÇAMENTO · JARDIM")
place_markers(S, 358, win, list(range(360, 369)), old_markers(OLD, 8))
S.T(370, LEG)
vista_items(S, V12, [
    ("Guarda-corpo do ático", "Vidro temperado laminado, perfil \"U\" inox (Briefing)"),
    ("Forro do beiral", "PVC cor Madeira Brun 02 (Plasbil) com spots"),
    ("Brises verticais", "Alumínio amadeirado, como na imagem (prancha 55 e memorial)"),
    ("Gradil condensadoras", "Metálico, cor da madeira do forro · sem tirar brises"),
    ("Floreiras", "Pintura Prata + textura rolada · vegetação pendente"),
    ("Pergolado", "Aço/alumínio, pintura eletrostática Nanquim (prancha 54)"),
    ("Paredes", "Pintura acrílica Prata + textura rolada Suvinil"),
    ("Sacadas", "Vidro com perfil \"U\" inox · piso Munari Branco (MOMO)"),
    ("Muro de pedra", "São Tomé retalho branca, como na imagem · borda cinza")])
S.P(398, FORRO); S.P(399, BRISE); S.P(400, PEDRA)
S.T(401, "FORRO · BRISE · PEDRA")
S.T(403, "CONFERÊNCIA", "Brise em TODAS as janelas, inclusive onde as condensadoras foram reposicionadas, e muro com a borda "
         "cinza da imagem (prancha 40).")

# 13 · VISTA LATERAL (tpl 11)
S = B.dup(11)
S.T(302, "Fachada lateral")
win = S.P(306, L4)
label_w(S, 307, "LANÇAMENTO · LATERAL 4")
place_markers(S, 306, win, list(range(308, 317)), old_markers(OLD, 9))
S.T(318, "LEITURA DA FACHADA LATERAL", "A lateral repete o ritmo de janelas com brises verticais e caixas amadeiradas (gradil das "
         "condensadoras) sob as janelas, com floreiras nos vértices. A conferência (pranchas 40 e 41) pede brise em todas as "
         "janelas e comunicação visual.")
S.T(320, LEG)
vista_items(S, V11, [
    ("Caixa d'água", "Logo Seazone iluminado, como na imagem"),
    ("Pergolado", "Metálico Nanquim"),
    ("Guarda-corpo", "Vidro com perfil \"U\" inox"),
    ("Floreiras", "Prata + vegetação na borda do ático"),
    ("Brises", "Alumínio amadeirado em todas as janelas"),
    ("Gradil AC", "Metálico cor da madeira do forro · sem tirar brises"),
    ("Vigas de bordo", "Nanquim + forro PVC madeira"),
    ("Paredes", "Prata + textura rolada"),
    ("Muro", "Claro com vegetação e borda cinza, como na imagem")])

# 14 · PONTOS DE ATENÇÃO FACHADA
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, F2, (0.5, 0.5))
label_w(S, 773, "LANÇAMENTO · FRONTAL 2")
checklist(S, [
    ("Brises em todas as janelas", "Alumínio amadeirado, como na imagem; incluir onde as condensadoras foram reposicionadas (pranchas 40 e 41)."),
    ("Comunicação visual", "Logo Seazone na caixa d'água e letreiro ROSA NORTE SPOT nas quatro elevações, com pontos elétricos."),
    ("Floreiras", "Seguir as floreiras das imagens: incluir as 3 que faltam na lateral direita (prancha 42)."),
    ("Guarda-corpo", "Vidro com perfil \"U\" inox (Briefing): compatibilizar o executivo (alumínio preto) antes do pedido."),
    ("Base e muro", "Pedra São Tomé branca e porcelanato amadeirado no hall, como na imagem; recuperar a borda cinza do muro."),
    ("Esquadrias", "Linha 32 (gold) preto fosco: indicar no quadro de esquadrias antes do pedido.")])

# 15 · DIVISOR 02
S = B.dup(14)
S.P(465, REC(2), (0.5, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito Seazone × decisão  ·  recepção e corredores  ·  rooftop interno e mini market  ·  rooftop externo e piscina  ·  pontos de atenção")

# 16 · CONCEITO × DECISÃO
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
concept_cards(S, [
    ("Recepção", REC(1), (0.5, 0.5),
     "Poltronas creme com estrutura preta + butterfly, pendente cônico de fibra, logo \"SPOT\" com toggle colorido",
     "Seguir a imagem (anteprojeto Studio Hall diverge: bouclé, bambu vazado, logo escuro)"),
    ("Corredores", REC(4), (0.5, 0.5),
     "Perfil de LED seguindo os quadros, portas brancas com numeração preta",
     "Perfil LED e painéis PVC (LOG) · arandelas nas portas"),
    ("Rooftop interno", RF(1), (0.5, 0.5),
     "Forro de madeira com trilhos pretos, gourmet grafite com tampo preto, mini market com vidro liso",
     "Gourmet em MDF Chumbo · mini market com portas de vidro liso"),
    ("Rooftop externo", RF(8), (0.5, 0.5),
     "Pérgola escura + jardim vertical, fire pit, espreguiçadeiras estofadas com almofadas terrosas",
     "Pérgola Nanquim (prancha 54) · piscina esverdeada 20×20 · sem balanço"),
    ("Projeto de interiores", RF(10), (0.5, 0.5),
     "Imagens Seazone (10 - Projeto de Interiores / 00 - Projeto Seazone)",
     "Anteprojeto Studio Hall com casca aderente (Análise de Aderência 13/07/2026)")])
S.T(507, "Análise de Aderência (Seazone, 13/07/2026): anteprojeto Studio Hall × imagens de venda. A \"casca\" está aderente; onde diverge, vale a imagem.")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Vale o conceito das imagens de venda Seazone. Onde o anteprojeto do Studio Hall diverge "
         "(mobiliário, tecidos, logo, portas do mini market, marcenaria da gourmet), seguir a imagem. Marcenaria da gourmet em MDF Chumbo.")

# 17 · FICHA RECEPÇÃO
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO E CORREDORES")
S.T(849, "Ficha: recepção, hall e corredores")
S.P(853, REC(2), (0.5, 0.5)); label_w(S, 854, "CONCEITO SEAZONE · RECEPÇÃO")
S.P(855, REC(4), (0.5, 0.5)); label_w(S, 856, "CONCEITO · CORREDOR")
S.T(857, "Painel ripado de madeira com logo Rosa Norte Spot; corredor com perfil de LED e quadros de ativação.")
rows(S, ROWS27, [
    ("Parede-destaque", "Painel ripado + painel liso em chapa de PVC imita madeira (Briefing), como na imagem, com logo"),
    ("Paredes", "Executivo: Botanica Ambar 19,7×120 + pintura Geada · briefing: cimento queimado areia"),
    ("Piso", "Munari Cimento 90×90 (executivo), tom conferido contra a imagem"),
    ("Rodapé", "Poliestireno branco de sobrepor, quina viva (executivo e briefing)"),
    ("Forro", "Gesso acartonado Branco Neve fosco na recepção · forro modular nos corredores (Briefing)"),
    ("Corredores", "Painéis em PVC imita madeira + perfil de LED seguindo os quadros (LOG) · pintura Outono Gelado (briefing)"),
    ("Portas", "Brancas, marcos de madeira, numeração preta sobre a porta (imagem)"),
    ("Iluminação", "Pendente de fibra + spots pretos · arandelas nas portas, como na imagem (incluir no elétrico)"),
    ("Mobiliário", "Poltronas creme com estrutura preta, poltrona butterfly, tapetes redondos taupe (imagem)")], PILLS27)

# 18 · MOODBOARD ROOFTOP (tpl 31, 11 imagens)
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP · MOODBOARD")
S.T(978, "Rooftop: gourmet, mini market, terraço e piscina")
for pic, img in zip(range(982, 993), [RF(1), RF(2), RF(3), RF(4), RF(5), RF(6), RF(7), RF(8), RF(9), RF(10), src("ROOF_PLANTA")]):
    S.P(pic, img)
S.T(993, "CONCEITO", "Forro de madeira com trilhos de spots pretos, gourmet em marcenaria grafite com tampo preto, mini market com "
         "letreiro preto sobre madeira e painel \"Seu lugar fora de casa\" com frisos de LED.",
         "TERRAÇO", "Piscina com revestimento esverdeado, banco com almofadas, fire pit, poltronas Acapulco pretas e jardim vertical "
         "sob a pérgola Nanquim.",
         "ESTILO", "Industrial: MDF Chumbo na gourmet, metal preto RAL 9005.")
for pic, lab, name, img in [(994, 995, "Forro", FORROROOF), (996, 997, "Painel LED", PAINELSEU),
                            (998, 999, "Jardim", JARDIM), (1000, 1001, "Chumbo", sw("amostra_chumbo.png"))]:
    S.P(pic, img); S.T(lab, name)

# 19 · FICHA ROOFTOP INTERNO
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP INTERNO")
S.T(939, "Ficha de premissas: gourmet interno e mini market")
S.P(943, RF(1), (0.5, 0.5)); label_w(S, 944, "CONCEITO · GOURMET")
S.P(945, RF(3), (0.8, 0.5)); label_w(S, 946, "PAINEL \"SEU LUGAR FORA DE CASA\"")
S.T(947, "Forro de madeira com trilhos de spots pretos; mini market com letreiro preto sobre madeira e prateleiras iluminadas.")
rows(S, ROWS30, [
    ("Forro", "PVC acabamento amadeirado (Briefing e imagem) · compatibilizar o executivo"),
    ("Paredes", "Cimento queimado cinza + painel em PVC imita madeira (briefing) · pintura Outono Gelado"),
    ("Painel", "\"Seu lugar fora de casa\" em madeira com frisos de LED verticais (aderente)"),
    ("Piso", "Porcelanato 90×90 (briefing) · Munari Cimento 90×90 Eliane (executivo)"),
    ("Layout", "Conferência: incluir o Mini Market e retirar uma das pias da planta"),
    ("Mini Market", "Letreiro preto sobre madeira, portas de vidro liso, pontos para geladeira e LED das prateleiras"),
    ("Gourmet", "Tampo preto, frigobar preto, micro-ondas em nicho de madeira · marcenaria em MDF Chumbo"),
    ("Iluminação", "Trilhos de sobrepor pretos (briefing e imagem)"),
    ("Mobiliário", "Mesas pretas, cadeiras creme com estrutura preta, banquetas altas (imagem)")], PILLS30)

# 20 · FICHA ROOFTOP EXTERNO
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP EXTERNO")
S.T(1008, "Ficha de premissas: rooftop externo e piscina")
S.P(1012, RF(8), (0.5, 0.5)); label_w(S, 1013, "CONCEITO · PISCINA")
S.P(1014, RF(5), (0.6, 0.5)); label_w(S, 1015, "PÉRGOLA + JARDIM VERTICAL")
S.T(1016, "Banco fixo com almofadas junto à piscina, fire pit, poltronas Acapulco pretas e jardim vertical sob a pérgola.")
rows(S, ROWS32, [
    ("Piscina", "Revestimento esverdeado 20×20 (Briefing) · porcelanato 20×20 Eliane no tom esverdeado"),
    ("Aquecimento", "Solar (Briefing e memorial 20.8) · incluir no hidrossanitário"),
    ("Piso externo", "Munari Cimento 90×90 Eliane (executivo) · briefing: porcelanato 90×90 antiderrapante"),
    ("Paredes", "Pintura Prata + textura rolada · jardim vertical (imagem e briefing: parede verde)"),
    ("Banco", "Banco em alvenaria junto à piscina (LOG, indicação 02) com tomadas"),
    ("Pérgola", "Metálica escura, pintura eletrostática Nanquim (prancha 54)"),
    ("Ducha", "Ducha de piso com ralo embutido (briefing e LOG, indicação 03)"),
    ("Balanço", "Sem balanço (Briefing: \"não deverá ser previsto\")"),
    ("Mobiliário", "Fire pit, Acapulco pretas, espreguiçadeiras estofadas, almofadas terrosas (imagem)")], PILLS32)

# 21 · PONTOS DE ATENÇÃO INTERIORES
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, RF(10), (0.6, 0.5))
label_w(S, 773, "ROOFTOP · JARDIM E LOGO")
checklist(S, [
    ("Seguir a imagem", "Onde o anteprojeto Studio Hall diverge (mobiliário, tecidos, logo, mini market, gourmet), vale a imagem Seazone."),
    ("Logos iluminados", "Recepção e rooftop: prever ponto elétrico e fixação antes dos painéis."),
    ("Corredores", "Perfil de LED seguindo os quadros e arandelas nas portas: incluir no elétrico."),
    ("Mini Market e gourmet", "Incluir o Mini Market na planta, retirar uma pia; gourmet em MDF Chumbo com tampo preto."),
    ("Piscina", "Esverdeado 20×20 (Briefing), aquecimento solar e ducha de piso no hidrossanitário."),
    ("Rooftop externo", "Sem balanço (Briefing); jardim vertical com sistema e irrigação no executivo.")])

# 22 · DIVISOR 03
S = B.dup(33)
S.P(1048, RF(10), (0.6, 0.5))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T24: amostras tiradas dos renders Seazone e padrões do catálogo Industrial para compra e aprovação")

# 23-25 · MATERIAIS
NOTE_M = "Amostras extraídas dos renders de interiores Seazone. Referência visual, não substitui amostra física."
for title, cards in [
    ("Materiais · recepção e corredores (1/3)", [
        ("T13", "Painel ripado de madeira", "Parede do logo na recepção.",
         "Painel ripado + painel liso de madeira (imagem) · briefing: chapa de PVC imita madeira.",
         "Prever ponto do logo iluminado antes do painel.", RIPADO),
        ("T14", "Perfil de LED + quadros", "Paredes dos corredores.",
         "Perfil de LED embutido seguindo os quadros de ativação (LOG, referência 01).",
         "Incluir no elétrico as arandelas nas portas e os pontos de LED.", LEDCORR),
        ("T15", "Pendente de fibra", "Recepção.",
         "Pendente cônico em fibra natural/juta, trama fechada (imagem).",
         "Seguir a imagem (não usar o bambu vazado do anteprojeto).", PENDENTE),
        ("T16", "Piso Munari Cimento", "Recepção, corredores e rooftop.",
         "Porcelanato Munari Cimento 90×90 Eliane (executivo).",
         "Conferir o tom contra a imagem antes da compra.", RF(1), (0.12, 0.86), 0.12)]),
    ("Materiais · rooftop (2/3)", [
        ("T17", "Forro de madeira", "Gourmet interno do rooftop.",
         "Forro PVC acabamento amadeirado com trilhos de spots pretos (briefing e imagem).",
         "Compatibilizar o forro especificado no executivo com a imagem.", FORROROOF),
        ("T18", "Painel \"Seu lugar fora de casa\"", "Rooftop interno.",
         "Painel de madeira com frisos de LED verticais e frase em preto (imagem e aderência).",
         "Prever ponto de LED embutido no painel.", PAINELSEU),
        ("T19", "Jardim vertical", "Rooftop externo, sob a pérgola.",
         "Jardim vertical (imagem) · briefing: painel em PVC imita madeira + parede verde na área externa.",
         "Sistema, irrigação e fixação detalhados no executivo.", JARDIM),
        ("T20", "Piscina", "Rooftop externo.",
         "Revestimento esverdeado 20×20 (Briefing) · porcelanato 20×20 Eliane no tom esverdeado.",
         "Aprovar amostra do revestimento antes do pedido.", RF(8), (0.8, 0.62), 0.16)])]:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title); S.T(1065, NOTE_M)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e metais (3/3)")
S.T(1146, "Padrões Berneck e cores de metal do catálogo Industrial Seazone (rev. 13/10/2026). Referência visual, não substitui amostra física.")
for ids, c in zip(MD, [
        ("T21", "MDF Carvalho Treviso", "Marcenaria amadeirada (moodboard de interiores Rosa Norte).",
         "MDF Berneck Carvalho Treviso (Design), amadeirado claro.",
         "Compatibilizar moodboard de interiores (Rícori Nativo); aprovar amostra física antes da compra.", sw("batel_amostra_carvalho_treviso.png")),
        ("T22", "MDF Cinza Cristal", "Marcenaria da copa e armários.",
         "MDF Berneck Cinza Cristal (Tx), liso.",
         "Já é do catálogo Industrial; aprovar amostra física antes da compra.", CINZACRISTAL),
        ("T23", "MDF Chumbo", "Marcenaria da gourmet do rooftop (grafite da imagem).",
         "MDF Berneck Chumbo (Micro), liso, com tampo preto.",
         "Não usar a marcenaria branca do anteprojeto Studio Hall; aprovar amostra física antes da compra.", sw("amostra_chumbo.png")),
        ("T24", "Metal preto RAL 9005", "Pergolado, estruturas de mobiliário, arara, estantes e metais.",
         "Metal pintado preto RAL 9005 (catálogo Industrial).",
         "Mesmo acabamento em todas as peças metálicas; aprovar amostra de cor.", sw("batel_amostra_metal_preto.png"))]):
    card(S, ids, *c)

# 26 · ENCERRAMENTO
S = B.dup(38)
S.P(1168, F1, (0.55, 0.5))
S.T(1170, "ROSA NORTE SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)

finish(B, OUT, "ROSA NORTE SPOT")
