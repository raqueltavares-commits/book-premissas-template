"""Imbassai Spot no layout v3 (template Campeche). Conteudo = decisoes aprovadas na v2 (13/10/2026).
python build_imbassai.py DADOS saida.pptx   (DADOS = trabalho/v3_dados/imbassai)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3"))
from bookv3 import *  # noqa: F401,F403

D, OUT = sys.argv[1], sys.argv[2]
OLD = os.path.join(REPO, "books", "legado", "IMBASSAI_SPOT_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
sw = lambda n: os.path.join(D, "sw", n)
O = lambda k: glob.glob(os.path.join(D, "old", k + "_*"))[0]
REC = lambda n: src(f"REC{n:02d}")
LZ = lambda n: src(f"LAZ{n:02d}")
FR, PI, LA, FU = src("FRONTAL"), src("PISCINA"), src("LATERAL"), src("FUNDO")
# amostras do legado
PILAR, BRISE, MUXA, JARDINEIRA, PEDRA, TELHADO = O("s05_12"), O("s05_19"), O("s05_26"), O("s05_33"), O("s05_40"), O("s05_47")
PAREDE, LETREIRO, RIPADO, PERG, DECK, CERCA = O("s06_12"), O("s06_19"), O("s06_26"), O("s06_33"), O("s06_40"), O("s06_47")
PEDRAINT, MUXAINT, PENDENTE, PEDRABRANCA, JARDIMV, FORROVIGAS = O("s17_12"), O("s17_19"), O("s17_26"), O("s17_33"), O("s17_40"), O("s17_47")
BANCADA, PASTILHA = O("s18_33"), O("s18_40")

B = Book(TPL, os.path.join(D, "cache"))

# 1 · CAPA
S = B.dup(1)
S.P(16, FR, (0.35, 0.5))
S.P(17, PI, (0.5, 0.5))
S.P(18, REC(2), (0.5, 0.5))
S.P(19, LZ(13), (0.5, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.sh(21).text_frame.paragraphs[0].runs[0].font.size = Pt(8.5)
S.T(22, "Imbassaí Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(36)
S.T(24, "Lançamento  ×  Projeto  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Clean")
S.T(25, "Alameda das Amendoeiras · Praia de Imbassaí · Mata de São João/BA", "DOCUMENTO TÉCNICO INTERNO")

# 2 · CONTEXTUALIZAÇÃO
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Imbassaí Spot a partir da leitura cruzada "
        "entre as imagens de lançamento, o projeto legal da Beth Câmara Arquitetura e o projeto de interiores da Seazone: o que "
        "a imagem vendeu, o que o projeto especifica, onde divergem e como a obra deve executar.")
S.P(37, FR, (0.3, 0.45), 0.6)
S.P(40, MUXA)
S.P(43, PI, (0.4, 0.75), 0.3)
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], [
        "O que as imagens de lançamento mostram na fachada?",
        "O que o projeto legal especifica para cada elemento?",
        "Onde o projeto diverge da imagem vendida?",
        "Como são os brises, o muxarabi e a madeira?",
        "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
        "O que não pode ser alterado sem validação?",
        "O que a obra precisa observar na execução?"]):
    S.T(sid, q)
S.T(62, "Regra do material: nenhuma informação foi inventada. Estilo de interiores: Clean. O projeto executivo ainda não foi "
        "emitido. Prevalecem a imagem de lançamento, o conceito Seazone e o Briefing; o projeto compatibiliza. Fontes (1.12 - "
        "[1797] Imbassaí Spot / 02 - Projetos): 08 - Projeto Arquitetônico (imagens 3D, projeto legal, interiores), 09 - "
        "Memorial Descritivo e 00 - Premissas de Projeto. Interiores das unidades, índices urbanísticos, áreas e custos não "
        "fazem parte deste Book.")

# 3 · COMO USAR
S = B.dup(3)
S.P(75, BRISE)
S.T(77, "T02")
S.T(78, "Código", "Código da textura (ex.: T02) para achar o item no projeto legal, no memorial e no Briefing.")
S.P(87, FR, (0.4, 0.5), None, 0.8)
S.P(90, REC(2), (0.5, 0.5))
S.P(93, LZ(13), (0.5, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas da piscina e lateral · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção · área de lazer e piscina · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T24 para compra e aprovação de amostras, com MDFs do estilo Clean")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# 4 · DIVISOR 01
S = B.dup(4)
S.P(101, FR, (0.3, 0.5))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# 5 · MOODBOARD
S = B.dup(5)
S.P(118, FR, (0.5, 0.5)); label_w(S, 119, "FACHADA FRONTAL")
S.P(120, PI, (0.5, 0.5)); label_w(S, 121, "FACHADA PISCINA")
S.P(122, LA, (0.3, 0.5)); label_w(S, 123, "LATERAL")
S.P(124, FU, (0.75, 0.5)); label_w(S, 125, "FUNDO")
S.P(126, LA, (0.75, 0.5)); label_w(S, 127, "LATERAL DIREITA")
S.P(128, src("ENTORNO"), (0.5, 0.5)); label_w(S, 129, "ENTORNO")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Madeira em toda a fachada", "pilares de eucalipto e painéis/brises amadeirados nas janelas"),
        ("Cobertura com beiral", "telha metálica escura com viga metálica no beiral"),
        ("Jardineiras", "faixa de vegetação pendente entre os pavimentos"),
        ("Acesso em muxarabi", "volume central do hall em madeira"),
        ("Pedra", "empena da esquina com logo Seazone e mureta na calçada")]):
    S.T(sid, a, b)

# 6 · PALETA
S = B.dup(6)
for pic, lab, txt, code, a, b, img in [
        (151, 152, 153, "T01", "Eucalipto roliço", "pilares e pergolado", PILAR),
        (154, 155, 156, "T02", "Brise em madeira", "amadeirado, todas as janelas", BRISE),
        (157, 158, 159, "T03", "Muxarabi de madeira", "volume do acesso", MUXA),
        (160, 161, 162, "T07", "Papel de Presente B208", "paredes claras (Suvinil)", PAREDE),
        (163, 164, 165, "T05", "Pedra natural", "empena da esquina + logo", PEDRA),
        (166, 167, 168, "T06", "Telha metálica escura", "cobertura com beiral", TELHADO)]:
    S.P(pic, img); S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = pilares de eucalipto + painéis/brises em madeira amadeirada em todas as janelas + paredes claras em "
                     "Papel de Presente Suvinil (B208) + jardineiras revestidas em madeira + cobertura escura com beiral + muxarabi "
                     "no acesso e empena de pedra com logo. Detalhe de cada material nas páginas de texturas T01–T12.")

# 7 · LEITURA GERAL
S = B.dup(7)
win = S.P(181, FR)
label_w(S, 182, "LANÇAMENTO · FACHADA FRONTAL")
place_markers(S, 181, win, list(range(183, 193)), old_markers(OLD, 4))
S.T(193, "A fachada vendida é um bloco de dois pavimentos com cobertura de grande beiral, pilares de eucalipto e painéis de "
         "madeira cobrindo as janelas, com jardineiras entre os pavimentos. O acesso central é marcado por um muxarabi e a "
         "esquina por uma empena de pedra com o logo Seazone.")
for sid, (a, b) in zip(range(196, 215, 2), [
        ("Logo-casinha Seazone", "na empena de pedra da esquina"),
        ("Empena em pedra", "pedra natural na ponta do bloco"),
        ("Cobertura com beiral", "telha metálica escura, viga metálica"),
        ("Pilares de madeira", "eucalipto roliço em toda a fachada"),
        ("Painéis e brises", "madeira amadeirada em frente a todas as janelas"),
        ("Jardineiras", "vegetação pendente nos pavimentos"),
        ("Muxarabi do acesso", "volume central em madeira"),
        ("Letreiro Imbassaí SPOT", "sobre ripado de madeira no jardim"),
        ("Muro de pedra", "mureta baixa na calçada"),
        ("Escada de acesso", "degraus claros até o hall")]):
    S.T(sid, a, b)

# 8-10 · TEXTURAS
NOTE = ("Amostras extraídas das imagens de lançamento (2023): referência visual, não substitui amostra física. "
        "Especificação: fachadas 08 e 09 do Projeto Legal R00, Memorial Descritivo e Premissas de Projeto.")
for title, cards in [
    ("Texturas da fachada · madeira e acesso (1/3)", [
        ("T01", "Pilares de eucalipto", "Pilares em toda a fachada e pergolado.",
         "Pilar em eucalipto (fachadas 08 e 09) · pergolado de eucalipto descoberto (planta do subsolo).",
         "Converge com a imagem: eucalipto tratado para área externa, acabamento como na imagem.", PILAR),
        ("T02", "Painéis e brises", "Frente das janelas e porta-janelas.",
         "Painéis/brises em madeira, tom amadeirado como na imagem, em todas as janelas; pode ser substituído por alumínio com o mesmo tom e acabamento.",
         "Preferência madeira tratada para área externa; se a obra optar por alumínio, manter o mesmo tom e acabamento; aprovar amostra contra a imagem.", BRISE),
        ("T03", "Muxarabi do acesso", "Volume central do hall de entrada.",
         "Muxarabi em madeira de lei tratada, desenho e tom como na imagem (fachada frontal PL).",
         "Mesmo desenho do muxarabi da recepção; detalhar espécie e fixação no executivo.", MUXA),
        ("T04", "Jardineiras", "Faixa entre os pavimentos.",
         "Jardineira em concreto revestida em madeira tratada, tom da imagem (fachadas PL).",
         "Prever dreno e impermeabilização antes do revestimento.", JARDINEIRA)]),
    ("Texturas da fachada · pedra, cobertura e paredes (2/3)", [
        ("T05", "Empena em pedra + logo", "Ponta do bloco na esquina.",
         "Pedra natural com logo-casinha Seazone iluminado, como na imagem.",
         "Incluir a pedra nas notas do projeto; prever ponto elétrico do logo.", PEDRA),
        ("T06", "Cobertura", "Telhado com beiral em todos os blocos.",
         "Telha metálica tipo sanduíche Isoeste (cobertura PL) · viga metálica no beiral (fachadas).",
         "Converge com a imagem (beiral escuro).", TELHADO),
        ("T07", "Paredes claras", "Paredes externas entre os pilares.",
         "Pintura sobre textura (PL), cor Papel de Presente Suvinil (B208), claro como na imagem · compatibilizar memorial.",
         "Aprovar painel-teste na fachada antes da compra; nunca Branco Puro.", PAREDE),
        ("T08", "Letreiro Imbassaí SPOT", "Jardim frontal, junto à calçada.",
         "Letras iluminadas sobre ripado de madeira, como na imagem · memorial: nome do edifício na fachada.",
         "Prever ponto elétrico do letreiro.", LETREIRO)]),
    ("Texturas da fachada · lazer e divisas (3/3)", [
        ("T09", "Ripado vertical", "Volume do acesso na lateral direita.",
         "Ripas verticais de madeira (imagem) · PL: muxarabi e brise fixo de madeira.",
         "Manter o ripado vertical, como na imagem; incluir no projeto.", RIPADO),
        ("T10", "Pergolado da piscina", "Área de lazer junto à piscina.",
         "Pergolado de madeira (imagem) · pergolado de eucalipto (PL, subsolo e térreo).",
         "Converge com a imagem.", PERG),
        ("T11", "Deck de madeira", "Entorno da piscina e áreas externas.",
         "Deck de madeira permeável (situação PL) · memorial: deck de madeira ou pisante em concreto.",
         "Deck de madeira permeável tratada para área externa, como na imagem.", DECK),
        ("T12", "Cerca e muro", "Divisas do terreno.",
         "Cerca de madeira com tela e mureta de pedra, como na imagem · PL: muro + cerca h=2,00 m.",
         "Executar como na imagem, com logo Seazone no acesso (Briefing).", CERCA)])]:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, c in zip(TEX, cards):
        card(S, ids, *c)

# 11 · VISTA PISCINA (tpl 11)
S = B.dup(11)
S.T(302, "Fachada da piscina")
win = S.P(306, PI)
label_w(S, 307, "LANÇAMENTO · FACHADA PISCINA")
place_markers(S, 306, win, list(range(308, 317)), old_markers(OLD, 7))
S.T(318, "LEITURA DA FACHADA DA PISCINA", "O fundo do bloco repete os painéis de madeira e as jardineiras, com o pergolado "
         "de madeira junto à piscina e deck contínuo. O projeto legal mantém pergolado e deck; os painéis em madeira vão em todas "
         "as janelas, como na imagem (ou alumínio no mesmo tom e acabamento).")
S.T(320, LEG)
vista_items(S, V11, [
    ("Beiral", "Telha sanduíche · viga metálica"),
    ("Painéis de madeira", "Madeira em todas as janelas, como na imagem · ou alumínio mesmo tom"),
    ("Jardineiras", "Concreto revestido em madeira tratada"),
    ("Pergolado", "Eucalipto (PL)"),
    ("Deck", "Deck de madeira permeável"),
    ("Piscina", "Esverdeado 20×20 (Briefing) · borda São Tomé (memorial)"),
    ("Espreguiçadeiras", "Madeira (imagem)"),
    ("Pisantes", "Pisante em concreto (memorial)"),
    ("Muro", "Cerca de madeira com tela + mureta de pedra (imagem) · logo Seazone")])

# 12 · VISTA LATERAL (tpl 11)
S = B.dup(11)
S.T(302, "Fachada lateral direita")
win = S.P(306, LA)
label_w(S, 307, "LANÇAMENTO · LATERAL DIREITA")
place_markers(S, 306, win, list(range(308, 317)), old_markers(OLD, 8))
S.T(318, "LEITURA DA FACHADA LATERAL", "Na lateral aparecem os dois blocos ligados pelo volume do acesso, com ripado vertical "
         "de madeira, painéis nas janelas e jardineiras contínuas. Na base, cerca de madeira com tela e mureta de pedra.")
S.T(320, LEG)
vista_items(S, V11, [
    ("Cobertura", "Telha metálica sanduíche Isoeste"),
    ("Painéis de madeira", "Madeira em todas as janelas · ou alumínio mesmo tom"),
    ("Ripado vertical", "Madeira, como na imagem · incluir no projeto"),
    ("Jardineiras", "Concreto revestido em madeira"),
    ("Pilares", "Eucalipto"),
    ("Paredes", "Pintura sobre textura, Papel de Presente (B208)"),
    ("Cerca", "Madeira com tela, como na imagem"),
    ("Mureta de pedra", "Pedra, como na imagem"),
    ("Volume do acesso", "Muxarabi em madeira de lei, como na imagem")])

# 13 · PONTOS DE ATENÇÃO FACHADA
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, FR, (0.18, 0.5))
label_w(S, 773, "LANÇAMENTO · FACHADA FRONTAL")
checklist(S, [
    ("Painéis e brises", "Madeira amadeirada em TODAS as janelas, como na imagem; pode ser substituído por alumínio com o mesmo tom e acabamento."),
    ("Paredes claras", "Papel de Presente Suvinil (B208), nunca Branco Puro: painel-teste antes da compra."),
    ("Empena de pedra e logo", "Incluir a pedra natural nas notas do projeto e o ponto elétrico do logo iluminado."),
    ("Muxarabi e ripado", "Madeira de lei, desenho e tom como na imagem; mesmo padrão do muxarabi da recepção."),
    ("Esquadrias e guarda-corpos", "Linha 32 imitando madeira (Briefing) · guarda-corpo de madeira com cabos de aço (PL e memorial)."),
    ("Divisas e letreiro", "Cerca de madeira com tela + mureta de pedra, logo Seazone no acesso e letreiro iluminado no jardim.")])

# 14 · DIVISOR 02
S = B.dup(14)
S.P(465, REC(2), (0.5, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito \"tranquilidade tropical\"  ·  recepção  ·  área de lazer e piscina  ·  pontos de atenção")

# 15 · CONCEITO × DECISÃO
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
concept_cards(S, [
    ("Conceito geral", LZ(5), (0.5, 0.5),
     "Inspiração balinesa: bege, marrom e amadeirado; palha, madeira e linho; textura com pigmento mineral",
     "Moodboard \"Tranquilidade tropical\" · estilo Clean"),
    ("Recepção", REC(2), (0.5, 0.5),
     "Painel muxarabi, parede em pedra, palha no teto, poltronas papasan, balanço, pendentes de palha",
     "Memorial da recepção (30/08/2023) · muxarabi em MDF Carvalho Treviso"),
    ("Área de lazer", LZ(13), (0.5, 0.5),
     "Pergolado de madeira, jardim vertical com frase, parede em pedra branca, balanços, bancada em concreto branco",
     "Moodboard \"Refúgio natural\" · piscina esverdeada 20×20 (Briefing)"),
    ("Revisão Seazone", REC(7), (0.5, 0.5),
     "Observações de 28/05/2023 sobre as imagens de interiores",
     "Retirar crochê e lã · cozinha menos amadeirada · spots no lugar do pendente"),
    ("Projeto de interiores", REC(8), (0.5, 0.5),
     "Imagens finais (10 - Projeto de Interiores / 02 - Imagens) e memoriais de repasse à Lofteria",
     "Sem executivo de interiores: memoriais e imagens valem como projeto")])
S.T(507, "Moodboards de interiores (05 - Imagens 3D / 04 - Moodboard Interno) e imagens finais (10 - Projeto de Interiores / 02 - Imagens).")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Vale o conceito das imagens finais de interiores e dos memoriais de repasse à Lofteria. "
         "O muxarabi da fachada se repete no painel da recepção: manter o mesmo padrão.")

# 16 · FICHA RECEPÇÃO
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO")
S.T(849, "Ficha: recepção")
S.P(853, REC(2), (0.5, 0.5)); label_w(S, 854, "INTERIORES · RECEPÇÃO")
S.P(855, REC(7), (0.5, 0.5)); label_w(S, 856, "LETREIRO IMBASSAÍ SPOT")
S.T(857, "Letreiro amadeirado com LED e frase \"Quando nos conectamos com a natureza, encontramos o equilíbrio que buscamos\".")
rows(S, ROWS27, [
    ("Parede-destaque", "Revestimento em pedra (briefing e imagem) + textura velvet areia (briefing)"),
    ("Muxarabi", "Painel muxarabi em MDF Carvalho Treviso (memorial da recepção)"),
    ("Palha", "Painel em palha na parede e no teto (memorial)"),
    ("Piso", "Porcelanato 90×90 natural/acetinado · rodapé branco 7 cm (briefing)"),
    ("Mobiliário", "Poltronas papasan, puff lounge em tricô (sem crochê nem lã), mesa lateral de vime, biombo de palha, balanço em corda náutica"),
    ("Iluminação", "Pendentes artesanais de bambu em diferentes alturas · postes de jardim"),
    ("Letreiros", "Letreiro e frase amadeirados com LED · ponto do logo iluminado (briefing)"),
    ("Acesso", "Porta em perfil metálico e vidro com reconhecimento facial (briefing)"),
    ("Climatização", "Infraestrutura de 22.000 BTU na parede central (briefing)")], PILLS27)

# 17 · MOODBOARD ÁREA DE LAZER (tpl 31, 11 imagens)
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · ÁREA DE LAZER · MOODBOARD")
S.T(978, "Área de lazer: integração, copa e piscina")
for pic, img in zip(range(982, 993), [LZ(13), LZ(11), LZ(2), LZ(1), LZ(4), LZ(6), LZ(7), LZ(9), LZ(10), LZ(12), LZ(14)]):
    S.P(pic, img)
S.T(993, "CONCEITO", "Refúgio natural: pergolado de madeira com forro de vigas e LED, jardim vertical com balanços e a frase "
         "\"Forget the maps, follow your instincts\", parede em pedra branca com frase.",
         "COPA", "Bancada em concreto branco, marcenaria em MDF Carvalho Treviso, revestimento Anticatto Olaria Lotus e pendentes de palha.",
         "ESTILO", "Clean: MDF Carvalho Treviso e Roble Catedral, palha, linho e pedra clara.")
for pic, lab, name, img in [(994, 995, "Pedra branca", PEDRABRANCA), (996, 997, "Jardim", JARDIMV),
                            (998, 999, "Palha", PENDENTE), (1000, 1001, "Concreto", BANCADA)]:
    S.P(pic, img); S.T(lab, name)

# 18 · FICHA ÁREA DE LAZER (interno)
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · ÁREA DE LAZER")
S.T(939, "Ficha de premissas: integração e copa")
S.P(943, LZ(5), (0.5, 0.5)); label_w(S, 944, "JARDIM VERTICAL + BALANÇOS")
S.P(945, LZ(7), (0.5, 0.5)); label_w(S, 946, "COPA")
S.T(947, "Jardim vertical com a frase \"Forget the maps, follow your instincts\" e parede em pedra branca com \"Na natureza, "
         "corpo e alma se entrelaçam em harmonia\".")
rows(S, ROWS30, [
    ("Forro", "Gesso com vigas de madeira e perfil de LED nas vigas (briefing) · vigas em MDF Carvalho Treviso"),
    ("Paredes", "Pedra natural branca + textura velvet (briefing) · pedra Moledo e cimento queimado (moodboard)"),
    ("Jardim vertical", "Jardim vertical preservado com frase (memorial e moodboard)"),
    ("Piso", "Porcelanato Urban Branco Savane (moodboard) · 90×90 antiderrapante nas áreas externas (briefing)"),
    ("Copa", "Bancada em concreto branco, MDF Carvalho Treviso, revestimento Anticatto Olaria Lotus (moodboard)"),
    ("Mobiliário", "Balanços Painho Tidelli, poltrona Dobra, puff Uni, banqueta Maya giratória"),
    ("Iluminação", "Luminária Cacimba em palha e pendentes na copa"),
    ("Balanços", "Prever reforço estrutural e pontos de tomada (Briefing)"),
    ("Frases", "Frases em letras pretas na pedra branca e iluminadas no jardim vertical (memorial)")], PILLS30)

# 19 · FICHA PISCINA E ÁREA EXTERNA
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · PISCINA")
S.T(1008, "Ficha de premissas: piscina e área externa")
S.P(1012, LZ(11), (0.5, 0.5)); label_w(S, 1013, "INTERIORES · PISCINA")
S.P(1014, PI, (0.4, 0.6)); label_w(S, 1015, "LANÇAMENTO · PISCINA")
S.T(1016, "Piscina com deck de madeira, pergolado de eucalipto e espreguiçadeiras de madeira, como na imagem.")
rows(S, ROWS32, [
    ("Piscina", "Revestimento esverdeado 20×20 (Briefing) · borda em pedra São Tomé (memorial)"),
    ("Aquecimento", "Solar para a piscina (briefing) · projeto no executivo"),
    ("Ducha", "Ducha de piso junto ao banheiro PNE (briefing)"),
    ("Deck", "Deck de madeira permeável tratada para área externa (PL e imagem)"),
    ("Pergolado", "Eucalipto (PL), como na imagem"),
    ("Pisantes", "Pisante em concreto (memorial)"),
    ("Espreguiçadeiras", "Madeira (imagem)"),
    ("Piso externo", "Porcelanato 90×90 antiderrapante (briefing)"),
    ("Divisas", "Cerca de madeira com tela + mureta de pedra (imagem) · logo Seazone no acesso")], PILLS32)

# 20 · PONTOS DE ATENÇÃO INTERIORES
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, REC(2), (0.6, 0.5))
label_w(S, 773, "RECEPÇÃO · MUXARABI E PEDRA")
checklist(S, [
    ("Sem executivo de interiores", "Memoriais de repasse (Lofteria) e imagens finais valem como projeto; trocas precisam de validação."),
    ("Muxarabi", "Mesmo desenho na fachada e na recepção; painel interno em MDF Carvalho Treviso."),
    ("Balanços", "Reforço estrutural e pontos de tomada na recepção e na área de lazer (Briefing)."),
    ("Piscina", "Esverdeado 20×20 (Briefing), não a pastilha azul da imagem; aquecimento solar e ducha de piso."),
    ("Revisão Seazone 28/05/2023", "Sem crochê nem lã nos puffs; copa menos amadeirada."),
    ("Letreiros e frases", "Letreiro amadeirado com LED e frases: prever pontos elétricos antes do revestimento.")])

# 21 · DIVISOR 03
S = B.dup(33)
S.P(1048, LZ(14), (0.5, 0.5))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T24: amostras tiradas das imagens finais e padrões do catálogo Clean para compra e aprovação")

# 22-24 · MATERIAIS
NOTE_M = "Amostras extraídas das imagens finais de interiores. Referência visual, não substitui amostra física."
for title, cards in [
    ("Materiais · recepção (1/3)", [
        ("T13", "Parede em pedra", "Parede-destaque da recepção.",
         "Revestimento em pedra (briefing) · imagem: pedra clara irregular + textura velvet areia.",
         "Pedra clara irregular, como na imagem; aprovar amostra e paginação antes da compra.", PEDRAINT),
        ("T14", "Muxarabi interno", "Painel da recepção.",
         "Painel muxarabi em MDF Berneck Carvalho Treviso (memorial da recepção).",
         "Mesmo desenho do muxarabi da fachada.", REC(2), (0.36, 0.28), 0.3),
        ("T15", "Palha e pendentes", "Recepção e copa da área de lazer.",
         "Painel em palha na parede e no teto · pendentes artesanais de bambu e luminária Cacimba em palha (memoriais).",
         "Prever pontos em diferentes alturas (briefing).", LZ(6), (0.33, 0.27), 0.16),
        ("T16", "Fibras naturais", "Mobiliário da recepção.",
         "Poltronas papasan em fibra, mesa lateral de vime e puffs em tricô (memorial da recepção).",
         "Puffs sem crochê nem lã (revisão Seazone 28/05/2023).", REC(2), (0.35, 0.72), 0.22)]),
    ("Materiais · área de lazer (2/3)", [
        ("T17", "Pedra branca", "Parede da área de lazer.",
         "Pedra natural branca + textura velvet (briefing) · frase em letras pretas.",
         "Aprovar amostra da pedra e fixação das letras antes da compra.", PEDRABRANCA),
        ("T18", "Jardim vertical", "Parede dos balanços na área de lazer.",
         "Jardim vertical preservado com frase (memorial da área de lazer).",
         "Prever reforço para os balanços e pontos de tomada (Briefing).", JARDIMV),
        ("T19", "Forro com vigas de madeira", "Área de integração.",
         "Gesso acartonado com vigas revestidas em MDF Carvalho Treviso e perfil de LED (briefing).",
         "Prever o perfil de LED nas vigas antes do fechamento do forro.", LZ(2), (0.42, 0.06), 0.12),
        ("T20", "Bancada em concreto branco", "Copa da área de lazer.",
         "Bancada em concreto branco (memorial da área de lazer).",
         "Concreto branco com impermeabilizante, acabamento como na imagem.", BANCADA)])]:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title); S.T(1065, NOTE_M)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e revestimentos (3/3)")
S.T(1146, "Padrões Berneck do catálogo Clean Seazone (rev. 13/10/2026) e revestimentos dos memoriais. Referência visual, não substitui amostra física.")
for ids, c in zip(MD, [
        ("T21", "MDF Carvalho Treviso", "Muxarabi e letreiro da recepção, vigas do forro e marcenaria da copa da área de lazer.",
         "MDF Berneck Carvalho Treviso (Design), amadeirado claro.",
         "Compatibilizar memoriais de interiores (Tauari); aprovar amostra física antes da compra.", sw("batel_amostra_carvalho_treviso.png")),
        ("T22", "MDF Roble Catedral", "Muxarabi das cabeceiras (memorial de interiores).",
         "MDF Berneck Roble Catedral (Grann), amadeirado.",
         "Compatibilizar memorial (Savana); aprovar amostra física antes da compra.", sw("batel_amostra_roble_catedral.png")),
        ("T23", "Piscina", "Piscina da área de lazer.",
         "Revestimento esverdeado 20×20 (Briefing) · borda em pedra São Tomé (memorial).",
         "A imagem mostra pastilha azul: seguir o Briefing e aprovar amostra antes do pedido.", PASTILHA),
        ("T24", "Revestimento da copa", "Parede da copa da área de lazer.",
         "Anticatto Olaria Lotus (moodboard).",
         "Aprovar amostra e paginação antes da compra.", LZ(7), (0.45, 0.42), 0.12)]):
    card(S, ids, *c)

# 25 · ENCERRAMENTO
S = B.dup(38)
S.P(1168, PI, (0.5, 0.5))
S.T(1170, "IMBASSAÍ SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)
S.T(1172, "Lançamento × Projeto  ·  Seazone")

finish(B, OUT, "IMBASSAÍ SPOT")
