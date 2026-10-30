"""Itacare Spot no layout v3 (template Campeche). Conteudo = decisoes aprovadas na v2 (13/10/2026).
python build_itacare.py DADOS saida.pptx   (DADOS = trabalho/v3_dados/itacare)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3"))
from bookv3 import *  # noqa: F401,F403

D, OUT = sys.argv[1], sys.argv[2]
OLD = os.path.join(REPO, "books", "legado", "ITACARE_SPOT_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
sw = lambda n: os.path.join(D, "sw", n)
O = lambda k: glob.glob(os.path.join(D, "old", k + "_*"))[0]
REC = lambda n: src(f"REC{n:02d}")
RF = lambda n: src(f"ROOF{n:02d}")
E = lambda n: src(f"E{n:02d}")
# amostras do legado
OFFWHITE, DIACHUVA, BRISE, ESQ, FAIXA, GCVIDRO = O("s05_12"), O("s05_19"), O("s05_26"), O("s05_33"), O("s05_40"), O("s05_47")
BIRIBA, RIPMED, PERG, LOGO, LETREIRO, PEDRA = O("s06_12"), O("s06_19"), O("s06_26"), O("s06_33"), O("s06_40"), O("s06_47")
TRELICA, PAINELLED, POLTRONA, FORRORIP, PASTILHA = O("s18_12"), O("s18_19"), O("s18_26"), O("s18_40"), O("s18_47")

B = Book(TPL, os.path.join(D, "cache"))

# 1 · CAPA
S = B.dup(1)
S.P(16, E(11), (0.55, 0.5))
S.P(17, E(1), (0.5, 0.5))
S.P(18, REC(2), (0.5, 0.5))
S.P(19, RF(2), (0.5, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.sh(21).text_frame.paragraphs[0].runs[0].font.size = Pt(8.5)
S.T(22, "Itacaré Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(36)
S.T(24, "Lançamento  ×  Projeto  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Biofílico")
S.T(25, "Rua Praia da Concha / Rua Hibisco · Conchas do Mar · Itacaré/BA", "DOCUMENTO TÉCNICO INTERNO")

# 2 · CONTEXTUALIZAÇÃO
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Itacaré Spot a partir da leitura cruzada entre "
        "as imagens de lançamento, o projeto legal da CMC Arquitetura, as indicações de fachada validadas pela diretoria e o "
        "projeto de interiores da Seazone: o que a imagem vendeu, o que o projeto traz e como a obra deve executar.")
S.P(37, E(11), (0.5, 0.45), 0.6)
S.P(40, BRISE)
S.P(43, RF(2), (0.6, 0.4), 0.3)
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], [
        "O que as imagens de lançamento mostram na fachada?",
        "Quais indicações de fachada a diretoria validou?",
        "Onde o projeto diverge da imagem vendida?",
        "Como são os brises, o biribá e o pergolado?",
        "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
        "O que não pode ser alterado sem validação?",
        "O que a obra precisa observar na execução?"]):
    S.T(sid, q)
S.T(62, "Regra do material: nenhuma informação foi inventada. Estilo de interiores: Biofílico. O executivo ainda não foi "
        "emitido. Prevalecem a imagem de lançamento, o conceito Seazone, o Briefing e o LOG validado pela diretoria; o projeto "
        "compatibiliza. Fontes (1.35 - [5624] Itacaré Spot / 02 - Projetos): 08 - Projeto Arquitetônico, 09 - Memorial "
        "Descritivo, 00 - Premissas de Projeto e Conferência de Projetos Executivos. Interiores das unidades, índices "
        "urbanísticos, áreas e custos não fazem parte deste Book.")

# 3 · COMO USAR
S = B.dup(3)
S.P(75, BIRIBA)
S.T(77, "IND 12")
S.T(78, "Código", "Código da textura (T07) ou da indicação do LOG de fachada (IND 12) para achar o item no projeto.")
S.P(87, E(11), (0.5, 0.5), None, 0.8)
S.P(90, REC(2), (0.5, 0.5))
S.P(93, RF(2), (0.5, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas da esquina e da rua · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção e corredores · rooftop, mini market e piscina · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T24 para compra e aprovação de amostras, com MDFs e metais do estilo Biofílico")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# 4 · DIVISOR 01
S = B.dup(4)
S.P(101, E(9), (0.5, 0.5))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# 5 · MOODBOARD
S = B.dup(5)
S.P(118, E(11), (0.5, 0.5)); label_w(S, 119, "FACHADA FRONTAL")
S.P(120, E(9), (0.5, 0.5)); label_w(S, 121, "ESQUINA")
S.P(122, E(13), (0.5, 0.5)); label_w(S, 123, "PRAIA")
S.P(124, E(6), (0.7, 0.5)); label_w(S, 125, "RUA")
S.P(126, E(1), (0.5, 0.5)); label_w(S, 127, "NOITE")
S.P(128, E(4), (0.5, 0.5)); label_w(S, 129, "ROOFTOP")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Brises ripados", "ripas verticais em madeira amadeirada em frente às janelas"),
        ("Cimento queimado", "bordas off white e paredes Dia de Chuva (LOG validado)"),
        ("Metais marrons", "esquadrias e guarda-corpos na textura marrom"),
        ("Gardens em biribá", "fechamento ripado de madeira no térreo"),
        ("Rooftop", "pergolado de madeira, vidro com floreira e letreiro seazone")]):
    S.T(sid, a, b)

# 6 · PALETA
S = B.dup(6)
for pic, lab, txt, code, a, b, img in [
        (151, 152, 153, "T01 · IND 05", "Cimento queimado off white", "bordas e pilares", OFFWHITE),
        (154, 155, 156, "T02 · IND 06", "Cimento queimado Dia de Chuva", "paredes (Suvinil)", DIACHUVA),
        (157, 158, 159, "T03", "Brise em madeira", "amadeirado, frente das janelas", BRISE),
        (160, 161, 162, "T04 · IND 13", "Esquadrias e metais marrons", "linha 32 bronze/marrom", ESQ),
        (163, 164, 165, "T07 · IND 12", "Biribá", "fechamento dos gardens", BIRIBA),
        (166, 167, 168, "T12 · IND 14", "Pedra moledo laguna branca", "paredes indicadas e rooftop", PEDRA)]:
    S.P(pic, img); S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = bordas e pilares em cimento queimado off white + paredes em cimento queimado Dia de Chuva + brises "
                     "ripados em madeira amadeirada + esquadrias e guarda-corpos marrons + gardens fechados em biribá + rooftop com vidro, "
                     "floreira e pergolado de madeira. Detalhe nas texturas T01–T12.")

# 7 · LEITURA GERAL
S = B.dup(7)
win = S.P(181, E(11))
label_w(S, 182, "LANÇAMENTO · FACHADA FRONTAL")
place_markers(S, 181, win, list(range(183, 193)), old_markers(OLD, 4))
S.T(193, "A fachada vendida tem dois pavimentos de studios com brises ripados verticais de madeira, pilares e bordas em "
         "cimento queimado claro, esquadrias marrons e gardens fechados por ripado biribá. No topo, rooftop com guarda-corpo de "
         "vidro, floreira, pergolado de madeira e letreiro seazone; na base, faixa amadeirada com o letreiro ITACARÉ SPOT.")
for sid, (a, b) in zip(range(196, 215, 2), [
        ("Pergolado do rooftop", "madeira sobre toda a extensão (IND 11)"),
        ("Letreiro seazone", "volume de concreto do rooftop"),
        ("Guarda-corpo do rooftop", "vidro h=1,50 m com floreira (IND 08/09)"),
        ("Brise ripado vertical", "madeira amadeirada em frente às janelas"),
        ("Pilares e bordas", "cimento queimado off white texturizado"),
        ("Faixa amadeirada", "viga de 30 cm para o nome (IND 10)"),
        ("Letreiro ITACARÉ SPOT", "letras brancas sobre a viga"),
        ("Logo-casinha", "branca no acesso"),
        ("Fechamento dos gardens", "ripado em madeira biribá (IND 12)"),
        ("Medidores e lixo", "painel ripado em madeira carvalho (IND 02)")]):
    S.T(sid, a, b)

# 8-10 · TEXTURAS
NOTE = ("Amostras extraídas das imagens de lançamento (jul/2026): referência visual, não substitui amostra física. "
        "Especificação: LOG de considerações de fachada (validado pela diretoria), Briefing, Premissas e Memorial.")
for title, cards in [
    ("Texturas da fachada · cores, brises e metais (1/3)", [
        ("T01", "Cimento queimado off white", "Lajes de borda e pilares.",
         "Textura cimento queimado off white, texturizada (LOG fachada · IND 05).",
         "Incluir a legenda de material no projeto e aprovar painel-teste.", OFFWHITE),
        ("T02", "Cimento queimado Dia de Chuva", "Paredes das unidades e divisas.",
         "Textura cimento queimado Dia de Chuva Suvinil, texturizada (LOG fachada · IND 06).",
         "Painel-teste antes da compra.", DIACHUVA),
        ("T03", "Brise ripado vertical", "Frente das janelas das unidades.",
         "Brise ripado vertical em madeira, tom amadeirado como na imagem, inclusive na condensadora; pode ser substituído por alumínio com o mesmo tom e acabamento.",
         "Preferência madeira tratada para área externa; se a obra optar por alumínio, manter o mesmo tom e acabamento; aprovar amostra contra a imagem.", BRISE),
        ("T04", "Esquadrias e metais marrons", "Janelas, portas e guarda-corpos das sacadas.",
         "Linha 32 cor bronze/marrom (Briefing e IND 13) · guarda-corpo em alumínio anodizado marrom com montante aparente.",
         "Completar a tabela de esquadrias antes do pedido.", ESQ)]),
    ("Texturas da fachada · acesso, rooftop e gardens (2/3)", [
        ("T05", "Faixa amadeirada + nome", "Acesso ao empreendimento.",
         "Pergolado no acesso com viga de 30 cm para fixar o nome (LOG · IND 10).",
         "Executar como na imagem; detalhar a viga e a fixação do nome no executivo.", FAIXA),
        ("T06", "Guarda-corpo do rooftop", "Borda do pavimento de lazer.",
         "Vidro h=1,50 m (LOG · IND 08) com perfil \"U\" inox (Briefing) e floreira (IND 09).",
         "Detalhar perfil e fixação no executivo antes do pedido.", GCVIDRO),
        ("T07", "Fechamento em biribá", "Fechamento de todos os gardens.",
         "Fechamento em madeira biribá (LOG · IND 12).",
         "Madeira biribá tratada para área externa; fixação detalhada no executivo.", BIRIBA),
        ("T08", "Painel ripado carvalho", "Fechamento dos medidores e do lixo.",
         "Painel ripado em madeira cor carvalho (LOG · IND 02); pode ser substituído por alumínio com o mesmo tom e acabamento.",
         "Preferência madeira tratada para área externa; se a obra optar por alumínio, manter o mesmo tom e acabamento; aprovar amostra contra a imagem.", RIPMED)]),
    ("Texturas da fachada · pergolado, letreiros e pedra (3/3)", [
        ("T09", "Pergolado do rooftop", "Extensão do rooftop.",
         "Pergolado de madeira em toda a extensão do rooftop (LOG · IND 11), como na imagem.",
         "Manter o pergolado; compatibilizar a cobertura aprovada (D11).", PERG),
        ("T10", "Letreiro seazone", "Volume do rooftop.",
         "Letras escuras iluminadas no volume do rooftop (imagem).",
         "Incluir ponto no elétrico (conferência 22).", LOGO),
        ("T11", "Letreiro ITACARÉ SPOT", "Faixa amadeirada do acesso.",
         "Letras brancas iluminadas na faixa amadeirada, como na imagem.",
         "Incluir ponto no elétrico antes do fechamento.", LETREIRO),
        ("T12", "Pedra moledo laguna branca", "Paredes indicadas na fachada e parede do rooftop.",
         "Pedra moledo laguna branca (LOG · IND 14).",
         "Paginação e locação no executivo.", PEDRA)])]:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, c in zip(TEX, cards):
        card(S, ids, *c)

# 11 · VISTA ESQUINA (tpl 11)
S = B.dup(11)
S.T(302, "Esquina e acesso")
win = S.P(306, E(9))
label_w(S, 307, "LANÇAMENTO · ESQUINA")
place_markers(S, 306, win, list(range(308, 317)), old_markers(OLD, 7))
S.T(318, "LEITURA DA ESQUINA", "A esquina marca o acesso com o volume ripado amadeirado, a logo-casinha e o letreiro ITACARÉ "
         "SPOT. O estacionamento recebe gradil com portão de correr e as vagas na via pública são em bloquete de concreto com grama.")
S.T(320, LEG)
vista_items(S, V11, [
    ("Platibanda", "Cimento queimado off white (IND 05)"),
    ("Brises", "Madeira ripada amadeirada, como na imagem · ou alumínio mesmo tom"),
    ("Parede", "Cimento queimado Dia de Chuva (IND 06)"),
    ("Letreiro", "ITACARÉ SPOT em letras brancas iluminadas"),
    ("Ripado", "Madeira ripada carvalho (IND 02) · ou alumínio mesmo tom"),
    ("Logo-casinha", "Branca iluminada no acesso, como na imagem"),
    ("Acesso", "Ripado de madeira + porta de vidro"),
    ("Gradil", "Gradil com portão de correr (IND 15)"),
    ("Vagas", "Bloquete + grama (IND 03) · paver no estacionamento interno (IND 04)")])

# 12 · VISTA RUA (tpl 11)
S = B.dup(11)
S.T(302, "Fachada da rua e gardens")
win = S.P(306, E(6))
label_w(S, 307, "LANÇAMENTO · RUA")
place_markers(S, 306, win, list(range(308, 317)), old_markers(OLD, 8))
S.T(318, "LEITURA DA FACHADA DA RUA", "Cada módulo de sacada tem spot sobre a janela, brise ripado e guarda-corpo marrom; os "
         "gardens do térreo ficam atrás do fechamento em biribá. Os pontos de iluminação de fachada entram no elétrico antes do fechamento.")
S.T(320, LEG)
vista_items(S, V11, [
    ("Pergolado", "Rooftop em toda a extensão (IND 11)"),
    ("Brises", "Madeira ripada amadeirada · ou alumínio mesmo tom"),
    ("Guarda-corpo", "Alumínio anodizado marrom com montante aparente (IND 13 e Briefing)"),
    ("Bordas", "Cimento queimado off white"),
    ("Fechamento", "Biribá nos gardens (IND 12)"),
    ("Iluminação", "Spot em cada módulo · incluir no elétrico"),
    ("Pisantes", "Placas de concreto sobre grama"),
    ("Esquadrias", "Marrom · porta de correr + muxarabi (P01)")])

# 13 · PONTOS DE ATENÇÃO FACHADA
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, E(11), (0.7, 0.5))
label_w(S, 773, "LANÇAMENTO · FACHADA FRONTAL")
checklist(S, [
    ("LOG de fachada validado", "As 15 indicações do LOG (diretoria) entram no projeto: cores, biribá, ripado carvalho, pedra, gradil e vagas."),
    ("Brises", "Madeira ripada amadeirada em todas as janelas, inclusive na condensadora; pode ser substituído por alumínio com o mesmo tom e acabamento."),
    ("Esquadrias e guarda-corpos", "Linha 32 bronze/marrom e alumínio anodizado marrom com montante aparente (IND 13 e Briefing)."),
    ("Rooftop", "Vidro h=1,50 m com perfil \"U\" inox e floreira; pergolado em toda a extensão: compatibilizar a cobertura aprovada (D11)."),
    ("Letreiros e logo", "ITACARÉ SPOT, logo-casinha e seazone iluminados: pontos no elétrico antes do fechamento."),
    ("Iluminação de fachada", "Spot em cada módulo de sacada, como na imagem: incluir no elétrico.")])

# 14 · DIVISOR 02
S = B.dup(14)
S.P(465, REC(2), (0.5, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito Seazone × decisão  ·  recepção e corredores  ·  rooftop e mini market  ·  piscina  ·  pontos de atenção")

# 15 · CONCEITO × DECISÃO
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
concept_cards(S, [
    ("Recepção", REC(2), (0.5, 0.5),
     "Painel de madeira com fitas de LED, treliça de bambu com vegetação e logo neon, poltronas creme, pendentes de fibra",
     "Poltronas de piso, como na imagem (sem poltrona balanço)"),
    ("Corredores", REC(3), (0.6, 0.5),
     "Papel picado, arandela sobre os números e rasgo de LED no forro",
     "Como na imagem e no Briefing"),
    ("Rooftop", RF(2), (0.5, 0.5),
     "Pergolado de madeira, parede de pedra com logo seazone, piscina, banco com floreira",
     "Volume e pergolado como na imagem · compatibilizar a cobertura aprovada"),
    ("Mini market", RF(3), (0.4, 0.5),
     "Letreiro, vending machine, adega embutida com LED",
     "Pontos elétricos do letreiro, vending e adega no elétrico"),
    ("Executivo de interiores", RF(6), (0.5, 0.5),
     "Contratar escritório especializado",
     "A contratar (ata 31/08/2026) · seguir este book")])
S.T(507, "Imagens de interiores Seazone × Conferência de Projetos Executivos (aba Interiores, 01/09/2026) e LOG de considerações.")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Não há executivo de interiores. Vale o conceito das imagens de venda (10 - Projeto de "
         "Interiores / 02 - Imagens) e do moodboard V02; logos iluminados, fitas de LED e pontos do mini market entram no elétrico.")

# 16 · FICHA RECEPÇÃO
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO E CORREDORES")
S.T(849, "Ficha: recepção e corredores")
S.P(853, REC(2), (0.5, 0.5)); label_w(S, 854, "CONCEITO · RECEPÇÃO")
S.P(855, REC(4), (0.5, 0.5)); label_w(S, 856, "PAINEL COM LED")
S.T(857, "Logo neon \"Itacaré SPOT\" sobre a treliça de bambu com vegetação; corredor com rasgo de LED no forro.")
rows(S, ROWS27, [
    ("Paredes", "Painel PVC amadeirado liso + painel quadriculado em bambu + textura velvet (briefing)"),
    ("Forro", "Gesso 2,60 m revestido no mesmo painel amadeirado da parede (briefing e LOG, indicação 10)"),
    ("Piso", "Porcelanato 90×90 natural/acetinado · rodapé 7 cm branco · carpete de alumínio embutido na entrada (LOG)"),
    ("Destaques", "Treliça de bambu com vegetação e logo neon · quadros de marketing com arandelas"),
    ("Mobiliário", "Poltronas creme com estrutura metálica, tapete de fibra (imagem) · sem poltrona balanço"),
    ("Iluminação", "Pendentes de fibra, spots e 3 fitas de LED verticais no painel (4000 K)"),
    ("Corredores", "Papel picado ou similar · arandela sobre os números e rasgo de LED no forro"),
    ("Portas", "Brancas, bandeira até o forro · rodapés brancos (LOG)"),
    ("Escada", "Corrimão metálico bronze (imagem)")], PILLS27)

# 17 · MOODBOARD ROOFTOP (tpl 31, 11 imagens)
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP · MOODBOARD")
S.T(978, "Rooftop: piscina, mini market e lazer")
for pic, img in zip(range(982, 993), [RF(2), RF(3), RF(4), RF(1), RF(5), RF(6), RF(7), RF(8), RF(9), RF(10), src("ROOF_PLANTA")]):
    S.P(pic, img)
S.T(993, "CONCEITO", "Parede de pedra com logo seazone iluminado, banco com almofadas e floreira, pergolado de madeira com forro "
         "ripado, mini market com letreiro, vending e adega com LED.",
         "PISCINA", "Piscina com borda em pedra São Tomé, revestimento esverdeado 20×20 e guarda-corpo de vidro h=1,50 m.",
         "ESTILO", "Biofílico: MDF Carvalho Treviso e Desert, metal bege RAL 1013, bambu e fibras naturais.")
for pic, lab, name, img in [(994, 995, "Pedra", PEDRA), (996, 997, "Forro ripado", FORRORIP),
                            (998, 999, "Piscina", PASTILHA), (1000, 1001, "Carvalho", sw("batel_amostra_carvalho_treviso.png"))]:
    S.P(pic, img); S.T(lab, name)

# 18 · FICHA ROOFTOP (lazer e mini market)
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP")
S.T(939, "Ficha de premissas: lazer, copa e mini market")
S.P(943, RF(4), (0.5, 0.5)); label_w(S, 944, "CONCEITO · ROOFTOP")
S.P(945, RF(7), (0.5, 0.5)); label_w(S, 946, "PEDRA + LOGO + BANCO")
S.T(947, "Parede de pedra com logo seazone iluminado, banco com almofadas e floreira, pergolado de madeira.")
rows(S, ROWS30, [
    ("Paredes", "Textura velvet + cimento queimado + pedra (briefing) · copa toda em pedra (LOG)"),
    ("Forro / pergolado", "Gesso 2,60 m com ripas revestidas em esteira de bambu (Briefing), visual ripado da imagem"),
    ("Piso", "Porcelanato 90×90 · externo antiderrapante · porcelanato imitando madeira 120×20"),
    ("Mini market", "Letreiro iluminado, vending machine, adega/frigobar com fita de LED (imagens)"),
    ("Banco", "Banco com floreira atrás (LOG)"),
    ("Logo", "seazone iluminado na parede de pedra: ponto elétrico e fixação antes da pedra"),
    ("Iluminação", "Arandelas na parede de pedra e fitas de LED no mini market, como na imagem"),
    ("Mobiliário", "Mesa de jantar com cadeiras de fibra, almofadões e mesas laterais brancas (imagem)"),
    ("BWC e DML", "Textura Branco Gatinho Coral (LOG fachada, indicação 07)")], PILLS30)

# 19 · FICHA PISCINA
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · PISCINA")
S.T(1008, "Ficha de premissas: piscina e borda do rooftop")
S.P(1012, RF(10), (0.5, 0.5)); label_w(S, 1013, "CONCEITO · PISCINA")
S.P(1014, src("E04"), (0.5, 0.5)); label_w(S, 1015, "LANÇAMENTO · ROOFTOP")
S.T(1016, "Piscina voltada para o mar, com borda em pedra, espreguiçadeiras e guarda-corpo de vidro.")
rows(S, ROWS32, [
    ("Piscina", "Esverdeado 20×20, ref. Portobello Barlavento Sea (premissa e LOG)"),
    ("Borda", "Pedra São Tomé (memorial)"),
    ("Piso externo", "Porcelanato 90×90 antiderrapante · porcelanato imitando madeira 120×20 (briefing)"),
    ("Guarda-corpo", "Vidro h=1,50 m (IND 08) com perfil \"U\" inox (Briefing)"),
    ("Floreira", "Espaço de floreira no rooftop (IND 09)"),
    ("Pergolado", "Madeira em toda a extensão (IND 11) · compatibilizar a cobertura aprovada (D11)"),
    ("Aquecimento", "Solar (Briefing) · projeto no executivo"),
    ("Ducha", "Ducha de parede + ralo · incluir no hidrossanitário"),
    ("Casa de máquinas", "Detalhada no executivo (conferência)")], PILLS32)

# 20 · PONTOS DE ATENÇÃO INTERIORES
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, RF(7), (0.5, 0.5))
label_w(S, 773, "ROOFTOP · PEDRA E LOGO")
checklist(S, [
    ("Executivo de interiores", "A contratar (ata 31/08/2026): até lá, vale este book e as imagens de venda."),
    ("Logos e LED", "Logo neon, seazone na pedra e 3 fitas de LED no painel: pontos no elétrico antes do fechamento."),
    ("Mini market", "Pontos do letreiro, vending machine e adega com LED no elétrico."),
    ("Recepção", "Poltronas de piso, como na imagem: sem poltrona balanço e sem reforço na laje."),
    ("Piscina", "Esverdeado 20×20 Barlavento Sea (imagem azul escuro): aprovar amostra; aquecimento solar no executivo."),
    ("Forro do rooftop", "Ripas com esteira de bambu (Briefing), visual ripado da imagem.")])

# 21 · DIVISOR 03
S = B.dup(33)
S.P(1048, RF(7), (0.5, 0.5))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T24: amostras tiradas das imagens Seazone e padrões do catálogo Biofílico para compra e aprovação")

# 22-24 · MATERIAIS
NOTE_M = "Amostras extraídas das imagens de interiores Seazone. Referência visual, não substitui amostra física."
for title, cards in [
    ("Materiais · recepção (1/3)", [
        ("T13", "Treliça de bambu", "Parede do logo na recepção.",
         "Painel quadriculado em bambu com vegetação e logo neon (briefing e imagem).",
         "Incluir no elétrico o ponto do logo iluminado.", TRELICA),
        ("T14", "Painel amadeirado com LED", "Recepção, parede e forro.",
         "Painel PVC amadeirado liso com 3 fitas de LED verticais (briefing e imagem).",
         "Mesmo painel no forro (LOG, indicação 10).", PAINELLED),
        ("T15", "Poltronas e tapete de fibra", "Recepção.",
         "Poltronas creme com estrutura metálica e tapete redondo de fibra (imagem).",
         "Seguir a imagem: poltronas de piso, sem poltrona balanço.", POLTRONA),
        ("T16", "Corrimão bronze", "Escada da recepção.",
         "Corrimão metálico bronze (imagem).",
         "Mesmo tom das esquadrias marrons; aprovar amostra de cor.", REC(1), (0.22, 0.35), 0.22)]),
    ("Materiais · rooftop (2/3)", [
        ("T17", "Parede de pedra", "Rooftop, parede do logo seazone.",
         "Pedra moledo laguna branca (LOG, indicação 14), aspecto rústico da imagem.",
         "Prever fixação e ponto do logo iluminado antes da pedra.", PEDRA),
        ("T18", "Forro ripado", "Rooftop.",
         "Ripas revestidas em esteira de bambu (Briefing), visual ripado da imagem.",
         "Compatibilizar o LOG (bambu ou palha) com o Briefing.", FORRORIP),
        ("T19", "Revestimento da piscina", "Piscina do rooftop.",
         "Esverdeado 20×20 ref. Portobello Barlavento Sea (premissa e LOG) · borda em pedra São Tomé.",
         "A imagem mostra azul escuro: seguir a premissa e aprovar amostra antes do pedido.", PASTILHA),
        ("T20", "Mini market com LED", "Rooftop, junto à copa.",
         "Adega/frigobar embutidos com fita de LED e letreiro iluminado (imagens).",
         "Pontos elétricos do letreiro, vending e adega antes do fechamento.", RF(3), (0.33, 0.55), 0.3)])]:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title); S.T(1065, NOTE_M)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e metais (3/3)")
S.T(1146, "Padrões Berneck e cores de metal do catálogo Biofílico Seazone (rev. 13/10/2026). Referência visual, não substitui amostra física.")
for ids, c in zip(MD, [
        ("T21", "MDF Carvalho Treviso", "Marcenaria amadeirada (moodboard V02).",
         "MDF Berneck Carvalho Treviso (Design), amadeirado claro.",
         "Já é do catálogo Biofílico; aprovar amostra física antes da compra.", sw("batel_amostra_carvalho_treviso.png")),
        ("T22", "MDF Desert", "Marcenaria clara, em composição com o Carvalho Treviso.",
         "MDF Berneck Desert (Linha Smart, Vel), liso.",
         "Já é do catálogo Biofílico; aprovar amostra física antes da compra.", sw("amostra_desert.png")),
        ("T23", "Metal bege RAL 1013", "Estruturas e metais do mobiliário.",
         "Metal pintado bege RAL 1013 (catálogo Biofílico).",
         "Mesmo acabamento em todas as peças metálicas; aprovar amostra de cor.", sw("amostra_metal_bege_ral1013.png")),
        ("T24", "Metal marrom", "Esquadrias e guarda-corpos das sacadas (fachada).",
         "Linha 32 bronze/marrom e alumínio anodizado marrom (Briefing e IND 13).",
         "Mantido na fachada, como na imagem; aprovar amostra de cor.", ESQ)]):
    card(S, ids, *c)

# 25 · ENCERRAMENTO
S = B.dup(38)
S.P(1168, E(1), (0.55, 0.5))
S.T(1170, "ITACARÉ SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)
S.T(1172, "Lançamento × Projeto  ·  Seazone")

finish(B, OUT, "ITACARÉ SPOT")
