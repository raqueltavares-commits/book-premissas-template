"""Salvador Spot no layout v3 (template Campeche). Fonte = book legado da Rachel (pptx) com as regras de 15/10/2026
(tudo DEFINIDO, unidades fora, estilo Clean, MDFs trocados: Carvalho Natural -> Carvalho Treviso, Areia Noronha -> Desert).
python build_salvador.py DADOS saida.pptx   (DADOS = trabalho/v3_dados/salvador)
"""
import glob
import os
import sys

sys.path.insert(0, os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3"))
from bookv3 import *  # noqa: F401,F403

D, OUT = sys.argv[1], sys.argv[2]
OLD = os.path.join(REPO, "books", "legado", "SALVADOR_SPOT_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
swc = lambda n: os.path.join(D, "..", "sw_common", n)
O = lambda k: glob.glob(os.path.join(D, "old", k + "_*.jpg"))[0]
E = lambda n: src(f"E{n:02d}")
REC = lambda n: src(f"REC{n:02d}")
RF = lambda n: src(f"ROOF{n:02d}")
FRENTE, ESQUINA, AEREA, RUA = E(2), E(3), E(4), E(1)
# amostras do legado (pic id = id do rotulo - 2)
TORRE, CINZA, BRISE, ESQ, PAINEL, FLOREIRA = O("s05_12"), O("s05_19"), O("s05_26"), O("s05_33"), O("s05_40"), O("s05_47")
PERG, LOGO, LETREIRO, ROOFT, TERREO, PISOEXT = O("s06_12"), O("s06_19"), O("s06_26"), O("s06_33"), O("s06_40"), O("s06_47")
MUXA, FITAS, CACOS, REBOCO, VIGAS, CEMENTO = O("s18_12"), O("s18_19"), O("s18_26"), O("s18_33"), O("s18_40"), O("s18_47")
TREVISO, DESERT = swc("amostra_carvalho_treviso_clean.png"), swc("amostra_desert.png")

B = Book(TPL, os.path.join(D, "cache"))


def vista_portrait(S, img, label, legacy_slide):
    """tpl 12 com imagem retrato inteira (torre alta): quadro no aspecto da imagem e lista puxada para a esquerda"""
    iw, ih = B.img(img).size
    x, y, w, h = S.pos(358)
    nw = h * iw / ih
    dx = w - nw
    S.geom(358, w=nw)
    for sid in [369, 370] + [n for t in V12 for n in t[:2]]:
        S.geom(sid, x=S.pos(sid)[0] - dx)
    for _, tid, _ in V12:
        S.geom(tid, w=S.pos(tid)[2] + dx)
    S.geom(370, w=S.pos(370)[2] + dx)
    win = S.P(358, img)
    label_w(S, 359, label)
    place_markers(S, 358, win, list(range(360, 369)), old_markers(OLD, legacy_slide))


# 1 · CAPA
S = B.dup(1)
S.P(16, FRENTE, (0.5, 0.5))
S.P(17, RUA, (0.5, 0.45))
S.P(18, REC(2), (0.5, 0.5))
S.P(19, RF(6), (0.5, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.sh(21).text_frame.paragraphs[0].runs[0].font.size = Pt(8.5)
S.T(22, "Salvador Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(36)
S.T(24, "Lançamento  ×  Executivo  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Clean")
S.T(25, "Rua Milton Oliveira · Barra · Salvador/BA", "DOCUMENTO TÉCNICO INTERNO")

# 2 · CONTEXTUALIZAÇÃO
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Salvador Spot a partir da leitura cruzada "
        "entre as imagens de lançamento (P124), o projeto executivo da GS Arquitetos, o Briefing de premissas, o LOG de "
        "considerações e o projeto de interiores da Seazone: o que a imagem vendeu, o que o executivo traz e como a obra deve executar.")
S.P(37, ESQUINA, (0.5, 0.3), 0.45)
S.P(40, CINZA)
S.P(43, RF(1), (0.75, 0.65), 0.6)
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], [
        "O que as imagens de lançamento mostram na fachada?",
        "Onde o executivo precisa ser ajustado à imagem vendida?",
        "Qual revestimento externo vale para a compra?",
        "Como são os brises, as floreiras e o pergolado?",
        "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
        "O que não pode ser alterado sem validação?",
        "O que a obra precisa observar na execução?"]):
    S.T(sid, q)
S.T(62, "Regra do material: nenhuma informação foi inventada. Estilo de interiores: Clean. Prevalecem a imagem de lançamento, "
        "o conceito Seazone, o Briefing e o LOG de considerações; o executivo compatibiliza. Fontes (1.14 - [2074] Salvador Spot / "
        "02 - Projetos): 05 - Imagens 3D, 07 - Projeto Executivo (PE-ARQ R11), 09 - Memorial Descritivo, 10 - Projeto de Interiores, "
        "00 - Premissas de Projeto e 12 - Conferência de Projetos. Interiores das unidades, índices urbanísticos, áreas e custos "
        "não fazem parte deste Book.")

# 3 · COMO USAR
S = B.dup(3)
S.P(75, BRISE)
S.T(77, "T03")
S.T(78, "Código", "Código da textura (T03) ou da indicação do LOG (LOG 01) para achar o item no projeto.")
S.P(87, RUA, (0.5, 0.45), None, 0.8)
S.P(90, REC(2), (0.5, 0.5))
S.P(93, RF(6), (0.5, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas da esquina e do rooftop · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção · rooftop interno e externo · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T24 para compra e aprovação de amostras, com MDFs do estilo Clean")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# 4 · DIVISOR 01
S = B.dup(4)
S.P(101, RUA, (0.45, 0.45))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# 5 · MOODBOARD FACHADA
S = B.dup(5)
S.P(118, RUA, (0.5, 0.45)); label_w(S, 119, "RUA")
S.P(120, RF(1), (0.5, 0.5)); label_w(S, 121, "ROOFTOP · PISCINA")
S.P(122, FRENTE, (0.5, 0.5)); label_w(S, 123, "FRENTE")
S.P(124, ESQUINA, (0.5, 0.5)); label_w(S, 125, "ESQUINA")
S.P(126, AEREA, (0.5, 0.22)); label_w(S, 127, "AÉREA")
S.P(128, RF(11), (0.5, 0.6)); label_w(S, 129, "ROOFTOP · DECK")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Volumes cinza claros", "textura cinza clara granulada nos volumes principais, pilares e bordas"),
        ("Torre escura", "volume vertical texturizado grafite/marrom com o logo-casinha iluminado no topo"),
        ("Brises e painéis amadeirados", "brises ripados em frente às janelas e painel amadeirado no fundo das sacadas"),
        ("Floreiras e pergolado", "vegetação na borda de todas as lajes e pergolado ripado sobre o volume frontal"),
        ("Rooftop e térreo", "piscina esverdeada com deck e vidro; térreo envidraçado com jardim frontal")]):
    S.T(sid, a, b)

# 6 · PALETA
S = B.dup(6)
for pic, lab, txt, code, a, b, img in [
        (151, 152, 153, "T02", "Textura cinza clara granulada", "volumes, pilares e bordas", CINZA),
        (154, 155, 156, "T01", "Textura escura Leeds Máximo", "torre, no tom da imagem", TORRE),
        (157, 158, 159, "T03", "Brise em madeira", "amadeirado, frente das janelas", BRISE),
        (160, 161, 162, "T04", "Esquadrias linha 32 pretas", "janelas e portas-janela", ESQ),
        (163, 164, 165, "T05", "Painel amadeirado", "fundo das sacadas", PAINEL),
        (166, 167, 168, "T06", "Floreiras com vegetação", "borda de todas as lajes", FLOREIRA)]:
    S.P(pic, img); S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = volumes em textura cinza clara granulada (T02) + torre escura texturizada (T01) + brises ripados "
                     "em madeira amadeirada (T03) + esquadrias pretas (T04) + painel amadeirado nas sacadas (T05) + floreiras em todas "
                     "as lajes (T06). Detalhe de cada material nas páginas de texturas T01–T12.")

# 7 · LEITURA GERAL (tres imagens retrato da torre; marcadores na frente, como no legado)
S = B.dup(7)
x, y, w, h = S.pos(181)
iw, ih = B.img(FRENTE).size
pw = h * iw / ih
gap = (w - 3 * pw) / 2
S.geom(181, w=pw)
win = S.P(181, FRENTE)
label_w(S, 182, "LANÇAMENTO · FRENTE")
for k, (img, lab) in enumerate([(ESQUINA, "ESQUINA"), (AEREA, "AÉREA")], 1):
    nid = S.clone(181, x=x + k * (pw + gap))
    drop = S.s.part.drop_rel
    S.s.part.drop_rel = lambda rId: None  # o clone divide a imagem com o 181: nao soltar a relacao
    S.P(nid, img)
    S.s.part.drop_rel = drop
    lid = S.clone(182, x=S.pos(182)[0] + k * (pw + gap))
    label_w(S, lid, lab)
place_markers(S, 181, win, list(range(183, 193)), old_markers(OLD, 4))
S.T(193, "A imagem vendida é uma torre cinza clara granulada com um volume vertical escuro que leva o logo-casinha iluminado no "
         "topo. Em frente às janelas, brises ripados de madeira e painéis amadeirados no fundo das sacadas; floreiras com vegetação "
         "em todas as lajes. No alto, pergolado ripado e rooftop com piscina; na base, térreo envidraçado e jardim frontal.")
for sid, (a, b) in zip(range(196, 215, 2), [
        ("Logo-casinha iluminado", "topo da torre escura (letreiro na caixa d'água)"),
        ("Torre escura texturizada", "Ibratin Leeds Máximo no tom grafite/marrom da imagem"),
        ("Textura cinza clara granulada", "volumes principais, pilares e bordas"),
        ("Brises móveis ripados", "madeira amadeirada em frente às janelas"),
        ("Esquadrias pretas", "linha 32, caixilho quadriculado, vidro incolor"),
        ("Floreiras em todas as lajes", "caqueiros com vegetação na borda"),
        ("Painel amadeirado", "fundo das sacadas, ao lado das janelas"),
        ("Pergolado ripado", "terraço sobre o volume frontal"),
        ("Térreo envidraçado", "recepção · esquadria imitando madeira (Briefing)"),
        ("Piso externo com grama", "placas claras, faixas verdes e bancos")]):
    S.T(sid, a, b)

# 8-10 · TEXTURAS DA FACHADA
NOTE = ("Amostras extraídas das imagens de lançamento (P124): referência visual, não substitui amostra física. "
        "Especificação: Executivo PE-ARQ R11 (fachadas), Briefing de premissas, LOG de considerações e Memorial.")
for title, cards in [
    ("Texturas da fachada · volumes, brises e esquadrias (1/3)", [
        ("T01", "Torre escura texturizada", "Volume vertical da circulação e caixa d'água.",
         "Textura Ibratin cor Leeds Máximo (executivo), no tom grafite/marrom da imagem de lançamento.",
         "Compatibilizar o tom com a imagem; aprovar painel-teste contra a imagem antes da compra.", TORRE),
        ("T02", "Textura cinza clara granulada", "Volumes principais, pilares, bordas e empenas.",
         "Textura cinza clara granulada, como na imagem de lançamento (P124).",
         "Seguir a imagem; compatibilizar o executivo (pastilha Atlas Rubi SG-8409 e Cinza Chumbo B-2134); aprovar amostra "
         "contra a imagem antes da compra.", CINZA),
        ("T03", "Brise móvel ripado", "Frente das janelas, em todos os pavimentos.",
         "Brise móvel ripado em madeira (memorial), tom amadeirado e ripas horizontais como na imagem; pode ser substituído por alumínio com o mesmo tom e acabamento.",
         "Preferência madeira tratada para área externa; se a obra optar por alumínio, manter o mesmo tom e acabamento; aprovar amostra contra a imagem.", BRISE),
        ("T04", "Esquadrias pretas", "Janelas e portas-janela dos pavimentos-tipo.",
         "Linha 32 (gold), PVC ou alumínio anodizado (Briefing), cor preta com caixilho quadriculado (imagem) · vidro incolor.",
         "Compatibilizar o executivo (sem linha e cor); vidro conforme laudo térmico-acústico.", ESQ)]),
    ("Texturas da fachada · sacadas e identidade (2/3)", [
        ("T05", "Painel amadeirado", "Fundo das sacadas, ao lado das janelas.",
         "Painel liso em madeira para área externa, tom amadeirado como na imagem; pode ser substituído por alumínio com o mesmo tom e acabamento.",
         "Sem indicação no executivo: detalhar no executivo; aprovar amostra contra a imagem.", PAINEL),
        ("T06", "Floreiras nas lajes", "Borda de todas as sacadas.",
         "Caqueiros com vegetação (imagem), pintura cimento queimado escuro (executivo).",
         "Paisagismo das floreiras: incluir no projeto complementar.", FLOREIRA),
        ("T07", "Pergolado ripado", "Terraço sobre o volume frontal.",
         "Ripas claras, como na imagem · estrutura metálica aço corten com madeira (memorial) · proteção ignifugante (executivo).",
         "Executar o visual ripado claro da imagem; detalhar material e fixação no executivo.", PERG),
        ("T08", "Logo-casinha iluminado", "Topo da torre escura.",
         "Logo branco com luz de contorno, como na imagem · letreiro na caixa d'água com ponto de força (memorial).",
         "Luminotécnico: Projetos entrega antes do fechamento de fachada.", LOGO)]),
    ("Texturas da fachada · rooftop, térreo e jardim (3/3)", [
        ("T09", "Letreiro seazone", "Empena cinza do volume frontal.",
         "Letras escuras sobre a textura clara, como na imagem (memorial: nome do edifício na fachada).",
         "Como na imagem; Projetos detalha fixação e ponto elétrico.", LETREIRO),
        ("T10", "Rooftop: vidro, deck e piscina", "Rooftop, 15º pavimento.",
         "Guarda-corpo de vidro com perfil \"U\" inox · deck em porcelanato imitando madeira 120×20 · piscina esverdeada 20×20 (Briefing).",
         "Compatibilizar o executivo (alumínio e vidros laminados) para perfil \"U\" inox antes do pedido.", ROOFT),
        ("T11", "Térreo envidraçado", "Recepção e acesso.",
         "Esquadrias das áreas comuns imitando madeira (Briefing e LOG) · porta com vidro incolor e painel ripado (executivo).",
         "Seguir Briefing e LOG (a imagem externa mostra caixilho preto); compatibilizar o executivo.", TERREO),
        ("T12", "Piso externo e jardins", "Recuo frontal e acesso.",
         "Placas claras com faixas de grama e bancos (imagem) · piso drenante (executivo) · floreira + banco de cobogó iluminado (Briefing).",
         "Paginação e paisagismo: detalhar no executivo / incluir no complementar.", PISOEXT)])]:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, c in zip(TEX, cards):
        card(S, ids, *c)

# 11 · VISTA ESQUINA (tpl 12, retrato)
S = B.dup(12)
S.T(354, "Esquina e volume frontal")
vista_portrait(S, ESQUINA, "LANÇAMENTO · ESQUINA", 7)
S.T(370, LEG)
vista_items(S, V12, [
    ("Torre escura", "Ibratin Leeds Máximo no tom grafite/marrom da imagem"),
    ("Letreiro seazone", "Letras escuras na empena · Projetos detalha fixação e ponto elétrico"),
    ("Empena", "Textura cinza clara granulada, como na imagem · compatibilizar o executivo"),
    ("Brises", "Madeira ripada amadeirada, como na imagem · ou alumínio mesmo tom"),
    ("Esquadrias", "Linha 32 preta, caixilho quadriculado (imagem e Briefing)"),
    ("Floreiras", "Caqueiros com vegetação · paisagismo no projeto complementar"),
    ("Pergolado", "Ripado claro (imagem) · estrutura aço corten + madeira (memorial)"),
    ("Jardim frontal", "Placas claras e grama (imagem) · piso drenante (executivo)")])
S.P(398, CINZA); S.P(399, TORRE); S.P(400, BRISE)
S.T(401, "T02 · T01 · T03")
S.T(403, "LEITURA DA ESQUINA", "A imagem vendida é cinza clara, com torre escura, brises de madeira e floreiras em todas as lajes. "
         "O executivo indica pastilha terracota (Rubi) e cinza chumbo: seguir a imagem, compatibilizar o executivo e aprovar "
         "amostra contra a imagem antes da compra.")

# 12 · VISTA AÉREA (tpl 12, retrato)
S = B.dup(12)
S.T(354, "Vista aérea: rooftop e terraços")
vista_portrait(S, AEREA, "LANÇAMENTO · AÉREA", 8)
S.T(370, LEG)
vista_items(S, V12, [
    ("Logo na torre", "Casinha iluminada · ponto de força no elétrico"),
    ("Integração", "Volume envidraçado, esquadria imitando madeira (Briefing)"),
    ("Deck", "Porcelanato imitando madeira 120×20 (Briefing)"),
    ("Piscina", "Revestimento esverdeado 20×20 (Briefing), como na aérea"),
    ("Guarda-corpo", "Vidro incolor + perfil \"U\" inox (Briefing) · compatibilizar o executivo"),
    ("Terraços", "Floreiras com vegetação, como na imagem"),
    ("Pergolado", "Ripado com vegetação sobre o volume frontal"),
    ("Acesso", "Escada e jardim frontal · paisagismo no projeto complementar")])
S.P(398, ROOFT, (0.2, 0.5)); S.P(399, ROOFT, (0.72, 0.6)); S.P(400, PERG)
S.T(401, "Deck · piscina · pergolado")
S.T(403, "LEITURA DO ROOFTOP", "A aérea confirma a piscina esverdeada do Briefing e o deck ripado. O rooftop foi redesenhado "
         "no executivo para atender aos afastamentos, mantendo o conceito: executar conforme as imagens e o Briefing.")

# 13 · PONTOS DE ATENÇÃO FACHADA
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, FRENTE, (0.5, 0.45))
label_w(S, 773, "LANÇAMENTO · FRENTE")
checklist(S, [
    ("Textura dos volumes (T02)", "Seguir a imagem: cinza clara granulada. Compatibilizar o executivo (pastilha Rubi e Cinza Chumbo); "
                                  "aprovar amostra contra a imagem antes da compra."),
    ("Torre escura (T01)", "Ibratin Leeds Máximo no tom grafite/marrom da imagem: aprovar painel-teste contra a imagem."),
    ("Brises e painel amadeirado", "Brise ripado e painel das sacadas em madeira amadeirada, como na imagem; pode ser substituído por alumínio com o mesmo tom e acabamento."),
    ("Esquadrias e guarda-corpos", "Pavimentos-tipo: linha 32 preta. Comuns: imitando madeira. Guarda-corpo comum vidro + \"U\" inox; "
                                   "privativo alumínio anodizado com montante aparente."),
    ("Identidade e luz", "Logo-casinha e letreiro seazone como na imagem. Luminotécnico: Projetos entrega antes do fechamento de fachada."),
    ("Muro, floreiras e jardim", "Muro claro Papel de Presente Suvinil (B208) com logo e nome Seazone; paisagismo no projeto complementar.")])

# 14 · DIVISOR 02
S = B.dup(14)
S.P(465, REC(2), (0.55, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito Seazone × decisão  ·  recepção  ·  rooftop interno: copa e mini market  ·  rooftop externo: piscina e deck  ·  pontos de atenção")

# 15 · CONCEITO × DECISÃO
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
concept_cards(S, [
    ("Recepção", REC(2), (0.6, 0.5),
     "Painel muxarabi com drywall curvo, reboco baiano, logo e letreiro iluminados, balanço, sofá nuvem, pendentes de fibra",
     "Seguir o conceito: drywall curvo sobre o muxarabi detalhado no executivo de interiores (LOG 01)"),
    ("Rooftop interno", src("ROOF00"), (0.6, 0.5),
     "Parede de fitinhas do Bonfim, piso com cacos, vigas de madeira, mini market, copa com bancada de concreto branco",
     "Seguir o conceito: copa e paginação do piso detalhadas no executivo de interiores"),
    ("Rooftop externo", RF(11), (0.6, 0.5),
     "Deck ripado, piscina esverdeada, banco de cobogó iluminado com floreira",
     "Piscina esverdeada 20×20 (Briefing), como na aérea; o render interno azul não vale"),
    ("Portas e esquadrias", REC(6), (0.45, 0.4),
     "Portas amadeiradas com bandeira até o forro e esquadrias imitando madeira (LOG e Briefing)",
     "Portas das áreas comuns amadeiradas (LOG) · compatibilizar a premissa (porta branca)"),
    ("Executivo de interiores", RF(9), (0.5, 0.5),
     "Imagens de venda, moodboard final e indicações 01 a 04 do LOG",
     "Seguir o conceito Seazone (imagens) até a entrega do executivo")])
S.T(507, "Imagens de lançamento interna × Conferência de Projetos Executivos (aba Interiores), LOG de considerações, moodboard final e memorial de interiores.")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Vale o conceito das imagens de venda (05 - Imagens 3D / 02 - Lançamento Interna), o moodboard "
         "final e o LOG 01 a 04. Muxarabi, reboco baiano, fitinhas do Bonfim, piso de cacos e vigas de madeira são conceito Seazone: "
         "manter. Logos, letreiros, LED e mini market entram no elétrico antes do fechamento.")

# 16 · FICHA RECEPÇÃO
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO")
S.T(849, "Ficha: recepção (pé-direito duplo)")
S.P(853, REC(2), (0.55, 0.5)); label_w(S, 854, "CONCEITO · RECEPÇÃO")
S.P(855, REC(1), (0.5, 0.5)); label_w(S, 856, "BALANÇO + JARDIM")
S.T(857, "Painel muxarabi com letreiro \"Salvador SPOT\", logo-casinha iluminado, balanço de corda e vasos de barro.")
rows(S, ROWS27, [
    ("Paredes", "Pedra Textura Velvet + painel em MDF muxarabi + reboco baiano (Briefing e LOG 02)"),
    ("Painel", "Muxarabi com drywall curvo sobreposto e fita de LED (LOG 01) · fixação detalhada no executivo de interiores"),
    ("Forro", "Gesso, pé-direito duplo (Briefing)"),
    ("Piso", "Porcelanato 90×90 natural/acetinado, ref. Cemento Grigio (moodboard) · rodapé branco de poliestireno"),
    ("Madeira", "MDF Berneck Carvalho Treviso (Design) · portas amadeiradas com bandeira até o forro (LOG)"),
    ("Mobiliário", "Sofá estofado Nuvem, puff Pumpkin, poltronas de bambu, balanço Cariri, tapete redondo Samba"),
    ("Identidade", "Letreiro Salvador SPOT, logo seazone e frase: amadeirados com LED"),
    ("Iluminação", "Pendentes artesanais de fibra (4000 K) · interruptor próprio para logos e frases"),
    ("Decoração", "Vasos de barro com cactos · plantas artificiais (premissa) com o visual das imagens")], PILLS27)

# 17 · MOODBOARD ROOFTOP (tpl 31)
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP · MOODBOARD")
S.T(978, "Rooftop: integração, copa, mini market e piscina")
for pic, img, foc in zip(range(982, 993),
                         [src("ROOF00"), RF(6), RF(8), RF(9), RF(1), RF(10), RF(5), RF(11), RF(4), src("LROOF02"), AEREA],
                         [(.6, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.45, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .22)]):
    S.P(pic, img, foc)
S.T(993, "CONCEITO", "Parede de fitinhas do Senhor do Bonfim com letreiro iluminado e balanço, reboco baiano com frase, piso de "
         "cacos, vigas de madeira, mesa central com pendentes de fibra, copa e mini market.",
         "PISCINA", "Revestimento esverdeado 20×20 (Briefing), deck em porcelanato imitando madeira e banco de cobogó iluminado "
         "com floreira.",
         "ESTILO", "Clean: marcenaria em MDF Carvalho Treviso e Desert.")
for pic, lab, name, img in [(994, 995, "Fitinhas", FITAS), (996, 997, "Cacos", CACOS),
                            (998, 999, "Vigas", VIGAS), (1000, 1001, "Reboco", REBOCO)]:
    S.P(pic, img); S.T(lab, name)

# 18 · FICHA ROOFTOP INTERNO (tpl 30)
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP INTERNO")
S.T(939, "Ficha de premissas: integração, copa e mini market")
S.P(943, RF(6), (0.5, 0.5)); label_w(S, 944, "PAREDE DE FITINHAS")
S.P(945, RF(9), (0.5, 0.5)); label_w(S, 946, "COPA + MINI MARKET")
S.T(947, "Fitinhas do Senhor do Bonfim, piso de caquinhos, vigas de madeira, balanço e letreiros iluminados.")
rows(S, ROWS30, [
    ("Paredes", "Reboco baiano + estrutura metálica revestida com fitas do Senhor do Bonfim (Briefing e LOG 02)"),
    ("Copa", "Bancada em concreto branco (moodboard, como na imagem) · parede em pedra branca canjiquinha (Briefing)"),
    ("Piso", "Porcelanato 90×90 mesclado com cacos de pisos diversos (Briefing e LOG 03) · paginação no executivo de interiores"),
    ("Forro", "Gesso com vigas de madeira aparentes (Briefing e LOG 04) · cortineiro nas portas e janelas"),
    ("Marcenaria", "MDF Berneck Carvalho Treviso (Design) · mini market com fita de LED e logo iluminado"),
    ("Mobiliário", "Mesa Flexa, cadeiras Doha, banco Diamantina, puff Pumpkin, poltronas de bambu, balanço Painho"),
    ("Letreiros", "Frases iluminadas e letreiro \"Salvador SPOT\" amadeirado com LED"),
    ("Iluminação", "Trilho de iluminação + pendentes de fibra sobre a mesa central"),
    ("Esquadrias", "Acabamento imitando madeira em toda a área comum (Briefing e LOG)")], PILLS30)

# 19 · FICHA ROOFTOP EXTERNO (tpl 32)
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · ROOFTOP EXTERNO")
S.T(1008, "Ficha de premissas: piscina e deck")
S.P(1012, RF(1), (0.5, 0.5)); label_w(S, 1013, "CONCEITO · PISCINA")
S.P(1014, RF(11), (0.5, 0.5)); label_w(S, 1015, "DECK + BANCO")
S.T(1016, "Deck ripado, banco de cobogó iluminado com almofadas e floreira com vegetação atrás.")
rows(S, ROWS32, [
    ("Piso externo", "Porcelanato que imita madeira 120×20, rodapé embutido (Briefing)"),
    ("Piscina", "Revestimento esverdeado 20×20 (Briefing), como na aérea · o render interno azul não vale"),
    ("Borda", "Pedra São Tomé, atérmica (memorial)"),
    ("Referência", "\"On Fire Glacier\" na página da piscina (moodboard final)"),
    ("Banco", "Floreira em alvenaria + banco de cobogó com iluminação interna (Briefing)"),
    ("Guarda-corpo", "Vidro incolor com perfil \"U\" inox (Briefing) · compatibilizar o executivo (alumínio e laminado)"),
    ("Mobiliário", "Espreguiçadeiras Selima e puffs Joy (memorial de interiores)"),
    ("Ducha", "Ducha de piso + ralo e torneira próxima (Briefing) · incluir no hidrossanitário"),
    ("Aquecimento", "Solar para a piscina (Briefing)")], PILLS32)

# 20 · PONTOS DE ATENÇÃO INTERIORES
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, src("ROOF00"), (0.72, 0.5))
label_w(S, 773, "ROOFTOP · FITINHAS")
checklist(S, [
    ("Recepção: muxarabi e drywall curvo", "Seguir o conceito (LOG 01): drywall curvo sobre o muxarabi, com fita de LED, detalhado no executivo de interiores."),
    ("Identidade baiana", "Muxarabi, reboco baiano, fitinhas do Bonfim, piso de cacos e vigas de madeira: manter como nas imagens."),
    ("Rooftop interno", "Paginação dos cacos, estrutura das fitinhas e vigas no executivo de interiores; painel-teste do reboco baiano."),
    ("Piscina", "Esverdeado 20×20 (Briefing), como na aérea; o render interno azul não vale. Aquecimento solar; aprovar amostra."),
    ("Portas e esquadrias comuns", "Portas amadeiradas com bandeira até o forro (LOG); esquadrias imitando madeira (Briefing e LOG)."),
    ("Executivo de interiores", "Seguir o conceito Seazone (imagens) até a entrega do executivo; logos, letreiros e LED no elétrico.")])

# 21 · DIVISOR 03
S = B.dup(33)
S.P(1048, RF(6), (0.45, 0.5))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T24: amostras tiradas das imagens Seazone e padrões do catálogo Clean para compra e aprovação")

# 22-23 · MATERIAIS
NOTE_M = "Amostras do \"Moodboard - Salvador Spot - Final\" e das imagens de venda. Referência visual, não substitui amostra física."
for title, cards in [
    ("Materiais · recepção e rooftop (1/3)", [
        ("T13", "Painel muxarabi", "Recepção, parede do letreiro.",
         "Painel em MDF muxarabi (Briefing, moodboard) com drywall curvo por cima (LOG 01) e fita de LED.",
         "Manter o recorte curvo da imagem; fixação do drywall detalhada no executivo de interiores.", MUXA),
        ("T14", "Fitinhas do Bonfim", "Rooftop, parede do balanço.",
         "Estrutura metálica revestida com fitas coloridas do Senhor do Bonfim (Briefing, moodboard).",
         "Estrutura e fixação no executivo de interiores; ponto do letreiro antes do revestimento.", FITAS),
        ("T15", "Piso de caquinhos", "Rooftop interno.",
         "Porcelanato mesclado com cacos cerâmicos, referência ao Pelourinho (Briefing, LOG 03, moodboard).",
         "Paginação no executivo de interiores, como na imagem.", CACOS),
        ("T16", "Reboco baiano", "Paredes da recepção e do rooftop.",
         "Acabamento em reboco baiano (Briefing e LOG 02), textura clara como nas imagens.",
         "Aprovar painel-teste antes da execução.", REBOCO)]),
    ("Materiais · forro, piso, luz e banco (2/3)", [
        ("T17", "Vigas de madeira", "Forro do rooftop.",
         "Vigas de madeira aparentes sob o forro de gesso (Briefing e LOG 04).",
         "Planta de forro com as vigas detalhada no executivo de interiores.", VIGAS),
        ("T18", "Porcelanato Cemento Grigio", "Piso da recepção e do rooftop interno.",
         "Porcelanato Cemento Grigio (moodboard), 90×90 natural/acetinado (Briefing).",
         "No rooftop, mesclar com os cacos (T15) conforme a paginação.", CEMENTO),
        ("T19", "Pendentes de fibra", "Pé-direito duplo da recepção e mesa central do rooftop.",
         "Pendentes artesanais de fibra, 4000 K (Briefing e imagens).",
         "Interruptor próprio para logos e frases (Briefing); pontos no forro conforme as imagens.", REC(3), (0.62, 0.18), 0.3),
        ("T20", "Banco de cobogó iluminado", "Rooftop externo, junto à floreira.",
         "Floreira em alvenaria + banco de cobogó com iluminação interna e almofadas (Briefing e imagem).",
         "Prever ponto elétrico da iluminação interna antes do revestimento.", RF(11), (0.78, 0.78), 0.3)])]:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title); S.T(1065, NOTE_M)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e amadeirados (3/3)")
S.T(1146, "Padrões Berneck do catálogo Clean Seazone (rev. 13/10/2026) e acabamentos amadeirados das imagens. Referência visual, não substitui amostra física.")
for ids, c in zip(MD, [
        ("T21", "MDF Carvalho Treviso", "Marcenaria das áreas comuns, copa e mini market.",
         "MDF Berneck Carvalho Treviso (Design), amadeirado claro.",
         "Compatibilizar memorial de interiores (Guararapes Carvalho Natural); aprovar amostra física antes da compra.", TREVISO),
        ("T22", "MDF Desert", "Marcenaria clara: aéreos e prateleiras, com o Carvalho Treviso.",
         "MDF Berneck Desert (Linha Smart, Vel), liso.",
         "Compatibilizar memorial de interiores (Areia Noronha); aprovar amostra física antes da compra.", DESERT),
        ("T23", "Portas amadeiradas", "Portas das áreas comuns.",
         "Kit porta pronta lisa, acabamento amadeirado, bandeira até o forro (LOG e observação do Briefing).",
         "Compatibilizar a premissa (porta branca); aprovar amostra junto com o Carvalho Treviso.", REC(2), (0.33, 0.55), 0.35),
        ("T24", "Esquadrias imitando madeira", "Esquadrias das áreas comuns (térreo e rooftop).",
         "Acabamento imitando madeira (Briefing e LOG), como nas imagens internas.",
         "Compatibilizar o executivo (alumínio); aprovar amostra de cor contra as imagens.", REC(1), (0.82, 0.6), 0.35)]):
    card(S, ids, *c)

# 25 · ENCERRAMENTO
S = B.dup(38)
S.P(1168, ESQUINA, (0.5, 0.4))
S.T(1170, "SALVADOR SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)
S.T(1172, "Lançamento × Executivo  ·  Seazone")

finish(B, OUT, "SALVADOR SPOT")
