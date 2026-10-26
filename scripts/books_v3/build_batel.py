"""Batel Spot no layout v3 (template Campeche). Conteudo = decisoes aprovadas na v2 (13/10/2026).

python build_batel.py DADOS saida.pptx
DADOS: pasta com src/ (renders e interiores do Drive, ja reduzidos), old/ (imagens extraidas do book legado)
e sw/ (amostras Berneck). Base: build_canasvieiras.py da Rachel.
"""
import os
import sys

KIT = os.path.expanduser("~/Claude/.claude/skills/book-premissas/scripts/v3")
sys.path.insert(0, KIT)
from tplkit import Book, E  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Pt  # noqa: E402

D = sys.argv[1]
OUT = sys.argv[2]
TPL = os.path.expanduser("~/Claude/projetos/book-premissas-template/template/"
                         "TEMPLATE_v3_CAMPECHE_SPOT_Book_de_Premissas_rev13-10-2026.pptx")
OLD = os.path.expanduser("~/Claude/projetos/book-premissas-template/books/legado/BATEL_SPOT_Book_de_Premissas_legado.pptx")
src = lambda n: os.path.join(D, "src", n + ".jpg")
old = lambda n: os.path.join(D, "old", n)
sw = lambda n: os.path.join(D, "sw", n)
R001, R002, R003, R004, R005, R006, R007, R008 = (src(f"R00{i}") for i in range(1, 9))
REC1, REC2, RECP = src("REC01"), src("REC02"), src("REC_PLANTA")
RF = lambda n: src(f"ROOF0{n}")
RFP = src("ROOF_PLANTA")
SWO = {"R1": "s05_12_6c0a7607.jpg", "R2": "s05_19_1f019240.jpg", "R3": "s05_26_801f5f76.jpg", "R4": "s05_33_5f5bfb8c.jpg",
       "esq": "s05_40_8d8b05c3.jpg", "gc": "s05_47_75ed4bd4.jpg", "flor": "s06_12_1c4a2e56.jpg", "letreiro": "s06_19_f962922f.jpg",
       "marquise": "s06_26_2221c0f3.jpg", "totem": "s06_33_fbef955b.jpg", "muro": "s06_40_9de61ba7.jpg", "arauc": "s06_47_d41e9832.jpg",
       "cobogo": "s19_12_a07a1ab3.jpg", "jardim": "s19_19_52b014a0.jpg", "vinilico": "s19_26_58678b3b.jpg",
       "superquadra": "s19_33_0de470c7.jpg", "saogabriel": "s19_40_d08c0afa.jpg", "paredeescura": "s19_47_a1bf8005.jpg",
       "pastilha": "s20_12_bbc2d3d7.jpg", "deck": "s20_19_4c129b24.jpg", "couro": "s20_40_81403315.jpg", "pendente": "s20_47_9601073e.jpg"}
S_ = lambda k: old(SWO[k])

D_ = "DEFINIDO"
LEG = "Status do item: DEFINIDO"


def old_markers(slide_no, pic_id=11):
    p = Presentation(OLD)
    sl = p.slides[slide_no - 1]
    pic = [s for s in sl.shapes if s.shape_id == pic_id][0]
    cl, cr, ct, cb = pic.crop_left, pic.crop_right, pic.crop_top, pic.crop_bottom
    out = {}
    for s in sl.shapes:
        if s.has_text_frame and s.text_frame.text.strip().isdigit() and s.width < 0.26 * E and \
           pic.left <= s.left <= pic.left + pic.width and pic.top <= s.top <= pic.top + pic.height:
            cx = s.left + s.width / 2; cy = s.top + s.height / 2
            u = cl + (cx - pic.left) / pic.width * (1 - cl - cr)
            v = ct + (cy - pic.top) / pic.height * (1 - ct - cb)
            out[int(s.text_frame.text)] = (u, v)
    return [out[k] for k in sorted(out)]


def place_markers(S, pic_id, win, marker_ids, pts):
    for mid, (u, v) in zip(marker_ids, pts):
        x, y, ok = S.map_pt(pic_id, win, u, v)
        w = S.pos(mid)[2]
        S.geom(mid, x - w / 2, y - w / 2)
        if not ok:
            print("  marcador fora do quadro:", mid)
    for mid in marker_ids[len(pts):]:
        S.delete(mid)


def label_w(S, sid, txt):
    S.T(sid, txt)
    S.geom(sid, w=0.30 + 0.078 * len(txt))


def card(S, ids, code, title, onde, spec, obra, img, focus=(0.5, 0.5), hfrac=None):
    pic, lab, txt = ids
    S.P(pic, img, focus, hfrac)
    S.T(lab, code)
    S.T(txt, title, "ONDE", onde, "ESPECIFICAÇÃO", spec, "ATENÇÃO NA OBRA", obra)


def rows(S, rws, data, pill_ids, full_w=5.98):
    for pid in pill_ids:
        S.delete(pid)
    for (lid, tid), (lab, txt) in zip(rws, data):
        S.T(lid, lab); S.T(tid, txt)
        S.geom(tid, w=full_w)


def checklist(S, data):
    rws = [(775, 776, 777), (778, 779, 780), (782, 783, 784), (785, 786, 787), (789, 790, 791), (792, 793, 794)]
    for (box, tid, pid), (title, body) in zip(rws, data):
        S.T(tid, title, body)
        S.T(box, "✓")
        y = S.pos(pid)[1]
        S.delete(pid)
        S.pill(D_, 12.71 - 0.90, y)


def vista_items(S, pairs, data):
    for (nid, tid, iid), (title, body) in zip(pairs, data):
        S.T(tid, title, body)
        S.status_icon(iid, D_)
    for (nid, tid, iid) in pairs[len(data):]:
        S.delete(nid, tid, iid)


B = Book(TPL, os.path.join(D, "cache"))

# ================= 1 · CAPA =================
S = B.dup(1)
S.P(16, R002, (0.5, 0.45))
S.P(17, R005, (0.45, 0.5))
S.P(18, REC1, (0.55, 0.5))
S.P(19, RF(4), (0.5, 0.5))
S.T(21, "BOOK DE PREMISSAS DE FACHADA E INTERIORES DAS ÁREAS COMUNS")
S.T(22, "Batel Spot")
S.sh(22).text_frame.paragraphs[0].runs[0].font.size = Pt(36)
S.T(24, "Lançamento  ×  Executivo  ·  Fachada  ·  Interiores das áreas comuns  ·  Estilo Bruma")
S.T(25, "Av. Visconde de Guarapuava, 5354 · Batel · Curitiba/PR", "DOCUMENTO TÉCNICO INTERNO")

# ================= 2 · CONTEXTUALIZAÇÃO =================
S = B.dup(2)
S.T(36, "Consolidar as premissas de fachada e de interiores das áreas comuns do Batel Spot a partir da leitura cruzada "
        "entre as imagens de lançamento (R03), o projeto executivo da HS Arquitetura e o projeto de interiores da Seazone: "
        "o que a imagem vendeu, o que o executivo especifica, onde o executivo precisa ser ajustado e o que a obra deve "
        "seguir na compra e na execução.")
S.P(37, R001, (0.5, 0.45), 0.6)
S.P(40, S_("R2"))
S.P(43, R004, (0.75, 0.55), 0.35)
S.T(45, "CONFIRA", "o que a obra precisa observar e seguir")
qs = ["O que as imagens de lançamento mostram na fachada?",
      "Qual cor e textura o executivo define para cada volume?",
      "Onde o executivo precisa ser ajustado à imagem vendida?",
      "Como são os brises, os guarda-corpos e os letreiros?",
      "Qual é o conceito de interiores das áreas comuns que deve ser seguido?",
      "O que não pode ser alterado sem validação?",
      "O que a obra precisa observar na execução?"]
for sid, q in zip([49, 51, 53, 55, 57, 59, 61], qs):
    S.T(sid, q)
S.T(62, "Regra do material: premissas validadas pela Seazone. Estilo de interiores: Bruma. Prevalecem a imagem de lançamento, "
        "o conceito Seazone e o Briefing; o executivo compatibiliza. Fontes (1.23 - [3299] Batel Spot / 02 - Projetos): "
        "imagens 3D R03, executivo HS 0598, Memorial Descritivo R01, projeto de interiores Seazone e Premissas de Projeto. "
        "Interiores das unidades, índices urbanísticos, áreas e custos não fazem parte deste Book.")

# ================= 3 · COMO USAR =================
S = B.dup(3)
S.P(75, S_("R3"))
S.T(77, "R3")
S.T(78, "Código", "Código do executivo/memorial (ex.: R1 a R4, GC1, T07) para achar o item no projeto.")
S.P(87, R005, (0.5, 0.45), None, 0.8)
S.P(90, RF(1), (0.5, 0.5))
S.P(93, RF(3), (0.3, 0.5))
S.T(89, "Fachada", "imagens de lançamento · paleta · leitura geral · texturas T01–T12 · vistas frontal, lateral e acesso · pontos de atenção")
S.T(92, "Interiores das áreas comuns", "conceito × decisão · recepção e corredores · ático · terraço e piscina · pontos de atenção")
S.T(95, "Materiais de interiores das áreas comuns", "texturas T13–T28 para compra e aprovação de amostras, com MDFs e metais do estilo Bruma")
S.T(79, "STATUS DOS ITENS")
S.T(81, "premissa validada pela Seazone: seguir na compra e na execução")
S.delete(82, 83, 84, 85)

# ================= 4 · DIVISOR 01 =================
S = B.dup(4)
S.P(101, R005, (0.4, 0.5))
S.T(105, "Imagens de lançamento  ·  paleta  ·  leitura geral  ·  texturas  ·  vistas com marcação e especificação  ·  pontos de atenção")

# ================= 5 · MOODBOARD FACHADA =================
S = B.dup(5)
S.P(118, R001, (0.45, 0.45)); label_w(S, 119, "001 · FRONTAL")
S.P(120, R005, (0.45, 0.5)); label_w(S, 121, "005 · PERSPECTIVA")
S.P(122, R002, (0.6, 0.5)); label_w(S, 123, "002")
S.P(124, R004, (0.5, 0.5)); label_w(S, 125, "004")
S.P(126, R008, (0.6, 0.5)); label_w(S, 127, "008 · ACESSO")
S.P(128, R006, (0.5, 0.5)); label_w(S, 129, "006 · ÁTICO")
for sid, (a, b) in zip([132, 134, 136, 138, 140], [
        ("Volumes bege e cinza", "R2 Suvinil C164 Ovelha nos volumes e R3 cimento queimado Imensidão Cinza nos painéis recuados"),
        ("Lajes marcadas", "bordas de laje em tom cinza (R1), como na imagem, desenhando cada pavimento"),
        ("Brises amadeirados", "faixas verticais de brises em alumínio tom amêndoa em frente às janelas"),
        ("Base e acesso", "araucárias preservadas, marquise preta com BATEL SPOT, totem da casa e muro claro com arandelas"),
        ("Coroamento e ático", "letreiro SEAZONE na caixa d'água, floreiras, guarda-corpo de vidro e piscina no terraço")]):
    S.T(sid, a, b)

# ================= 6 · PALETA =================
S = B.dup(6)
pal = [(151, 152, 153, "R2", "Suvinil C164 Ovelha", "volumes principais", S_("R2")),
       (154, 155, 156, "R3", "Cimento queimado Imensidão Cinza", "painéis recuados e caixa d'água", S_("R3")),
       (157, 158, 159, "R1", "Cimento queimado em tom cinza", "bordas de laje (como na imagem)", S_("R1")),
       (160, 161, 162, "R4", "Brise alumínio amadeirado", "faixas em frente às janelas", S_("R4")),
       (163, 164, 165, "ESQUADRIAS", "Linha 32 preto fosco", "janelas e porta-janelas", S_("esq")),
       (166, 167, 168, "GC1–GC3", "Vidro + perfil \"U\" inox", "guarda-corpos do ático", S_("gc"))]
for pic, lab, txt, code, a, b, img in pal:
    S.P(pic, img); S.T(lab, code); S.T(txt, a, b)
S.T(170, "COMO LER", "Fachada = volumes bege (R2) + painéis cinza (R3) + bordas de laje cinza (R1) + brises em alumínio amadeirado (R4) + "
                     "esquadria preta + guarda-corpo de vidro com perfil \"U\" inox. Detalhe de cada material nas páginas de texturas T01–T12.")

# ================= 7 · LEITURA GERAL =================
S = B.dup(7)
win = S.P(181, R005)
label_w(S, 182, "LANÇAMENTO · 005 (PERSPECTIVA)")
place_markers(S, 181, win, list(range(183, 193)), old_markers(4))
S.T(193, "A fachada é lida em três camadas: volumes em tom bege (R2) e cinza (R3) com lajes marcadas em cinza (R1); faixas "
         "verticais de brises em alumínio amadeirado (R4) em frente às janelas; e uma base ajardinada com araucárias preservadas, "
         "marquise com letreiro BATEL SPOT e totem da casa Seazone. No ático, o volume da caixa d'água leva o letreiro SEAZONE.")
leg = [("Letreiro SEAZONE no ático", "letras iluminadas, como na imagem"),
       ("Vegetação da cobertura", "floreiras dos terraços · impermeabilização antirraiz"),
       ("Brises verticais das janelas", "R4 brise em alumínio tom amêndoa"),
       ("Volumes em tom bege", "R2 pintura Suvinil C164 Ovelha"),
       ("Painéis recuados cinza", "R3 cimento queimado Imensidão Cinza"),
       ("Bordas das lajes", "R1 cimento queimado em tom cinza, como na imagem"),
       ("Esquadrias", "linha 32 (gold), preto fosco"),
       ("Floreiras de fachada", "alvenaria/concreto com vegetação"),
       ("Marquise com letreiro", "BATEL SPOT iluminado sobre marquise metálica preta"),
       ("Totem de acesso", "casa Seazone iluminada no jardim")]
for sid, (a, b) in zip(range(196, 215, 2), leg):
    S.T(sid, a, b)

# ================= 8-10 · TEXTURAS DA FACHADA =================
CARDS = [(227, 228, 229), (231, 232, 233), (235, 236, 237), (239, 240, 241)]
NOTE = ("Amostras extraídas das imagens de lançamento: referência visual, não substitui amostra física. "
        "Especificação: legenda \"Revestimentos das fachadas\" (PE 11 a 14), Memorial Descritivo R01 e Briefing.")
tex = [
    ("Texturas da fachada · cores e revestimentos (1/3)", [
        ("T01", "R1 · Bordas de laje", "Bordas de laje e faixas horizontais das quatro elevações.",
         "Efeito cimento queimado Coral em tom cinza, como na imagem de lançamento (PE 11–14 indica Branco Gatinho).",
         "Compatibilizar o executivo para o tom cinza da imagem; aprovar amostra contra a imagem antes da compra.", S_("R1")),
        ("T02", "R2 · Pintura Ovelha", "Volumes principais em tom bege das quatro fachadas.",
         "Pintura acrílica Suvinil, cor C164 Ovelha, sobre textura/massa acrílica (memorial 23.2).",
         "Painel-teste na fachada; mesma cor e lote em todas as faces.", S_("R2")),
        ("T03", "R3 · Cimento queimado cinza", "Painéis recuados entre janelas e volume da caixa d'água.",
         "Efeito cimento queimado Coral, cor Imensidão Cinza (PE 11–14 · memorial 26.3.3).",
         "Efeito texturizado: aplicar por equipe única e aprovar amostra em painel.", S_("R3")),
        ("T04", "R4 · Brise amêndoa", "Faixas verticais em frente às janelas, do 2º pavimento ao ático.",
         "Brise em alumínio com acabamento amadeirado tom amêndoa, como na imagem · BR1 a BR4 (DT 09).",
         "Executivo indica alumínio: manter, com acabamento amadeirado no tom da imagem; aprovar amostra antes da compra.", S_("R4"))]),
    ("Texturas da fachada · esquadrias e vegetação (2/3)", [
        ("T05", "Esquadrias", "Janelas e porta-janelas de todas as unidades e áreas comuns.",
         "Linha 32 (gold), alumínio ou PVC, pintura preto fosco, vidro incolor/laminado (caderno de esquadrias · memorial 11).",
         "Contramarco antes do reboco; boneca mínima de 10 cm (Briefing).", S_("esq")),
        ("T06", "Guarda-corpos GC1–GC3", "Terraços do ático e terraços descobertos.",
         "Mureta de alvenaria h=0,80 m + vidro h=1,00 m, perfil \"U\" inox, sem alumínio (Briefing).",
         "DT 06–08 indica perfil Infinity: compatibilizar para \"U\" inox antes do pedido.", S_("gc")),
        ("T07", "Floreiras de fachada", "Bordas de laje dos terraços e floreiras do térreo.",
         "Floreiras em alvenaria/concreto impermeabilizadas com pintura antirraiz (memorial 17.2 e 26.3.3).",
         "Teste de estanqueidade antes do plantio; paisagismo conforme projeto específico.", S_("flor")),
        ("T08", "Araucárias preservadas", "Jardim frontal, em frente ao acesso.",
         "Preservação das duas araucárias existentes (implantação · memorial 26.3.3).",
         "Proteger raízes e copas durante toda a obra; não remover.", S_("arauc"))]),
    ("Texturas da fachada · identidade e acesso (3/3)", [
        ("T09", "Letreiro SEAZONE", "Volume da caixa d'água, frente e lateral.",
         "Letras iluminadas, como na imagem de lançamento (memorial 24: nome do edifício na fachada).",
         "Projetos detalha fixação e ponto elétrico antes do fechamento da fachada.", S_("letreiro")),
        ("T10", "Marquise + BATEL SPOT", "Acesso principal do térreo.",
         "Marquise metálica preta com spots e letras brancas iluminadas BATEL SPOT (render 008).",
         "Executar como na imagem; Projetos detalha estrutura, letras e iluminação.", S_("marquise")),
        ("T11", "Totem casa Seazone", "Jardim frontal, ao lado do acesso.",
         "Totem em forma de casa, iluminado, sobre base preta (renders 001, 004 e 008).",
         "Prever base e ponto elétrico no jardim.", S_("totem")),
        ("T12", "Muro frontal", "Fechamento frontal e lateral do lote.",
         "Muro h=1,80/2,20 m claro com arandelas (imagem), com logo e nome Seazone (Briefing).",
         "Prever pontos das arandelas e da iluminação do logo.", S_("muro"))])]
for title, cards in tex:
    S = B.dup(8)
    S.T(221, title); S.T(225, NOTE)
    for ids, (code, ttl, onde, spec, obra, img) in zip(CARDS, cards):
        card(S, ids, code, ttl, onde, spec, obra, img)

# ================= 11 · VISTA FRONTAL (tpl 11) =================
S = B.dup(11)
S.T(302, "Fachada frontal (Av. Visconde de Guarapuava)")
win = S.P(306, R001)
label_w(S, 307, "LANÇAMENTO · 001 (FRONTAL)")
place_markers(S, 306, win, list(range(308, 317)), old_markers(7))
S.T(318, "LEITURA DA FACHADA FRONTAL", "A fachada frontal é simétrica: faixa central de brises amadeirados entre volumes em tom "
         "bege, lajes marcadas e coroamento com o letreiro SEAZONE. Na base, a marquise com BATEL SPOT, o totem da casa e as "
         "araucárias preservadas formam a identidade do acesso.")
S.T(320, LEG)
V11 = [(321, 322, 323), (324, 325, 326), (327, 328, 329), (330, 331, 332), (333, 334, 335), (336, 337, 338),
       (339, 340, 341), (342, 343, 344), (345, 346, 347)]
vista_items(S, V11, [
    ("Letreiro SEAZONE", "Letras iluminadas no volume da caixa d'água, como no render"),
    ("Guarda-corpo do ático", "Mureta 0,80 m + vidro 1,00 m, perfil \"U\" inox (Briefing)"),
    ("Brises", "Alumínio amadeirado tom amêndoa, como na imagem"),
    ("Floreiras", "Alvenaria/concreto com pintura antirraiz (memorial 17.2)"),
    ("Volumes bege", "R2 pintura Suvinil C164 Ovelha"),
    ("Araucárias", "Preservadas na implantação (memorial 26.3.3)"),
    ("Letreiro BATEL SPOT", "Letras iluminadas sobre marquise metálica preta, como no render"),
    ("Totem casa Seazone", "Totem iluminado no jardim, como no render"),
    ("Muro frontal", "Muro h=1,80/2,20 m, claro com arandelas, logo e nome Seazone")])

# ================= 12 · VISTA LATERAL (tpl 12) =================
S = B.dup(12)
S.T(354, "Esquina e fachada lateral")
win = S.P(358, R002)
label_w(S, 359, "LANÇAMENTO · 002")
place_markers(S, 358, win, list(range(360, 369)), old_markers(8))
S.T(370, LEG)
V12 = [(371, 372, 373), (374, 375, 376), (377, 378, 379), (380, 381, 382), (383, 384, 385),
       (386, 387, 388), (389, 390, 391), (392, 393, 394), (395, 396, 397)]
vista_items(S, V12, [
    ("Letreiro SEAZONE", "Repetido na face lateral do volume da caixa d'água, como no render"),
    ("Vegetação da cobertura", "Floreiras do ático impermeabilizadas com pintura antirraiz"),
    ("Brises", "Alumínio amadeirado tom amêndoa, como na imagem"),
    ("Volume bege", "R2 pintura Suvinil C164 Ovelha"),
    ("Face lateral cinza", "R3 efeito cimento queimado Coral Imensidão Cinza"),
    ("Esquadrias", "Linha 32 preto fosco (caderno de esquadrias)"),
    ("Araucárias", "Preservação na implantação; proteger durante a obra"),
    ("Letreiro BATEL SPOT", "Letras iluminadas sobre a marquise do acesso, como no render")])
S.P(398, S_("R2")); S.P(399, S_("R3")); S.P(400, S_("R1"))
S.T(401, "R2 · R3 · R1")
S.T(403, "NAS QUATRO ELEVAÇÕES", "As elevações PE 11 a 14 repetem os códigos R1 a R4. Seguir as imagens: volumes em bege-rosado "
         "(R2), painéis cinza (R3) e lajes em tom cinza (R1); aprovar amostras antes da compra.")

# ================= 13 · VISTA ACESSO (tpl 13) =================
S = B.dup(13)
S.T(410, "Térreo: acesso, marquise e totem")
win = S.P(414, R008)
label_w(S, 415, "LANÇAMENTO · 008 (ACESSO)")
place_markers(S, 414, win, list(range(416, 426)), old_markers(9))
S.T(427, "ATENÇÃO · IDENTIDADE DO ACESSO", "Marquise, letreiro, totem e iluminação do jardim são a \"cara\" do empreendimento "
         "no lançamento: executar como nas imagens. Projetos detalha estrutura, elétrica e fornecedor antes do fechamento de fachada.")
S.T(429, LEG)
V13 = [(430, 431, 432), (433, 434, 435), (436, 437, 438), (439, 440, 441), (442, 443, 444),
       (445, 446, 447), (448, 449, 450), (451, 452, 453), (454, 455, 456), (457, 458, 459)]
vista_items(S, V13, [
    ("Letreiro BATEL SPOT", "Letras brancas iluminadas sobre a marquise"),
    ("Marquise", "Estrutura metálica preta com spots, como no render"),
    ("Totem casa Seazone", "Casa iluminada sobre base preta, no jardim"),
    ("Porta de acesso", "Perfil metálico + vidro temperado · reconhecimento facial (memorial 15.2)"),
    ("Guarda-corpo da rampa", "Vidro com perfis pretos (render) · corrimão conforme NBR 9050"),
    ("Floreira com arandelas", "Mureta/floreira clara com arandelas de facho duplo, como no render"),
    ("Brises", "Alumínio tom amêndoa sobre os terraços do 2º pavimento"),
    ("Araucárias", "Preservadas no jardim frontal"),
    ("Escada e piso externo", "Piso claro antiderrapante no acesso, como no render")])

# ================= 14 · PONTOS DE ATENÇÃO FACHADA (tpl 24) =================
S = B.dup(24)
S.T(767, "01 · FACHADA · PONTOS DE ATENÇÃO")
S.T(768, "Fachada: pontos de atenção na obra")
S.P(772, R004, (0.6, 0.5))
label_w(S, 773, "LANÇAMENTO · 004")
checklist(S, [
    ("Bordas de laje (R1)", "Seguir a imagem: tom cinza. Compatibilizar o executivo (Branco Gatinho) e aprovar amostra antes da compra."),
    ("Brises (R4)", "Alumínio com acabamento amadeirado tom amêndoa, como na imagem: aprovar amostra antes de fabricar BR1 a BR4."),
    ("Guarda-corpos GC1–GC3", "Vidro com perfil \"U\" inox, sem alumínio (Briefing): compatibilizar DT 06–08 (perfil Infinity)."),
    ("Letreiros, marquise e totem", "Executar como nas imagens; prever fixação e pontos elétricos antes do fechamento da fachada."),
    ("Muro frontal", "Claro com arandelas (imagem), com logo e nome Seazone (Briefing); prever pontos de iluminação."),
    ("Araucárias", "Proteger raízes e copas durante toda a obra; não remover sem autorização.")])

# ================= 15 · DIVISOR 02 =================
S = B.dup(14)
S.P(465, RF(1), (0.55, 0.5))
S.T(468, "Interiores das áreas comuns")
S.T(469, "Conceito Seazone × decisão  ·  recepção e corredores  ·  ático: lounge, copa e jantar  ·  terraço e piscina  ·  pontos de atenção")

# ================= 16 · CONCEITO × DECISÃO (tpl 15) =================
S = B.dup(15)
S.T(477, "02 · INTERIORES DAS ÁREAS COMUNS")
S.T(478, "Conceito Seazone × decisão registrada")
cards15 = [(483, 484, 485, 486), (488, 489, 490, 491), (493, 494, 495, 496), (498, 499, 500, 501), (503, 504, 505, 506)]
data15 = [
    ("Recepção / hall", REC1, (0.6, 0.5),
     "Jardim vertical (treliça preta + plantas) com letreiro Batel Spot; poltronas terracota",
     "Seguir o conceito: sem muxarabi, parede cinza-escura ou poltrona preta/dourada da prévia HS"),
    ("Ático · lounge e copa", RF(1), (0.5, 0.5),
     "Cobogó bege, ilha faceada em cobogó com tampo preto, mesa preta, cadeiras couro caramelo, pendente esfera dourada",
     "Seguir o conceito: sem muxarabi, paredes verdes nem ilha em pedra escura"),
    ("Corredores", RF(0), (0.5, 0.5),
     "Cobogó bege claro com piso amadeirado; shafts embutidos no padrão da vedação",
     "Piso amadeirado do conceito (compatibilizar memorial: Superquadra)"),
    ("Terraço + piscina", RF(4), (0.5, 0.5),
     "Pastilha verde, deck de madeira, poltronas Acapulco pretas, almofadões brancos, guarda-corpo de vidro",
     "Fiel ao projeto Seazone · aprovado nas prévias HS"),
    ("Adega do ático", RF(2), (0.5, 0.5),
     "Espaço fechado, como no render do conceito",
     "Adega fechada · marcenaria em MDF Crômio e Roble Catedral (estilo Bruma)")]
for (pic, ttl, body, pill), (name, img, foc, conc, sit) in zip(cards15, data15):
    S.P(pic, img, foc)
    S.T(ttl, name)
    S.T(body, "CONCEITO SEAZONE (SEGUIR)", conc, "DECISÃO REGISTRADA", sit)
    x, y, w, h = S.pos(pill)
    S.delete(pill)
    S.pill(D_, x, y)
S.T(507, "Documento \"Divergências Interiores (Seazone × HS)\", 27/07/2026: a Seazone já criou o projeto de interiores; o escopo "
         "da HS é transformá-lo em executivo, sem reinterpretar o design.")
S.T(509, "ATENÇÃO · REGRA PARA A OBRA", "Onde a prévia da HS diverge, vale o conceito Seazone (10 - Projeto de Interiores / "
         "00 - Projeto Seazone / 02 - Imagens). Não executar muxarabi, paredes verdes nem ilha em pedra escura.")

# ================= 17 · FICHA RECEPÇÃO (tpl 27) =================
S = B.dup(27)
S.T(848, "02 · INTERIORES DAS ÁREAS COMUNS · RECEPÇÃO E CORREDORES")
S.T(849, "Ficha: recepção, hall e corredores")
S.P(853, REC1, (0.55, 0.5)); label_w(S, 854, "CONCEITO · RECEPÇÃO")
S.P(855, REC2, (0.5, 0.5)); label_w(S, 856, "CONCEITO · JARDIM VERTICAL")
S.T(857, "Corredor: cobogó bege e piso amadeirado (conceito Seazone). Recepção entregue mobiliada e equipada (memorial 27).")
rows(S, [(859, 860), (861, 862), (865, 866), (867, 868), (870, 871), (873, 874), (876, 877), (878, 879), (881, 882)], [
    ("Parede-destaque", "Jardim vertical em treliça metálica preta com plantas + letreiro Batel Spot iluminado (ponto H 1,20–1,40 m)"),
    ("Paredes", "Destaque Cinza Chumbo e demais Crômio (Suvinil, memorial)"),
    ("Piso", "Recepção: Superquadra Concreto Natural 90×90 · corredor: piso amadeirado (conceito; compatibilizar memorial)"),
    ("Teto", "Gesso acartonado, Suvinil Branco Neve · corredores: forro modular 2,60 m (Briefing)"),
    ("Rodapé", "Poliestireno branco h=7 cm, laca PU, quina viva (memorial e Briefing)"),
    ("Mobiliário", "Duas poltronas terracota em couro ecológico com metal preto (conceito e memorial de interiores)"),
    ("Iluminação", "Spots sobrepostos + perfil de LED na parede · corredores: arandela ao lado de cada porta"),
    ("Portas e números", "Porta lisa branca com bandeira até o forro · números soltos pretos com arandela/spot"),
    ("Acesso", "Porta em perfil metálico + vidro temperado · reconhecimento facial Intelbras SS 3532 MF ou similar")],
    [863, 872])

# ================= 18 · MOODBOARD ÁTICO E TERRAÇO (tpl 31) =================
S = B.dup(31)
S.T(977, "02 · INTERIORES DAS ÁREAS COMUNS · ÁTICO · MOODBOARD")
S.T(978, "Ático: lounge, copa, jantar, terraço e piscina")
for pic, img, foc in zip(range(982, 993),
                         [RF(1), RF(3), RF(2), RF(0), RF(4), R003, RF(6), RF(5), RF(7), RF(8), RFP],
                         [(.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.45, .4), (.5, .5), (.5, .5), (.5, .5), (.5, .5), (.5, .5)]):
    S.P(pic, img, foc)
S.T(993, "CONCEITO", "Cobogó bege na divisória e na face da ilha com tampo preto, mesa retangular preta com cadeiras couro "
         "caramelo, sofá Togo e pendente esfera trançada dourada sobre a mesa.",
         "TERRAÇO", "Piscina com revestimento esverdeado, deck de madeira, poltronas Acapulco pretas, almofadões brancos e "
         "guarda-corpo de vidro.",
         "ESTILO", "Bruma: marcenaria em MDF Crômio e Roble Catedral com metal preto.")
for pic, lab, name, img in [(994, 995, "Cobogó", S_("cobogo")), (996, 997, "Vinílico", S_("vinilico")),
                            (998, 999, "Pastilha", S_("pastilha")), (1000, 1001, "Deck", S_("deck"))]:
    S.P(pic, img); S.T(lab, name)

# ================= 19 · FICHA ÁTICO (tpl 30) =================
S = B.dup(30)
S.T(938, "02 · INTERIORES DAS ÁREAS COMUNS · ÁTICO")
S.T(939, "Ficha de premissas: lounge, copa e jantar")
S.P(943, RF(1), (0.5, 0.5)); label_w(S, 944, "CONCEITO SEAZONE · ÁTICO")
S.P(945, RF(3), (0.5, 0.5)); label_w(S, 946, "ILHA FACEADA EM COBOGÓ")
S.T(947, "Indicações 01, 02 e 03 (00 - Premissas de Projeto) marcam o cobogó, a ilha e o trilho de iluminação desta imagem.")
rows(S, [(949, 950), (952, 953), (955, 956), (957, 958), (960, 961), (962, 963), (965, 966), (967, 968), (970, 971)], [
    ("Divisória", "Cobogó bege (conceito e memorial de interiores) · não executar o muxarabi do memorial R01"),
    ("Ilha / bar", "Faceada em cobogó com tampo em granito São Gabriel polido · pontos elétricos e hidráulicos saindo do piso"),
    ("Piso", "Vinílico em réguas Carvalho Natural Claro (memorial R01 e conceito)"),
    ("Paredes", "Gerais Ouro Branco fosco · destaque Cinza Chumbo (memorial)"),
    ("Teto", "Gesso acartonado Branco Neve, pintura setorizada conforme interiores"),
    ("Mobiliário", "Mesa retangular preta, cadeiras couro caramelo, sofá Togo couro preto, puffs (conceito Seazone)"),
    ("Marcenaria", "MDF Berneck Crômio e Roble Catedral com metal preto fosco (estilo Bruma)"),
    ("Iluminação", "Pendente esfera trançada dourada + trilho de spots · interruptor separado para o pendente (Briefing)"),
    ("Adega", "Espaço fechado, como no render do conceito Seazone")],
    [951])

# ================= 20 · FICHA TERRAÇO E PISCINA (tpl 32) =================
S = B.dup(32)
S.T(1007, "02 · INTERIORES DAS ÁREAS COMUNS · TERRAÇO DO ÁTICO")
S.T(1008, "Ficha de premissas: terraço e piscina")
S.P(1012, RF(4), (0.5, 0.5)); label_w(S, 1013, "CONCEITO · PISCINA (APROVADO)")
S.P(1014, RF(6), (0.5, 0.5)); label_w(S, 1015, "CONCEITO · TERRAÇO")
S.T(1016, "Terraço e piscina são o único ambiente aprovado nas prévias HS: avançar para o executivo seguindo estas imagens.")
rows(S, [(1018, 1019), (1021, 1022), (1024, 1025), (1026, 1027), (1029, 1030), (1031, 1032), (1034, 1035), (1037, 1038), (1041, 1042)], [
    ("Revestimento", "Esverdeado 20×20, como nas imagens (Briefing) · compatibilizar memorial (azul-turquesa)"),
    ("Borda / banco", "Porcelanato que imita madeira 120×20 (Briefing), visual de deck da imagem"),
    ("Rejunte e escada", "Epóxi Cinza Platina (Portobello) · escada/barra de apoio em aço inox"),
    ("Aquecimento", "Solar (Briefing) · compatibilizar memorial 20.8"),
    ("Piso do terraço", "Superquadra Concreto Natural 90×90 (memorial) · porcelanato 90×90 antiderrapante"),
    ("Guarda-corpo", "Mureta alvenaria 0,80 m + vidro 1,00 m, perfil \"U\" inox (GC1–GC3)"),
    ("Mobiliário", "Poltronas Acapulco pretas, espreguiçadeiras, almofadões brancos, balanço instalado na laje (Briefing)"),
    ("Pontos", "Logo Seazone iluminado H 1,40 m · tomadas no banco de concreto, junto às mesas e fire places"),
    ("Floreiras", "Impermeabilização com pintura antirraiz e teste de estanqueidade (memorial 17.2)")],
    [1020, 1036, 1039])

# ================= 21 · PONTOS DE ATENÇÃO INTERIORES (tpl 24) =================
S = B.dup(24)
S.T(767, "02 · INTERIORES DAS ÁREAS COMUNS · PONTOS DE ATENÇÃO")
S.T(768, "Áreas comuns: pontos de atenção na obra")
S.P(772, REC2, (0.5, 0.5))
label_w(S, 773, "CONCEITO · RECEPÇÃO")
checklist(S, [
    ("Prévia HS × conceito Seazone", "Vale o conceito: não executar muxarabi, paredes verdes nem ilha em pedra escura."),
    ("Jardim vertical da recepção", "Treliça preta com plantas e letreiro: detalhar fixação, irrigação, impermeabilização e iluminação."),
    ("Piso do corredor", "Amadeirado, como no conceito: compatibilizar o memorial (Superquadra) antes da compra."),
    ("Piscina do ático", "Esverdeado 20×20 (Briefing), rejunte epóxi Cinza Platina; aquecimento solar."),
    ("Deck dos gardens e borda", "Deck de madeira nos gardens fixado sem perfurar a impermeabilização; borda em porcelanato madeira 120×20."),
    ("Marcenaria do ático", "MDF Crômio e Roble Catedral com metal preto (estilo Bruma): compatibilizar o memorial de interiores.")])

# ================= 22 · DIVISOR 03 =================
S = B.dup(33)
S.P(1048, RF(3), (0.4, 0.45))
S.T(1051, "Materiais de interiores das áreas comuns")
S.geom(1052, y=S.pos(1052)[1] + 0.75)
S.T(1052, "Mapa de materialidade T13–T28: amostras tiradas dos renders Seazone e padrões do catálogo Bruma para compra e aprovação")

# ================= 23-26 · MATERIAIS =================
MC = [(1067, 1068, 1069), (1071, 1072, 1073), (1075, 1076, 1077), (1079, 1080, 1081)]
NOTE_M = "Amostras extraídas dos renders de interiores Seazone. Referência visual, não substitui amostra física."
mats = [
    ("Materiais · paredes e pisos (1/4)", NOTE_M, [
        ("T13", "Cobogó bege", "Divisória e face da ilha do ático; corredores (conceito Seazone).",
         "Bloco vazado bege (divergências 27/07/2026) · memorial de interiores: parede em cobogó da ilha.",
         "Não executar o muxarabi do memorial R01: vale o cobogó do conceito.", S_("cobogo")),
        ("T14", "Jardim vertical", "Parede-destaque da recepção, com letreiro Batel Spot.",
         "Treliça metálica preta com plantas (conceito) · memorial de interiores: metal preto fosco na parede.",
         "Detalhar fixação, irrigação, impermeabilização e iluminação.", S_("jardim")),
        ("T15", "Vinílico Carvalho Natural Claro", "Ático interno e corredores (piso amadeirado do conceito).",
         "Piso vinílico em réguas, tonalidade amadeirada clara, fosco (memorial 26) · rodapé branco 7 cm.",
         "Vinílico no ático, como no memorial e na imagem do conceito.", S_("vinilico")),
        ("T16", "Superquadra Concreto Natural", "Recepção, hall e terraço externo.",
         "Porcelanato retificado 90×90 Portobello (memorial 26) · antiderrapante no terraço.",
         "Peças > 40×40 com argamassa no tardoz.", S_("superquadra"))]),
    ("Materiais · pedra, pintura e piscina (2/4)", NOTE_M, [
        ("T17", "Granito São Gabriel", "Bancadas dos WCs comuns e ilha do ático.",
         "Granito São Gabriel polido, espessura conforme detalhamento (memorial 26).",
         "Deixar pontos hidráulicos e elétricos saindo do piso na ilha (Briefing).", S_("saogabriel")),
        ("T18", "Parede escura", "Parede-destaque da recepção e do ático.",
         "Pintura acrílica Suvinil Cinza Chumbo (memorial).",
         "Cinza Chumbo no destaque; aprovar amostra antes da compra.", S_("paredeescura")),
        ("T19", "Pastilha da piscina", "Piscina do ático.",
         "Revestimento esverdeado 20×20 (Briefing e imagem) · rejunte epóxi Cinza Platina.",
         "Compatibilizar memorial (azul-turquesa) e aprovar amostra contra a imagem antes do pedido.", RF(6), (0.45, 0.62), 0.12),
        ("T20", "Deck de madeira", "Borda/banco da piscina e gardens.",
         "Deck de madeira nos gardens (Briefing) · borda em porcelanato que imita madeira 120×20 (Briefing).",
         "Fixar sem perfurar a impermeabilização.", S_("deck"))]),
    ("Materiais · mobiliário e MDFs do ático (3/4)", "Amostras dos renders Seazone e padrões Berneck do catálogo Bruma (rev. 13/10/2026). Referência visual, não substitui amostra física.", [
        ("T21", "Couro caramelo / terracota", "Cadeiras do ático e poltronas da recepção.",
         "Couro caramelo no ático · poltronas terracota em couro ecológico na recepção (conceito Seazone).",
         "Prévia HS trocou por veludo verde e poltrona preta/dourada: seguir o conceito.", S_("couro")),
        ("T22", "Pendente esfera dourada", "Mesa do ático.",
         "Pendente esfera trançada dourada + trilho de spots (conceito · Indicação 03).",
         "Prever interruptor separado para o pendente central (Briefing).", S_("pendente")),
        ("T23", "MDF Crômio", "Marcenaria do ático, em composição com o Roble Catedral.",
         "MDF Berneck Crômio (Linha Smart, Vel), liso, com metal preto fosco.",
         "Compatibilizar memorial de interiores (Frapê); aprovar amostra física antes da compra.", sw("batel_amostra_cromio.png")),
        ("T24", "MDF Roble Catedral", "Marcenaria do ático, em composição com o Crômio.",
         "MDF Berneck Roble Catedral (Grann), amadeirado.",
         "Compatibilizar memorial de interiores (Nogueira Persa); aprovar amostra física antes da compra.", sw("batel_amostra_roble_catedral.png"))]),
]
for title, note, cards in mats:
    S = B.dup(34)
    S.T(1060, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
    S.T(1061, title)
    S.T(1065, note)
    for ids, c in zip(MC, cards):
        card(S, ids, *c)

S = B.dup(37)
S.T(1141, "03 · INTERIORES DAS ÁREAS COMUNS · MAPA DE MATERIALIDADE")
S.T(1142, "Materiais · MDFs e metais (4/4)")
S.T(1146, "Padrões Berneck e metal do catálogo Bruma Seazone (rev. 13/10/2026). Referência visual, não substitui amostra física.")
MD = [(1148, 1149, 1150), (1152, 1153, 1154), (1156, 1157, 1158), (1160, 1161, 1162)]
for ids, (code, ttl, onde, spec, obra, img) in zip(MD, [
        ("T25", "MDF Carvalho Treviso", "Cabeceira, nichos e marcenaria das unidades.",
         "MDF Berneck Carvalho Treviso (Design), amadeirado claro, com metal preto fosco.",
         "Compatibilizar memorial de interiores (Rícori Nativo); aprovar amostra física antes da compra.", sw("batel_amostra_carvalho_treviso.png")),
        ("T26", "MDF Ópera", "Balcão da copa das unidades.",
         "MDF Berneck Ópera (Micro), liso, com Carvalho Treviso.",
         "Compatibilizar memorial de interiores (Grafito Chess); aprovar amostra física antes da compra.", sw("batel_amostra_opera.png")),
        ("T27", "MDF Cinza Cobalto", "Marcenaria das unidades.",
         "MDF Berneck Cinza Cobalto (Tx), liso.",
         "Compatibilizar memorial de interiores (Cinza Cristal); aprovar amostra física antes da compra.", sw("batel_amostra_cinza_cobalto.png")),
        ("T28", "Metal preto fosco", "Marcenaria, jardim vertical da recepção e estruturas de mobiliário.",
         "Metal pintado preto fosco (estilo Bruma).",
         "Mesmo acabamento em todas as peças metálicas; aprovar amostra de cor junto com os MDFs.", sw("batel_amostra_metal_preto.png"))]):
    card(S, ids, code, ttl, onde, spec, obra, img)

# ================= 27 · ENCERRAMENTO =================
S = B.dup(38)
S.P(1168, R007, (0.45, 0.45))
S.T(1170, "BATEL SPOT")
S.T(1171, "Book de Premissas de Fachada e Interiores das Áreas Comuns")
S.geom(1172, y=S.pos(1172)[1] + 0.55)

B.finish(OUT, "BATEL SPOT")
# rodape das aberturas de secao (caixa estreita): fonte menor para caber em 1 linha
P = Presentation(OUT)
for sl in P.slides:
    for sh in sl.shapes:
        if sh.has_text_frame and "SPOT  ·  BOOK DE PREMISSAS" in sh.text_frame.text and sh.width < 5 * E:
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(5.5)
P.save(OUT)
print("ok", OUT)
