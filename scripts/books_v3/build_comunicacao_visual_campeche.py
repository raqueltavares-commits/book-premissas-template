# Builds section "04 · Comunicação visual" of the Campeche Spot book (template v3 layout)
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tplkit import Book, E, set_paras
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from PIL import Image, ImageDraw, ImageFont

HERE = r"C:/Users/Seazone/Claude/seazone/campeche-book-v3"
TPL = r"C:/Users/Seazone/Claude/projetos/book-premissas-template/template/TEMPLATE_v3_CAMPECHE_SPOT_Book_de_Premissas_rev13-10-2026.pptx"
SCR = sys.argv[1]                      # scratch dir with extracted template images (img/sNN_ID.ext)
OUT = os.path.join(HERE, "entregas", "CAMPECHE SPOT - Book de Premissas - 04 Comunicacao visual.pptx")
FIRST_PAGE = 38                        # section enters before the closing slide of the v3 book
IMG = lambda n: os.path.join(SCR, "img", n)
GEN = os.path.join(SCR, "gen"); os.makedirs(GEN, exist_ok=True)
R107 = r"C:/Users/Seazone/Claude/projetos/renders-realistas/recepcao-nova/r2/IMAGEM 107.png"
REF = os.path.join(HERE, "entregas", "numero-unidade-ref.png")

NAVY, CORAL, GREY = (0, 20, 61), (252, 96, 88), (111, 120, 145)
ARIAL, ARIALB = r"C:/Windows/Fonts/arial.ttf", r"C:/Windows/Fonts/arialbd.ttf"
EYEBROW = "04 · COMUNICAÇÃO VISUAL"


def pt(v):
    return v / 72


def geom(S, sid, x=None, y=None, w=None, h=None):
    S.geom(sid, *(None if v is None else pt(v) for v in (x, y, w, h)))


def add_pic(S, path, x, y, w=None, h=None):
    S.s.shapes.add_picture(path, Emu(int(pt(x) * E)), Emu(int(pt(y) * E)),
                           None if w is None else Emu(int(pt(w) * E)), None if h is None else Emu(int(pt(h) * E)))


# ---------- generated drawings ----------
def elevation_png():
    """Unit door elevation to scale (5 px = 1 cm)."""
    s = 5
    W, H = 205 * s, 340 * s
    floor = 300 * s
    X = lambda cm: int(cm * s)
    Y = lambda h_cm: int(floor - h_cm * s)
    im = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(im)
    f_dim = ImageFont.truetype(ARIALB, 30)
    f_lbl = ImageFont.truetype(ARIAL, 26)
    d.rectangle([0, Y(285), W, floor], fill=(242, 244, 248))
    # door: frame 7 cm, leaf 80 x 210
    d.rectangle([X(40), Y(217), X(134), floor], fill="white", outline=(150, 156, 170), width=3)
    d.rectangle([X(47), Y(210), X(127), floor], fill="white", outline=(150, 156, 170), width=2)
    d.rectangle([X(52), Y(105), X(60), Y(103)], fill=(150, 156, 170))
    # number 107, 100 mm, right-aligned with frame, base 40 mm above frame
    fnum = ImageFont.truetype(ARIALB, 70)
    bb = d.textbbox((0, 0), "107", font=fnum)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    d.text((X(134) - tw - bb[0], Y(231) - bb[1]), "107", font=fnum, fill=NAVY)
    # wall lamp Slide: body 1.82 → 2.59 m, 160 mm from frame
    d.rectangle([X(150), Y(259), X(153), Y(182)], fill=(201, 169, 122))
    head = Image.new("RGBA", (X(15), X(6)), (0, 0, 0, 0))
    ImageDraw.Draw(head).rounded_rectangle([0, 0, X(15) - 1, X(6) - 1], radius=10, fill=NAVY)
    head = head.rotate(28, expand=True)
    im.paste(head, (X(141), Y(252)), head)
    d.line([X(151.5), Y(182), X(150), Y(165), X(152), Y(150)], fill=NAVY, width=2)
    d.line([0, floor, W, floor], fill=NAVY, width=4)

    def vdim(xcm, h0, h1, label, side="left"):
        x = X(xcm)
        d.line([x, Y(h0), x, Y(h1)], fill=CORAL, width=3)
        for h in (h0, h1):
            d.line([x - 10, Y(h), x + 10, Y(h)], fill=CORAL, width=3)
        tb = d.textbbox((0, 0), label, font=f_dim)
        ty = (Y(h0) + Y(h1)) // 2 - (tb[3] - tb[1]) // 2
        tx = x - (tb[2] - tb[0]) - 14 if side == "left" else x + 14
        d.text((tx, ty), label, font=f_dim, fill=CORAL)

    def hdim(ycm, x0, x1, label):
        y = Y(ycm)
        d.line([X(x0), y, X(x1), y], fill=CORAL, width=3)
        for x in (x0, x1):
            d.line([X(x), y - 10, X(x), y + 10], fill=CORAL, width=3)
        tb = d.textbbox((0, 0), label, font=f_dim)
        d.text(((X(x0) + X(x1)) // 2 - (tb[2] - tb[0]) // 2, y + 12), label, font=f_dim, fill=CORAL)

    vdim(98, 221, 231, "100 mm")
    vdim(98, 217, 221, "40 mm")
    vdim(14, 0, 221, "2,21 m", "right")
    vdim(36, 0, 210, "2,10 m", "left")
    vdim(176, 0, 182, "1,82 m", "right")
    vdim(176, 182, 259, "0,78 m", "right")
    hdim(175, 134, 150, "160 mm")
    hdim(-12, 47, 127, "folha 80 cm")
    d.text((X(55), Y(150)), "PORTA BRANCA", font=f_lbl, fill=GREY)
    d.text((X(55), Y(143)), "guarnição 7 cm", font=f_lbl, fill=GREY)
    d.text((X(158), Y(272)), "ARANDELA SLIDE", font=f_lbl, fill=GREY)
    # graphic scale 0–50 cm
    d.rectangle([X(4), Y(-24), X(29), Y(-26)], fill=NAVY)
    d.rectangle([X(29), Y(-24), X(54), Y(-26)], outline=NAVY, width=2)
    d.text((X(4), Y(-27)), "0", font=f_lbl, fill=NAVY)
    d.text((X(48), Y(-27)), "50 cm", font=f_lbl, fill=NAVY)
    im = im.crop((0, Y(285), W, Y(-34)))
    out = os.path.join(GEN, "elevacao.png")
    im.save(out)
    return out


def digits_png(name, text, mm, bg=(242, 244, 248), fg=NAVY):
    """Same mm scale on every card (2.4 px/mm) so heights compare."""
    pxmm = 1.8
    W = H = 800
    im = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(im)
    f = ImageFont.truetype(ARIALB, int(mm * pxmm / 0.716))
    bb = d.textbbox((0, 0), text, font=f)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    fl = ImageFont.truetype(ARIALB, 34)
    lw = d.textbbox((0, 0), f"{mm} mm", font=fl)[2]
    x0, y0 = (W - tw + lw + 70) // 2, (H - th) // 2
    d.text((x0 - bb[0], y0 - bb[1]), text, font=f, fill=fg)
    xd = x0 - 50
    dc = (255, 140, 130) if bg != (242, 244, 248) else CORAL
    d.line([xd, y0, xd, y0 + th], fill=dc, width=5)
    for y in (y0, y0 + th):
        d.line([xd - 14, y, xd + 14, y], fill=dc, width=5)
    lb = f"{mm} mm"
    d.text((xd - lw - 20, y0 + th // 2 - 18), lb, font=fl, fill=dc)
    out = os.path.join(GEN, name)
    im.save(out)
    return out


# ---------- slide builders ----------
import time
B = Book(TPL, os.path.join(SCR, f"cache_{int(time.time())}"))


def opener():
    S = B.dup(33)
    S.P(1048, R107, (0.55, 0.42))
    S.T(1050, "04")
    S.T(1051, "Comunicação visual")
    S.T(1052, "Número das unidades, letreiros, sinalização dos pavimentos, rooftop e segurança: "
              "medida, material e posição de cada peça.")


def fachada():
    S = B.dup(13)
    S.T(409, EYEBROW + " · FACHADA E ACESSO")
    S.T(410, "Fachada e acesso: letreiros e identificação")
    for sid, (x, y) in {419: (270, 318), 420: (409, 246), 421: (205, 345)}.items():
        geom(S, sid, x, y)
    S.T(419, "4"); S.T(420, "5"); S.T(421, "6")
    S.delete(422, 423, 424, 425)
    rows = [(430, 431, 432), (433, 434, 435), (436, 437, 438), (439, 440, 441),
            (442, 443, 444), (445, 446, 447), (448, 449, 450), (451, 452, 453), (454, 455, 456), (457, 458, 459)]
    data = [("Logo-casinha", "Retroiluminado LED 3000 K, na pilastra REV-06 escura"),
            ("Letreiro CV-01", "\"Campeche SPOT\" letra caixa em aço inox com LED 3000 K"),
            ("LED da marquise", "Fita LED 3000 K contínua na marquise e floreiras"),
            ("Número predial", "200 mm, aço branco fosco RAL 9003 na pilastra escura"),
            ("Faixa das lojas L1 e L2", "REV-06 contínuo; letreiro do lojista dentro da faixa"),
            ("Capacho de entrada", "Treadmaster Alpha Step 2, 2200 × 497 mm, rebaixo 12 mm"),
            ("Logo \"C\" Seazone na empena", "ACM branco recortado sobre REV-06 (fora desta vista)")]
    for (c, t, k), (a, b) in zip(rows, data):
        S.T(t, a, b)
    for c, t, k in rows[len(data):]:
        S.delete(c, t, k)
    S.T(427, "ATENÇÃO · IDENTIDADE DO ACESSO",
        "Pilastra escura com logo-casinha iluminado e letreiro Campeche SPOT, como na imagem de lançamento. "
        "O executivo mostra pilastra clara e só o texto CV-01: corrigir na errata e detalhar elétrica e fixação. "
        "Excluir o letreiro \"seazone café\" da prancha 028.")


def cards8(eyebrow, title, source, cards):
    S = B.dup(19)
    S.T(603, eyebrow); S.T(604, title); S.T(608, source)
    S.delete(612, 625, 630, 635, 640, 645)
    ids = [(610, 611, 613), (615, 616, 617), (619, 620, 621), (623, 624, 626),
           (628, 629, 631), (633, 634, 636), (638, 639, 641), (643, 644, 646)]
    for (pic, q, txt), (img, focus, hfrac, qtd, cat, name, desc) in zip(ids, cards):
        S.P(pic, img, focus, hfrac)
        S.T(q, f"QTD {qtd}")
        S.T(txt, cat, name, desc)
        S.geom(q, w=pt(30 + 6.2 * len(f"QTD {qtd}")))


def recepcao_corredor():
    cards8(EYEBROW + " · RECEPÇÃO E CORREDORES", "Recepção e corredores: peças de comunicação visual",
           "Fonte: memorial da nova recepção e do corredor PAV 1 (rev. 07/10/2026) e espelho de unidades (66 unidades, PAV 1 a 6). Todos os itens DEFINIDOS.",
           [(IMG("s19_633.jpg"), (0.5, 0.5), None, 1, "ILUMINAÇÃO · RECEPÇÃO", "Letreiro luminoso Campeche SPOT",
             "≈970 × 610 mm, letra caixa com LED 3000 K, sobre a grelha de metalon e o cimento queimado."),
            (IMG("s20_672.jpg"), (0.5, 0.5), None, 4, "DECORAÇÃO · RECEPÇÃO", "Quadros de fotografia praia",
             "500 × 700 × 15 mm, moldura preta, paspatur branco, impressão fine art. Parede de velvet do hall 107."),
            (IMG("s20_676.jpg"), (0.5, 0.5), None, 1, "DECORAÇÃO · RECEPÇÃO", "Quadro grande abstrato",
             "570 × 870 mm, moldura preta fina, arte abstrata azul e areia."),
            (IMG("s19_638.jpg"), (0.5, 0.5), None, 2, "TAPETES · ACESSO", "Capacho de alumínio (entrada)",
             "2200 × 497 mm, rebaixo de 12 mm. Treadmaster Alpha Step 2: alumínio anodizado + carpete preto."),
            (R107, (0.545, 0.26), 0.16, 66, "SINALIZAÇÃO · CORREDOR", "Número da unidade",
             "Aço preto fosco RAL 9005, dígitos de 100 mm, 40 mm acima da guarnição. Detalhe nas próximas páginas."),
            (R107, (0.58, 0.24), 0.28, 66, "ILUMINAÇÃO · CORREDOR", "Arandela Slide",
             "Preta com haste de madeira, luz quente, a 160 mm da guarnição, base a 1,82 m do piso."),
            (R107, (0.5, 0.8), 0.16, 66, "TAPETES · CORREDOR", "Capacho Seazone (portas)",
             "450 × 600 mm, azul com logo Seazone, um por porta."),
            (IMG("s26_837.jpg"), (0.5, 0.45), None, 6, "SINALIZAÇÃO · HALL", "Pavimento e direcional",
             "De frente para o elevador: número do pavimento 250 mm + faixa de unidades com seta, letras de 40 mm.")])


def numero_referencia():
    S = B.dup(26)
    S.T(830, EYEBROW + " · NÚMERO DA UNIDADE")
    S.T(831, "Número da unidade: referência e posição")
    S.T(840, "CONCEITO",
        "Número preto em relevo, sem placa de fundo, acima da porta branca e junto à arandela Slide, como no render do hall 107.",
        "ESPECIFICAÇÃO",
        "Dígitos de 100 mm em aço inox 2 mm, corte a laser, pintura eletrostática preto fosco RAL 9005, fonte Montserrat SemiBold, afastados 10 mm da parede.",
        "SITUAÇÃO REGISTRADA",
        "Porta, batente e rodapé BRANCOS (07/10/2026); o número continua preto. 66 unidades, do PAV 1 ao PAV 6, numeradas como no espelho de vendas.")
    geom(S, 841, 36, 372)
    S.T(842, "Imagem grande: render do hall 107 (porta preta no render; na obra, branca). Abaixo: referência, detalhe do número e corredor de lançamento.")
    geom(S, 842, 36, 400, 245, 50)
    geom(S, 835, 299, 101, 625, 250); S.P(835, R107, (0.6, 0.42), 0.62)
    S.delete(836)
    geom(S, 837, 299, 357, 205, 136); S.P(837, REF, (0.5, 0.5))
    geom(S, 838, 509, 357, 205, 136); S.P(838, R107, (0.55, 0.27), 0.2)
    geom(S, 839, 719, 357, 205, 136); S.P(839, IMG("s26_838.jpg"), (0.4, 0.4), 0.8)


def numero_elevacao():
    S = B.dup(27)
    S.T(848, EYEBROW + " · NÚMERO DA UNIDADE")
    S.T(849, "Número da unidade: elevação cotada e especificação")
    S.delete(853, 854, 855, 856, 863, 872)
    png = elevation_png()
    w, h = Image.open(png).size
    bh = 345; bw = bh * w / h
    add_pic(S, png, 36 + (302 - bw) / 2, 101, bw, bh)
    S.T(857, "Elevação em escala a partir do render do hall 107 (folha 80 × 210 cm, guarnição 7 cm). Medidas do piso acabado.")
    geom(S, 857, 36, 452, 302, 42)
    rows = [(859, 860), (861, 862), (865, 866), (867, 868), (870, 871), (873, 874), (876, 877), (878, 879), (881, 882)]
    data = [("Tipo", "Dígitos avulsos recortados, em relevo, sem placa de fundo"),
            ("Material", "Aço inox 2 mm, corte a laser, pintura eletrostática preto fosco RAL 9005"),
            ("Tamanho", "Dígitos de 100 mm de altura · número de 3 dígitos com ≈ 210 mm"),
            ("Fonte", "Montserrat SemiBold"),
            ("Fixação", "2 pinos roscados por dígito, afastamento 10 mm, furo 4 mm com adesivo PU · gabarito de furação em papel"),
            ("Posição", "Base 40 mm acima da guarnição (2,21 m do piso), alinhado pela direita com a guarnição, do lado da arandela"),
            ("Arandela", "Slide preta a 160 mm da guarnição · base 1,82 m e topo 2,59 m do piso"),
            ("Porta", "Folha 80 × 210 cm, guarnição 7 cm · folha, batente e rodapé BRANCOS"),
            ("Quantidade", "66 números (PAV 1 a 6) + L1 e L2 · 3 dígitos de reposição de cada algarismo")]
    for (a, b), (l, t) in zip(rows, data):
        S.T(a, l); S.T(b, t)
    geom(S, 862, w=431); geom(S, 871, w=431)


def familia():
    S = B.dup(21)
    S.T(697, EYEBROW + " · FAMÍLIA DE NÚMEROS")
    S.T(698, "Família de números: um desenho, quatro alturas")
    S.T(702, "Mesmo desenho em todo o prédio: aço inox 2 mm, pintura eletrostática preto fosco RAL 9005, Montserrat SemiBold, afastado 10 mm da parede. "
              "Muda só a altura. Imagens na mesma escala.")
    S.delete(710, 720, 721, 722, 723, 724)
    cards = [((703, 704, 705, 706), digits_png("f100.png", "107", 100), 68, "UNIDADES E LOJAS", "100 mm",
              "Número de cada porta, 40 mm acima da guarnição. 66 unidades + L1 e L2."),
             ((707, 708, 709, 711), digits_png("f150.png", "3", 150), 8, "ESCADA", "150 mm",
              "Número do pavimento em cada patamar, com face fotoluminescente."),
             ((712, 713, 714, 715), digits_png("f200.png", "75", 200, (74, 74, 72), (255, 255, 255)), 1, "NÚMERO PREDIAL", "200 mm",
              "Na pilastra do acesso, abaixo do letreiro. Branco fosco RAL 9003 por estar sobre a pilastra escura REV-06."),
             ((716, 717, 718, 719), digits_png("f250.png", "3", 250), 6, "HALL DO ELEVADOR", "250 mm",
              "Número do pavimento de frente para o elevador, centro a 1,60 m do piso.")]
    w = (888 - 3 * 16) / 4
    for i, ((bg, pic, q, txt), img, qtd, cat, name, desc) in enumerate(cards):
        x = 36 + i * (w + 16)
        geom(S, bg, x, 115, w, 378)
        geom(S, pic, x, 115, w, w); S.P(pic, img)
        geom(S, q, x + 6, 121)
        geom(S, txt, x + 8, 115 + w + 6, w - 16, 378 - w - 10)
        S.T(q, f"QTD {qtd}"); S.T(txt, cat, name, desc)


def rooftop():
    cards8(EYEBROW + " · ROOFTOP", "Rooftop: frase, logo wall, ambientes e piscina",
           "Fonte: renders do rooftop, Memorial Descritivo R01 (Pedra Canjiquinha com a frase) e diretrizes Spot (logo wall, ativação de marca). Todos os itens DEFINIDOS.",
           [(IMG("s28_901.jpg"), (0.32, 0.25), 0.45, 1, "SINALIZAÇÃO · COPA", "Frase \"Seu lugar fora de casa\"",
             "Letra caixa em aço preto fosco, 100 mm, afastada 20 mm com LED 3000 K atrás, sobre a Pedra Canjiquinha."),
            (IMG("s29_927.jpg"), (0.5, 0.4), 0.6, 1, "ILUMINAÇÃO · HALL", "Logo wall SPOT",
             "Letreiro SPOT retroiluminado LED 3000 K na parede amadeirada, com o balanço à frente."),
            (IMG("s28_907.jpg"), (0.5, 0.3), 0.5, 1, "SINALIZAÇÃO · MINI MARKET", "Mini market",
             "Letras recortadas em aço preto fosco, 60 mm, sobre o painel ripado."),
            (IMG("s28_899.jpg"), (0.68, 0.45), 0.7, 1, "SINALIZAÇÃO · HALL", "Pavimento ROOFTOP",
             "Palavra ROOFTOP em letras de 120 mm, aço preto fosco, de frente para o elevador."),
            (IMG("s29_929.jpg"), (0.5, 0.45), None, 2, "SINALIZAÇÃO · LAVABOS", "Placa de lavabo",
             "Aço preto fosco 150 × 150 mm, pictograma + texto 25 mm, a 1,50 m do piso. Rooftop e recepção."),
            (IMG("s31_988.jpg"), (0.5, 0.55), None, 1, "SEGURANÇA · PISCINA", "Regras da piscina",
             "ACM preto fosco 600 × 400 mm, texto em vinil branco, no acesso à piscina, a 1,50 m do piso."),
            (IMG("s31_987.jpg"), (0.5, 0.6), None, 2, "SEGURANÇA · PISCINA", "Profundidade da piscina",
             "Pastilha contrastante na borda indicando a profundidade, nas duas extremidades."),
            (IMG("s31_989.jpg"), (0.5, 0.6), None, 1, "ATIVAÇÃO DE MARCA", "Kit Seazone do deck",
             "Capas de espreguiçadeira e puffs navy, guarda-sóis, fâmula vertical no deck, baldes de gelo com logo.")])


def seguranca():
    S = B.dup(27)
    S.T(848, EYEBROW + " · SEGURANÇA E REGRAS GERAIS")
    S.T(849, "Segurança, acessibilidade e regras gerais")
    S.P(853, IMG("s28_899.jpg"), (0.6, 0.45)); S.T(854, "HALL DO ELEVADOR")
    S.P(855, IMG("s26_837.jpg"), (0.5, 0.5)); S.T(856, "CORREDOR")
    S.T(857, "Placas obrigatórias seguem a norma e o PPCI aprovado no CBMSC. O acabamento preto fosco só vale onde a norma permite.")
    S.delete(863, 872)
    rows = [(859, 860), (861, 862), (865, 866), (867, 868), (870, 871), (873, 874), (876, 877), (878, 879), (881, 882)]
    data = [("Emergência", "Placas fotoluminescentes de saída, rota de fuga, extintor, hidrante e alarme, nas posições do PPCI aprovado"),
            ("Escada", "Número do pavimento 150 mm com face fotoluminescente em cada patamar (8 un)"),
            ("Elevador", "\"Em caso de incêndio não use o elevador\" ao lado da botoeira, em cada parada (8 un)"),
            ("Áreas técnicas", "\"Acesso restrito\" em aço preto fosco 150 × 150 mm: casa de máquinas, medidores, lixeira e DML (4 un)"),
            ("Piso tátil", "Alerta e direcional no acesso e em frente ao elevador, conforme projeto de acessibilidade"),
            ("Família", "Toda a sinalização decorativa em aço preto fosco RAL 9005, Montserrat SemiBold, afastada 10 mm da parede"),
            ("Iluminação", "Letreiros, logos e frase em LED 3000 K, sem ponto de luz aparente"),
            ("Exceção de cor", "Sobre fundo escuro (pilastra REV-06), mesma peça em branco fosco RAL 9003"),
            ("Fornecedor", "Entrega o executivo de comunicação visual (vistas, gabaritos e fixação) para aprovação antes da produção")]
    for (a, b), (l, t) in zip(rows, data):
        S.T(a, l); S.T(b, t)
    geom(S, 862, w=431); geom(S, 871, w=431)


opener()
fachada()
recepcao_corredor()
numero_referencia()
numero_elevacao()
familia()
rooftop()
seguranca()
B.finish(OUT, "CAMPECHE SPOT")

# page numbers continue the v3 book
from pptx import Presentation
p = Presentation(OUT)
for n, sl in enumerate(p.slides, FIRST_PAGE):
    for sh in sl.shapes:
        if sh.has_text_frame and "BOOK DE PREMISSAS DE FACHADA E INTERIORES" in sh.text_frame.text:
            set_paras(sh, [f"CAMPECHE SPOT  ·  BOOK DE PREMISSAS DE FACHADA E INTERIORES  ·  {n:02d}"])
p.save(OUT)
print(OUT)
