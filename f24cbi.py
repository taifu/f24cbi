#!/usr/bin/env python3
"""Converte deleghe F24 in PDF in un flusso CBI-F24 (F4-EF, standard CBI-F24-001)
importabile nel remote banking.

Legge i PDF del modello F24 ordinario dell'Agenzia delle Entrate: quelli del
gestionale, il modello AdE compilato e le rate Sogei. Tutti usano la stessa
griglia, quindi i campi si leggono per posizione sulla pagina.

Il PDF si legge con pdfminer, che per ogni carattere dice anche il font: il
testo prestampato del modulo usa i font in FONT_MODULO, i dati compilati quelli
in FONT_DATI. Un font che non sta in nessuno dei due elenchi ferma lo script.

Ogni dato compilato deve finire in un campo gestito o in una zona ignorata di
proposito (ZONE_IGNORATE). Se resta testo compilato non letto, lo script si
ferma e dice dove: meglio un errore che un versamento sbagliato. Si ferma anche
per righe senza importo, importi illeggibili o totali stampati che non tornano.

Sezioni gestite: Erario, INPS, Regioni, IMU e altri tributi locali, INAIL.
Nella sezione IMU le caselle ravvedimento/immobili variati/acconto/saldo, il
numero immobili, la detrazione e l'identificativo operazione non sono letti:
se sono compilati lo script si ferma.

Per default e' una prova: stampa il flusso e non scrive nulla.
Con --commit scrive il file.
"""

import argparse
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field, replace
from datetime import date, datetime
from pathlib import Path

RECORD_LEN = 120
IBAN_DEFAULT = "IT10L0840410802000000240306"
SIA_DEFAULT = "CJBA1"

FONT_MODULO = ("Futura", "Times", "AvantGarde", "CourierNewPS")
FONT_DATI = ("Verdana", "DejaVuSansMono", "DejaVuSansCondensed", "DejaVuSerif", "Helvetica")

# Griglia del modello F24 (punti PDF, origine in alto a sinistra).
Y_CF = (100, 118)
Y_ANAGRAFICA = (125, 142)
Y_NASCITA = (150, 166)
Y_DOMICILIO = (173, 190)
Y_CF2 = (192, 215)
Y_ERARIO = (237, 310)
Y_ERARIO_TOT = (310, 322)
Y_INPS = (345, 395)
Y_INPS_TOT = (395, 406)
Y_REGIONI = (429, 479)
Y_REGIONI_TOT = (479, 490)
Y_IMU = (513, 563)
Y_IMU_TOT = (563, 575)
Y_INAIL = (596, 634)
Y_INAIL_TOT = (634, 646)
Y_SALDO_FINALE = (703, 718)
# Righe TOTALE delle sezioni non gestite: vi si accetta solo il segno "+" stampato a vuoto.
Y_TOT_NON_GESTITI = [(680, 694)]

X_DEBITO = (340, 403)
X_CREDITO = (420, 488)
X_SALDO = (500, 580)
X_SEGNO = (478, 500)

# Zone il cui contenuto non serve al flusso CBI: (banda y, banda x).
ZONE_IGNORATE = {
    "intestazione del programma": ((0, 45), (0, 600)),
    "banca delegata": ((45, 95), (230, 600)),
    "firma": ((695, 725), (0, 330)),
    "estremi del versamento": ((725, 900), (0, 600)),
}

# Per dire in quale parte del modulo sta un testo non letto: (y fino a, nome).
ZONE_MODULO = [
    (100, "intestazione"), (215, "contribuente"), (322, "sezione Erario"),
    (406, "sezione INPS"), (490, "sezione Regioni"), (577, "sezione IMU e altri tributi locali"),
    (648, "sezione INAIL"), (696, "sezione Altri enti previdenziali"), (900, "saldo finale/firma"),
]

ROW_TOLERANCE = 3.0
LINE_TOLERANCE = 2.0
WORD_GAP = 1.0
SPACE_GAP_SOGEI = 4.2

SEZIONI_MAX = {"erario": 6, "inps": 4, "regioni": 4, "imu": 4, "inail": 3}


class F24Error(Exception):
    pass


@dataclass
class Word:
    x0: float
    y0: float
    x1: float
    text: str
    used: bool = False


@dataclass
class Erario:
    tributo: str
    rateazione: str
    anno: str
    debito: int
    credito: int


@dataclass
class Inps:
    sede: str
    causale: str
    matricola: str
    da: str
    a: str
    debito: int
    credito: int


@dataclass
class Regione:
    regione: str
    tributo: str
    rateazione: str
    anno: str
    debito: int
    credito: int


@dataclass
class Imu:
    ente: str
    tributo: str
    rateazione: str
    anno: str
    debito: int
    credito: int


@dataclass
class Inail:
    sede: str
    ditta: str
    cc: str
    riferimento: str
    causale: str
    debito: int
    credito: int


@dataclass
class Delega:
    origine: str
    cf: str
    cognome: str
    nome: str
    sesso: str = ""
    nascita_data: str = ""
    nascita_comune: str = ""
    nascita_prov: str = ""
    comune: str = ""
    prov: str = ""
    indirizzo: str = ""
    cf2: str = ""
    codice_identificativo: str = ""
    codice_ufficio: str = ""
    codice_atto: str = ""
    data_pagamento: date | None = None
    erario: list[Erario] = field(default_factory=list)
    inps: list[Inps] = field(default_factory=list)
    regioni: list[Regione] = field(default_factory=list)
    imu: list[Imu] = field(default_factory=list)
    inail: list[Inail] = field(default_factory=list)

    def chiave(self):
        return repr(replace(self, origine=""))

    @property
    def righe(self):
        return self.erario + self.inps + self.regioni + self.imu + self.inail

    @property
    def tot_debito(self):
        return sum(r.debito for r in self.righe)

    @property
    def tot_credito(self):
        return sum(r.credito for r in self.righe)

    @property
    def saldo(self):
        return self.tot_debito - self.tot_credito


def read_pages(pdf: Path) -> list[tuple[list[Word], bool]]:
    """Per ogni pagina: le parole compilate e se la pagina ha il modulo prestampato come testo."""
    try:
        from pdfminer.high_level import extract_pages
        from pdfminer.layout import LTChar
    except ImportError:
        raise F24Error("manca pdfminer: apt install python3-pdfminer (oppure pip install pdfminer.six)")

    def chars(obj):
        if isinstance(obj, LTChar):
            yield obj
        elif hasattr(obj, "__iter__"):
            for child in obj:
                yield from chars(child)

    pages = []
    try:
        layouts = list(extract_pages(str(pdf)))
    except Exception as e:
        raise F24Error(f"{pdf.name}: PDF illeggibile ({e})")
    for n, layout in enumerate(layouts, 1):
        dati, modulo = [], False
        for c in chars(layout):
            t = c.get_text()
            if not t.strip():
                continue
            font = c.fontname.split("+", 1)[-1]
            if font.startswith(FONT_MODULO):
                modulo = True
            elif font.startswith(FONT_DATI):
                dati.append((layout.height - c.y1, c.x0, c.x1, t))
            else:
                raise F24Error(f"{pdf.name} p.{n}: font sconosciuto {font!r} (testo {t!r}): "
                               "non so se e' modulo o dato, va aggiunto a FONT_MODULO o FONT_DATI")
        pages.append((build_words(dati), modulo))
    return pages


def build_words(chars) -> list[Word]:
    lines = []
    for ch in sorted(chars):
        if lines and ch[0] - lines[-1][0][0] <= LINE_TOLERANCE:
            lines[-1].append(ch)
        else:
            lines.append([ch])
    words = []
    for line in lines:
        top = min(ch[0] for ch in line)
        for y, x0, x1, t in sorted(line, key=lambda ch: ch[1]):
            if words and words[-1].y0 == top and x0 - words[-1].x1 <= WORD_GAP:
                words[-1].text += t
                words[-1].x1 = max(words[-1].x1, x1)
            else:
                words.append(Word(x0, top, x1, t))
    return words


def pick(words, y, x=(0, 1000)):
    ws = sorted((w for w in words if y[0] <= w.y0 < y[1] and x[0] <= w.x0 < x[1]),
                key=lambda w: w.x0)
    for w in ws:
        w.used = True
    return ws


def text(words, y, x=(0, 1000), spaces=True, gap=0.0):
    ws = pick(words, y, x)
    if not spaces:
        return "".join(w.text for w in ws)
    out, last = "", None
    for w in ws:
        if last is not None and w.x0 - last > gap:
            out += " "
        out += w.text
        last = w.x1
    return out.strip()


def rows(words, y):
    """Raggruppa in righe le parole compilate nella banda y."""
    band = sorted((w for w in words if y[0] <= w.y0 < y[1]), key=lambda w: w.y0)
    groups = []
    for w in band:
        if groups and w.y0 - groups[-1][0] <= ROW_TOLERANCE:
            groups[-1][1].append(w)
        else:
            groups.append((w.y0, [w]))
    return [(ws, (y0 - ROW_TOLERANCE, y0 + ROW_TOLERANCE)) for y0, ws in groups]


def amount(words, y, col) -> int | None:
    """Importo in centesimi, None se il campo e' vuoto.

    Il gestionale scrive parte intera e decimali come due testi ai lati della
    virgola prestampata ('4.441' '07'); Sogei scrive '2390,18'.
    """
    ws = pick(words, y, col)
    if not ws:
        return None
    raw = "".join(w.text for w in ws)
    if "," not in raw and len(ws) == 2:
        raw = f"{ws[0].text},{ws[1].text}"
    m = re.fullmatch(r"(\d{1,3}(?:\.\d{3})+|\d+),(\d{2})", raw)
    if not m:
        raise F24Error(f"importo illeggibile: {raw!r}")
    return int(m.group(1).replace(".", "")) * 100 + int(m.group(2))


def importi_riga(ws, r, sezione):
    debito, credito = amount(ws, r, X_DEBITO), amount(ws, r, X_CREDITO)
    if not debito and not credito:
        raise F24Error(f"{sezione}: riga senza importo")
    return debito or 0, credito or 0


def saldo_con_segno(words, y) -> int | None:
    segno = text(words, y, X_SEGNO, spaces=False)
    if segno not in ("", "+", "-", "\u2013"):
        raise F24Error(f"segno del saldo illeggibile: {segno!r}")
    saldo = amount(words, y, X_SALDO)
    if saldo is not None and segno in ("-", "\u2013"):
        saldo = -saldo
    return saldo


def leggi_totali(words, y):
    return amount(words, y, X_DEBITO), amount(words, y, X_CREDITO), saldo_con_segno(words, y)


def confronta_totali(letti, righe, sezione):
    attesi = (sum(r.debito for r in righe), sum(r.credito for r in righe))
    attesi += (attesi[0] - attesi[1],)
    for letto, atteso, nome in zip(letti, attesi, ("debito", "credito", "saldo")):
        if letto is not None and letto != atteso:
            raise F24Error(f"{sezione}: totale {nome} stampato {euro(letto)}, righe lette {euro(atteso)}")


def digits(s):
    return re.sub(r"\D", "", s)


def zona(w):
    return next(nome for limite, nome in ZONE_MODULO if w.y0 < limite)


def parse_delega(words, modulo, origine) -> Delega | None:
    if not words:
        return None
    # Le rate Sogei non hanno il modulo come testo e spaziano le lettere.
    gap = 0.0 if modulo else SPACE_GAP_SOGEI
    cf = text(words, Y_CF, (110, 460), spaces=False).upper()
    if not re.fullmatch(r"[A-Z0-9]{11}|[A-Z0-9]{16}", cf):
        raise F24Error(f"codice fiscale del contribuente illeggibile: {cf!r}")
    d = Delega(origine=origine, cf=cf,
               cognome=text(words, Y_ANAGRAFICA, (110, 410), gap=gap),
               nome=text(words, Y_ANAGRAFICA, (410, 580), gap=gap))
    nascita = digits(text(words, Y_NASCITA, (110, 230), spaces=False))
    if nascita:
        if len(nascita) != 8:
            raise F24Error(f"data di nascita illeggibile: {nascita!r}")
        d.nascita_data = nascita[4:] + nascita[2:4] + nascita[:2]
    d.sesso = text(words, Y_NASCITA, (235, 262), spaces=False)
    d.nascita_comune = text(words, Y_NASCITA, (262, 530), gap=gap)
    d.nascita_prov = text(words, Y_NASCITA, (530, 580), spaces=False)
    d.comune = text(words, Y_DOMICILIO, (110, 328), gap=gap)
    d.prov = text(words, Y_DOMICILIO, (328, 362), spaces=False)
    d.indirizzo = text(words, Y_DOMICILIO, (362, 580), gap=gap)
    d.cf2 = text(words, Y_CF2, (180, 480), spaces=False).upper()
    d.codice_identificativo = text(words, Y_CF2, (535, 580), spaces=False)
    if d.cf2 and not re.fullmatch(r"[A-Z0-9]{11}|[A-Z0-9]{16}", d.cf2):
        raise F24Error(f"secondo codice fiscale illeggibile: {d.cf2!r}")
    if d.cf2 or d.codice_identificativo:
        # Il flag firmatario del record 50-01 dipende dal tipo di coobbligato: da verificare
        # sul tracciato prima di gestirlo.
        raise F24Error(f"coobbligato {d.cf2!r} con codice identificativo {d.codice_identificativo!r}: "
                       "non ancora gestito")

    for ws, r in rows(words, Y_ERARIO):
        tributo = text(ws, r, (150, 215), spaces=False)
        rateazione = text(ws, r, (215, 272), spaces=False)
        anno = text(ws, r, (272, 315), spaces=False)
        d.erario.append(Erario(tributo, rateazione, anno, *importi_riga(ws, r, "sezione Erario")))
    d.codice_ufficio = text(words, Y_ERARIO_TOT, (10, 72), spaces=False)
    d.codice_atto = text(words, Y_ERARIO_TOT, (72, 250), spaces=False)
    if d.codice_atto and not d.codice_atto.isdigit():
        raise F24Error(f"codice atto illeggibile: {d.codice_atto!r}")
    totali = [(leggi_totali(words, Y_ERARIO_TOT), d.erario, "sezione Erario")]

    for ws, r in rows(words, Y_INPS):
        sede = text(ws, r, (10, 50), spaces=False)
        causale = text(ws, r, (50, 88), spaces=False)
        matricola = text(ws, r, (88, 215), gap=gap)
        da = digits(text(ws, r, (215, 266), spaces=False))
        a = digits(text(ws, r, (266, 318), spaces=False))
        d.inps.append(Inps(sede, causale, matricola, da, a, *importi_riga(ws, r, "sezione INPS")))
    totali.append((leggi_totali(words, Y_INPS_TOT), d.inps, "sezione INPS"))

    for ws, r in rows(words, Y_REGIONI):
        regione = digits(text(ws, r, (10, 60), spaces=False))
        tributo = text(ws, r, (150, 215), spaces=False)
        rateazione = text(ws, r, (215, 272), spaces=False)
        anno = text(ws, r, (272, 315), spaces=False)
        d.regioni.append(Regione(regione, tributo, rateazione, anno,
                                 *importi_riga(ws, r, "sezione Regioni")))
    totali.append((leggi_totali(words, Y_REGIONI_TOT), d.regioni, "sezione Regioni"))

    for ws, r in rows(words, Y_IMU):
        ente = text(ws, r, (10, 70), spaces=False)
        tributo = text(ws, r, (150, 215), spaces=False)
        rateazione = text(ws, r, (215, 272), spaces=False)
        anno = text(ws, r, (272, 315), spaces=False)
        d.imu.append(Imu(ente, tributo, rateazione, anno,
                         *importi_riga(ws, r, "sezione IMU e altri tributi locali")))
    totali.append((leggi_totali(words, Y_IMU_TOT), d.imu, "sezione IMU e altri tributi locali"))

    for ws, r in rows(words, Y_INAIL):
        sede = text(ws, r, (90, 135), spaces=False)
        ditta = text(ws, r, (135, 200), spaces=False)
        cc = text(ws, r, (200, 230), spaces=False)
        riferimento = text(ws, r, (230, 280), spaces=False)
        causale = text(ws, r, (280, 330), spaces=False)
        d.inail.append(Inail(sede, ditta, cc, riferimento, causale,
                             *importi_riga(ws, r, "sezione INAIL")))
    totali.append((leggi_totali(words, Y_INAIL_TOT), d.inail, "sezione INAIL"))

    for nome, limite in SEZIONI_MAX.items():
        if len(getattr(d, nome)) > limite:
            raise F24Error(f"sezione {nome}: {len(getattr(d, nome))} righe, massimo {limite}")

    saldo_finale = saldo_con_segno(words, Y_SALDO_FINALE)
    for y in Y_TOT_NON_GESTITI:
        for w in [w for w in words if y[0] <= w.y0 < y[1] and X_SEGNO[0] <= w.x0 < X_SEGNO[1]]:
            if w.text == "+":
                w.used = True

    top = text(words, ZONE_IGNORATE["intestazione del programma"][0], gap=gap)
    m = re.search(r"Scadenza(?:\s+rata|\s+del)?\s*:?\s*(\d{2})[/-](\d{2})[/-](\d{4})", top, re.I)
    if m:
        d.data_pagamento = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))

    for y, x in ZONE_IGNORATE.values():
        pick(words, y, x)
    resto = [w for w in words if not w.used]
    if resto:
        dettaglio = "; ".join(f"{zona(w)}: {w.text!r} (x={w.x0:.0f}, y={w.y0:.0f})" for w in resto)
        raise F24Error(f"testo compilato in campi che lo script non gestisce: {dettaglio}")

    if saldo_finale is None and all(v is None for letti, _, _ in totali for v in letti):
        raise F24Error("nessun totale stampato con cui verificare gli importi letti")
    for letti, righe, sezione in totali:
        confronta_totali(letti, righe, sezione)
    if saldo_finale is not None and saldo_finale != d.saldo:
        raise F24Error(f"saldo finale stampato {euro(saldo_finale)}, righe lette {euro(d.saldo)}")
    return d


def parse_pdf(pdf: Path) -> list[Delega]:
    deleghe, viste = [], set()
    for n, (words, modulo) in enumerate(read_pages(pdf), 1):
        origine = f"{pdf.name} p.{n}"
        try:
            d = parse_delega(words, modulo, origine)
        except F24Error as e:
            raise F24Error(f"{origine}: {e}")
        if d is None or d.chiave() in viste:
            continue
        viste.add(d.chiave())
        if d.data_pagamento is None:
            m = re.match(r"(\d{4}-\d{2}-\d{2})", pdf.name)
            if m:
                d.data_pagamento = date.fromisoformat(m.group(1))
        deleghe.append(d)
    if not deleghe:
        raise F24Error(f"{pdf.name}: nessuna delega F24 riconosciuta")
    return deleghe


def cbi_text(s, n):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper()
    s = re.sub(r"[^A-Z0-9 .,'/\-]", " ", s)
    return s[:n].ljust(n)


def num(v, n):
    s = str(v)
    if not s.isdigit() or len(s) > n:
        raise F24Error(f"valore numerico non valido per un campo di {n} cifre: {v!r}")
    return s.zfill(n)


class Record:
    def __init__(self):
        self.buf = [" "] * RECORD_LEN

    def set(self, pos, value):
        start = pos - 1
        for i, ch in enumerate(value):
            if self.buf[start + i] != " ":
                raise AssertionError(f"campo sovrapposto in posizione {pos + i}")
            self.buf[start + i] = ch
        return self

    def __str__(self):
        s = "".join(self.buf)
        assert len(s) == RECORD_LEN
        return s


def riferimento(rateazione, anno):
    if not re.fullmatch(r"(\d{4})?", rateazione) or not re.fullmatch(r"(\d{4})?", anno):
        raise F24Error(f"rateazione {rateazione!r} o anno {anno!r} illeggibili")
    return (rateazione or "0000") + (anno or "0000")


@dataclass
class Banca:
    sia: str
    iban: str
    cf_titolare: str
    cf_mittente: str

    @property
    def abi(self):
        return self.iban[5:10]

    @property
    def cab(self):
        return self.iban[10:15]

    @property
    def conto(self):
        return self.iban[15:27]

    @property
    def cin(self):
        return self.iban[4]


def build_flusso(deleghe, banca: Banca, oggi: date, nome_supporto: str) -> list[str]:
    head = (Record().set(2, "F4").set(4, cbi_text(banca.sia, 5)).set(9, num(banca.abi, 5))
            .set(14, oggi.strftime("%d%m%y")).set(20, cbi_text(nome_supporto, 20))
            .set(105, "2$").set(107, num(banca.abi, 5)).set(114, "E"))
    out = [str(head)]

    for i, d in enumerate(deleghe, 1):
        prog = num(i, 7)
        persona_fisica = len(d.cf) == 16
        data_pag = d.data_pagamento.strftime("%Y%m%d")

        r10 = (Record().set(2, "10").set(4, prog).set(11, cbi_text(d.cf, 16))
               .set(27, cbi_text(d.cognome, 24)).set(107, prog))
        if persona_fisica:
            r10.set(51, cbi_text(d.nome, 20)).set(71, cbi_text(d.sesso, 1))
            r10.set(72, cbi_text(d.nascita_comune, 25)).set(97, cbi_text(d.nascita_prov, 2))
            r10.set(99, num(d.nascita_data, 8))
        else:
            resto = d.cognome[24:] + ((" " + d.nome) if d.nome else "")
            if resto.strip():
                r10.set(51, cbi_text(resto.strip(), 20))
        out.append(str(r10))

        r20 = (Record().set(2, "20").set(4, prog).set(11, cbi_text(d.comune, 25))
               .set(36, cbi_text(d.prov, 2)).set(38, cbi_text(d.indirizzo, 35))
               .set(73, data_pag).set(81, "0"))
        if d.cf2:
            r20.set(82, cbi_text(d.cf2, 16)).set(98, cbi_text(d.codice_identificativo, 2))
        out.append(str(r20))

        if d.erario:
            for j, t in enumerate(d.erario, 1):
                r = (Record().set(2, "40").set(4, prog).set(11, "01").set(13, num(j, 2))
                     .set(15, cbi_text(t.tributo, 4)).set(19, cbi_text(riferimento(t.rateazione, t.anno), 8))
                     .set(27, num(t.debito, 15)).set(42, num(t.credito, 15)))
                if d.codice_ufficio:
                    r.set(57, cbi_text(d.codice_ufficio, 3))
                if d.codice_atto:
                    r.set(60, num(d.codice_atto, 11))
                out.append(str(r))
            out.append(saldo_record(prog, "02", 13, d.erario))

        if d.inps:
            for j, t in enumerate(d.inps, 1):
                r = (Record().set(2, "40").set(4, prog).set(11, "03").set(13, num(j, 2))
                     .set(15, num(t.sede, 4)).set(19, cbi_text(t.causale, 4))
                     .set(23, cbi_text(t.matricola, 17)).set(40, num(t.da or 0, 6))
                     .set(52, num(t.debito, 15)).set(67, num(t.credito, 15)))
                if t.a:
                    r.set(46, num(t.a, 6))
                out.append(str(r))
            out.append(saldo_record(prog, "04", 13, d.inps))

        if d.regioni:
            for j, t in enumerate(d.regioni, 1):
                out.append(str(
                    Record().set(2, "40").set(4, prog).set(11, "05").set(13, num(t.regione, 2))
                    .set(15, num(j, 2)).set(17, cbi_text(t.tributo, 4))
                    .set(21, cbi_text(riferimento(t.rateazione, t.anno), 8))
                    .set(29, num(t.debito, 15)).set(44, num(t.credito, 15))))
            out.append(saldo_record(prog, "06", 15, d.regioni))

        if d.imu:
            for j, t in enumerate(d.imu, 1):
                out.append(str(
                    Record().set(2, "40").set(4, prog).set(11, "07").set(13, cbi_text(t.ente, 4))
                    .set(17, num(j, 2)).set(19, cbi_text(t.tributo, 4))
                    .set(23, cbi_text(riferimento(t.rateazione, t.anno), 8))
                    .set(31, num(t.debito, 15)).set(46, num(t.credito, 15))
                    .set(61, "0000").set(65, num(0, 3)).set(68, num(0, 15))))
            out.append(saldo_record(prog, "08", 17, d.imu))

        if d.inail:
            for j, t in enumerate(d.inail, 1):
                out.append(str(
                    Record().set(2, "40").set(4, prog).set(11, "09").set(13, num(j, 2))
                    .set(15, num(t.sede, 5)).set(20, num(t.ditta, 8)).set(28, num(t.cc, 2))
                    .set(30, cbi_text(t.causale, 1)).set(35, num(t.riferimento, 6))
                    .set(41, num(t.debito, 15)).set(56, num(t.credito, 15))))
            out.append(saldo_record(prog, "10", 13, d.inail))

        titolare = "2" if d.cf == banca.cf_titolare else "3"
        out.append(str(
            Record().set(2, "50").set(4, prog).set(11, "01").set(13, num(banca.abi, 5))
            .set(18, num(banca.cab, 5)).set(23, cbi_text(banca.conto, 12)).set(35, banca.cin)
            .set(36, num(d.saldo, 15)).set(51, "1" if d.cf2 else "0")
            .set(54, cbi_text(banca.cf_titolare, 16)).set(70, titolare).set(71, data_pag)
            .set(79, num(d.tot_credito, 15)).set(96, "3")
            .set(97, banca.iban[:2]).set(99, banca.iban[2:4])))

        out.append(str(
            Record().set(2, "50").set(4, prog).set(11, "02").set(13, cbi_text(banca.cf_mittente, 16))
            .set(33, num(banca.abi, 5)).set(38, num(banca.cab, 5)).set(63, "1")))

    tail = (Record().set(2, "EF").set(4, cbi_text(banca.sia, 5)).set(9, num(banca.abi, 5))
            .set(14, oggi.strftime("%d%m%y")).set(20, cbi_text(nome_supporto, 20))
            .set(46, num(len(deleghe), 7)).set(53, num(sum(d.saldo for d in deleghe), 15))
            .set(68, num(0, 15)).set(83, num(len(out) + 1, 7)).set(114, "E"))
    out.append(str(tail))
    return out


def saldo_record(prog, subtipo, pos, righe):
    deb = sum(r.debito for r in righe)
    cred = sum(r.credito for r in righe)
    saldo = deb - cred
    return str(Record().set(2, "40").set(4, prog).set(11, subtipo)
               .set(pos, num(deb, 15)).set(pos + 15, num(cred, 15))
               .set(pos + 30, "N" if saldo < 0 else "P").set(pos + 31, num(abs(saldo), 15)))


def valida_iban(iban):
    iban = iban.replace(" ", "").upper()
    if not re.fullmatch(r"IT\d{2}[A-Z]\d{10}[A-Z0-9]{12}", iban):
        raise F24Error(f"IBAN italiano non valido: {iban!r}")
    n = int("".join(str(int(c, 36)) for c in iban[4:] + iban[:4]))
    if n % 97 != 1:
        raise F24Error(f"IBAN con check digit errato: {iban}")
    return iban


def euro(cent):
    return f"{cent / 100:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def riepilogo(d: Delega):
    print(f"# {d.origine}: {d.cf} {d.cognome} {d.nome}".rstrip(), file=sys.stderr)
    print(f"#   pagamento {d.data_pagamento:%d/%m/%Y}  saldo {euro(d.saldo)}", file=sys.stderr)
    for t in d.erario:
        print(f"#   erario  {t.tributo} {t.rateazione or '-':>4} {t.anno}  {euro(t.debito):>10} {euro(t.credito):>10}",
              file=sys.stderr)
    if d.codice_atto or d.codice_ufficio:
        print(f"#   ufficio {d.codice_ufficio or '-'}  atto {d.codice_atto or '-'}", file=sys.stderr)
    for t in d.inps:
        print(f"#   inps    {t.sede} {t.causale} {t.matricola} {t.da}-{t.a or ''}  {euro(t.debito):>10} {euro(t.credito):>10}",
              file=sys.stderr)
    for t in d.regioni:
        print(f"#   regione {t.regione} {t.tributo} {t.rateazione} {t.anno}  {euro(t.debito):>10} {euro(t.credito):>10}",
              file=sys.stderr)
    for t in d.imu:
        print(f"#   imu     {t.ente} {t.tributo} {t.rateazione or '-':>4} {t.anno}  {euro(t.debito):>10} {euro(t.credito):>10}",
              file=sys.stderr)
    for t in d.inail:
        print(f"#   inail   {t.sede} {t.ditta} {t.cc} {t.riferimento} {t.causale}  {euro(t.debito):>10} {euro(t.credito):>10}",
              file=sys.stderr)


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("pdf", nargs="+", type=Path, help="PDF delle deleghe F24")
    p.add_argument("--sia", default=os.environ.get("F24_SIA", SIA_DEFAULT),
                   help=f"codice SIA dell'azienda mittente (default: $F24_SIA, poi {SIA_DEFAULT})")
    p.add_argument("--iban", default=os.environ.get("F24_IBAN", IBAN_DEFAULT),
                   help=f"IBAN di addebito; quello stampato sul PDF e' ignorato (default: $F24_IBAN, poi {IBAN_DEFAULT})")
    p.add_argument("--cf-titolare", default=os.environ.get("F24_CF_TITOLARE"),
                   help="codice fiscale del titolare del conto (default: $F24_CF_TITOLARE, poi il contribuente)")
    p.add_argument("--cf-mittente", default=os.environ.get("F24_CF_MITTENTE"),
                   help="codice fiscale dell'azienda titolare del codice SIA (default: $F24_CF_MITTENTE, poi il titolare del conto)")
    p.add_argument("--data", type=date.fromisoformat,
                   help="data di pagamento AAAA-MM-GG (default: scadenza sul PDF, poi data nel nome file)")
    p.add_argument("-o", "--output", type=Path,
                   help="file di uscita (default: accanto al primo PDF, estensione .cbi)")
    p.add_argument("--commit", action="store_true", help="scrive davvero il file")
    args = p.parse_args(argv)

    try:
        deleghe = [d for pdf in args.pdf for d in parse_pdf(pdf)]
        oggi = date.today()
        for d in deleghe:
            if args.data:
                d.data_pagamento = args.data
            if d.data_pagamento is None:
                raise F24Error(f"{d.origine}: data di pagamento sconosciuta, usa --data")
            if d.data_pagamento < oggi:
                raise F24Error(f"{d.origine}: data di pagamento {d.data_pagamento} gia' passata, usa --data")
            if d.saldo <= 0:
                raise F24Error(f"{d.origine}: saldo {euro(d.saldo)} non positivo")

        if not re.fullmatch(r"[A-Z0-9]{5}", args.sia.upper()):
            raise F24Error(f"codice SIA non valido: {args.sia!r}")
        iban = valida_iban(args.iban)

        cf_titolare = (args.cf_titolare or deleghe[0].cf).upper()
        cf_mittente = (args.cf_mittente or cf_titolare).upper()
        banca = Banca(sia=args.sia.upper(), iban=iban, cf_titolare=cf_titolare, cf_mittente=cf_mittente)
        nome = f"F24{datetime.now():%H%M%S}"
        records = build_flusso(deleghe, banca, oggi, nome)
    except F24Error as e:
        print(f"errore: {e}", file=sys.stderr)
        return 1

    for d in deleghe:
        riepilogo(d)
    totale = sum(d.saldo for d in deleghe)
    print(f"# {len(deleghe)} deleghe, totale {euro(totale)}, addebito su {iban}, {len(records)} record",
          file=sys.stderr)

    output = args.output or args.pdf[0].with_suffix(".cbi")
    if not args.commit:
        print("\n".join(records))
        print(f"# prova: nessun file scritto. Per scrivere {output} rilancia con --commit", file=sys.stderr)
        return 0
    output.write_bytes(("\r\n".join(records) + "\r\n").encode("ascii"))
    print(f"# scritto {output}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
