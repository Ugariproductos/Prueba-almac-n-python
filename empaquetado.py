"""Calcula en qué cajas de envío va cada pedido y cómo colocar la mercancía.

Lee un Excel con dos hojas:
  - Articulos: Referencia | Descripción | Largo | Ancho | Alto | Uds por caja
               [| Largo ud | Ancho ud | Alto ud]   (cm)
    Medidas de la caja del artículo y unidades que trae. Si se rellenan las
    medidas de una unidad se usan; si no, se estiman dividiendo la caja.
  - Cajas:     Caja | Largo | Ancho | Alto | Solo si no cabe   (cajas de envío, en cm)
    "Solo si no cabe" = sí: esa caja (XL) solo se usa si una unidad no cabe en otra.

Cada unidad es un bloque que se coloca por capas dentro de la caja: rejilla
nx × ny en el suelo, varias alturas, y la siguiente referencia encima.
Se mezclan artículos y se busca usar el menor número de cajas.
Es el mismo algoritmo que el artefacto "Cajas por pedido".
"""

import os
import re
import sys
from itertools import permutations

from openpyxl import load_workbook

CARPETA = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(CARPETA, "cajas.xlsx")
EPS = 1e-9


# ---------- lectura del Excel ----------

def _col(cabeceras, *claves):
    for i, c in enumerate(cabeceras):
        texto = str(c or "").strip().lower()
        if any(k in texto for k in claves):
            return i
    return None


def _col_ud(cabeceras, clave):
    for i, c in enumerate(cabeceras):
        texto = str(c or "").strip().lower()
        if clave in texto and re.search(r"\bud\b|unidad", texto):
            return i
    return None


def _num(v):
    if v is None or v == "":
        return None
    return float(str(v).replace(",", "."))


def dims_unidad(l, a, h, uds):
    """Medidas estimadas de una unidad: se reparte la caja entre sus unidades."""
    d = [l, a, h]
    n = int(round(uds or 1))
    factores, p = [], 2
    while n > 1:
        if n % p == 0:
            factores.append(p)
            n //= p
        else:
            p += 1
    for p in sorted(factores, reverse=True):
        i = d.index(max(d))
        d[i] /= p
    return d


CAJAS_UGARI = [  # se usan si el Excel no trae hoja de cajas
    {"nombre": "S", "l": 35, "a": 30, "h": 30, "vol": 35 * 30 * 30, "solo_si_no_cabe": False},
    {"nombre": "M", "l": 40, "a": 30, "h": 30, "vol": 40 * 30 * 30, "solo_si_no_cabe": False},
    {"nombre": "L", "l": 60, "a": 40, "h": 40, "vol": 60 * 40 * 40, "solo_si_no_cabe": False},
    {"nombre": "XL", "l": 80, "a": 60, "h": 50, "vol": 80 * 60 * 50, "solo_si_no_cabe": True},
]


def leer_excel(ruta=EXCEL, cajas_por_defecto=None):
    """Devuelve (articulos, cajas). Detecta las hojas y la fila de cabecera por su contenido.

    Admite medidas en cm o en mm (si la cabecera dice "mm"), columna EAN opcional y
    artículos sin medidas (quedan marcados con "sin_medidas" y no se pueden empaquetar).
    Si no hay hoja de cajas se usan cajas_por_defecto (o las cajas S/M/L/XL de Ugari).
    """
    wb = load_workbook(ruta, data_only=True, read_only=True)
    articulos, cajas = {}, []
    for hoja in wb.worksheets:
        todas = list(hoja.iter_rows(values_only=True))
        ic = next((i for i, f in enumerate(todas[:15])
                   if None not in (_col(f, "largo"), _col(f, "ancho"), _col(f, "alto"))), None)
        if ic is None:
            continue
        cab, filas = todas[ic], todas[ic:]
        i_l, i_a, i_h = _col(cab, "largo"), _col(cab, "ancho"), _col(cab, "alto")

        def med(f, i):
            v = _num(f[i]) if i is not None and i < len(f) else None
            return None if v is None else v * (0.1 if re.search(r"\bmm\b", str(cab[i]).lower()) else 1)

        i_ref = _col(cab, "ref", "código", "codigo")
        if i_ref is not None:  # hoja de artículos
            i_desc = _col(cab, "desc", "nombre")
            i_uds = _col(cab, "uds", "unidades", "cantidad")
            i_ean = _col(cab, "ean")
            u_l, u_a, u_h = _col_ud(cab, "largo"), _col_ud(cab, "ancho"), _col_ud(cab, "alto")
            for f in filas[1:]:
                if not f or f[i_ref] in (None, ""):
                    continue
                ref = str(f[i_ref]).strip()
                l, a, h = med(f, i_l), med(f, i_a), med(f, i_h)
                uds = (_num(f[i_uds]) if i_uds is not None else 1) or 1
                ean = str(f[i_ean]).strip() if i_ean is not None and f[i_ean] not in (None, "") else ""
                desc = str(f[i_desc] or "") if i_desc is not None else ""
                sin = not (l and a and h)
                ud = [med(f, i) for i in (u_l, u_a, u_h)]
                articulos[ref] = {
                    "ref": ref, "ean": ean, "desc": desc or (f"EAN {ean}" if ean else ""),
                    "sin_medidas": sin,
                    "vol_ud": 0 if sin else l * a * h / uds,
                    "u": None if sin else (ud if all(ud) else dims_unidad(l, a, h, uds)),
                }
        else:  # hoja de cajas de envío
            i_nom = _col(cab, "caja", "nombre", "tipo")
            if i_nom is None:
                i_nom = 0
            i_solo = _col(cab, "solo", "oblig", "necesari")
            for f in filas[1:]:
                if not f or f[i_nom] in (None, ""):
                    continue
                l, a, h = med(f, i_l), med(f, i_a), med(f, i_h)
                if not (l and a and h):
                    continue
                nombre = str(f[i_nom]).strip()
                if i_solo is not None:
                    solo = str(f[i_solo] or "").strip().lower() in ("sí", "si", "x", "1", "true", "yes")
                else:
                    solo = nombre.upper() == "XL"
                cajas.append({"nombre": nombre, "l": l, "a": a, "h": h, "vol": l * a * h,
                              "solo_si_no_cabe": solo})
    wb.close()
    if not articulos:
        raise ValueError("No encuentro la hoja de artículos (falta columna Referencia).")
    if not cajas:
        cajas = [dict(c) for c in (cajas_por_defecto or CAJAS_UGARI)]
    return articulos, cajas


def buscar_ref(articulos, texto):
    """Referencia del maestro que corresponde a un texto (referencia o EAN), o None."""
    t = str(texto or "").strip()
    if t in articulos:
        return t
    n = re.sub(r"[\s.\-_/]", "", t).upper()
    for k, a in articulos.items():
        if re.sub(r"[\s.\-_/]", "", k).upper() == n or (a.get("ean") and a["ean"] == n):
            return k
    return None


# ---------- cálculo ----------

def _orientaciones(u):
    return list(dict.fromkeys(permutations(u)))


def _llenar(caja, items):
    """Coloca por capas lo que quepa de items [(ref, cant, dims)]. Devuelve (capas, puestos)."""
    z, capas, puestos = 0.0, [], {}
    for ref, q, u in items:
        while q > 0:
            mejor = None
            for ox, oy, oz in _orientaciones(u):
                nx = int(caja["l"] / ox + EPS)
                ny = int(caja["a"] / oy + EPS)
                alt = int((caja["h"] - z) / oz + EPS)
                per = nx * ny
                if not per or alt < 1:
                    continue
                k = min(q, per * alt)
                niveles = -(-k // per)
                usado = niveles * oz
                if mejor is None or k > mejor["k"] or (k == mejor["k"] and usado < mejor["usado"]):
                    mejor = {"o": (ox, oy, oz), "nx": nx, "ny": ny, "per": per,
                             "k": k, "niveles": niveles, "usado": usado}
            if mejor is None:
                break
            capas.append({"ref": ref, "z": z, **mejor})
            puestos[ref] = puestos.get(ref, 0) + mejor["k"]
            z += mejor["usado"]
            q -= mejor["k"]
    return capas, puestos


def _cabe(u, caja):
    return any(x <= caja["l"] + EPS and y <= caja["a"] + EPS and z <= caja["h"] + EPS
               for x, y, z in _orientaciones(u))


def empaquetar(pedido, articulos, cajas):
    """pedido: lista de (referencia, cantidad). Devuelve lista de bultos.

    Las cajas marcadas "solo si no cabe" (XL) solo se usan para unidades que no
    caben en la caja normal más grande (L); el resto va en L o en una menor.
    """
    tipos = sorted(cajas, key=lambda c: c["vol"])
    mayor = tipos[-1]
    normales = [t for t in tipos if not t.get("solo_si_no_cabe")] or tipos
    principal = normales[-1]
    pend = {}
    for ref, cant in pedido:
        if ref not in articulos:
            raise ValueError(f"Referencia desconocida: {ref}")
        if articulos[ref].get("sin_medidas"):
            raise ValueError(f"{ref}: faltan las medidas en el Excel maestro.")
        pend[ref] = pend.get(ref, 0) + int(cant)
    orden = sorted(pend, key=lambda r: articulos[r]["vol_ud"], reverse=True)
    for ref in orden:
        if not _cabe(articulos[ref]["u"], mayor):
            raise ValueError(f"{ref}: una unidad no cabe ni en la caja {mayor['nombre']}.")

    bultos = []

    def meter(tipo, refs, candidatas):
        items = [(r, pend[r], articulos[r]["u"]) for r in refs if pend[r] > 0]
        capas, puestos = _llenar(tipo, items)
        if not puestos:
            raise ValueError("No se ha podido colocar la mercancía.")
        cont = [(r, puestos[r], articulos[r]["u"]) for r in orden if puestos.get(r)]
        elegido = tipo
        for t in candidatas:  # la caja más pequeña donde cabe lo mismo
            if t["vol"] >= tipo["vol"]:
                break
            c2, p2 = _llenar(t, cont)
            if all(p2.get(r) == q for r, q, _ in cont):
                elegido, capas, puestos = t, c2, p2
                break
        for r, q in puestos.items():
            pend[r] -= q
        usado = sum(q * articulos[r]["vol_ud"] for r, q in puestos.items())
        bultos.append({"caja": elegido, "capas": capas, "contenido": puestos, "usado": usado})

    # 1) solo lo obligatorio va en caja grande (XL); el hueco que sobra se aprovecha
    while True:
        grandes = [r for r in orden if pend[r] > 0 and not _cabe(articulos[r]["u"], principal)]
        if not grandes:
            break
        tipo = next(t for t in tipos if _cabe(articulos[grandes[0]]["u"], t))
        resto = [r for r in orden if pend[r] > 0 and r not in grandes]
        meter(tipo, grandes + resto, tipos)
    # 2) el resto, en la caja normal más grande (L) o en una más pequeña si basta
    while any(pend[r] > 0 for r in orden):
        meter(principal, orden, normales)
    return bultos


def colocacion(bulto, ref):
    partes = []
    for k in bulto["capas"]:
        if k["ref"] != ref:
            continue
        llenos, resto = divmod(k["k"], k["per"])
        t = f"{llenos} altura{'s' if llenos > 1 else ''} de {k['nx']}x{k['ny']}" if llenos else ""
        if resto:
            t += f" + {resto} encima" if t else f"1 capa de {resto} uds"
        partes.append(t)
    return " · ".join(partes)


def _cm(v):
    return f"{round(v, 1):g}".replace(".", ",")


def pasos_datos(bulto):
    """Pasos de colocación de abajo arriba: una entrada por capa completa (con sus alturas) o capa parcial."""
    orden = list(bulto["contenido"])
    out = []
    for k in bulto["capas"]:
        llenos, resto = divmod(k["k"], k["per"])
        base = {"ref": k["ref"], "n": orden.index(k["ref"]) + 1, "nx": k["nx"], "ny": k["ny"], "o": tuple(k["o"])}
        if llenos:
            out.append({**base, "uds": llenos * k["per"], "alturas": llenos, "cuantos": k["per"],
                        "z": k["z"], "parcial": False})
        if resto:
            out.append({**base, "uds": resto, "alturas": 1, "cuantos": resto,
                        "z": k["z"] + llenos * k["o"][2], "parcial": True})
    return out


def texto_paso(p):
    ox, oy, oz = p["o"]
    donde = "En el fondo de la caja" if p["z"] < EPS else f"Encima, a {_cm(p['z'])} cm del fondo"
    if p["parcial"]:
        que = f"{p['uds']} uds de {p['ref']}, empezando por la esquina delantera izquierda y fila a fila"
    else:
        que = (f"{p['uds']} uds de {p['ref']}: {p['alturas']} altura{'s' if p['alturas'] > 1 else ''} de "
               f"{p['nx']} x {p['ny']} ({p['nx']} a lo largo, {p['ny']} a lo ancho)")
    orient = f"Cada unidad: {_cm(ox)} cm a lo largo · {_cm(oy)} cm a lo ancho · {_cm(oz)} cm de alto"
    return donde, que, orient


def celdas_paso(p):
    """Rectángulos en planta (x a lo largo, y desde el frente) de las unidades de un paso."""
    ox, oy, _ = p["o"]
    return [((i % p["nx"]) * ox, (i // p["nx"]) * oy, ox, oy) for i in range(p["cuantos"])]


def pasos(bulto):
    """Instrucciones de colocación en texto, de abajo arriba."""
    out = []
    for p in pasos_datos(bulto):
        donde, que, orient = texto_paso(p)
        out.append(f"{donde}: [{p['n']}] {que}. {orient}")
    return out


def imprimir(bultos, articulos, albaran="", pedido=""):
    if not bultos:
        print("Pedido vacío.")
        return
    total = len(bultos)
    print(f"\nAlbarán Ugari: {albaran or '-'}   Pedido cliente: {pedido or '-'}")
    resumen = {}
    for b in bultos:
        resumen[b["caja"]["nombre"]] = resumen.get(b["caja"]["nombre"], 0) + 1
    print(f"Total: {total} bulto(s): " + ", ".join(f"{n} x {k}" for k, n in resumen.items()))
    for i, b in enumerate(bultos, 1):
        c = b["caja"]
        llenado = 100 * b["usado"] / c["vol"]
        print(f"\nBULTO {i}/{total}  caja {c['nombre']} ({c['l']:g}x{c['a']:g}x{c['h']:g} cm)  {llenado:.0f}% lleno")
        print("  Colocar de abajo arriba:")
        for j, (ref, uds) in enumerate(b["contenido"].items(), 1):
            print(f"   {j}. {uds:>4} x {ref}  {articulos[ref]['desc']}  [{colocacion(b, ref)}]")
        print("  Colocación paso a paso:")
        for n, t in enumerate(pasos(b), 1):
            print(f"   Paso {n}. {t}")


# ---------- menú ----------

def pedir_pedido(articulos):
    print("Escribe referencia y cantidad (ej: REF001 12). Línea vacía para terminar.")
    pedido = []
    while True:
        linea = input("> ").strip()
        if not linea:
            return pedido
        partes = linea.replace(";", " ").split()
        if len(partes) != 2 or not partes[1].isdigit():
            print("Formato: REFERENCIA CANTIDAD")
            continue
        ref = buscar_ref(articulos, partes[0])
        if ref is None:
            print(f"No existe la referencia {partes[0]}.")
            continue
        if articulos[ref].get("sin_medidas"):
            print(f"{ref}: faltan las medidas en el Excel maestro.")
            continue
        pedido.append((ref, int(partes[1])))


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else EXCEL
    articulos, cajas = leer_excel(ruta)
    print(f"Cargados {len(articulos)} artículos y {len(cajas)} cajas de {os.path.basename(ruta)}")
    menu = """
=== EMPAQUETADO ===
1. Calcular bultos de un pedido
2. Ver artículos
3. Ver cajas de envío
4. Recargar Excel
0. Salir
> """
    while True:
        op = input(menu).strip()
        if op == "1":
            albaran = input("Nº albarán Ugari: ").strip()
            pedido_cli = input("Nº pedido cliente: ").strip()
            imprimir(empaquetar(pedir_pedido(articulos), articulos, cajas), articulos, albaran, pedido_cli)
        elif op == "2":
            for a in articulos.values():
                u = "FALTAN MEDIDAS" if a["u"] is None else "unidad " + " x ".join(f"{v:g}" for v in a["u"]) + " cm"
                print(f"{a['ref']:<14}{a['desc']:<40} {u}")
        elif op == "3":
            for c in sorted(cajas, key=lambda c: c["vol"]):
                print(f"{c['nombre']:<8}{c['l']:g} x {c['a']:g} x {c['h']:g} cm")
        elif op == "4":
            articulos, cajas = leer_excel(ruta)
            print(f"Recargados {len(articulos)} artículos y {len(cajas)} cajas.")
        elif op == "0":
            break
        else:
            print("Opción no válida.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        input("Pulsa Enter para cerrar...")
