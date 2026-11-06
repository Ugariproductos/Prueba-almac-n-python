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


def _mejor(u, w, d, h, q):
    """Mejor colocación en rejilla de q unidades u dentro de un espacio w x d x h (cm)."""
    mejor = None
    for ox, oy, oz in _orientaciones(u):
        nx, ny, alt = int(w / ox + EPS), int(d / oy + EPS), int(h / oz + EPS)
        per = nx * ny
        if not per or alt < 1:
            continue
        k = min(q, per * alt)
        niveles = -(-k // per)
        usado = niveles * oz
        if mejor is None or k > mejor["k"] or (k == mejor["k"] and usado < mejor["usado"]):
            mejor = {"o": (ox, oy, oz), "nx": nx, "ny": ny, "per": per, "k": k, "niveles": niveles, "usado": usado}
    return mejor


def _llenar(caja, items, huecos_grandes_primero=False):
    """Coloca lo que quepa de items [(ref, cant, dims)] por capas, aprovechando huecos.

    Cada colocación es una rejilla nx x ny con varias alturas que empieza en (x0, y0, z).
    Los huecos que deja (franja a la derecha, franja al fondo y lo que falta de una capa
    incompleta) se rellenan con las referencias siguientes antes de seguir apilando.
    Devuelve (capas, puestos).
    """
    L, A, H = caja["l"], caja["a"], caja["h"]
    arriba = [0.0, 0.0, L, A, 0.0, H]          # espacio libre por encima de todo lo colocado
    huecos, capas, puestos = [], [], {}

    def colocar(reg, ref, q, u, es_hueco):
        x, y, w, d, z, h = reg
        m = _mejor(u, w, d, h, q)
        if not m:
            return 0, []
        ox, oy, oz = m["o"]
        capas.append({"ref": ref, "x0": x, "y0": y, "z": z, "hueco": es_hueco, **m})
        llenos, resto = divmod(m["k"], m["per"])
        nuevos = []
        if w - m["nx"] * ox > EPS:
            nuevos.append([x + m["nx"] * ox, y, w - m["nx"] * ox, d, z, m["usado"]])
        if d - m["ny"] * oy > EPS:
            nuevos.append([x, y + m["ny"] * oy, m["nx"] * ox, d - m["ny"] * oy, z, m["usado"]])
        if resto:
            zp = z + llenos * oz
            fila, col = divmod(resto, m["nx"])
            if col:
                nuevos.append([x + col * ox, y + fila * oy, (m["nx"] - col) * ox, oy, zp, oz])
                fila += 1
            if m["ny"] - fila > 0:
                nuevos.append([x, y + fila * oy, m["nx"] * ox, (m["ny"] - fila) * oy, zp, oz])
        if es_hueco and h - m["usado"] > EPS:
            nuevos.append([x, y, m["nx"] * ox, m["ny"] * oy, z + m["usado"], h - m["usado"]])
        return m, nuevos

    for ref, q, u in items:
        while q > 0:
            puesto = False
            clave = ((lambda r: -r[2] * r[3] * r[5]) if huecos_grandes_primero
                     else (lambda r: (r[4], r[2] * r[3] * r[5])))
            for reg in sorted(huecos, key=clave):
                m, nuevos = colocar(reg, ref, q, u, True)
                if m:
                    huecos.remove(reg)
                    huecos.extend(nuevos)
                    puesto = True
                    break
            if not puesto:
                m, nuevos = colocar(arriba, ref, q, u, False)
                if not m:
                    break
                huecos.extend(nuevos)
                arriba = [0.0, 0.0, L, A, arriba[4] + m["usado"], arriba[5] - m["usado"]]
            puestos[ref] = puestos.get(ref, 0) + m["k"]
            q -= m["k"]
    return capas, puestos


def _llenar_mejor(caja, items):
    """Prueba varios órdenes de colocación y se queda con el que mete más volumen."""
    vol = lambda it: it[2][0] * it[2][1] * it[2][2]
    ordenes = [
        list(items),
        sorted(items, key=lambda it: -max(it[2])),
        sorted(items, key=lambda it: -min(it[2])),
        sorted(items, key=lambda it: -vol(it) * it[1]),
        sorted(items, key=lambda it: vol(it)),
    ]
    mejor = None
    for orden, grandes in [(o, g) for g in (False, True) for o in ordenes]:
        capas, puestos = _llenar(caja, orden, grandes)
        v = sum(puestos[r] * vol(it) for it in items for r in [it[0]] if r in puestos)
        if mejor is None or v > mejor[0] + EPS or (abs(v - mejor[0]) <= EPS and len(capas) < len(mejor[1])):
            mejor = (v, capas, puestos)
        if all(puestos.get(r) == q for r, q, _ in items):
            break
    return mejor[1], mejor[2]


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
        capas, puestos = _llenar_mejor(tipo, items)
        if not puestos:
            raise ValueError("No se ha podido colocar la mercancía.")
        cont = [(r, puestos[r], articulos[r]["u"]) for r in orden if puestos.get(r)]
        elegido = tipo
        for t in candidatas:  # la caja más pequeña donde cabe lo mismo
            if t["vol"] >= tipo["vol"]:
                break
            c2, p2 = _llenar_mejor(t, cont)
            if all(p2.get(r) == q for r, q, _ in cont):
                elegido, capas, puestos = t, c2, p2
                break
        for r, q in puestos.items():
            pend[r] -= q
        usado = sum(q * articulos[r]["vol_ud"] for r, q in puestos.items())
        bultos.append({"caja": elegido, "capas": capas, "contenido": _ordenar(capas, puestos), "usado": usado})

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

    # 3) intentar quitar el último bulto metiendo su contenido en otro ya hecho
    def reempaquetar(contenido, candidatas):
        items = [(r, contenido[r], articulos[r]["u"]) for r in orden if contenido.get(r)]
        for t in candidatas:
            capas, puestos = _llenar_mejor(t, items)
            if all(puestos.get(r) == q for r, q, _ in items):
                return {"caja": t, "capas": capas, "contenido": _ordenar(capas, puestos),
                        "usado": sum(q * articulos[r]["vol_ud"] for r, q in puestos.items())}
        return None

    while len(bultos) > 1:
        ultimo = bultos[-1]
        hecho = False
        for i in range(len(bultos) - 2, -1, -1):
            junto = dict(bultos[i]["contenido"])
            for r, q in ultimo["contenido"].items():
                junto[r] = junto.get(r, 0) + q
            tope = bultos[i]["caja"] if bultos[i]["caja"]["vol"] > principal["vol"] else principal
            nuevo = reempaquetar(junto, [t for t in tipos if t["vol"] <= tope["vol"]
                                         and (not t.get("solo_si_no_cabe") or t is bultos[i]["caja"])])
            if nuevo:
                bultos[i] = nuevo
                bultos.pop()
                hecho = True
                break
        if not hecho:
            break
    return bultos


def _ordenar(capas, puestos):
    """Contenido en orden de colocación, de abajo arriba (por la altura de la primera capa de cada ref)."""
    pos = {}
    for i, k in enumerate(capas):
        z, j = pos.get(k["ref"], (k["z"], i))
        pos[k["ref"]] = (min(z, k["z"]), j)
    return {r: puestos[r] for r in sorted(puestos, key=lambda r: pos.get(r, (0, 0)))}


def colocacion(bulto, ref):
    partes = []
    for k in bulto["capas"]:
        if k["ref"] != ref:
            continue
        llenos, resto = divmod(k["k"], k["per"])
        t = f"{llenos} altura{'s' if llenos > 1 else ''} de {k['nx']}x{k['ny']}" if llenos else ""
        if resto:
            t += f" + {resto} encima" if t else f"1 capa de {resto} uds"
        if k.get("hueco"):
            t += " en el hueco libre"
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
        base = {"ref": k["ref"], "n": orden.index(k["ref"]) + 1, "nx": k["nx"], "ny": k["ny"], "o": tuple(k["o"]),
                "x0": k.get("x0", 0), "y0": k.get("y0", 0), "hueco": k.get("hueco", False)}
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
    if p.get("x0", 0) > EPS or p.get("y0", 0) > EPS:
        donde += f", en el hueco a {_cm(p['x0'])} cm del lado izquierdo y {_cm(p['y0'])} cm del frente"
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
    x0, y0 = p.get("x0", 0), p.get("y0", 0)
    return [(x0 + (i % p["nx"]) * ox, y0 + (i // p["nx"]) * oy, ox, oy) for i in range(p["cuantos"])]


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
