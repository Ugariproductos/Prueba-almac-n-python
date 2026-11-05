"""Calcula en qué cajas de envío va cada pedido y cómo colocar la mercancía.

Lee un Excel con dos hojas:
  - Articulos: Referencia | Descripción | Largo | Ancho | Alto | Uds por caja
               [| Largo ud | Ancho ud | Alto ud]   (cm)
    Medidas de la caja del artículo y unidades que trae. Si se rellenan las
    medidas de una unidad se usan; si no, se estiman dividiendo la caja.
  - Cajas:     Caja | Largo | Ancho | Alto   (cajas de envío, en cm)

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


def leer_excel(ruta=EXCEL):
    """Devuelve (articulos, cajas). Detecta las hojas por sus cabeceras."""
    wb = load_workbook(ruta, data_only=True, read_only=True)
    articulos, cajas = {}, []
    for hoja in wb.worksheets:
        filas = list(hoja.iter_rows(values_only=True))
        if not filas:
            continue
        cab = filas[0]
        i_l, i_a, i_h = _col(cab, "largo"), _col(cab, "ancho"), _col(cab, "alto")
        if None in (i_l, i_a, i_h):
            continue
        i_ref = _col(cab, "ref", "código", "codigo")
        if i_ref is not None:  # hoja de artículos
            i_desc = _col(cab, "desc", "nombre")
            i_uds = _col(cab, "uds", "unidades", "cantidad")
            u_l, u_a, u_h = _col_ud(cab, "largo"), _col_ud(cab, "ancho"), _col_ud(cab, "alto")
            for f in filas[1:]:
                if not f or f[i_ref] in (None, ""):
                    continue
                ref = str(f[i_ref]).strip()
                l, a, h = _num(f[i_l]), _num(f[i_a]), _num(f[i_h])
                uds = (_num(f[i_uds]) if i_uds is not None else 1) or 1
                ud = [_num(f[i]) if i is not None else None for i in (u_l, u_a, u_h)]
                articulos[ref] = {
                    "ref": ref,
                    "desc": str(f[i_desc] or "") if i_desc is not None else "",
                    "vol_ud": l * a * h / uds,
                    "u": ud if all(ud) else dims_unidad(l, a, h, uds),
                }
        else:  # hoja de cajas de envío
            i_nom = _col(cab, "caja", "nombre", "tipo")
            if i_nom is None:
                i_nom = 0
            for f in filas[1:]:
                if not f or f[i_nom] in (None, ""):
                    continue
                l, a, h = _num(f[i_l]), _num(f[i_a]), _num(f[i_h])
                cajas.append({"nombre": str(f[i_nom]).strip(), "l": l, "a": a, "h": h, "vol": l * a * h})
    wb.close()
    if not articulos:
        raise ValueError("No encuentro la hoja de artículos (falta columna Referencia).")
    if not cajas:
        raise ValueError("No encuentro la hoja de cajas de envío.")
    return articulos, cajas


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


def empaquetar(pedido, articulos, cajas):
    """pedido: lista de (referencia, cantidad). Devuelve lista de bultos."""
    tipos = sorted(cajas, key=lambda c: c["vol"])
    grande = tipos[-1]
    pend = {}
    for ref, cant in pedido:
        if ref not in articulos:
            raise ValueError(f"Referencia desconocida: {ref}")
        pend[ref] = pend.get(ref, 0) + int(cant)
    orden = sorted(pend, key=lambda r: articulos[r]["vol_ud"], reverse=True)
    for ref in orden:
        u = articulos[ref]["u"]
        if not any(x <= grande["l"] + EPS and y <= grande["a"] + EPS and z <= grande["h"] + EPS
                   for x, y, z in _orientaciones(u)):
            raise ValueError(f"{ref}: una unidad no cabe ni en la caja {grande['nombre']}.")

    bultos = []
    while any(pend[r] > 0 for r in orden):
        items = [(r, pend[r], articulos[r]["u"]) for r in orden if pend[r] > 0]
        capas, puestos = _llenar(grande, items)
        if not puestos:
            raise ValueError("No se ha podido colocar la mercancía.")
        cont = [(r, puestos[r], articulos[r]["u"]) for r in orden if puestos.get(r)]
        tipo = grande
        for t in tipos[:-1]:  # la caja más pequeña donde cabe lo mismo
            c2, p2 = _llenar(t, cont)
            if all(p2.get(r) == q for r, q, _ in cont):
                tipo, capas, puestos = t, c2, p2
                break
        for r, q in puestos.items():
            pend[r] -= q
        usado = sum(q * articulos[r]["vol_ud"] for r, q in puestos.items())
        bultos.append({"caja": tipo, "capas": capas, "contenido": puestos, "usado": usado})
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
        if partes[0] not in articulos:
            print(f"No existe la referencia {partes[0]}.")
            continue
        pedido.append((partes[0], int(partes[1])))


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
                u = " x ".join(f"{v:g}" for v in a["u"])
                print(f"{a['ref']:<14}{a['desc']:<40} unidad {u} cm")
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
