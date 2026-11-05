"""Calcula en qué cajas de envío cabe cada pedido.

Lee un Excel con dos hojas:
  - Artículos: Referencia | Descripción | Largo | Ancho | Alto | Uds por caja
    (medidas en cm de la caja del artículo y unidades que trae esa caja)
  - Cajas:     Caja | Largo | Ancho | Alto   (cajas de envío, en cm)

El volumen de una unidad = volumen de la caja del artículo / uds por caja.
Se mezclan artículos en la misma caja y se busca usar el menor número de cajas.
Para ampliar productos o cajas basta con añadir filas al Excel.
"""

import os
import sys

from openpyxl import load_workbook

CARPETA = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(CARPETA, "cajas.xlsx")
APROVECHAMIENTO = 0.90  # fracción del volumen de la caja que se puede llenar
EPS = 1e-9


# ---------- lectura del Excel ----------

def _col(cabeceras, *claves):
    for i, c in enumerate(cabeceras):
        texto = str(c or "").strip().lower()
        if any(k in texto for k in claves):
            return i
    return None


def _num(v):
    if v is None or v == "":
        return None
    return float(str(v).replace(",", "."))


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
            for f in filas[1:]:
                if not f or f[i_ref] in (None, ""):
                    continue
                ref = str(f[i_ref]).strip()
                uds = _num(f[i_uds]) if i_uds is not None else 1
                vol_caja = _num(f[i_l]) * _num(f[i_a]) * _num(f[i_h])
                articulos[ref] = {
                    "ref": ref,
                    "desc": str(f[i_desc] or "") if i_desc is not None else "",
                    "vol_ud": vol_caja / (uds or 1),
                }
        else:  # hoja de cajas de envío
            i_nom = _col(cab, "caja", "nombre", "tipo")
            if i_nom is None:
                i_nom = 0
            for f in filas[1:]:
                if not f or f[i_nom] in (None, ""):
                    continue
                cajas.append({
                    "nombre": str(f[i_nom]).strip(),
                    "medidas": f"{f[i_l]}x{f[i_a]}x{f[i_h]}",
                    "vol": _num(f[i_l]) * _num(f[i_a]) * _num(f[i_h]),
                })
    wb.close()
    if not articulos:
        raise ValueError("No encuentro la hoja de artículos (falta columna Referencia).")
    if not cajas:
        raise ValueError("No encuentro la hoja de cajas de envío.")
    return articulos, cajas


# ---------- cálculo ----------

def empaquetar(pedido, articulos, cajas, aprovechamiento=APROVECHAMIENTO):
    """pedido: lista de (referencia, cantidad).

    Devuelve lista de cajas: {"caja", "medidas", "capacidad", "usado", "contenido": {ref: uds}}.
    Llena primero cajas grandes (menos cajas) y después cambia cada una
    por la caja más pequeña en la que cabe su contenido.
    """
    tipos = sorted(cajas, key=lambda c: c["vol"])  # de pequeña a grande
    grande = tipos[-1]
    cap_grande = grande["vol"] * aprovechamiento

    lineas = []
    for ref, cant in pedido:
        if ref not in articulos:
            raise ValueError(f"Referencia desconocida: {ref}")
        vol = articulos[ref]["vol_ud"]
        if vol > cap_grande + EPS:
            raise ValueError(f"{ref}: una unidad no cabe ni en la caja {grande['nombre']}.")
        lineas.append((ref, int(cant), vol))
    lineas.sort(key=lambda x: x[2], reverse=True)  # primero lo más voluminoso

    abiertas = []
    for ref, cant, vol in lineas:
        for caja in abiertas:  # rellenar huecos de cajas ya abiertas
            if cant == 0:
                break
            caben = int((cap_grande - caja["usado"] + EPS) // vol)
            n = min(caben, cant)
            if n:
                caja["contenido"][ref] = caja["contenido"].get(ref, 0) + n
                caja["usado"] += n * vol
                cant -= n
        por_caja = int((cap_grande + EPS) // vol)
        while cant:
            n = min(por_caja, cant)
            abiertas.append({"contenido": {ref: n}, "usado": n * vol})
            cant -= n

    resultado = []
    for caja in abiertas:
        tipo = next(t for t in tipos if caja["usado"] <= t["vol"] * aprovechamiento + EPS)
        resultado.append({
            "caja": tipo["nombre"],
            "medidas": tipo["medidas"],
            "capacidad": tipo["vol"],
            "usado": caja["usado"],
            "contenido": caja["contenido"],
        })
    return resultado


def imprimir(resultado, articulos):
    if not resultado:
        print("Pedido vacío.")
        return
    print(f"\nTotal: {len(resultado)} caja(s)")
    resumen = {}
    for c in resultado:
        resumen[c["caja"]] = resumen.get(c["caja"], 0) + 1
    print("  " + ", ".join(f"{n} x {k}" for k, n in resumen.items()))
    for i, c in enumerate(resultado, 1):
        llenado = 100 * c["usado"] / c["capacidad"]
        print(f"\nCaja {i}: {c['caja']} ({c['medidas']} cm) - ocupación {llenado:.0f}%")
        for ref, uds in c["contenido"].items():
            print(f"   {uds:>5} x {ref}  {articulos[ref]['desc']}")


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
        ref = partes[0]
        if ref not in articulos:
            print(f"No existe la referencia {ref}.")
            continue
        pedido.append((ref, int(partes[1])))


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else EXCEL
    articulos, cajas = leer_excel(ruta)
    print(f"Cargados {len(articulos)} artículos y {len(cajas)} cajas de {os.path.basename(ruta)}")
    menu = """
=== EMPAQUETADO ===
1. Calcular cajas de un pedido
2. Ver artículos
3. Ver cajas de envío
4. Recargar Excel
0. Salir
> """
    while True:
        op = input(menu).strip()
        if op == "1":
            imprimir(empaquetar(pedir_pedido(articulos), articulos, cajas), articulos)
        elif op == "2":
            for a in articulos.values():
                print(f"{a['ref']:<14}{a['desc']:<35}{a['vol_ud']:>10.0f} cm³/ud")
        elif op == "3":
            for c in sorted(cajas, key=lambda c: c["vol"]):
                print(f"{c['nombre']:<14}{c['medidas']:<16}{c['vol']:>10.0f} cm³")
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
