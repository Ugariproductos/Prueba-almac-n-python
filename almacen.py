import os
import sqlite3

# La base de datos se guarda junto a este archivo
DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "almacen.db")


def conectar():
    con = sqlite3.connect(DB)
    con.execute("""
        CREATE TABLE IF NOT EXISTS productos (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre   TEXT NOT NULL UNIQUE,
            cantidad INTEGER NOT NULL DEFAULT 0,
            precio   REAL NOT NULL DEFAULT 0
        )
    """)
    return con


def anadir(con, nombre, cantidad, precio):
    try:
        con.execute(
            "INSERT INTO productos (nombre, cantidad, precio) VALUES (?, ?, ?)",
            (nombre, cantidad, precio),
        )
        con.commit()
        print(f"Añadido: {nombre}")
    except sqlite3.IntegrityError:
        print("Ese producto ya existe.")


def listar(con):
    filas = con.execute(
        "SELECT id, nombre, cantidad, precio FROM productos ORDER BY nombre"
    ).fetchall()
    if not filas:
        print("Almacén vacío.")
        return
    print(f"{'ID':<4}{'Nombre':<20}{'Cant.':>6}{'Precio':>10}")
    for id_, nombre, cant, precio in filas:
        print(f"{id_:<4}{nombre:<20}{cant:>6}{precio:>10.2f}")


def mover_stock(con, nombre, delta):
    cur = con.execute(
        "UPDATE productos SET cantidad = cantidad + ? "
        "WHERE nombre = ? AND cantidad + ? >= 0",
        (delta, nombre, delta),
    )
    con.commit()
    if cur.rowcount:
        print("Stock actualizado.")
    else:
        print("Producto no encontrado o stock insuficiente.")


def eliminar(con, nombre):
    cur = con.execute("DELETE FROM productos WHERE nombre = ?", (nombre,))
    con.commit()
    print("Eliminado." if cur.rowcount else "No existe.")


def valor_total(con):
    total = con.execute(
        "SELECT COALESCE(SUM(cantidad * precio), 0) FROM productos"
    ).fetchone()[0]
    print(f"Valor total del almacén: {total:.2f} €")


def main():
    con = conectar()
    menu = """
=== ALMACÉN ===
1. Añadir producto
2. Listar productos
3. Entrada de stock
4. Salida de stock
5. Eliminar producto
6. Valor total
0. Salir
> """
    while True:
        op = input(menu).strip()
        try:
            if op == "1":
                anadir(con, input("Nombre: "), int(input("Cantidad: ")),
                       float(input("Precio: ").replace(",", ".")))
            elif op == "2":
                listar(con)
            elif op == "3":
                mover_stock(con, input("Nombre: "), int(input("Unidades: ")))
            elif op == "4":
                mover_stock(con, input("Nombre: "), -int(input("Unidades: ")))
            elif op == "5":
                eliminar(con, input("Nombre: "))
            elif op == "6":
                valor_total(con)
            elif op == "0":
                break
            else:
                print("Opción no válida.")
        except ValueError:
            print("Número no válido.")
    con.close()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        input("Pulsa Enter para cerrar...")
