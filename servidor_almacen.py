"""Servidor del almacén: deja que la app Android use almacen.db por Wi-Fi.

Haz doble clic en este archivo y deja la ventana abierta mientras uses el móvil.
"""
import json
import socket
import sqlite3
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from almacen import conectar

PUERTO = 8765
CAMPOS = ["id", "nombre", "codigo", "cantidad", "precio"]


def producto_por_codigo(con, codigo):
    fila = con.execute(
        "SELECT id, nombre, codigo, cantidad, precio FROM productos WHERE codigo = ?",
        (codigo,),
    ).fetchone()
    return dict(zip(CAMPOS, fila)) if fila else None


class Manejador(BaseHTTPRequestHandler):
    def responder(self, datos, estado=200):
        cuerpo = json.dumps(datos, ensure_ascii=False).encode("utf-8")
        self.send_response(estado)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(cuerpo)))
        self.end_headers()
        self.wfile.write(cuerpo)

    def error(self, mensaje, estado=400):
        self.responder({"ok": False, "error": mensaje}, estado)

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/ping":
            return self.responder({"ok": True})
        if url.path != "/producto":
            return self.error("Ruta no encontrada", 404)
        codigo = parse_qs(url.query).get("codigo", [""])[0].strip()
        con = conectar()
        try:
            p = producto_por_codigo(con, codigo)
        finally:
            con.close()
        self.responder({"ok": True, "encontrado": p is not None, "producto": p})

    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length", 0))
            datos = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            return self.error("JSON no válido")
        if not isinstance(datos, dict):
            return self.error("JSON no válido")

        con = conectar()
        try:
            codigo = str(datos.get("codigo", "")).strip()
            if not codigo:
                return self.error("Falta el código de barras")

            if self.path == "/producto":
                nombre = str(datos.get("nombre", "")).strip()
                if not nombre:
                    return self.error("Falta el nombre")
                try:
                    con.execute(
                        "INSERT INTO productos (nombre, codigo, cantidad, precio) "
                        "VALUES (?, ?, ?, ?)",
                        (nombre, codigo, int(datos.get("cantidad", 0)),
                         float(datos.get("precio", 0))),
                    )
                    con.commit()
                except sqlite3.IntegrityError:
                    return self.error("Ya existe un producto con ese nombre o código", 409)

            elif self.path == "/stock":
                delta = int(datos.get("delta", 0))
                cur = con.execute(
                    "UPDATE productos SET cantidad = cantidad + ? "
                    "WHERE codigo = ? AND cantidad + ? >= 0",
                    (delta, codigo, delta),
                )
                con.commit()
                if not cur.rowcount:
                    return self.error("Producto no encontrado o stock insuficiente", 409)
            else:
                return self.error("Ruta no encontrada", 404)

            self.responder({"ok": True, "producto": producto_por_codigo(con, codigo)})
        except (TypeError, ValueError):
            self.error("Datos no válidos")
        finally:
            con.close()

    def log_message(self, formato, *args):
        print(f"[{self.address_string()}] {formato % args}")


def ip_local():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))  # no envía nada, solo elige la interfaz de red
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def main():
    conectar().close()  # crea o actualiza la base de datos
    servidor = ThreadingHTTPServer(("0.0.0.0", PUERTO), Manejador)
    print("=== SERVIDOR DEL ALMACÉN ===")
    print(f"En la app del móvil escribe:  {ip_local()}:{PUERTO}")
    print("El móvil tiene que estar en la misma red Wi-Fi que este PC.")
    print("Deja esta ventana abierta. Para parar, ciérrala o pulsa Ctrl+C.\n")
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        input("Pulsa Enter para cerrar...")
