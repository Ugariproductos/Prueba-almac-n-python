package com.ugari.almacen

import android.os.Bundle
import android.view.View
import android.widget.Button
import android.widget.EditText
import android.widget.LinearLayout
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import com.journeyapps.barcodescanner.ScanContract
import com.journeyapps.barcodescanner.ScanOptions
import org.json.JSONObject
import java.io.IOException
import java.net.HttpURLConnection
import java.net.URL
import java.net.URLEncoder
import java.util.Locale
import kotlin.concurrent.thread

class MainActivity : AppCompatActivity() {

    private lateinit var etServidor: EditText
    private lateinit var tvResultado: TextView
    private lateinit var panelExistente: LinearLayout
    private lateinit var panelNuevo: LinearLayout
    private lateinit var etUnidades: EditText
    private lateinit var etNombre: EditText
    private lateinit var etPrecio: EditText
    private lateinit var etCantidad: EditText

    private var codigoActual: String? = null

    private val escaner = registerForActivityResult(ScanContract()) { resultado ->
        resultado.contents?.let { buscar(it) }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        etServidor = findViewById(R.id.etServidor)
        tvResultado = findViewById(R.id.tvResultado)
        panelExistente = findViewById(R.id.panelExistente)
        panelNuevo = findViewById(R.id.panelNuevo)
        etUnidades = findViewById(R.id.etUnidades)
        etNombre = findViewById(R.id.etNombre)
        etPrecio = findViewById(R.id.etPrecio)
        etCantidad = findViewById(R.id.etCantidad)

        val prefs = getSharedPreferences("config", MODE_PRIVATE)
        etServidor.setText(prefs.getString("servidor", ""))

        findViewById<Button>(R.id.btnEscanear).setOnClickListener {
            if (etServidor.text.isBlank()) {
                aviso("Escribe primero la dirección del PC")
                return@setOnClickListener
            }
            prefs.edit().putString("servidor", etServidor.text.toString().trim()).apply()
            escaner.launch(
                ScanOptions()
                    .setPrompt("Apunta al código de barras")
                    .setBeepEnabled(true)
                    .setOrientationLocked(false)
            )
        }
        findViewById<Button>(R.id.btnEntrada).setOnClickListener { moverStock(1) }
        findViewById<Button>(R.id.btnSalida).setOnClickListener { moverStock(-1) }
        findViewById<Button>(R.id.btnAlta).setOnClickListener { darDeAlta() }
    }

    private fun buscar(codigo: String) {
        codigoActual = codigo
        panelExistente.visibility = View.GONE
        panelNuevo.visibility = View.GONE
        tvResultado.text = "Buscando $codigo…"
        val ruta = "/producto?codigo=" + URLEncoder.encode(codigo, "UTF-8")
        peticion("GET", ruta, null) { json ->
            if (json.getBoolean("encontrado")) {
                mostrarProducto(json.getJSONObject("producto"))
            } else {
                tvResultado.text = "El código $codigo no está en el almacén.\nDalo de alta:"
                etNombre.setText("")
                etPrecio.setText("")
                etCantidad.setText("")
                panelNuevo.visibility = View.VISIBLE
            }
        }
    }

    private fun moverStock(signo: Int) {
        val codigo = codigoActual ?: return
        val unidades = etUnidades.text.toString().toIntOrNull()
        if (unidades == null || unidades <= 0) {
            aviso("Indica un número de unidades mayor que 0")
            return
        }
        val cuerpo = JSONObject().put("codigo", codigo).put("delta", signo * unidades)
        peticion("POST", "/stock", cuerpo) { json ->
            mostrarProducto(json.getJSONObject("producto"))
            aviso("Stock actualizado")
        }
    }

    private fun darDeAlta() {
        val codigo = codigoActual ?: return
        val nombre = etNombre.text.toString().trim()
        if (nombre.isEmpty()) {
            aviso("Falta el nombre")
            return
        }
        val precio = etPrecio.text.toString().replace(',', '.').toDoubleOrNull() ?: 0.0
        val cantidad = etCantidad.text.toString().toIntOrNull() ?: 0
        val cuerpo = JSONObject()
            .put("codigo", codigo)
            .put("nombre", nombre)
            .put("precio", precio)
            .put("cantidad", cantidad)
        peticion("POST", "/producto", cuerpo) { json ->
            panelNuevo.visibility = View.GONE
            mostrarProducto(json.getJSONObject("producto"))
            aviso("Producto dado de alta")
        }
    }

    private fun mostrarProducto(p: JSONObject) {
        tvResultado.text = String.format(
            Locale("es", "ES"),
            "%s\nCódigo: %s\nStock: %d\nPrecio: %.2f €",
            p.getString("nombre"), p.getString("codigo"),
            p.getInt("cantidad"), p.getDouble("precio")
        )
        panelExistente.visibility = View.VISIBLE
    }

    private fun direccionBase(): String {
        var s = etServidor.text.toString().trim().removePrefix("http://")
        if (!s.contains(":")) s += ":8765"
        return "http://$s"
    }

    /** Hace la petición al PC en segundo plano y llama a [alTerminar] si la respuesta es ok. */
    private fun peticion(metodo: String, ruta: String, cuerpo: JSONObject?, alTerminar: (JSONObject) -> Unit) {
        val url = direccionBase() + ruta
        thread {
            try {
                val con = URL(url).openConnection() as HttpURLConnection
                con.requestMethod = metodo
                con.connectTimeout = 5000
                con.readTimeout = 5000
                if (cuerpo != null) {
                    con.doOutput = true
                    con.setRequestProperty("Content-Type", "application/json; charset=utf-8")
                    con.outputStream.use { it.write(cuerpo.toString().toByteArray(Charsets.UTF_8)) }
                }
                val flujo = if (con.responseCode < 400) con.inputStream else con.errorStream
                val texto = flujo?.bufferedReader()?.use { it.readText() }
                    ?: throw IOException("HTTP ${con.responseCode}")
                val json = JSONObject(texto)
                runOnUiThread {
                    if (json.optBoolean("ok")) alTerminar(json)
                    else tvResultado.text = "Error: " + json.optString("error", "desconocido")
                }
            } catch (e: Exception) {
                runOnUiThread {
                    tvResultado.text = "No se pudo conectar con el PC ($url).\n" +
                        "¿Está abierto servidor_almacen.py y el móvil en la misma Wi-Fi?\n\n${e.message}"
                }
            }
        }
    }

    private fun aviso(texto: String) = Toast.makeText(this, texto, Toast.LENGTH_SHORT).show()
}
