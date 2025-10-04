}from flask import Flask, render_template, request, redirect, send_file
import openpyxl
import os

app = Flask(__name__)

# Nombre del archivo Excel
excel_file = "clientes.xlsx"

# Si no existe, lo crea con encabezados
if not os.path.exists(excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Clientes"
    sheet.append(["Cédula", "Nombres", "Apellidos", "Dirección", "Teléfono", "Correo"])
    workbook.save(excel_file)

@app.route("/")
def formulario():
    return render_template("registro.html")

@app.route("/procesar-registro", methods=["POST"])
def procesar_registro():
    cedula = request.form["cedula"]
    nombres = request.form["nombres"]
    apellidos = request.form["apellidos"]
    direccion = request.form["direccion"]
    telefono = request.form["telefono"]
    correo = request.form["correo"]

    # Abrir y guardar datos en Excel
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    sheet.append([cedula, nombres, apellidos, direccion, telefono, correo])
    workbook.save(excel_file)

    return redirect("/")

# ✅ Nueva ruta para descargar el Excel
@app.route("/descargar-excel")
def descargar_excel():
    return send_file(excel_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
