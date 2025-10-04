from flask import Flask, render_template, request, redirect, send_file
import openpyxl
import os

app = Flask(__name__)

# Nombre del archivo Excel donde se guardarán los registros
EXCEL_FILE = "clientes.xlsx"

# Crear archivo Excel si no existe
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Clientes"
    sheet.append(["Cédula", "Nombres", "Apellidos", "Dirección", "Teléfono", "Correo"])
    workbook.save(EXCEL_FILE)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/procesar-registro', methods=["POST"])
def procesar_registro():
    cedula = request.form["cedula"]
    nombres = request.form["nombres"]
    apellidos = request.form["apellidos"]
    direccion = request.form["direccion"]
    telefono = request.form["telefono"]
    correo = request.form["correo"]

    # Abrir Excel y guardar datos
    workbook = openpyxl.load_workbook(EXCEL_FILE)
    sheet = workbook.active
    sheet.append([cedula, nombres, apellidos, direccion, telefono, correo])
    workbook.save(EXCEL_FILE)

    return redirect("/")

@app.route('/descargar-excel')
def descargar_excel():
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)