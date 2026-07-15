import sqlite3
import os
import json
import secrets
from functools import wraps
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, session
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.json.ensure_ascii = False
DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "flota.db")
EXCEL_FILE = "flota_vehicular.xlsx"


@app.template_filter("tojson_filter")
def tojson_filter(obj):
    return json.dumps(dict(obj), ensure_ascii=False, default=str)


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            nombre TEXT,
            rol TEXT NOT NULL DEFAULT 'operador',
            fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS vehiculos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            placa TEXT NOT NULL UNIQUE,
            chasis TEXT,
            motor TEXT,
            marca TEXT,
            modelo TEXT,
            anio INTEGER,
            kilometraje INTEGER DEFAULT 0,
            ubicacion TEXT,
            estado TEXT DEFAULT 'Activo',
            creado_por TEXT,
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS mantenimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vehiculo_id INTEGER NOT NULL,
            fecha TEXT NOT NULL,
            tipo TEXT,
            descripcion TEXT,
            kilometraje INTEGER,
            costo REAL DEFAULT 0,
            taller TEXT,
            FOREIGN KEY (vehiculo_id) REFERENCES vehiculos(id) ON DELETE CASCADE
        )
    """)
    admin = conn.execute("SELECT id FROM usuarios WHERE rol='admin'").fetchone()
    if not admin:
        conn.execute(
            "INSERT INTO usuarios (usuario, password, nombre, rol) VALUES (?, ?, ?, ?)",
            ("admin", generate_password_hash("Willian098"), "Administrador", "admin")
        )
    try:
        conn.execute("ALTER TABLE vehiculos ADD COLUMN creado_por TEXT")
    except sqlite3.OperationalError:
        pass
    try:
        conn.execute("ALTER TABLE vehiculos ADD COLUMN mecanica TEXT DEFAULT 'Multimarcas'")
    except sqlite3.OperationalError:
        pass
    conn.commit()
    conn.close()


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        if session.get("rol") != "admin":
            return redirect(url_for("index"))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        usuario = request.form["usuario"].strip()
        password = request.form["password"]
        conn = get_db()
        user = conn.execute("SELECT * FROM usuarios WHERE usuario=?", (usuario,)).fetchone()
        conn.close()
        if user and check_password_hash(user["password"], password):
            session["user_id"] = user["id"]
            session["usuario"] = user["usuario"]
            session["nombre"] = user["nombre"]
            session["rol"] = user["rol"]
            return redirect(url_for("index"))
        error = "Usuario o contrasena incorrectos"
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/usuarios")
@admin_required
def gestionar_usuarios():
    conn = get_db()
    usuarios = conn.execute("SELECT * FROM usuarios ORDER BY id").fetchall()
    conn.close()
    return render_template("usuarios.html", usuarios=usuarios)


@app.route("/usuarios/crear", methods=["POST"])
@admin_required
def crear_usuario():
    usuario = request.form["usuario"].strip()
    password = request.form["password"]
    nombre = request.form["nombre"].strip()
    rol = request.form["rol"]
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO usuarios (usuario, password, nombre, rol) VALUES (?, ?, ?, ?)",
            (usuario, generate_password_hash(password), nombre, rol)
        )
        conn.commit()
    except sqlite3.IntegrityError:
        pass
    conn.close()
    return redirect(url_for("gestionar_usuarios"))


@app.route("/usuarios/eliminar/<int:id>")
@admin_required
def eliminar_usuario(id):
    if id == session.get("user_id"):
        return redirect(url_for("gestionar_usuarios"))
    conn = get_db()
    conn.execute("DELETE FROM usuarios WHERE id=? AND rol != 'admin'", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for("gestionar_usuarios"))


@app.route("/")
@login_required
def index():
    conn = get_db()
    mecanica_filter = request.args.get("mecanica", "")
    if mecanica_filter:
        vehiculos = conn.execute("SELECT * FROM vehiculos WHERE mecanica=? ORDER BY id DESC", (mecanica_filter,)).fetchall()
    else:
        vehiculos = conn.execute("SELECT * FROM vehiculos ORDER BY id DESC").fetchall()
    mecanicas = conn.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica").fetchall()
    conn.close()
    return render_template("index.html", vehiculos=vehiculos, mecanicas=mecanicas, mecanica_actual=mecanica_filter)


@app.route("/agregar", methods=["POST"])
@login_required
def agregar():
    conn = get_db()
    conn.execute("""
        INSERT INTO vehiculos (placa, chasis, motor, marca, modelo, anio, kilometraje, ubicacion, estado, creado_por, mecanica)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        request.form["placa"].upper(),
        request.form["chasis"],
        request.form["motor"],
        request.form["marca"],
        request.form["modelo"],
        request.form["anio"],
        request.form["kilometraje"],
        request.form["ubicacion"],
        request.form["estado"],
        session.get("usuario", ""),
        request.form.get("mecanica", "Multimarcas")
    ))
    conn.commit()
    conn.close()
    return redirect(url_for("index"))


@app.route("/editar/<int:id>", methods=["POST"])
@admin_required
def editar(id):
    conn = get_db()
    conn.execute("""
        UPDATE vehiculos
        SET placa=?, chasis=?, motor=?, marca=?, modelo=?, anio=?, kilometraje=?, ubicacion=?, estado=?, mecanica=?
        WHERE id=?
    """, (
        request.form["placa"].upper(),
        request.form["chasis"],
        request.form["motor"],
        request.form["marca"],
        request.form["modelo"],
        request.form["anio"],
        request.form["kilometraje"],
        request.form["ubicacion"],
        request.form["estado"],
        request.form.get("mecanica", "Multimarcas"),
        id
    ))
    conn.commit()
    conn.close()
    return redirect(url_for("index"))


@app.route("/eliminar/<int:id>")
@admin_required
def eliminar(id):
    conn = get_db()
    conn.execute("DELETE FROM vehiculos WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for("index"))


@app.route("/vehiculo/<int:id>")
@login_required
def detalle(id):
    conn = get_db()
    vehiculo = conn.execute("SELECT * FROM vehiculos WHERE id=?", (id,)).fetchone()
    historial = conn.execute(
        "SELECT * FROM mantenimientos WHERE vehiculo_id=? ORDER BY fecha DESC", (id,)
    ).fetchall()
    conn.close()
    return render_template("detalle.html", vehiculo=vehiculo, historial=historial)


@app.route("/mantenimiento/<int:vehiculo_id>", methods=["POST"])
@login_required
def agregar_mantenimiento(vehiculo_id):
    conn = get_db()
    conn.execute("""
        INSERT INTO mantenimientos (vehiculo_id, fecha, tipo, descripcion, kilometraje, costo, taller)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        vehiculo_id,
        request.form["fecha"],
        request.form["tipo"],
        request.form["descripcion"],
        request.form["kilometraje"],
        request.form["costo"],
        request.form["taller"]
    ))
    conn.execute("UPDATE vehiculos SET kilometraje=? WHERE id=?", (request.form["kilometraje"], vehiculo_id))
    conn.commit()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/mantenimiento/editar/<int:id>", methods=["POST"])
@admin_required
def editar_mantenimiento(id):
    conn = get_db()
    mant = conn.execute("SELECT vehiculo_id FROM mantenimientos WHERE id=?", (id,)).fetchone()
    if not mant:
        conn.close()
        return redirect(url_for("index"))
    vehiculo_id = mant["vehiculo_id"]
    conn.execute("""
        UPDATE mantenimientos
        SET fecha=?, tipo=?, descripcion=?, kilometraje=?, costo=?, taller=?
        WHERE id=?
    """, (
        request.form["fecha"],
        request.form["tipo"],
        request.form["descripcion"],
        request.form["kilometraje"],
        request.form["costo"],
        request.form["taller"],
        id
    ))
    conn.execute("UPDATE vehiculos SET kilometraje=? WHERE id=?", (request.form["kilometraje"], vehiculo_id))
    conn.commit()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/mantenimiento/eliminar/<int:id>")
@admin_required
def eliminar_mantenimiento(id):
    conn = get_db()
    mant = conn.execute("SELECT vehiculo_id FROM mantenimientos WHERE id=?", (id,)).fetchone()
    if not mant:
        conn.close()
        return redirect(url_for("index"))
    vehiculo_id = mant["vehiculo_id"]
    conn.execute("DELETE FROM mantenimientos WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/buscar")
@login_required
def buscar():
    q = request.args.get("q", "")
    mecanica_filter = request.args.get("mecanica", "")
    conn = get_db()
    query = "SELECT * FROM vehiculos WHERE (placa LIKE ? OR marca LIKE ? OR ubicacion LIKE ? OR chasis LIKE ?)"
    params = [f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%"]
    if mecanica_filter:
        query += " AND mecanica=?"
        params.append(mecanica_filter)
    query += " ORDER BY id DESC"
    vehiculos = conn.execute(query, params).fetchall()
    mecanicas = conn.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica").fetchall()
    conn.close()
    return render_template("index.html", vehiculos=vehiculos, busqueda=q, mecanicas=mecanicas, mecanica_actual=mecanica_filter)


@app.route("/api/estadisticas")
@login_required
def estadisticas():
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) as c FROM vehiculos").fetchone()["c"]
    activos = conn.execute("SELECT COUNT(*) as c FROM vehiculos WHERE estado='Activo'").fetchone()["c"]
    inactivos = conn.execute("SELECT COUNT(*) as c FROM vehiculos WHERE estado='Inactivo'").fetchone()["c"]
    taller = conn.execute("SELECT COUNT(*) as c FROM vehiculos WHERE estado='En Taller'").fetchone()["c"]
    remate = conn.execute("SELECT COUNT(*) as c FROM vehiculos WHERE estado='Remate'").fetchone()["c"]
    conn.close()
    return jsonify({
        "total": total,
        "activos": activos,
        "inactivos": inactivos,
        "en_taller": taller,
        "remate": remate
    })


@app.route("/exportar-excel")
@login_required
def exportar():
    buffer = exportar_excel()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"flota_vehicular_{timestamp}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/exportar-vehiculo/<int:id>")
@login_required
def exportar_vehiculo(id):
    conn = get_db()
    vehiculo = conn.execute("SELECT * FROM vehiculos WHERE id=?", (id,)).fetchone()
    if not vehiculo:
        conn.close()
        return redirect(url_for("index"))
    historial = conn.execute(
        "SELECT * FROM mantenimientos WHERE vehiculo_id=? ORDER BY fecha DESC", (id,)
    ).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = vehiculo["placa"]
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    fill_blue = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    fill_green = PatternFill(start_color="22C55E", end_color="22C55E", fill_type="solid")

    info = [
        ("Placa", vehiculo["placa"]), ("Marca", vehiculo["marca"]),
        ("Modelo", vehiculo["modelo"]), ("Anio", vehiculo["anio"]),
        ("Chasis", vehiculo["chasis"]), ("Motor", vehiculo["motor"]),
        ("Kilometraje", vehiculo["kilometraje"]), ("Ubicacion", vehiculo["ubicacion"]),
        ("Estado", vehiculo["estado"]), ("Mecanica", vehiculo["mecanica"])
    ]
    for row_idx, (label, val) in enumerate(info, 1):
        cell_label = ws.cell(row=row_idx, column=1, value=label)
        cell_label.font = Font(bold=True, color="FFFFFF")
        cell_label.fill = fill_blue
        cell_label.border = thin_border
        cell_val = ws.cell(row=row_idx, column=2, value=val or "-")
        cell_val.border = thin_border
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 30

    mant_row = len(info) + 3
    ws.cell(row=mant_row, column=1, value="HISTORIAL DE MANTENIMIENTO").font = Font(bold=True, size=14, color="22C55E")
    ws.merge_cells(start_row=mant_row, start_column=1, end_row=mant_row, end_column=7)
    mant_row += 1

    mant_headers = ["Fecha", "Tipo", "Descripcion", "Kilometraje", "Costo", "Taller"]
    for col, h in enumerate(mant_headers, 1):
        cell = ws.cell(row=mant_row, column=col, value=h)
        cell.font = header_font
        cell.fill = fill_green
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
    mant_row += 1

    for m in historial:
        for col, key in enumerate(["fecha", "tipo", "descripcion", "kilometraje", "costo", "taller"], 1):
            val = m[key] if m[key] else "-"
            cell = ws.cell(row=mant_row, column=col, value=val)
            cell.border = thin_border
            if key == "costo":
                cell.number_format = "#,##0.00"
            if key == "kilometraje":
                cell.number_format = "#,##0"
        mant_row += 1

    if not historial:
        ws.cell(row=mant_row, column=1, value="Sin registros de mantenimiento").font = Font(italic=True, color="64748B")

    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 25

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{vehiculo['placa']}_historial.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/importar-excel", methods=["POST"])
@admin_required
def importar():
    if "archivo" not in request.files:
        return redirect(url_for("index"))
    file = request.files["archivo"]
    if file.filename == "":
        return redirect(url_for("index"))
    if not file.filename.endswith((".xlsx", ".xls")):
        return redirect(url_for("index"))
    filepath = os.path.join("uploads", file.filename)
    os.makedirs("uploads", exist_ok=True)
    file.save(filepath)
    imported, updated = importar_excel(filepath)
    os.remove(filepath)
    return redirect(url_for("index"))


@app.route("/sync-excel", methods=["POST"])
@login_required
def sync_excel():
    buffer = exportar_excel()
    filepath = os.path.join(os.path.dirname(os.path.abspath(DB)), EXCEL_FILE)
    with open(filepath, "wb") as f:
        f.write(buffer.getvalue())
    return jsonify({"status": "ok", "file": filepath, "message": f"Excel guardado en: {filepath}"})


def exportar_excel():
    conn = get_db()
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    veh_headers = ["ID", "Placa", "Chasis", "Motor", "Marca", "Modelo", "Anio", "Kilometraje", "Ubicacion", "Estado", "Mecanica", "Creado Por", "Fecha Registro"]
    veh_cols = ["id", "placa", "chasis", "motor", "marca", "modelo", "anio", "kilometraje", "ubicacion", "estado", "mecanica", "creado_por", "fecha_registro"]
    mant_headers = ["ID", "Vehiculo ID", "Placa", "Fecha", "Tipo", "Descripcion", "Kilometraje", "Costo", "Taller"]
    mant_cols = ["id", "vehiculo_id", "placa", "fecha", "tipo", "descripcion", "kilometraje", "costo", "taller"]
    fill_mant = PatternFill(start_color="22C55E", end_color="22C55E", fill_type="solid")

    TODAS_LAS_MECANICAS = [
        "Multimarcas", "Kia", "Ambacar", "Chevrolet", "Pesados",
        "Great Wall 2025", "Great Wall 2026", "Motos Multimarcas",
        "Honda", "Morini", "Kia Tasman"
    ]

    vehiculos = conn.execute("SELECT * FROM vehiculos ORDER BY id").fetchall()
    ws_all = wb.active
    ws_all.title = "Todos los Vehiculos"
    fill_all = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    for col, h in enumerate(veh_headers, 1):
        cell = ws_all.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = fill_all
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
    for row_idx, v in enumerate(vehiculos, 2):
        for col_idx, key in enumerate(veh_cols, 1):
            try:
                val = v[key]
            except (IndexError, KeyError):
                val = ""
            cell = ws_all.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if key == "kilometraje":
                cell.number_format = "#,##0"
    for col in range(1, len(veh_headers) + 1):
        ws_all.column_dimensions[get_column_letter(col)].width = 18

    colores = ["16A34A", "EA580C", "7C3AED", "DC2626", "0891B2", "CA8A04", "DB2777", "4F46E5", "059669", "9333EA", "B91C1C"]

    for idx, nombre in enumerate(TODAS_LAS_MECANICAS):
        ws = wb.create_sheet(title=nombre[:31])
        fill = PatternFill(start_color=colores[idx], end_color=colores[idx], fill_type="solid")

        vehiculos_mec = conn.execute(
            "SELECT * FROM vehiculos WHERE mecanica=? ORDER BY id", (nombre,)
        ).fetchall()

        ws.cell(row=1, column=1, value=f"VEHICULOS - {nombre.upper()}").font = Font(bold=True, size=14, color=colores[idx])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(veh_headers))

        for col, h in enumerate(veh_headers, 1):
            cell = ws.cell(row=2, column=col, value=h)
            cell.font = header_font
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border

        row_start = 3
        for row_idx, v in enumerate(vehiculos_mec):
            for col_idx, key in enumerate(veh_cols, 1):
                try:
                    val = v[key]
                except (IndexError, KeyError):
                    val = ""
                cell = ws.cell(row=row_start + row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if key == "kilometraje":
                    cell.number_format = "#,##0"

        if not vehiculos_mec:
            ws.cell(row=3, column=1, value="Sin vehiculos registrados").font = Font(italic=True, color="64748B")

        mant_row = row_start + len(vehiculos_mec) + 2
        mant_mec = conn.execute("""
            SELECT m.*, v.placa FROM mantenimientos m
            LEFT JOIN vehiculos v ON m.vehiculo_id = v.id
            WHERE v.mecanica=?
            ORDER BY m.fecha DESC
        """, (nombre,)).fetchall()

        ws.cell(row=mant_row, column=1, value=f"HISTORIAL MANTENIMIENTO - {nombre.upper()}").font = Font(bold=True, size=14, color="22C55E")
        ws.merge_cells(start_row=mant_row, start_column=1, end_row=mant_row, end_column=len(mant_headers))
        mant_row += 1

        for col, h in enumerate(mant_headers, 1):
            cell = ws.cell(row=mant_row, column=col, value=h)
            cell.font = header_font
            cell.fill = fill_mant
            cell.alignment = Alignment(horizontal="center")
            cell.border = thin_border
        mant_row += 1

        for row_idx, m in enumerate(mant_mec):
            for col_idx, key in enumerate(mant_cols, 1):
                try:
                    val = m[key]
                except (IndexError, KeyError):
                    val = ""
                cell = ws.cell(row=mant_row + row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if key == "costo":
                    cell.number_format = "#,##0.00"

        if not mant_mec:
            ws.cell(row=mant_row, column=1, value="Sin mantenimientos registrados").font = Font(italic=True, color="64748B")

        for col in range(1, max(len(veh_headers), len(mant_headers)) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18

    ws_mant_all = wb.create_sheet("Todos Mantenimientos")
    for col, h in enumerate(mant_headers, 1):
        cell = ws_mant_all.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = fill_mant
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
    mant_all = conn.execute("""
        SELECT m.*, v.placa FROM mantenimientos m
        LEFT JOIN vehiculos v ON m.vehiculo_id = v.id
        ORDER BY m.fecha DESC
    """).fetchall()
    for row_idx, m in enumerate(mant_all, 2):
        for col_idx, key in enumerate(mant_cols, 1):
            try:
                val = m[key]
            except (IndexError, KeyError):
                val = ""
            cell = ws_mant_all.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if key == "costo":
                cell.number_format = "#,##0.00"
    for col in range(1, len(mant_headers) + 1):
        ws_mant_all.column_dimensions[get_column_letter(col)].width = 20

    conn.close()
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def importar_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb["Vehiculos"]
    conn = get_db()
    imported = 0
    updated = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        placa = str(row[1]).upper().strip()
        existing = conn.execute("SELECT id FROM vehiculos WHERE placa=?", (placa,)).fetchone()
        data = {
            "chasis": row[2] or "", "motor": row[3] or "", "marca": row[4] or "",
            "modelo": row[5] or "", "anio": row[6] or 0, "kilometraje": row[7] or 0,
            "ubicacion": row[8] or "", "estado": row[9] or "Activo", "mecanica": row[10] or "Multimarcas"
        }
        if existing:
            conn.execute("""
                UPDATE vehiculos SET chasis=?, motor=?, marca=?, modelo=?, anio=?, kilometraje=?, ubicacion=?, estado=?, mecanica=?
                WHERE placa=?
            """, (data["chasis"], data["motor"], data["marca"], data["modelo"], data["anio"], data["kilometraje"], data["ubicacion"], data["estado"], data["mecanica"], placa))
            updated += 1
        else:
            conn.execute("""
                INSERT INTO vehiculos (placa, chasis, motor, marca, modelo, anio, kilometraje, ubicacion, estado, mecanica)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (placa, data["chasis"], data["motor"], data["marca"], data["modelo"], data["anio"], data["kilometraje"], data["ubicacion"], data["estado"], data["mecanica"]))
            imported += 1
    if "Mantenimientos" in wb.sheetnames:
        ws2 = wb["Mantenimientos"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if not row[1]:
                continue
            vehiculo_id = row[1]
            v = conn.execute("SELECT id FROM vehiculos WHERE id=?", (vehiculo_id,)).fetchone()
            if v:
                conn.execute("""
                    INSERT INTO mantenimientos (vehiculo_id, fecha, tipo, descripcion, kilometraje, costo, taller)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (vehiculo_id, str(row[3] or ""), row[4] or "", row[5] or "", row[6] or 0, row[7] or 0, row[8] or ""))
    conn.commit()
    conn.close()
    return imported, updated


init_db()

@app.route("/health")
def health():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
