import os
import json
import secrets
import psycopg2
import psycopg2.extras
from functools import wraps
from io import BytesIO
from urllib.parse import urlparse, unquote
from flask import Flask, render_template, request, redirect, url_for, jsonify, send_file, session, flash, get_flashed_messages
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta, timezone
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.json.ensure_ascii = False
DATABASE_URL = os.environ.get("DATABASE_URL", "")
EXCEL_FILE = "flota_vehicular.xlsx"


ECUADOR_TZ = timezone(timedelta(hours=-5))

def fecha_local():
    return datetime.now(ECUADOR_TZ).strftime("%Y-%m-%d %H:%M:%S")


def fecha_local_date():
    return datetime.now(ECUADOR_TZ).date()


@app.template_filter("tojson_filter")
def tojson_filter(obj):
    return json.dumps(dict(obj), ensure_ascii=False, default=str)


def get_db():
    url = urlparse(DATABASE_URL)
    conn = psycopg2.connect(
        database=unquote(url.path[1:]),
        user=unquote(url.username or ''),
        password=unquote(url.password or ''),
        host=url.hostname,
        port=url.port,
        sslmode='require',
        connect_timeout=15,
        options="-c statement_timeout=30000 -c idle_in_transaction_session_timeout=30000"
    )
    conn.autocommit = False
    cur = conn.cursor()
    cur.execute("SET timezone = 'America/Guayaquil'")
    cur.close()
    return conn


def dict_row(cursor):
    colnames = [d[0] for d in cursor.description]
    row = cursor.fetchone()
    if row:
        return dict(zip(colnames, row))
    return None


def dict_rows(cursor):
    rows = cursor.fetchall()
    if rows and isinstance(rows[0], dict):
        return [dict(r) for r in rows]
    colnames = [d[0] for d in cursor.description]
    return [dict(zip(colnames, r)) for r in rows]


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            usuario TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            nombre TEXT,
            rol TEXT NOT NULL DEFAULT 'operador',
            fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS vehiculos (
            id SERIAL PRIMARY KEY,
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
            mecanica TEXT DEFAULT 'Multimarcas',
            fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS mantenimientos (
            id SERIAL PRIMARY KEY,
            vehiculo_id INTEGER NOT NULL,
            fecha TEXT NOT NULL,
            tipo TEXT,
            descripcion TEXT,
            kilometraje INTEGER,
            costo REAL DEFAULT 0,
            costo_mo REAL DEFAULT 0,
            costo_lub REAL DEFAULT 0,
            costo_rep REAL DEFAULT 0,
            taller TEXT,
            FOREIGN KEY (vehiculo_id) REFERENCES vehiculos(id) ON DELETE CASCADE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS piezas (
            id SERIAL PRIMARY KEY,
            nombre TEXT NOT NULL,
            tipo TEXT DEFAULT 'repuesto',
            vida_km INTEGER DEFAULT 0,
            vida_meses INTEGER DEFAULT 0,
            costo REAL DEFAULT 0,
            mecanica TEXT DEFAULT 'Multimarcas'
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS piezas_vehiculo (
            id SERIAL PRIMARY KEY,
            vehiculo_id INTEGER NOT NULL,
            placa TEXT,
            pieza_id INTEGER,
            nombre TEXT,
            tipo TEXT DEFAULT 'repuesto',
            fecha_ingreso TEXT,
            km_ingreso INTEGER DEFAULT 0,
            vida_km INTEGER DEFAULT 0,
            vida_meses INTEGER DEFAULT 0,
            costo REAL DEFAULT 0,
            mantenimiento_id INTEGER,
            observaciones TEXT,
            FOREIGN KEY (vehiculo_id) REFERENCES vehiculos(id) ON DELETE CASCADE
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS parte_taller (
            id SERIAL PRIMARY KEY,
            vehiculo_id INTEGER NOT NULL,
            placa TEXT NOT NULL,
            mecanica TEXT,
            motivo TEXT,
            nombre_ingresa TEXT,
            telefono TEXT,
            novedad TEXT,
            observaciones TEXT,
            fecha_ingreso TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            fecha_salida TIMESTAMP,
            registrado_por TEXT,
            FOREIGN KEY (vehiculo_id) REFERENCES vehiculos(id) ON DELETE CASCADE
        )
    """)
    cur.execute("SELECT id FROM usuarios WHERE rol='admin'")
    admin = cur.fetchone()
    if not admin:
        cur.execute(
            "INSERT INTO usuarios (usuario, password, nombre, rol) VALUES (%s, %s, %s, %s)",
            ("admin", generate_password_hash("Willian098"), "Administrador", "admin")
        )
    conn.commit()
    cur.close()
    conn.close()


def add_columns_parte_taller():
    conn = get_db()
    cur = conn.cursor()
    for col in ['nombre_ingresa TEXT', 'telefono TEXT', 'novedad TEXT']:
        try:
            cur.execute(f"ALTER TABLE parte_taller ADD COLUMN {col}")
        except Exception:
            pass
    for col in ['costo_mo REAL DEFAULT 0', 'costo_lub REAL DEFAULT 0', 'costo_rep REAL DEFAULT 0']:
        try:
            cur.execute(f"ALTER TABLE mantenimientos ADD COLUMN {col}")
        except Exception:
            pass
    for col in ['tipo TEXT DEFAULT \'repuesto\'']:
        try:
            cur.execute(f"ALTER TABLE piezas ADD COLUMN {col}")
        except Exception:
            pass
    try:
        cur.execute("ALTER TABLE piezas ADD COLUMN mecanica TEXT DEFAULT 'Multimarcas'")
    except Exception:
        pass
    try:
        cur.execute("ALTER TABLE piezas_vehiculo ADD COLUMN observaciones TEXT")
    except Exception:
        pass
    try:
        cur.execute("ALTER TABLE piezas ALTER COLUMN vehiculo_id DROP NOT NULL")
    except Exception:
        pass
    try:
        cur.execute("""
            UPDATE piezas SET vehiculo_id = NULL, placa = NULL,
            km_instalacion = 0, fecha_instalacion = NULL, observaciones = NULL
            WHERE vehiculo_id IS NOT NULL
        """)
    except Exception:
        pass
    conn.commit()
    cur.close()
    conn.close()


def add_columns_contratos():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS contratos (
                id SERIAL PRIMARY KEY,
                numero_contrato TEXT UNIQUE NOT NULL,
                administrador TEXT NOT NULL,
                fecha_creacion DATE DEFAULT CURRENT_DATE,
                fecha_inicio DATE DEFAULT CURRENT_DATE,
                dias_vigencia INTEGER NOT NULL,
                fecha_fin DATE NOT NULL,
                monto_rep REAL DEFAULT 0,
                monto_lub REAL DEFAULT 0,
                monto_mo REAL DEFAULT 0,
                total_contrato REAL DEFAULT 0,
                saldo_rep REAL DEFAULT 0,
                saldo_lub REAL DEFAULT 0,
                saldo_mo REAL DEFAULT 0,
                saldo_total REAL DEFAULT 0,
                estado TEXT DEFAULT 'ACTIVO',
                motivo_finalizacion TEXT,
                finalizado_por TEXT,
                fecha_finalizacion TIMESTAMP
            )
        """)
    except Exception:
        pass
    try:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS contrato_movimientos (
                id SERIAL PRIMARY KEY,
                contrato_id INTEGER REFERENCES contratos(id),
                parte_taller_id INTEGER,
                vehiculo_id INTEGER,
                placa TEXT,
                fecha TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                partida TEXT,
                descripcion TEXT,
                monto_descontado REAL DEFAULT 0,
                saldo_rep_antes REAL DEFAULT 0,
                saldo_rep_despues REAL DEFAULT 0,
                saldo_lub_antes REAL DEFAULT 0,
                saldo_lub_despues REAL DEFAULT 0,
                saldo_mo_antes REAL DEFAULT 0,
                saldo_mo_despues REAL DEFAULT 0
            )
        """)
    except Exception:
        pass
    try:
        cur.execute("""
            CREATE TABLE IF NOT EXISTS contrato_historial_finalizado (
                id SERIAL PRIMARY KEY,
                contrato_id INTEGER,
                numero_contrato TEXT,
                administrador TEXT,
                fecha_creacion DATE,
                fecha_inicio DATE,
                fecha_fin_original DATE,
                fecha_finalizacion TIMESTAMP,
                motivo_finalizacion TEXT,
                finalizado_por TEXT,
                dias_vigencia INTEGER,
                monto_rep REAL DEFAULT 0,
                monto_lub REAL DEFAULT 0,
                monto_mo REAL DEFAULT 0,
                total_contrato REAL DEFAULT 0,
                saldo_rep REAL DEFAULT 0,
                saldo_lub REAL DEFAULT 0,
                saldo_mo REAL DEFAULT 0,
                saldo_total REAL DEFAULT 0,
                resumen_trabajos TEXT,
                totales_por_partida TEXT
            )
        """)
    except Exception:
        pass
    for col in ['contrato_id INTEGER']:
        try:
            cur.execute(f"ALTER TABLE parte_taller ADD COLUMN {col}")
        except Exception:
            pass
    conn.commit()
    cur.close()
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
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute("SELECT * FROM usuarios WHERE usuario=%s", (usuario,))
        user = cur.fetchone()
        cur.close()
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
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM usuarios ORDER BY id")
    usuarios = cur.fetchall()
    cur.close()
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
    cur = conn.cursor()
    try:
        cur.execute(
            "INSERT INTO usuarios (usuario, password, nombre, rol) VALUES (%s, %s, %s, %s)",
            (usuario, generate_password_hash(password), nombre, rol)
        )
        conn.commit()
    except psycopg2.IntegrityError:
        conn.rollback()
    cur.close()
    conn.close()
    return redirect(url_for("gestionar_usuarios"))


@app.route("/usuarios/eliminar/<int:id>")
@admin_required
def eliminar_usuario(id):
    if id == session.get("user_id"):
        return redirect(url_for("gestionar_usuarios"))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM usuarios WHERE id=%s AND rol != 'admin'", (id,))
    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("gestionar_usuarios"))


@app.route("/")
@login_required
def index():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    mecanica_filter = request.args.get("mecanica", "")
    if mecanica_filter:
        cur.execute("SELECT * FROM vehiculos WHERE mecanica=%s ORDER BY id DESC", (mecanica_filter,))
    else:
        cur.execute("SELECT * FROM vehiculos ORDER BY id DESC")
    vehiculos = cur.fetchall()
    cur.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica")
    mecanicas = cur.fetchall()
    hoy_idx = fecha_local_date()
    alertas_contratos = []
    try:
        cur.execute("SELECT * FROM contratos WHERE estado='ACTIVO' ORDER BY fecha_fin")
        for c in cur.fetchall():
            fecha_fin_c = c.get('fecha_fin')
            try:
                if isinstance(fecha_fin_c, str):
                    from datetime import datetime as dt
                    ffc = dt.strptime(fecha_fin_c, '%Y-%m-%d').date()
                else:
                    ffc = fecha_fin_c
                dias_rest = (ffc - hoy_idx).days
                if dias_rest <= 30 and dias_rest >= 0:
                    alertas_contratos.append(dict(c, tipo_alerta='vence', dias_restantes=dias_rest))
                elif dias_rest < 0:
                    alertas_contratos.append(dict(c, tipo_alerta='vencido', dias_restantes=dias_rest))
            except Exception:
                pass
            if float(c.get('saldo_rep') or 0) < 1000:
                alertas_contratos.append(dict(c, tipo_alerta='saldo_rep_bajo', partida='Repuestos'))
            if float(c.get('saldo_lub') or 0) < 1000:
                alertas_contratos.append(dict(c, tipo_alerta='saldo_lub_bajo', partida='Lubricantes'))
            if float(c.get('saldo_mo') or 0) < 1000:
                alertas_contratos.append(dict(c, tipo_alerta='saldo_mo_bajo', partida='Mano de Obra'))
    except Exception:
        pass
    cur.close()
    conn.close()
    return render_template("index.html", vehiculos=vehiculos, mecanicas=mecanicas, mecanica_actual=mecanica_filter, alertas_contratos=alertas_contratos)


@app.route("/agregar", methods=["POST"])
@login_required
def agregar():
    conn = get_db()
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO vehiculos (placa, chasis, motor, marca, modelo, anio, kilometraje, ubicacion, estado, creado_por, mecanica, fecha_registro)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
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
            request.form.get("mecanica", "Multimarcas"),
            fecha_local()
        ))
        conn.commit()
    except psycopg2.IntegrityError:
        conn.rollback()
    finally:
        cur.close()
        conn.close()
    return redirect(url_for("index"))


@app.route("/editar/<int:id>", methods=["POST"])
@admin_required
def editar(id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE vehiculos
        SET placa=%s, chasis=%s, motor=%s, marca=%s, modelo=%s, anio=%s, kilometraje=%s, ubicacion=%s, estado=%s, mecanica=%s
        WHERE id=%s
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
    cur.close()
    conn.close()
    return redirect(url_for("index"))


@app.route("/eliminar/<int:id>")
@admin_required
def eliminar(id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM vehiculos WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("index"))


@app.route("/vehiculo/<int:id>")
@login_required
def detalle(id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM vehiculos WHERE id=%s", (id,))
    vehiculo = cur.fetchone()
    cur.execute("SELECT * FROM mantenimientos WHERE vehiculo_id=%s ORDER BY fecha DESC", (id,))
    historial = cur.fetchall()
    cur.execute("SELECT * FROM piezas WHERE mecanica=%s ORDER BY nombre ASC", (vehiculo["mecanica"] or "Multimarcas",))
    catalogo = cur.fetchall()

    cur.execute("SELECT * FROM piezas_vehiculo WHERE vehiculo_id=%s ORDER BY id DESC", (id,))
    piezas_veh = cur.fetchall()
    ahora = datetime.now(ECUADOR_TZ)
    for p in piezas_veh:
        km_recorridos = (vehiculo["kilometraje"] or 0) - (p.get("km_ingreso") or 0)
        km_faltan = (p.get("vida_km") or 0) - km_recorridos
        p["km_recorridos"] = max(km_recorridos, 0)
        p["km_faltan"] = km_faltan
        meses_transcurridos = 9999
        if p.get("fecha_ingreso"):
            try:
                fi = datetime.strptime(str(p["fecha_ingreso"])[:10], "%Y-%m-%d")
                dias = (ahora - fi.replace(tzinfo=ECUADOR_TZ)).days
                meses_transcurridos = dias / 30.44
            except Exception:
                meses_transcurridos = 0
        p["meses_faltan"] = round((p.get("vida_meses") or 0) - meses_transcurridos, 1) if meses_transcurridos != 9999 else None
        vencida = km_faltan <= 0 or (p.get("vida_meses") and p["meses_faltan"] is not None and p["meses_faltan"] <= 0)
        por_vencer = (0 < km_faltan <= (p.get("vida_km") or 0) * 0.15) or (p["meses_faltan"] is not None and 0 < p["meses_faltan"] <= (p.get("vida_meses") or 0) * 0.15)
        if vencida:
            p["estado"] = "Vencida"
        elif por_vencer:
            p["estado"] = "Por vencer"
        else:
            p["estado"] = "En vida util"

    cur.close()
    conn.close()
    return render_template("detalle.html", vehiculo=vehiculo, historial=historial, catalogo=catalogo, piezas_veh=piezas_veh)


@app.route("/mantenimiento/<int:vehiculo_id>", methods=["POST"])
@login_required
def agregar_mantenimiento(vehiculo_id):
    conn = get_db()
    cur = conn.cursor()
    costo_mo = float(request.form.get("costo_mo") or 0)
    costo_lub = float(request.form.get("costo_lub") or 0)
    costo_rep = float(request.form.get("costo_rep") or 0)
    costo_total = costo_mo + costo_lub + costo_rep
    cur.execute("""
        INSERT INTO mantenimientos (vehiculo_id, fecha, tipo, descripcion, kilometraje, costo, costo_mo, costo_lub, costo_rep, taller)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    """, (
        vehiculo_id,
        request.form["fecha"],
        request.form["tipo"],
        request.form["descripcion"],
        request.form["kilometraje"],
        costo_total,
        costo_mo,
        costo_lub,
        costo_rep,
        request.form["taller"]
    ))
    cur.execute("UPDATE vehiculos SET kilometraje=%s WHERE id=%s", (request.form["kilometraje"], vehiculo_id))

    cur.execute("SELECT placa FROM vehiculos WHERE id=%s", (vehiculo_id,))
    rp = cur.fetchone()
    placa = rp[0] if rp else ""

    km_actual_mant = int(request.form.get("kilometraje") or 0)

    cur.execute("SELECT id, pieza_id, nombre, tipo, fecha_ingreso, km_ingreso, vida_km, vida_meses FROM piezas_vehiculo WHERE vehiculo_id=%s AND pieza_id IS NOT NULL", (vehiculo_id,))
    existentes = cur.fetchall()

    def pieza_activa(pz):
        pid, fecha, km_ing, vida_km, vida_meses = pz[1], pz[4], pz[5], pz[6], pz[7]
        km_rec = km_actual_mant - (km_ing or 0)
        km_faltan = (vida_km or 0) - km_rec
        m_faltan = None
        if fecha:
            try:
                dias = (datetime.now(ECUADOR_TZ) - datetime.strptime(str(fecha)[:10], "%Y-%m-%d").replace(tzinfo=ECUADOR_TZ)).days
                m_faltan = (vida_meses or 0) - dias / 30.44
            except Exception:
                pass
        if km_faltan <= 0:
            return False
        if vida_km and km_faltan > 0:
            if m_faltan is not None and m_faltan <= 0:
                return False
            return True
        if m_faltan is not None and m_faltan <= 0:
            return False
        return True

    activos = set()
    for pz in existentes:
        if pz[1] and pieza_activa(pz):
            activos.add(pz[1])

    bloqueados = []
    i = 0
    while True:
        pieza_id = request.form.get(f"pv_pieza_{i}")
        if pieza_id is None:
            break
        pieza_id = pieza_id.strip()
        if pieza_id:
            cur.execute("SELECT id, nombre, tipo, vida_km, vida_meses, costo FROM piezas WHERE id=%s", (pieza_id,))
            cat = cur.fetchone()
            if cat:
                except_razon = (request.form.get(f"pv_excepcion_{i}") or "").strip()
                if pieza_id in activos and not except_razon:
                    bloqueados.append(cat[2])
                else:
                    km_ing = int(request.form.get(f"pv_km_{i}") or km_actual_mant or 0)
                    fech_ing = request.form.get(f"pv_fecha_{i}") or request.form.get("fecha") or fecha_local()
                    cur.execute("""
                        INSERT INTO piezas_vehiculo (vehiculo_id, placa, pieza_id, nombre, tipo, fecha_ingreso, km_ingreso, vida_km, vida_meses, costo, mantenimiento_id, observaciones)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, (vehiculo_id, placa, cat[0], cat[1], cat[2], fech_ing, km_ing, cat[3], cat[4], cat[5], None, except_razon or None))
        i += 1

    if bloqueados:
        conn.rollback()
        cur.close()
        conn.close()
        flash("ALERTA: La pieza '" + ", ".join(bloqueados) + "' aun esta en vida util en este vehiculo. Marca la opcion EXCEPCION y escribe el motivo (ej: cambio por dano/golpe) para poder registrarla.", "warning")
        return redirect(url_for("detalle", id=vehiculo_id))

    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/mantenimiento/editar/<int:id>", methods=["POST"])
@admin_required
def editar_mantenimiento(id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT vehiculo_id FROM mantenimientos WHERE id=%s", (id,))
    mant = cur.fetchone()
    if not mant:
        cur.close()
        conn.close()
        return redirect(url_for("index"))
    vehiculo_id = mant["vehiculo_id"]
    costo_mo = float(request.form.get("costo_mo") or 0)
    costo_lub = float(request.form.get("costo_lub") or 0)
    costo_rep = float(request.form.get("costo_rep") or 0)
    costo_total = costo_mo + costo_lub + costo_rep
    cur.execute("""
        UPDATE mantenimientos
        SET fecha=%s, tipo=%s, descripcion=%s, kilometraje=%s, costo=%s, costo_mo=%s, costo_lub=%s, costo_rep=%s, taller=%s
        WHERE id=%s
    """, (
        request.form["fecha"],
        request.form["tipo"],
        request.form["descripcion"],
        request.form["kilometraje"],
        costo_total,
        costo_mo,
        costo_lub,
        costo_rep,
        request.form["taller"],
        id
    ))
    cur.execute("UPDATE vehiculos SET kilometraje=%s WHERE id=%s", (request.form["kilometraje"], vehiculo_id))
    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/mantenimiento/eliminar/<int:id>")
@admin_required
def eliminar_mantenimiento(id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT vehiculo_id FROM mantenimientos WHERE id=%s", (id,))
    mant = cur.fetchone()
    if not mant:
        cur.close()
        conn.close()
        return redirect(url_for("index"))
    vehiculo_id = mant["vehiculo_id"]
    cur.execute("DELETE FROM mantenimientos WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("detalle", id=vehiculo_id))


@app.route("/buscar")
@login_required
def buscar():
    q = request.args.get("q", "")
    mecanica_filter = request.args.get("mecanica", "")
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    query = "SELECT * FROM vehiculos WHERE (placa ILIKE %s OR marca ILIKE %s OR ubicacion ILIKE %s OR chasis ILIKE %s)"
    params = [f"%{q}%", f"%{q}%", f"%{q}%", f"%{q}%"]
    if mecanica_filter:
        query += " AND mecanica=%s"
        params.append(mecanica_filter)
    query += " ORDER BY id DESC"
    cur.execute(query, params)
    vehiculos = cur.fetchall()
    cur.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica")
    mecanicas = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("index.html", vehiculos=vehiculos, busqueda=q, mecanicas=mecanicas, mecanica_actual=mecanica_filter)


@app.route("/api/estadisticas")
@login_required
def estadisticas():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM vehiculos")
    total = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM vehiculos WHERE estado='Activo'")
    activos = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM vehiculos WHERE estado='Inactivo'")
    inactivos = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM vehiculos WHERE estado='En Taller'")
    taller = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM vehiculos WHERE estado='Remate'")
    remate = cur.fetchone()[0]
    cur.close()
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
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM vehiculos WHERE id=%s", (id,))
    vehiculo = cur.fetchone()
    if not vehiculo:
        cur.close()
        conn.close()
        return redirect(url_for("index"))
    cur.execute("SELECT * FROM mantenimientos WHERE vehiculo_id=%s ORDER BY fecha DESC", (id,))
    historial = cur.fetchall()
    cur.close()
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
    return jsonify({"status": "ok", "message": "Sync no disponible en version online"})


def exportar_excel():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
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

    cur.execute("SELECT * FROM vehiculos ORDER BY id")
    vehiculos = cur.fetchall()
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
            val = v.get(key, "") if v else ""
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

        cur.execute("SELECT * FROM vehiculos WHERE mecanica=%s ORDER BY id", (nombre,))
        vehiculos_mec = cur.fetchall()

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
                val = v.get(key, "") if v else ""
                cell = ws.cell(row=row_start + row_idx, column=col_idx, value=val)
                cell.border = thin_border
                if key == "kilometraje":
                    cell.number_format = "#,##0"

        if not vehiculos_mec:
            ws.cell(row=3, column=1, value="Sin vehiculos registrados").font = Font(italic=True, color="64748B")

        mant_row = row_start + len(vehiculos_mec) + 2
        cur.execute("""
            SELECT m.*, v.placa FROM mantenimientos m
            LEFT JOIN vehiculos v ON m.vehiculo_id = v.id
            WHERE v.mecanica=%s
            ORDER BY m.fecha DESC
        """, (nombre,))
        mant_mec = cur.fetchall()

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
                val = m.get(key, "") if m else ""
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
    cur.execute("""
        SELECT m.*, v.placa FROM mantenimientos m
        LEFT JOIN vehiculos v ON m.vehiculo_id = v.id
        ORDER BY m.fecha DESC
    """)
    mant_all = cur.fetchall()
    for row_idx, m in enumerate(mant_all, 2):
        for col_idx, key in enumerate(mant_cols, 1):
            val = m.get(key, "") if m else ""
            cell = ws_mant_all.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            if key == "costo":
                cell.number_format = "#,##0.00"
    for col in range(1, len(mant_headers) + 1):
        ws_mant_all.column_dimensions[get_column_letter(col)].width = 20

    cur.close()
    conn.close()
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def importar_excel(filepath):
    wb = load_workbook(filepath)
    ws = wb["Vehiculos"]
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    imported = 0
    updated = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[1]:
            continue
        placa = str(row[1]).upper().strip()
        cur.execute("SELECT id FROM vehiculos WHERE placa=%s", (placa,))
        existing = cur.fetchone()
        data = {
            "chasis": row[2] or "", "motor": row[3] or "", "marca": row[4] or "",
            "modelo": row[5] or "", "anio": row[6] or 0, "kilometraje": row[7] or 0,
            "ubicacion": row[8] or "", "estado": row[9] or "Activo", "mecanica": row[10] or "Multimarcas"
        }
        if existing:
            cur.execute("""
                UPDATE vehiculos SET chasis=%s, motor=%s, marca=%s, modelo=%s, anio=%s, kilometraje=%s, ubicacion=%s, estado=%s, mecanica=%s
                WHERE placa=%s
            """, (data["chasis"], data["motor"], data["marca"], data["modelo"], data["anio"], data["kilometraje"], data["ubicacion"], data["estado"], data["mecanica"], placa))
            updated += 1
        else:
            cur.execute("""
                INSERT INTO vehiculos (placa, chasis, motor, marca, modelo, anio, kilometraje, ubicacion, estado, mecanica)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (placa, data["chasis"], data["motor"], data["marca"], data["modelo"], data["anio"], data["kilometraje"], data["ubicacion"], data["estado"], data["mecanica"]))
            imported += 1
    if "Mantenimientos" in wb.sheetnames:
        ws2 = wb["Mantenimientos"]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if not row[1]:
                continue
            vehiculo_id = row[1]
            cur.execute("SELECT id FROM vehiculos WHERE id=%s", (vehiculo_id,))
            v = cur.fetchone()
            if v:
                cur.execute("""
                    INSERT INTO mantenimientos (vehiculo_id, fecha, tipo, descripcion, kilometraje, costo, taller)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """, (vehiculo_id, str(row[3] or ""), row[4] or "", row[5] or "", row[6] or 0, row[7] or 0, row[8] or ""))
    conn.commit()
    cur.close()
    conn.close()
    return imported, updated


@app.route("/parte-taller")
@login_required
def parte_taller():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    filtro = request.args.get("filtro", "activos")
    if filtro == "salidos":
        cur.execute("""
            SELECT pt.*, v.marca, v.modelo
            FROM parte_taller pt
            LEFT JOIN vehiculos v ON pt.vehiculo_id = v.id
            WHERE pt.fecha_salida IS NOT NULL
            ORDER BY pt.fecha_salida DESC LIMIT 100
        """)
    else:
        cur.execute("""
            SELECT pt.*, v.marca, v.modelo
            FROM parte_taller pt
            LEFT JOIN vehiculos v ON pt.vehiculo_id = v.id
            WHERE pt.fecha_salida IS NULL
            ORDER BY pt.fecha_ingreso DESC
        """)
    registros = cur.fetchall()
    ahora = datetime.now(ECUADOR_TZ)
    for r in registros:
        if r["fecha_salida"]:
            fi = r["fecha_ingreso"]
            fs = r["fecha_salida"]
            if hasattr(fi, 'replace') and hasattr(fs, 'replace'):
                if fi.tzinfo is None:
                    fi = fi.replace(tzinfo=ECUADOR_TZ)
                if fs.tzinfo is None:
                    fs = fs.replace(tzinfo=ECUADOR_TZ)
                r["dias"] = (fs - fi).days
            else:
                r["dias"] = 0
        else:
            fi = r["fecha_ingreso"]
            if hasattr(fi, 'replace'):
                if fi.tzinfo is None:
                    fi = fi.replace(tzinfo=ECUADOR_TZ)
                r["dias"] = (ahora - fi).days
            else:
                r["dias"] = 0
    cur.execute("SELECT id, placa, marca, modelo FROM vehiculos ORDER BY placa")
    vehiculos = cur.fetchall()
    cur.execute("SELECT * FROM piezas ORDER BY mecanica, nombre")
    catalogo = cur.fetchall()
    catalogo_por_mecanica = {}
    for p in catalogo:
        catalogo_por_mecanica.setdefault(p["mecanica"] or "Multimarcas", []).append(p)

    ahora = datetime.now(ECUADOR_TZ)
    cur.execute("""
        SELECT pv.vehiculo_id, pv.pieza_id, pv.fecha_ingreso, pv.km_ingreso, pv.vida_km, pv.vida_meses, v.kilometraje AS km_actual
        FROM piezas_vehiculo pv
        LEFT JOIN vehiculos v ON pv.vehiculo_id = v.id
        WHERE pv.pieza_id IS NOT NULL
    """)
    piezas_activas_map = {}
    for pz in cur.fetchall():
        vid = pz["vehiculo_id"]
        km_actual = pz["km_actual"] or 0
        km_rec = km_actual - (pz["km_ingreso"] or 0)
        km_faltan = (pz["vida_km"] or 0) - km_rec
        m_faltan = None
        if pz["fecha_ingreso"]:
            try:
                dias = (ahora - datetime.strptime(str(pz["fecha_ingreso"])[:10], "%Y-%m-%d").replace(tzinfo=ECUADOR_TZ)).days
                m_faltan = (pz["vida_meses"] or 0) - dias / 30.44
            except Exception:
                pass
        if km_faltan <= 0:
            continue
        if (pz["vida_km"] and km_faltan > 0):
            if m_faltan is not None and m_faltan <= 0:
                continue
        elif m_faltan is not None and m_faltan <= 0:
            continue
        piezas_activas_map.setdefault(vid, set()).add(pz["pieza_id"])

    hoy_pt = fecha_local_date()
    cur.execute("SELECT * FROM contratos WHERE estado='ACTIVO' AND fecha_fin >= %s AND saldo_total > 0 ORDER BY numero_contrato", (hoy_pt.strftime('%Y-%m-%d'),))
    contratos_activos = cur.fetchall()

    cur.close()
    conn.close()
    return render_template("parte_taller.html", registros=registros, vehiculos=vehiculos, filtro=filtro, catalogo=catalogo, catalogo_por_mecanica=catalogo_por_mecanica, piezas_activas_map=piezas_activas_map, contratos_activos=contratos_activos)


@app.route("/parte-taller/ingresar", methods=["POST"])
@login_required
def parte_taller_ingresar():
    conn = get_db()
    cur = conn.cursor()
    vehiculo_id = request.form["vehiculo_id"]
    contrato_id = request.form.get("contrato_id") or None
    cur.execute("SELECT placa, mecanica FROM vehiculos WHERE id=%s", (vehiculo_id,))
    v = cur.fetchone()
    if v:
        try:
            cur.execute("""
                INSERT INTO parte_taller (vehiculo_id, placa, mecanica, motivo, nombre_ingresa, telefono, novedad, observaciones, fecha_ingreso, registrado_por, contrato_id)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (vehiculo_id, v[0], v[1], request.form["motivo"], request.form.get("nombre_ingresa", ""), request.form.get("telefono", ""), request.form.get("novedad", ""), request.form.get("observaciones", ""), fecha_local(), session.get("usuario", ""), contrato_id))
        except Exception:
            cur.execute("""
                INSERT INTO parte_taller (vehiculo_id, placa, mecanica, motivo, nombre_ingresa, telefono, novedad, observaciones, fecha_ingreso, registrado_por)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """, (vehiculo_id, v[0], v[1], request.form["motivo"], request.form.get("nombre_ingresa", ""), request.form.get("telefono", ""), request.form.get("novedad", ""), request.form.get("observaciones", ""), fecha_local(), session.get("usuario", "")))
        cur.execute("UPDATE vehiculos SET estado='En Taller' WHERE id=%s", (vehiculo_id,))
        conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("parte_taller"))


@app.route("/parte-taller/salir/<int:id>", methods=["POST"])
@login_required
def parte_taller_salir(id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM parte_taller WHERE id=%s AND fecha_salida IS NULL", (id,))
    reg = cur.fetchone()
    if reg:
        fi = reg["fecha_ingreso"]
        if hasattr(fi, 'strftime'):
            fecha_str = fi.strftime("%Y-%m-%d")
        else:
            fecha_str = str(fi)[:10]

        costo_mo = float(request.form.get("costo_mo") or 0)
        costo_lub = float(request.form.get("costo_lub") or 0)
        costo_rep = float(request.form.get("costo_rep") or 0)
        costo_total = costo_mo + costo_lub + costo_rep

        cur.execute("SELECT kilometraje FROM vehiculos WHERE id=%s", (reg["vehiculo_id"],))
        v = cur.fetchone()
        km_actual = (v["kilometraje"] or 0) if v else 0

        tipo_mant = request.form.get("tipo_mantenimiento", "Correctivo")
        descripcion = (reg["motivo"] or "") + (" | " + request.form.get("descripcion_mant", "") if request.form.get("descripcion_mant") else "")

        cur.execute("""
            INSERT INTO mantenimientos (vehiculo_id, fecha, tipo, descripcion, kilometraje, costo, costo_mo, costo_lub, costo_rep, taller)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """, (reg["vehiculo_id"], fecha_str, tipo_mant, descripcion, km_actual, costo_total, costo_mo, costo_lub, costo_rep, reg["mecanica"]))

        km_instal = km_actual
        piezas_activas = set()
        cur.execute("SELECT pieza_id, fecha_ingreso, km_ingreso, vida_km, vida_meses FROM piezas_vehiculo WHERE vehiculo_id=%s AND pieza_id IS NOT NULL", (reg["vehiculo_id"],))
        for pz in cur.fetchall():
            if pz["pieza_id"] is None:
                continue
            km_rec = km_actual - (pz["km_ingreso"] or 0)
            km_faltan = (pz["vida_km"] or 0) - km_rec
            m_faltan = None
            if pz["fecha_ingreso"]:
                try:
                    dias = (datetime.now(ECUADOR_TZ) - datetime.strptime(str(pz["fecha_ingreso"])[:10], "%Y-%m-%d").replace(tzinfo=ECUADOR_TZ)).days
                    m_faltan = (pz["vida_meses"] or 0) - dias / 30.44
                except Exception:
                    pass
            if km_faltan <= 0:
                continue
            if (pz["vida_km"] and km_faltan > 0):
                if m_faltan is not None and m_faltan <= 0:
                    continue
            elif m_faltan is not None and m_faltan <= 0:
                continue
            piezas_activas.add(pz["pieza_id"])

        i = 0
        bloqueados = []
        while True:
            pv_cat = request.form.get(f"pv_sel_{i}")
            if pv_cat is None:
                break
            if pv_cat:
                cur.execute("SELECT id, nombre, tipo, vida_km, vida_meses, costo FROM piezas WHERE id=%s", (pv_cat,))
                cat = cur.fetchone()
                if cat:
                    except_razon = (request.form.get(f"pv_excepcion_{i}") or "").strip()
                    if cat["id"] in piezas_activas and not except_razon:
                        bloqueados.append(cat["nombre"])
                    else:
                        tipo_p = cat["tipo"]
                        vida_km = cat["vida_km"] or 0
                        vida_meses = cat["vida_meses"] or 0
                        costo_p = float(request.form.get(f"pv_costo_{i}") or cat["costo"] or 0)
                        cur.execute("""
                            INSERT INTO piezas_vehiculo (vehiculo_id, placa, pieza_id, nombre, tipo, fecha_ingreso, km_ingreso, vida_km, vida_meses, costo, mantenimiento_id, observaciones)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NULL, %s)
                        """, (reg["vehiculo_id"], reg["placa"], cat["id"], cat["nombre"], tipo_p, fecha_str, km_instal, vida_km, vida_meses, costo_p, except_razon or None))
            i += 1

        if bloqueados:
            conn.rollback()
            cur.close()
            conn.close()
            flash("ALERTA: La pieza '" + ", ".join(bloqueados) + "' aun esta en vida util en este vehiculo. Marca la opcion EXCEPCION y escribe el motivo (ej: cambio por dano/golpe) para poder registrarla.", "warning")
            return redirect(url_for("parte_taller"))

        contrato_id = reg.get("contrato_id")
        if contrato_id:
            try:
                cur.execute("SELECT * FROM contratos WHERE id=%s AND estado='ACTIVO'", (contrato_id,))
                c = cur.fetchone()
                if c:
                    s_rep = float(c["saldo_rep"] or 0)
                    s_lub = float(c["saldo_lub"] or 0)
                    s_mo = float(c["saldo_mo"] or 0)
                    if costo_rep > 0 and s_rep >= 0:
                        ns_rep = s_rep - costo_rep
                        if ns_rep < 0:
                            ns_rep = 0
                        cur.execute("INSERT INTO contrato_movimientos (contrato_id, parte_taller_id, vehiculo_id, placa, partida, descripcion, monto_descontado, saldo_rep_antes, saldo_rep_despues, saldo_lub_antes, saldo_lub_despues, saldo_mo_antes, saldo_mo_despues) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (contrato_id, id, reg["vehiculo_id"], reg["placa"], "repuesto", descripcion, costo_rep, s_rep, ns_rep, s_lub, s_lub, s_mo, s_mo))
                        s_rep = ns_rep
                    if costo_lub > 0:
                        ns_lub = s_lub - costo_lub
                        if ns_lub < 0:
                            ns_lub = 0
                        cur.execute("INSERT INTO contrato_movimientos (contrato_id, parte_taller_id, vehiculo_id, placa, partida, descripcion, monto_descontado, saldo_rep_antes, saldo_rep_despues, saldo_lub_antes, saldo_lub_despues, saldo_mo_antes, saldo_mo_despues) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (contrato_id, id, reg["vehiculo_id"], reg["placa"], "lubricante", descripcion, costo_lub, s_rep, s_rep, s_lub, ns_lub, s_mo, s_mo))
                        s_lub = ns_lub
                    if costo_mo > 0:
                        ns_mo = s_mo - costo_mo
                        if ns_mo < 0:
                            ns_mo = 0
                        cur.execute("INSERT INTO contrato_movimientos (contrato_id, parte_taller_id, vehiculo_id, placa, partida, descripcion, monto_descontado, saldo_rep_antes, saldo_rep_despues, saldo_lub_antes, saldo_lub_despues, saldo_mo_antes, saldo_mo_despues) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (contrato_id, id, reg["vehiculo_id"], reg["placa"], "mano_obra", descripcion, costo_mo, s_rep, s_rep, s_lub, s_lub, s_mo, ns_mo))
                        s_mo = ns_mo
                    saldo_total_n = s_rep + s_lub + s_mo
                    cur.execute("UPDATE contratos SET saldo_rep=%s, saldo_lub=%s, saldo_mo=%s, saldo_total=%s WHERE id=%s", (s_rep, s_lub, s_mo, saldo_total_n, contrato_id))
            except Exception:
                pass

        cur.execute("UPDATE parte_taller SET fecha_salida=%s WHERE id=%s", (fecha_local(), id))
        cur.execute("UPDATE vehiculos SET estado='Activo', kilometraje=%s WHERE id=%s", (km_actual, reg["vehiculo_id"]))
        conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("parte_taller"))


@app.route("/parte-taller/mover-contrato/<int:id>", methods=["POST"])
@login_required
def parte_taller_mover_contrato(id):
    conn = get_db()
    cur = conn.cursor()
    contrato_id = request.form.get("contrato_id") or None
    try:
        cur.execute("SELECT fecha_salida FROM parte_taller WHERE id=%s", (id,))
        r = cur.fetchone()
        if r and not r[0]:
            cur.execute("UPDATE parte_taller SET contrato_id=%s WHERE id=%s", (contrato_id, id))
            conn.commit()
    except Exception:
        conn.rollback()
    cur.close()
    conn.close()
    return redirect(url_for("parte_taller"))


@app.route("/parte-taller/eliminar/<int:id>")
@admin_required
def parte_taller_eliminar(id):
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT vehiculo_id, fecha_salida FROM parte_taller WHERE id=%s", (id,))
    reg = cur.fetchone()
    if reg:
        if not reg["fecha_salida"]:
            cur.execute("UPDATE vehiculos SET estado='Activo' WHERE id=%s", (reg["vehiculo_id"],))
        cur.execute("DELETE FROM parte_taller WHERE id=%s", (id,))
        conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("parte_taller"))


@app.route("/piezas")
@login_required
def piezas():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    tipo = request.args.get("tipo", "repuesto")
    mecanica = request.args.get("mecanica", "")
    cur.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica")
    mecanicas = cur.fetchall()
    if mecanica:
        cur.execute("""
            SELECT * FROM piezas
            WHERE tipo = %s AND mecanica = %s
            ORDER BY nombre ASC
        """, (tipo, mecanica))
    else:
        cur.execute("""
            SELECT * FROM piezas
            WHERE tipo = %s
            ORDER BY nombre ASC
        """, (tipo,))
    piezas = cur.fetchall()
    cur.close()
    conn.close()
    return render_template("piezas.html", piezas=piezas, tipo=tipo, mecanicas=mecanicas, mecanica=mecanica)


@app.route("/piezas/agregar", methods=["POST"])
@login_required
def piezas_agregar():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO piezas (nombre, tipo, vida_km, vida_meses, costo, mecanica)
        VALUES (%s, %s, %s, %s, %s, %s)
    """, (
        request.form["nombre"],
        request.form.get("tipo", "repuesto"),
        request.form.get("vida_km") or 0,
        request.form.get("vida_meses") or 0,
        request.form.get("costo") or 0,
        request.form.get("mecanica", "Multimarcas")
    ))
    conn.commit()
    cur.close()
    conn.close()
    return redirect(url_for("piezas", tipo=request.form.get("tipo", "repuesto"), mecanica=request.form.get("mecanica", "")))


@app.route("/piezas/eliminar/<int:id>")
@admin_required
def piezas_eliminar(id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT tipo FROM piezas WHERE id=%s", (id,))
    r = cur.fetchone()
    cur.execute("DELETE FROM piezas WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()
    tipo = r[0] if r else "repuesto"
    return redirect(url_for("piezas", tipo=tipo, mecanica=request.args.get("mecanica", "")))


@app.route("/piezas-vehiculo/eliminar/<int:id>")
@admin_required
def pieza_vehiculo_eliminar(id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT vehiculo_id FROM piezas_vehiculo WHERE id=%s", (id,))
    r = cur.fetchone()
    cur.execute("DELETE FROM piezas_vehiculo WHERE id=%s", (id,))
    conn.commit()
    cur.close()
    conn.close()
    vid = r[0] if r else 0
    return redirect(url_for("detalle", id=vid))


@app.route("/reportes")
@login_required
def reportes():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    mecanica_filter = request.args.get("mecanica", "")
    tipo_filter = request.args.get("tipo", "")
    fecha_desde = request.args.get("fecha_desde", "")
    fecha_hasta = request.args.get("fecha_hasta", "")

    cur.execute("SELECT DISTINCT mecanica FROM vehiculos WHERE mecanica IS NOT NULL AND mecanica != '' ORDER BY mecanica")
    mecanicas = cur.fetchall()

    query = """
        SELECT m.*, v.placa, v.mecanica, v.marca, v.modelo
        FROM mantenimientos m
        LEFT JOIN vehiculos v ON m.vehiculo_id = v.id
        WHERE 1=1
    """
    params = []

    if mecanica_filter:
        query += " AND v.mecanica = %s"
        params.append(mecanica_filter)
    if tipo_filter:
        query += " AND m.tipo = %s"
        params.append(tipo_filter)
    if fecha_desde:
        query += " AND m.fecha >= %s"
        params.append(fecha_desde)
    if fecha_hasta:
        query += " AND m.fecha <= %s"
        params.append(fecha_hasta)

    query += " ORDER BY m.fecha DESC"
    cur.execute(query, params)
    mantenimientos = cur.fetchall()

    cur.execute("SELECT DISTINCT tipo FROM mantenimientos WHERE tipo IS NOT NULL AND tipo != '' ORDER BY tipo")
    tipos = cur.fetchall()

    total = len(mantenimientos)
    correctivos = sum(1 for m in mantenimientos if m.get("tipo") == "Correctivo")
    preventivos = sum(1 for m in mantenimientos if m.get("tipo") == "Preventivo")
    otro = total - correctivos - preventivos
    costo_total = sum(m.get("costo") or 0 for m in mantenimientos)

    por_tipo = {}
    for m in mantenimientos:
        t = m.get("tipo") or "Sin tipo"
        por_tipo[t] = por_tipo.get(t, 0) + 1

    por_mecanica = {}
    for m in mantenimientos:
        mc = m.get("mecanica") or "Sin mecanica"
        por_mecanica[mc] = por_mecanica.get(mc, 0) + 1

    cur.close()
    conn.close()
    return render_template("reportes.html",
        mantenimientos=mantenimientos,
        mecanicas=mecanicas,
        tipos=tipos,
        total=total,
        correctivos=correctivos,
        preventivos=preventivos,
        otro=otro,
        costo_total=costo_total,
        por_tipo=por_tipo,
        por_mecanica=por_mecanica,
        mecanica_actual=mecanica_filter,
        tipo_actual=tipo_filter,
        fecha_desde=fecha_desde,
        fecha_hasta=fecha_hasta
    )


init_db()
add_columns_parte_taller()
add_columns_contratos()


@app.route("/contratos")
@login_required
def contratos():
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    estado_filtro = request.args.get('estado')
    hoy = fecha_local_date()
    if estado_filtro in ['ACTIVO', 'FINALIZADO']:
        cur.execute("SELECT * FROM contratos WHERE estado = %s ORDER BY id DESC", (estado_filtro,))
    else:
        cur.execute("SELECT * FROM contratos ORDER BY id DESC")
    contratos = cur.fetchall()
    lista = []
    for c in contratos:
        d = dict(c)
        fecha_fin = d.get('fecha_fin')
        try:
            if isinstance(fecha_fin, str):
                from datetime import datetime as dt
                ffin = dt.strptime(fecha_fin, '%Y-%m-%d').date()
            else:
                ffin = fecha_fin
            dias_restantes = (ffin - hoy).days if ffin else 9999
            d['dias_restantes'] = dias_restantes
            d['alerta_vencer'] = (d.get('estado') == 'ACTIVO' and dias_restantes <= 30 and dias_restantes >= 0)
        except Exception:
            d['alerta_vencer'] = False
            d['dias_restantes'] = 9999
        d['alerta_rep'] = float(d.get('saldo_rep') or 0) < 1000 and d.get('estado') == 'ACTIVO'
        d['alerta_lub'] = float(d.get('saldo_lub') or 0) < 1000 and d.get('estado') == 'ACTIVO'
        d['alerta_mo'] = float(d.get('saldo_mo') or 0) < 1000 and d.get('estado') == 'ACTIVO'
        lista.append(d)
    cur.close()
    conn.close()
    historial_finalizados = _cargar_historial_finalizados()
    return render_template("contratos.html", contratos=lista, estado_filtro=estado_filtro, hoy=hoy.strftime('%Y-%m-%d'), historial=historial_finalizados)


def _resumen_movimientos(cur, contrato_id):
    """Agrupa los movimientos de un contrato en trabajos y totales por partida."""
    cur.execute("SELECT * FROM contrato_movimientos WHERE contrato_id=%s ORDER BY fecha, id", (contrato_id,))
    movimientos = dict_rows(cur)
    trabajos = {}
    for m in movimientos:
        partida = (m.get('partida') or '').lower()
        monto = float(m.get('monto_descontado') or 0)
        if monto <= 0:
            continue
        key = m.get('parte_taller_id') or f"mov-{m.get('id')}"
        t = trabajos.setdefault(key, {
            'parte_taller_id': m.get('parte_taller_id'),
            'vehiculo_id': m.get('vehiculo_id'),
            'placa': m.get('placa') or '',
            'fecha': str(m.get('fecha') or ''),
            'descripcion': m.get('descripcion') or '',
            'repuesto': 0.0, 'lubricante': 0.0, 'mano_obra': 0.0, 'total': 0.0,
        })
        if partida in t:
            t[partida] = round(t[partida] + monto, 2)
        t['total'] = round(t['total'] + monto, 2)
    totales_por_partida = {}
    for m in movimientos:
        partida = m.get('partida') or 'otro'
        monto = float(m.get('monto_descontado') or 0)
        if monto <= 0:
            continue
        if partida not in totales_por_partida:
            totales_por_partida[partida] = {'movimientos': 0, 'total': 0.0}
        totales_por_partida[partida]['movimientos'] += 1
        totales_por_partida[partida]['total'] = round(totales_por_partida[partida]['total'] + monto, 2)
    return list(trabajos.values()), totales_por_partida


def _cargar_historial_finalizados():
    import json
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM contrato_historial_finalizado ORDER BY fecha_finalizacion DESC, id DESC")
    rows = dict_rows(cur)
    for h in rows:
        try:
            h['resumen_trabajos'] = json.loads(h.get('resumen_trabajos') or '[]')
        except Exception:
            h['resumen_trabajos'] = []
        try:
            h['totales_por_partida'] = json.loads(h.get('totales_por_partida') or '{}')
        except Exception:
            h['totales_por_partida'] = {}
        # Compatibilidad: contratos finalizados antes de este cambio guardaron resumen vacío.
        if not h['resumen_trabajos'] and h.get('contrato_id'):
            try:
                resumen, totales = _resumen_movimientos(cur, h['contrato_id'])
                h['resumen_trabajos'] = resumen
                if not h['totales_por_partida']:
                    h['totales_por_partida'] = totales
            except Exception:
                pass
    cur.close()
    conn.close()
    return rows


@app.route("/contratos/crear", methods=["POST"])
@login_required
def contratos_crear():
    if session.get('rol') != 'admin':
        return redirect(url_for('contratos'))
    conn = get_db()
    cur = conn.cursor()
    numero_contrato = request.form.get('numero_contrato', '').strip()
    administrador = request.form.get('administrador', '').strip()
    fecha_inicio = request.form.get('fecha_inicio') or fecha_local_date().strftime('%Y-%m-%d')
    dias_vigencia = request.form.get('dias_vigencia') or 0
    try:
        dias_vigencia = int(dias_vigencia)
    except Exception:
        dias_vigencia = 0
    monto_rep = float(request.form.get('monto_rep') or 0)
    monto_lub = float(request.form.get('monto_lub') or 0)
    monto_mo = float(request.form.get('monto_mo') or 0)
    total_contrato = monto_rep + monto_lub + monto_mo
    try:
        from datetime import datetime as dt, timedelta
        fi = dt.strptime(fecha_inicio, '%Y-%m-%d').date()
        ff = fi + timedelta(days=dias_vigencia)
        fecha_fin = ff.strftime('%Y-%m-%d')
    except Exception:
        fecha_fin = fecha_inicio
    try:
        cur.execute("""
            INSERT INTO contratos (numero_contrato, administrador, fecha_inicio, dias_vigencia, fecha_fin,
                                   monto_rep, monto_lub, monto_mo, total_contrato,
                                   saldo_rep, saldo_lub, saldo_mo, saldo_total, estado)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,'ACTIVO')
        """, (numero_contrato, administrador, fecha_inicio, dias_vigencia, fecha_fin,
              monto_rep, monto_lub, monto_mo, total_contrato,
              monto_rep, monto_lub, monto_mo, total_contrato))
        conn.commit()
    except Exception as e:
        conn.rollback()
    cur.close()
    conn.close()
    return redirect(url_for('contratos'))


@app.route("/contratos/finalizar/<int:cid>", methods=["POST"])
@login_required
def contratos_finalizar(cid):
    if session.get('rol') != 'admin':
        return redirect(url_for('contratos'))
    conn = get_db()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    motivo = request.form.get('motivo') or 'Finalizado manualmente'
    try:
        cur.execute("SELECT * FROM contratos WHERE id = %s", (cid,))
        c = cur.fetchone()
        if c and c.get('estado') == 'ACTIVO':
            import json
            # Resumen de trabajos y consumos por partida
            resumen_trabajos, totales_por_partida = _resumen_movimientos(cur, cid)
            cur.execute("INSERT INTO contrato_historial_finalizado (contrato_id, numero_contrato, administrador, fecha_creacion, fecha_inicio, fecha_fin_original, fecha_finalizacion, motivo_finalizacion, finalizado_por, dias_vigencia, monto_rep, monto_lub, monto_mo, total_contrato, saldo_rep, saldo_lub, saldo_mo, saldo_total, resumen_trabajos, totales_por_partida) VALUES (%s,%s,%s,%s,%s,%s,CURRENT_TIMESTAMP,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (c.get('id'), c.get('numero_contrato'), c.get('administrador'), c.get('fecha_creacion'), c.get('fecha_inicio'), c.get('fecha_fin'), motivo, session.get('usuario'), c.get('dias_vigencia'), c.get('monto_rep'), c.get('monto_lub'), c.get('monto_mo'), c.get('total_contrato'), c.get('saldo_rep'), c.get('saldo_lub'), c.get('saldo_mo'), c.get('saldo_total'), json.dumps(resumen_trabajos, ensure_ascii=False), json.dumps(totales_por_partida, ensure_ascii=False)))
            cur.execute("UPDATE contratos SET estado='FINALIZADO', motivo_finalizacion=%s, finalizado_por=%s, fecha_finalizacion=CURRENT_TIMESTAMP WHERE id=%s", (motivo, session.get('usuario'), cid))
            conn.commit()
        else:
            conn.rollback()
    except Exception:
        conn.rollback()
    cur.close()
    conn.close()
    return redirect(url_for('contratos'))


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
