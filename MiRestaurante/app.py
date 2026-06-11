# -*- coding: utf-8 -*-
"""
================================================================================
SISTEMA PROFESIONAL DE RESTAURANTE - VERSIÓN 16.0 (CORREGIDA Y EXTENDIDA)
================================================================================
- Corregido error de base de datos cerrada en cafeteria.
- Añadida ruta ticket_cafeteria faltante.
- Mejorado manejo de errores y logs.
- Código extenso (>1200 líneas) con documentación completa.
- Incluye helpers, validadores, exportadores y más.
================================================================================
"""

import sys
import os
import sqlite3
import json
import logging
import traceback
import hashlib
import secrets
import shutil
from datetime import datetime, timedelta
from functools import wraps
from io import BytesIO
import io as io_lib


# Configuración de logs detallados
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('restaurante.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)
logger.info("Iniciando sistema de restaurante versión 16.0")

# Imports de Flask
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, make_response, abort, jsonify
)

# Imports de librerías externas
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from werkzeug.utils import secure_filename
import xhtml2pdf.pisa as pisa

# ==================== CONFIGURACIÓN DE LA APLICACIÓN ====================
app = Flask(__name__)
app.secret_key = 'clave_super_secreta_restaurante_2025_mas_segura_con_unicode_✨'
app.config['SESSION_COOKIE_SECURE'] = False  # En producción debe ser True
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

# ==================== CONFIGURACIÓN DE SUBIDA DE ARCHIVOS ====================
UPLOAD_FOLDER = 'static/uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Verifica si el archivo tiene una extensión permitida."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def save_uploaded_file(file):
    """
    Guarda un archivo subido en el servidor y retorna el nombre del archivo.
    Si no hay archivo o no es válido, retorna None.
    """
    if file and file.filename and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            # Añadir timestamp para evitar duplicados
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{name}_{timestamp}{ext}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            logger.info(f"Archivo guardado: {filepath}")
            return filename
        except Exception as e:
            logger.error(f"Error al guardar archivo: {e}")
            return None
    return None

def delete_old_uploads(days=30):
    """
    Elimina archivos de uploads que no se han modificado en más de 'days' días.
    Útil para limpiar imágenes no utilizadas.
    """
    try:
        now = datetime.now()
        for filename in os.listdir(UPLOAD_FOLDER):
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(filepath):
                mtime = datetime.fromtimestamp(os.path.getmtime(filepath))
                if (now - mtime).days > days:
                    os.remove(filepath)
                    logger.info(f"Archivo antiguo eliminado: {filepath}")
    except Exception as e:
        logger.error(f"Error al limpiar uploads: {e}")

# ==================== CLASES AUXILIARES EXTENDIDAS ====================
class ValidadorDatos:
    """Clase centralizada para validar datos de entrada."""
    @staticmethod
    def validar_mesa(mesa):
        try:
            mesa_int = int(mesa)
            if 0 <= mesa_int <= 10:
                return mesa_int
            return None
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def validar_cantidad(cantidad):
        try:
            cant = int(cantidad)
            if 1 <= cant <= 99:
                return cant
            return 1
        except (ValueError, TypeError):
            return 1
    
    @staticmethod
    def validar_precio(precio):
        try:
            p = float(precio)
            return round(p, 2) if p >= 0 else 0
        except (ValueError, TypeError):
            return 0
    
    @staticmethod
    def validar_texto(texto, max_len=100):
        if not texto or not isinstance(texto, str):
            return ""
        return texto.strip()[:max_len]
    
    @staticmethod
    def validar_email(email):
        import re
        patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(patron, email)) if email else False

class GestorLogs:
    """Manejo de logs del sistema."""
    @staticmethod
    def log_evento(accion, usuario, detalles):
        logger.info(f"EVENTO: {accion} | Usuario: {usuario} | Detalles: {detalles}")
    
    @staticmethod
    def log_error(accion, error):
        logger.error(f"ERROR: {accion} | {error}")
        traceback.print_exc()
    
    @staticmethod
    def log_debug(mensaje):
        logger.debug(mensaje)

class HelperFechas:
    """Utilidades de fechas y horas."""
    @staticmethod
    def obtener_fecha_hoy():
        return datetime.now().strftime('%Y-%m-%d')
    
    @staticmethod
    def obtener_hora_actual():
        return datetime.now().strftime('%H:%M:%S')
    
    @staticmethod
    def obtener_fecha_hora_completa():
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    @staticmethod
    def formatear_fecha_para_humano(fecha):
        try:
            return datetime.strptime(fecha, '%Y-%m-%d').strftime('%d/%m/%Y')
        except:
            return fecha

class ExportadorReportes:
    """Clase para manejar exportaciones a Excel y PDF."""
    @staticmethod
    def generar_excel(almuerzos, cafeterias, hoy):
        wb = Workbook()
        ws_alm = wb.active
        ws_alm.title = "Almuerzos"
        ws_alm.append(["ID", "Mesa", "Segundo", "Cantidad", "Tipo", "Precio Unitario", "Total", "Para llevar", "Hora"])
        for a in almuerzos:
            total = a['cantidad'] * a['precio_unitario']
            ws_alm.append([a["id"], a["mesa_id"], a["segundo_elegido"], a["cantidad"], a["tipo_pedido"], a["precio_unitario"], total, "Sí" if a["para_llevar"] else "No", a["hora"]])
        ws_cafe = wb.create_sheet("Cafetería")
        ws_cafe.append(["ID", "Mesa", "Items", "Total", "Hora"])
        for c in cafeterias:
            ws_cafe.append([c["id"], c["mesa_id"], c["items"], c["total"], c["hora"]])
        output = io_lib.BytesIO()
        wb.save(output)
        output.seek(0)
        return output
    
    @staticmethod
    def generar_pdf(html_string):
        result = BytesIO()
        pdf = pisa.CreatePDF(BytesIO(html_string.encode('UTF-8')), dest=result)
        if not pdf.err:
            return result.getvalue()
        return None

# ==================== FUNCIONES DE BASE DE DATOS ====================
def get_db():
    """Devuelve conexión a SQLite con row_factory configurado."""
    try:
        conn = sqlite3.connect('restaurante.db')
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn
    except sqlite3.Error as e:
        logger.error(f"Error de conexión a BD: {e}")
        raise

def migrar_base_datos():
    """Verifica y añade columnas/tablas faltantes de forma automática."""
    conn = get_db()
    cursor = conn.cursor()
    logger.info("Iniciando migración de base de datos")
    try:
        # Tabla pedidos_almuerzo
        cursor.execute("PRAGMA table_info(pedidos_almuerzo)")
        columnas = [col[1] for col in cursor.fetchall()]
        if 'cantidad' not in columnas:
            cursor.execute("ALTER TABLE pedidos_almuerzo ADD COLUMN cantidad INTEGER DEFAULT 1")
            logger.info("Añadida columna cantidad a pedidos_almuerzo")
        if 'tipo_pedido' not in columnas:
            cursor.execute("ALTER TABLE pedidos_almuerzo ADD COLUMN tipo_pedido TEXT DEFAULT 'completo'")
            logger.info("Añadida columna tipo_pedido a pedidos_almuerzo")
        if 'precio_unitario' not in columnas:
            cursor.execute("ALTER TABLE pedidos_almuerzo ADD COLUMN precio_unitario REAL DEFAULT 0")
            logger.info("Añadida columna precio_unitario a pedidos_almuerzo")
        
        # Tabla menu_dia
        cursor.execute("PRAGMA table_info(menu_dia)")
        columnas_menu = [col[1] for col in cursor.fetchall()]
        if 'precio_almuerzo_completo' not in columnas_menu:
            cursor.execute("ALTER TABLE menu_dia ADD COLUMN precio_almuerzo_completo REAL DEFAULT 10")
            cursor.execute("ALTER TABLE menu_dia ADD COLUMN precio_segundo_suelto1 REAL DEFAULT 6")
            cursor.execute("ALTER TABLE menu_dia ADD COLUMN precio_segundo_suelto2 REAL DEFAULT 6")
            logger.info("Añadidas columnas de precio a menu_dia")
        
        # Tabla pedidos_cafeteria
        cursor.execute("PRAGMA table_info(pedidos_cafeteria)")
        columnas_cafe = [col[1] for col in cursor.fetchall()]
        if 'usuario_id' not in columnas_cafe:
            cursor.execute("ALTER TABLE pedidos_cafeteria ADD COLUMN usuario_id INTEGER")
            logger.info("Añadida columna usuario_id a pedidos_cafeteria")
        
        # Tabla productos_cafeteria (columna imagen)
        cursor.execute("PRAGMA table_info(productos_cafeteria)")
        columnas_prod = [col[1] for col in cursor.fetchall()]
        if 'imagen' not in columnas_prod:
            cursor.execute("ALTER TABLE productos_cafeteria ADD COLUMN imagen TEXT")
            logger.info("Añadida columna imagen a productos_cafeteria")
        if 'descripcion' not in columnas_prod:
            cursor.execute("ALTER TABLE productos_cafeteria ADD COLUMN descripcion TEXT")
            logger.info("Añadida columna descripcion a productos_cafeteria")
        if 'categoria' not in columnas_prod:
            cursor.execute("ALTER TABLE productos_cafeteria ADD COLUMN categoria TEXT")
            logger.info("Añadida columna categoria a productos_cafeteria")
        
        # Tabla producto_variantes
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS producto_variantes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                producto_id INTEGER NOT NULL,
                nombre_variante TEXT,
                precio REAL NOT NULL,
                imagen TEXT,
                stock INTEGER DEFAULT 999,
                activo BOOLEAN DEFAULT 1,
                FOREIGN KEY (producto_id) REFERENCES productos_cafeteria(id) ON DELETE CASCADE
            )
        ''')
        logger.info("Tabla producto_variantes verificada/creada")
        
        conn.commit()
        logger.info("Migración completada exitosamente")
    except Exception as e:
        logger.error(f"Error en migración: {e}")
        conn.rollback()
    finally:
        conn.close()

# ==================== DECORADORES ====================
def login_required(f):
    """Decorador para rutas que requieren autenticación."""
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            flash('Por favor inicia sesión', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    """Decorador para restringir acceso por rol."""
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('rol') not in roles:
                flash('No tienes permiso para acceder a esta sección', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated
    return decorator




@app.template_filter('from_json')
def from_json_filter(value):
    if value is None:
        return []
    try:
        return json.loads(value) if value else []
    except (json.JSONDecodeError, TypeError):
        return []
# ==================== CONTEXT PROCESSORS ====================
@app.context_processor
def inject_now():
    """Inyecta fecha y hora actual en todas las plantillas."""
    return {
        'now': datetime.now,
        'app_version': '16.0',
        'current_year': datetime.now().year
    }

# ==================== RUTAS DE AUTENTICACIÓN ====================
@app.route('/login', methods=['GET', 'POST'])
def login():
    """Página de inicio de sesión."""
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if not username or not password:
            flash('Usuario y contraseña son requeridos', 'danger')
            return redirect(url_for('login'))
        try:
            conn = get_db()
            user = conn.execute(
                "SELECT * FROM usuarios WHERE username = ? AND password = ?",
                (username, password)
            ).fetchone()
            conn.close()
            if user:
                session.permanent = True
                session['user_id'] = user['id']
                session['username'] = user['username']
                session['rol'] = user['rol']
                session['nombre'] = user['nombre']
                flash(f'Bienvenido {user["nombre"]}', 'success')
                logger.info(f"Login exitoso: {username} ({user['rol']})")
                return redirect(url_for('index'))
            else:
                flash('Usuario o contraseña incorrectos', 'danger')
                logger.warning(f"Intento de login fallido para usuario: {username}")
        except Exception as e:
            logger.error(f"Error en login: {e}")
            flash('Error interno, contacte al administrador', 'danger')
    return render_template('login.html')

@app.route('/logout')
def logout():
    """Cierra la sesión del usuario."""
    username = session.get('username', 'anon')
    session.clear()
    flash('Sesión cerrada correctamente', 'info')
    logger.info(f"Logout: {username}")
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    """Dashboard principal."""
    return render_template('index.html', rol=session.get('rol'))

# ==================== ALMUERZO (SIMPLE) ====================
@app.route('/almuerzo', methods=['GET', 'POST'])
@login_required
def almuerzo():
    """Gestión de pedidos de almuerzo simple (un solo segundo)."""
    hoy = HelperFechas.obtener_fecha_hoy()
    conn = get_db()
    cursor = conn.cursor()
    
    # Obtener o crear menú del día
    cursor.execute("SELECT * FROM menu_dia WHERE fecha = ?", (hoy,))
    menu = cursor.fetchone()
    if not menu:
        cursor.execute('''
            INSERT INTO menu_dia (fecha, entrada, sopa, segundo1_nombre, segundo1_stock, segundo2_nombre, segundo2_stock, salida, precio_almuerzo_completo, precio_segundo_suelto1, precio_segundo_suelto2)
            VALUES (?, 'Ensalada mixta', 'Sopa de verduras', 'Pollo asado', 20, 'Pescado a la plancha', 15, 'Fruta del tiempo', 10, 6, 6)
        ''', (hoy,))
        conn.commit()
        cursor.execute("SELECT * FROM menu_dia WHERE fecha = ?", (hoy,))
        menu = cursor.fetchone()
        logger.info(f"Menú creado para hoy: {hoy}")
    
    if request.method == 'POST':
        try:
            mesa_input = request.form.get('mesa', '').strip()
            segundo = request.form.get('segundo', '').strip()
            cantidad = int(request.form.get('cantidad', 1))
            tipo_pedido = request.form.get('tipo_pedido', 'completo')
            para_llevar = 1 if request.form.get('para_llevar') == '1' else 0
            
            # Validaciones
            if not segundo or segundo not in (menu['segundo1_nombre'], menu['segundo2_nombre']):
                flash('Segundo no válido', 'danger')
                return redirect(url_for('almuerzo'))
            
            if not para_llevar:
                mesa_id = ValidadorDatos.validar_mesa(mesa_input)
                if mesa_id is None or mesa_id <= 0:
                    flash('Seleccione una mesa o marque "Para llevar"', 'danger')
                    return redirect(url_for('almuerzo'))
            else:
                mesa_id = 0
            
            # Precio unitario
            if tipo_pedido == 'completo':
                precio = menu['precio_almuerzo_completo']
            else:
                precio = menu['precio_segundo_suelto1'] if segundo == menu['segundo1_nombre'] else menu['precio_segundo_suelto2']
            
            # Verificar y descontar stock
            if segundo == menu['segundo1_nombre']:
                if menu['segundo1_stock'] < cantidad:
                    flash(f'Stock insuficiente de {menu["segundo1_nombre"]}', 'danger')
                    return redirect(url_for('almuerzo'))
                cursor.execute("UPDATE menu_dia SET segundo1_stock = segundo1_stock - ? WHERE fecha = ?", (cantidad, hoy))
            else:
                if menu['segundo2_stock'] < cantidad:
                    flash(f'Stock insuficiente de {menu["segundo2_nombre"]}', 'danger')
                    return redirect(url_for('almuerzo'))
                cursor.execute("UPDATE menu_dia SET segundo2_stock = segundo2_stock - ? WHERE fecha = ?", (cantidad, hoy))
            
            # Insertar pedido
            cursor.execute('''
                INSERT INTO pedidos_almuerzo (mesa_id, segundo_elegido, cantidad, tipo_pedido, precio_unitario, para_llevar, fecha, hora, estado, usuario_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (mesa_id, segundo, cantidad, tipo_pedido, precio, para_llevar, hoy, HelperFechas.obtener_hora_actual(), 'pendiente', session['user_id']))
            conn.commit()
            pedido_id = cursor.lastrowid
            conn.close()
            flash('Pedido registrado exitosamente', 'success')
            GestorLogs.log_evento("NUEVO_PEDIDO_ALMUERZO", session.get('username'), f"id={pedido_id}")
            return redirect(url_for('ticket_almuerzo', pedido_id=pedido_id))
        except Exception as e:
            conn.close()
            GestorLogs.log_error("REGISTRO_PEDIDO_ALMUERZO", e)
            flash(f'Error al registrar pedido: {str(e)}', 'danger')
            return redirect(url_for('almuerzo'))
    
    mesas = conn.execute("SELECT numero FROM mesas ORDER BY numero").fetchall()
    conn.close()
    return render_template('almuerzo.html', menu=menu, mesas=mesas)

@app.route('/ticket_almuerzo/<int:pedido_id>')
@login_required
def ticket_almuerzo(pedido_id):
    """Muestra el ticket de un pedido de almuerzo simple."""
    conn = get_db()
    pedido = conn.execute('''
        SELECT p.*, m.numero
        FROM pedidos_almuerzo p
        LEFT JOIN mesas m ON p.mesa_id = m.id
        WHERE p.id = ?
    ''', (pedido_id,)).fetchone()
    conn.close()
    if not pedido:
        flash('Pedido no encontrado', 'danger')
        return redirect(url_for('almuerzo'))
    return render_template('ticket.html', pedido=pedido, tipo='almuerzo')

# ==================== ALMUERZO MÚLTIPLE ====================
@app.route('/almuerzo_multiple', methods=['POST'])
@login_required
def almuerzo_multiple():
    """Recibe un pedido con múltiples líneas (JSON) y genera ticket combinado."""
    try:
        data = request.get_json(silent=True)
        if not data or 'items' not in data:
            flash('Datos inválidos', 'danger')
            return redirect(url_for('almuerzo'))
        
        items = data['items']
        mesa_id = data.get('mesa', 0)
        para_llevar = data.get('para_llevar', 0)
        
        if len(items) == 0:
            flash('El pedido está vacío', 'danger')
            return redirect(url_for('almuerzo'))
        
        hoy = HelperFechas.obtener_fecha_hoy()
        hora_actual = HelperFechas.obtener_hora_actual()
        conn = get_db()
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM menu_dia WHERE fecha = ?", (hoy,))
        row = cursor.fetchone()
        if not row:
            cursor.execute('''
                INSERT INTO menu_dia (fecha, entrada, sopa, segundo1_nombre, segundo1_stock, segundo2_nombre, segundo2_stock, salida,
                precio_almuerzo_completo, precio_segundo_suelto1, precio_segundo_suelto2)
                VALUES (?, 'Ensalada mixta', 'Sopa de verduras', 'Pollo asado', 20, 'Pescado a la plancha', 15, 'Fruta del tiempo', 10, 6, 6)
            ''', (hoy,))
            conn.commit()
            cursor.execute("SELECT * FROM menu_dia WHERE fecha = ?", (hoy,))
            row = cursor.fetchone()
        menu = dict(row)
        
        # Validar stock y preparar items para ticket
        stock_seg1 = menu['segundo1_stock']
        stock_seg2 = menu['segundo2_stock']
        items_ticket = []
        total_general = 0.0
        pedido_ids = []
        
        for item in items:
            nombre = item['segundo_nombre']
            tipo = item['tipo']
            cantidad = int(item['cantidad'])
            
            if nombre == menu['segundo1_nombre']:
                if stock_seg1 < cantidad:
                    flash(f'Stock insuficiente de {nombre}', 'danger')
                    conn.close()
                    return redirect(url_for('almuerzo'))
                stock_seg1 -= cantidad
                precio = menu['precio_segundo_suelto1'] if tipo == 'suelto' else menu['precio_almuerzo_completo']
            elif nombre == menu['segundo2_nombre']:
                if stock_seg2 < cantidad:
                    flash(f'Stock insuficiente de {nombre}', 'danger')
                    conn.close()
                    return redirect(url_for('almuerzo'))
                stock_seg2 -= cantidad
                precio = menu['precio_segundo_suelto2'] if tipo == 'suelto' else menu['precio_almuerzo_completo']
            else:
                flash(f'Segundo no reconocido: {nombre}', 'danger')
                conn.close()
                return redirect(url_for('almuerzo'))
            
            total_general += precio * cantidad
            items_ticket.append({
                'cantidad': cantidad,
                'nombre': nombre,
                'tipo': 'Completo' if tipo == 'completo' else 'Suelto',
                'precio_unitario': precio
            })
        
        # Actualizar stock e insertar pedidos
        cursor.execute("SELECT segundo1_stock, segundo2_stock FROM menu_dia WHERE fecha = ?", (hoy,))
        stocks = cursor.fetchone()
        stock_seg1_db = stocks['segundo1_stock']
        stock_seg2_db = stocks['segundo2_stock']
        
        for item in items_ticket:
            nombre = item['nombre']
            cantidad = item['cantidad']
            tipo = 'completo' if item['tipo'] == 'Completo' else 'suelto'
            precio = item['precio_unitario']
            
            if nombre == menu['segundo1_nombre']:
                nuevo_stock = stock_seg1_db - cantidad
                cursor.execute("UPDATE menu_dia SET segundo1_stock = ? WHERE fecha = ?", (nuevo_stock, hoy))
                stock_seg1_db = nuevo_stock
            else:
                nuevo_stock = stock_seg2_db - cantidad
                cursor.execute("UPDATE menu_dia SET segundo2_stock = ? WHERE fecha = ?", (nuevo_stock, hoy))
                stock_seg2_db = nuevo_stock
            
            cursor.execute('''
                INSERT INTO pedidos_almuerzo (mesa_id, segundo_elegido, cantidad, tipo_pedido, precio_unitario, para_llevar, fecha, hora, estado, usuario_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (mesa_id, nombre, cantidad, tipo, precio, para_llevar, hoy, hora_actual, 'pendiente', session['user_id']))
            pedido_ids.append(cursor.lastrowid)
        
        conn.commit()
        conn.close()
        
        flash(f'Pedido múltiple registrado ({len(pedido_ids)} líneas)', 'success')
        return render_template('ticket.html',
                               pedido_multiple=True,
                               items=items_ticket,
                               total=total_general,
                               mesa=mesa_id,
                               para_llevar=para_llevar,
                               fecha=hoy,
                               hora=hora_actual,
                               pedido_ids=pedido_ids,
                               tipo='')
    except Exception as e:
        logger.error(f"Error en almuerzo_multiple: {e}")
        traceback.print_exc()
        flash('Error al procesar el pedido múltiple', 'danger')
        return redirect(url_for('almuerzo'))

# ==================== CAFETERÍA ====================
@app.route('/cafeteria', methods=['GET', 'POST'])
@login_required
def cafeteria():
    """Gestión de pedidos de cafetería (con variantes y stock)."""
    if request.method == 'POST':
        conn = None
        try:
            mesa_id = request.form.get('mesa', '0')
            items_json = request.form.get('items', '[]')
            total = float(request.form.get('total', 0))
            items = json.loads(items_json)
            
            if not items:
                flash('No hay productos en el pedido', 'danger')
                return redirect(url_for('cafeteria'))

            # =============== VALIDAR Y ESTANDARIZAR ESTRUCTURA DE ITEMS ===============
            items_validados = []
            for item in items:
                items_validados.append({
                    'nombre': str(item.get('nombre', item.get('producto_nombre', 'Sin nombre'))),
                    'variante_nombre': str(item.get('variante_nombre', '')),
                    'variante_id': int(item.get('variante_id', 0)) if item.get('variante_id') else 0,
                    'precio': float(item.get('precio', 0)),
                    'cantidad': int(item.get('cantidad', 1))
                })

            logger.info(f"Items validados para cafetería: {items_validados}")

            conn = get_db()
            cursor = conn.cursor()

            # Descontar stock de cada ítem (variante o producto base)
            for item in items_validados:
                cantidad = item.get('cantidad', 1)
                variante_id = item.get('variante_id', 0)
                nombre = item.get('nombre', '')
                
                if variante_id and variante_id != 0:
                    cursor.execute('''
                        UPDATE producto_variantes 
                        SET stock = stock - ? 
                        WHERE id = ? AND stock >= ?
                    ''', (cantidad, variante_id, cantidad))
                    if cursor.rowcount == 0:
                        raise Exception(f'Stock insuficiente para la variante ID {variante_id}')
                else:
                    cursor.execute('''
                        UPDATE productos_cafeteria 
                        SET stock = stock - ? 
                        WHERE nombre = ? AND stock >= ?
                    ''', (cantidad, nombre, cantidad))
                    if cursor.rowcount == 0:
                        raise Exception(f'Stock insuficiente para el producto {nombre}')

            # Insertar pedido en pedidos_cafeteria con items validados
            cursor.execute('''
                INSERT INTO pedidos_cafeteria (mesa_id, items, total, fecha, hora, estado, usuario_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (mesa_id, json.dumps(items_validados), total, HelperFechas.obtener_fecha_hoy(), 
                  HelperFechas.obtener_hora_actual(), 'pendiente', session['user_id']))
            conn.commit()
            pedido_id = cursor.lastrowid
            
            flash('Pedido de cafetería registrado', 'success')
            logger.info(f"Pedido de cafetería creado: ID={pedido_id}, Mesa={mesa_id}, Items={len(items_validados)}")
            return redirect(url_for('ticket_cafeteria', pedido_id=pedido_id))

        except Exception as e:
            if conn:
                conn.rollback()
            logger.error(f"Error POST /cafeteria: {str(e)}")
            traceback.print_exc()
            flash(f'Error al registrar pedido: {str(e)}', 'danger')
            return redirect(url_for('cafeteria'))
        finally:
            if conn:
                conn.close()

    # GET: mostrar productos y mesas
    conn = get_db()
    productos = conn.execute('''
        SELECT id, nombre, imagen, stock, precio
        FROM productos_cafeteria 
        WHERE activo = 1 
        ORDER BY nombre
    ''').fetchall()
    mesas = conn.execute('SELECT numero FROM mesas ORDER BY numero').fetchall()
    conn.close()
    return render_template('cafeteria.html', productos=productos, mesas=mesas)

# ==================== TICKET CAFETERÍA (CORREGIDO) ====================
@app.route('/ticket_cafeteria/<int:pedido_id>')
@login_required
def ticket_cafeteria(pedido_id):
    """Muestra el ticket de un pedido de cafetería."""
    conn = None
    try:
        conn = get_db()
        pedido = conn.execute('''
            SELECT p.*, m.numero
            FROM pedidos_cafeteria p
            LEFT JOIN mesas m ON p.mesa_id = m.id
            WHERE p.id = ?
        ''', (pedido_id,)).fetchone()
        
        if not pedido:
            flash('Pedido no encontrado', 'danger')
            return redirect(url_for('cafeteria'))
        
        # Convertir a diccionario mutable para poder asignar valores
        pedido_dict = dict(pedido)
        
        # Asegurar que items sea un string JSON válido, incluso si es None
        if pedido_dict['items'] is None:
            pedido_dict['items'] = '[]'
        
        # Validar que el JSON sea válido
        try:
            json.loads(pedido_dict['items'])
        except json.JSONDecodeError:
            pedido_dict['items'] = '[]'
        
        logger.info(f"Ticket cafetería #={pedido_id} generado exitosamente")
        return render_template('ticket.html', pedido=pedido_dict, tipo='cafeteria')
    
    except Exception as e:
        logger.error(f"Error al generar ticket de cafetería: {e}")
        traceback.print_exc()
        flash(f'Error al generar el ticket: {str(e)}', 'danger')
        return redirect(url_for('cafeteria'))
    finally:
        if conn:
            conn.close()

# ==================== COMANDAS ====================
@app.route('/comanda/<tipo>/<int:pedido_id>')
@login_required
def comanda(tipo, pedido_id):
    """Genera la comanda para cocina (almuerzo) o barra (cafetería)."""
    conn = get_db()
    if tipo == 'almuerzo':
        pedido = conn.execute('''
            SELECT p.*, m.numero
            FROM pedidos_almuerzo p
            LEFT JOIN mesas m ON p.mesa_id = m.id
            WHERE p.id = ?
        ''', (pedido_id,)).fetchone()
        conn.close()
        return render_template('comanda_cocina.html', pedido=pedido)
    elif tipo == 'cafeteria':
        pedido = conn.execute('''
            SELECT p.*, m.numero
            FROM pedidos_cafeteria p
            LEFT JOIN mesas m ON p.mesa_id = m.id
            WHERE p.id = ?
        ''', (pedido_id,)).fetchone()
        conn.close()
        return render_template('comanda_barra.html', pedido=pedido)
    abort(404)

@app.route('/comanda_multiple', methods=['POST'])
@login_required
def comanda_multiple():
    """Genera comanda para pedido múltiple (recibe JSON)."""
    data = request.get_json()
    if not data:
        abort(400)
    return render_template('comanda_cocina.html',
                           pedido_multiple=True,
                           items=data.get('items', []),
                           total=data.get('total', 0),
                           mesa=data.get('mesa', 0),
                           para_llevar=data.get('para_llevar', 0),
                           fecha=data.get('fecha', ''),
                           hora=data.get('hora', ''),
                           pedido_ids=data.get('pedido_ids', []))

# ==================== ADMIN: CRUD PRODUCTOS CAFETERÍA ====================
@app.route('/admin/productos', methods=['GET', 'POST'])
@login_required
@role_required('admin')
def admin_productos():
    """CRUD completo de productos de cafetería y sus variantes."""
    conn = get_db()
    cursor = conn.cursor()
    
    # Obtener productos con sus variantes
    cursor.execute('''
        SELECT p.*, 
               (SELECT json_group_array(json_object('id', v.id, 'nombre_variante', v.nombre_variante, 'precio', v.precio, 'imagen', v.imagen, 'stock', v.stock))
                FROM producto_variantes v WHERE v.producto_id = p.id AND v.activo = 1) as variantes_json
        FROM productos_cafeteria p
        WHERE p.activo = 1
        ORDER BY p.nombre
    ''')
    productos = cursor.fetchall()
    
    # Procesar JSON de variantes para cada producto
    productos_con_variantes = []
    for p in productos:
        prod = dict(p)
        if prod['variantes_json']:
            prod['variantes'] = json.loads(prod['variantes_json']) if prod['variantes_json'] else []
        else:
            prod['variantes'] = []
        productos_con_variantes.append(prod)
    
    if request.method == 'POST':
        action = request.form.get('action')
        try:
            # ---------- AGREGAR PRODUCTO ----------
            if action == 'add_producto':
                nombre = ValidadorDatos.validar_texto(request.form.get('nombre', ''))
                stock = int(request.form.get('stock', 0))
                imagen = None
                if 'imagen' in request.files:
                    imagen = save_uploaded_file(request.files['imagen'])
                cursor.execute('''
                    INSERT INTO productos_cafeteria (nombre, stock, imagen) VALUES (?, ?, ?)
                ''', (nombre, stock, imagen))
                conn.commit()
                flash('Producto agregado correctamente', 'success')
            
            # ---------- EDITAR PRODUCTO ----------
            elif action == 'edit_producto':
                prod_id = request.form.get('id')
                nombre = ValidadorDatos.validar_texto(request.form.get('nombre', ''))
                stock = int(request.form.get('stock', 0))
                imagen = None
                if 'imagen' in request.files and request.files['imagen'].filename:
                    imagen = save_uploaded_file(request.files['imagen'])
                    if imagen:
                        cursor.execute('UPDATE productos_cafeteria SET imagen = ? WHERE id = ?', (imagen, prod_id))
                cursor.execute('UPDATE productos_cafeteria SET nombre = ?, stock = ? WHERE id = ?', (nombre, stock, prod_id))
                conn.commit()
                flash('Producto actualizado', 'success')
            
            # ---------- ELIMINAR PRODUCTO ----------
            elif action == 'delete_producto':
                prod_id = request.form.get('id')
                cursor.execute('UPDATE productos_cafeteria SET activo = 0 WHERE id = ?', (prod_id,))
                conn.commit()
                flash('Producto desactivado', 'warning')
            
            # ---------- AGREGAR VARIANTE ----------
            elif action == 'add_variante':
                producto_id = request.form.get('producto_id')
                nombre_variante = ValidadorDatos.validar_texto(request.form.get('nombre_variante', ''))
                precio = ValidadorDatos.validar_precio(request.form.get('precio', 0))
                stock = int(request.form.get('stock', 0))
                imagen = None
                if 'imagen' in request.files and request.files['imagen'].filename:
                    imagen = save_uploaded_file(request.files['imagen'])
                cursor.execute('''
                    INSERT INTO producto_variantes (producto_id, nombre_variante, precio, stock, imagen)
                    VALUES (?, ?, ?, ?, ?)
                ''', (producto_id, nombre_variante, precio, stock, imagen))
                conn.commit()
                flash('Variante agregada', 'success')
            
            # ---------- EDITAR VARIANTE ----------
            elif action == 'edit_variante':
                variante_id = request.form.get('variante_id')
                nombre_variante = ValidadorDatos.validar_texto(request.form.get('nombre_variante', ''))
                precio = ValidadorDatos.validar_precio(request.form.get('precio', 0))
                stock = int(request.form.get('stock', 0))
                imagen = None
                if 'imagen' in request.files and request.files['imagen'].filename:
                    imagen = save_uploaded_file(request.files['imagen'])
                    if imagen:
                        cursor.execute('UPDATE producto_variantes SET imagen = ? WHERE id = ?', (imagen, variante_id))
                cursor.execute('''
                    UPDATE producto_variantes SET nombre_variante = ?, precio = ?, stock = ? WHERE id = ?
                ''', (nombre_variante, precio, stock, variante_id))
                conn.commit()
                flash('Variante actualizada', 'success')
            
            # ---------- ELIMINAR VARIANTE ----------
            elif action == 'delete_variante':
                variante_id = request.form.get('variante_id')
                cursor.execute('UPDATE producto_variantes SET activo = 0 WHERE id = ?', (variante_id,))
                conn.commit()
                flash('Variante eliminada', 'warning')
        
        except Exception as e:
            conn.rollback()
            GestorLogs.log_error("ADMIN_PRODUCTOS", e)
            flash(f'Error: {e}', 'danger')
        return redirect(url_for('admin_productos'))
    
    conn.close()
    return render_template('admin_productos.html', productos=productos_con_variantes)

# ==================== RUTA PARA OBTENER VARIANTES (AJAX) ====================
@app.route('/get_variantes/<int:producto_id>')
@login_required
def get_variantes(producto_id):
    """Devuelve las variantes de un producto en formato JSON (para AJAX)."""
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, nombre_variante, precio, imagen, stock
        FROM producto_variantes
        WHERE producto_id = ? AND activo = 1
    ''', (producto_id,))
    variantes = cursor.fetchall()
    conn.close()
    return jsonify({'variantes': [dict(v) for v in variantes]})

# ==================== ADMIN: CONFIGURAR MENÚ DEL DÍA ====================
@app.route('/admin/menu', methods=['GET', 'POST'])
@login_required
@role_required('admin')
def admin_menu():
    """Configuración del menú del día (platos, stock y precios)."""
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn = get_db()
    cursor = conn.cursor()
    if request.method == 'POST':
        try:
            entrada = request.form.get('entrada', '').strip()
            sopa = request.form.get('sopa', '').strip()
            segundo1_nombre = request.form.get('segundo1_nombre', '').strip()
            segundo1_stock = int(request.form.get('segundo1_stock', 0))
            segundo2_nombre = request.form.get('segundo2_nombre', '').strip()
            segundo2_stock = int(request.form.get('segundo2_stock', 0))
            salida = request.form.get('salida', '').strip()
            precio_almuerzo = ValidadorDatos.validar_precio(request.form.get('precio_almuerzo', 10))
            precio_suelto1 = ValidadorDatos.validar_precio(request.form.get('precio_suelto1', 6))
            precio_suelto2 = ValidadorDatos.validar_precio(request.form.get('precio_suelto2', 6))
            
            cursor.execute('''
                INSERT OR REPLACE INTO menu_dia (fecha, entrada, sopa, segundo1_nombre, segundo1_stock, segundo2_nombre, segundo2_stock, salida, precio_almuerzo_completo, precio_segundo_suelto1, precio_segundo_suelto2)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (hoy, entrada, sopa, segundo1_nombre, segundo1_stock, segundo2_nombre, segundo2_stock, salida, precio_almuerzo, precio_suelto1, precio_suelto2))
            conn.commit()
            flash('Menú del día actualizado correctamente', 'success')
            logger.info(f"Menú actualizado para {hoy}")
        except Exception as e:
            conn.rollback()
            GestorLogs.log_error("ADMIN_MENU", e)
            flash(f'Error: {e}', 'danger')
        return redirect(url_for('almuerzo'))
    menu = cursor.execute('SELECT * FROM menu_dia WHERE fecha = ?', (hoy,)).fetchone()
    conn.close()
    return render_template('admin_menu.html', menu=menu)

# ==================== CAJA DIARIA ====================
@app.route('/caja', methods=['GET', 'POST'])
@login_required
@role_required('admin')
def caja_diaria():
    """Control de caja diaria: apertura, retiros, ingresos extra y cierre."""
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn = get_db()
    cursor = conn.cursor()
    
    # Asegurar registro de caja para hoy
    cursor.execute('SELECT * FROM caja_diaria WHERE fecha = ?', (hoy,))
    caja = cursor.fetchone()
    if not caja:
        cursor.execute('INSERT INTO caja_diaria (fecha, apertura, estado, retiros, ingresos_extra) VALUES (?, 0, "abierta", "[]", "[]")', (hoy,))
        conn.commit()
        cursor.execute('SELECT * FROM caja_diaria WHERE fecha = ?', (hoy,))
        caja = cursor.fetchone()
    
    if request.method == 'POST':
        action = request.form.get('action')
        try:
            if action == 'retiro':
                monto = ValidadorDatos.validar_precio(request.form.get('monto', 0))
                concepto = request.form.get('concepto', '').strip()
                retiros = json.loads(caja['retiros'])
                retiros.append({'hora': datetime.now().strftime('%H:%M'), 'monto': monto, 'concepto': concepto})
                cursor.execute('UPDATE caja_diaria SET retiros = ? WHERE fecha = ?', (json.dumps(retiros), hoy))
                conn.commit()
                flash(f'Retiro de {monto} Bs registrado', 'success')
            elif action == 'ingreso_extra':
                monto = ValidadorDatos.validar_precio(request.form.get('monto', 0))
                concepto = request.form.get('concepto', '').strip()
                extras = json.loads(caja['ingresos_extra'])
                extras.append({'hora': datetime.now().strftime('%H:%M'), 'monto': monto, 'concepto': concepto})
                cursor.execute('UPDATE caja_diaria SET ingresos_extra = ? WHERE fecha = ?', (json.dumps(extras), hoy))
                conn.commit()
                flash(f'Ingreso extra de {monto} Bs registrado', 'success')
            elif action == 'cerrar_caja':
                monto_real = ValidadorDatos.validar_precio(request.form.get('monto_cierre', 0))
                cursor.execute('UPDATE caja_diaria SET cierre = ?, estado = "cerrada" WHERE fecha = ?', (monto_real, hoy))
                conn.commit()
                flash('Caja cerrada exitosamente', 'success')
        except Exception as e:
            conn.rollback()
            GestorLogs.log_error("CAJA_DIARIA", e)
            flash(f'Error: {e}', 'danger')
        return redirect(url_for('caja_diaria'))
    
    # Datos para el resumen
    total_almuerzos = cursor.execute('SELECT COUNT(*) as total FROM pedidos_almuerzo WHERE fecha = ?', (hoy,)).fetchone()['total'] or 0
    total_almuerzo = cursor.execute('SELECT SUM(precio_unitario * cantidad) as total_ingreso FROM pedidos_almuerzo WHERE fecha = ?', (hoy,)).fetchone()['total_ingreso'] or 0
    total_cafe = cursor.execute('SELECT SUM(total) as suma FROM pedidos_cafeteria WHERE fecha = ?', (hoy,)).fetchone()['suma'] or 0
    ingresos_automaticos = total_almuerzo + total_cafe
    
    retiros = json.loads(caja['retiros'])
    extras = json.loads(caja['ingresos_extra'])
    total_retiros = sum(r['monto'] for r in retiros)
    total_extras = sum(e['monto'] for e in extras)
    esperado = caja['apertura'] + ingresos_automaticos + total_extras - total_retiros
    diferencia = caja['cierre'] - esperado if caja['cierre'] else 0
    
    conn.close()
    return render_template('caja.html', caja=caja, total_almuerzos=total_almuerzos, total_almuerzo=total_almuerzo,
                          total_cafe=total_cafe, ingresos_automaticos=ingresos_automaticos,
                          retiros=retiros, extras=extras, esperado=esperado, diferencia=diferencia)

# ==================== CONFIGURACIÓN (SETTINGS) ====================
@app.route('/admin/settings', methods=['GET', 'POST'])
@login_required
@role_required('admin')
def admin_settings():
    """Configuración general del sistema."""
    conn = get_db()
    cursor = conn.cursor()
    if request.method == 'POST':
        precio = ValidadorDatos.validar_precio(request.form.get('precio_almuerzo', 10))
        cursor.execute('UPDATE settings SET value = ? WHERE key = "precio_almuerzo"', (str(precio),))
        conn.commit()
        flash('Configuración guardada correctamente', 'success')
        return redirect(url_for('admin_settings'))
    precio = cursor.execute('SELECT value FROM settings WHERE key = "precio_almuerzo"').fetchone()
    conn.close()
    return render_template('admin_settings.html', precio_almuerzo=float(precio['value']) if precio else 10)

# ==================== REPORTES ====================
@app.route('/reportes')
@login_required
@role_required('admin')
def reportes():
    """Panel de reportes del día."""
    conn = get_db()
    cursor = conn.cursor()
    hoy = datetime.now().strftime('%Y-%m-%d')
    almuerzos = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN para_llevar=1 THEN 1 ELSE 0 END) as para_llevar FROM pedidos_almuerzo WHERE fecha = ?', (hoy,)).fetchone()
    cafeterias = cursor.execute('SELECT COUNT(*) as total, SUM(total) as ingreso FROM pedidos_cafeteria WHERE fecha = ?', (hoy,)).fetchone()
    # Top productos más vendidos en cafetería
    pedidos = cursor.execute('SELECT items FROM pedidos_cafeteria WHERE fecha = ?', (hoy,)).fetchall()
    contador = {}
    for p in pedidos:
        try:
            items = json.loads(p['items'])
            for item in items:
                nombre = item.get('nombre', 'Desconocido')
                contador[nombre] = contador.get(nombre, 0) + 1
        except:
            pass
    top = sorted(contador.items(), key=lambda x: x[1], reverse=True)[:5]
    conn.close()
    return render_template('reportes.html', almuerzos=almuerzos, cafeterias=cafeterias, top=top)

# ==================== EXPORTACIÓN A EXCEL ====================
@app.route('/exportar_excel')
@login_required
@role_required('admin')
def exportar_excel():
    """Exporta reportes del día a archivo Excel."""
    conn = get_db()
    hoy = datetime.now().strftime('%Y-%m-%d')
    almuerzos = conn.execute("SELECT * FROM pedidos_almuerzo WHERE fecha = ?", (hoy,)).fetchall()
    cafeterias = conn.execute("SELECT * FROM pedidos_cafeteria WHERE fecha = ?", (hoy,)).fetchall()
    conn.close()
    
    output = ExportadorReportes.generar_excel(almuerzos, cafeterias, hoy)
    logger.info(f"Exportado Excel para {hoy}")
    return send_file(output, download_name=f"reporte_{hoy}.xlsx", as_attachment=True)

# ==================== EXPORTACIÓN A PDF ====================
@app.route('/exportar_pdf')
@login_required
@role_required('admin')
def exportar_pdf():
    """Exporta reportes del día a PDF."""
    hoy = datetime.now().strftime('%Y-%m-%d')
    conn = get_db()
    almuerzos = conn.execute("SELECT * FROM pedidos_almuerzo WHERE fecha = ?", (hoy,)).fetchall()
    cafeterias = conn.execute("SELECT * FROM pedidos_cafeteria WHERE fecha = ?", (hoy,)).fetchall()
    conn.close()
    rendered_html = render_template("reporte_pdf.html", hoy=hoy, almuerzos=almuerzos, cafeterias=cafeterias)
    pdf_data = ExportadorReportes.generar_pdf(rendered_html)
    if pdf_data:
        response = make_response(pdf_data)
        response.headers['Content-Type'] = 'application/pdf'
        response.headers['Content-Disposition'] = f'attachment; filename=reporte_{hoy}.pdf'
        logger.info(f"Exportado PDF para {hoy}")
        return response
    else:
        flash('Error al generar el PDF', 'danger')
        return redirect(url_for('reportes'))

# ==================== MANEJADORES DE ERRORES ====================
@app.errorhandler(404)
def not_found(e):
    logger.warning(f"404: {request.url}")
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(e):
    logger.error(f"500: {e}")
    return render_template('500.html'), 500



# ==================== INICIALIZACIÓN ====================
if __name__ == '__main__':
    import database as db
    db.init_db()
    migrar_base_datos()
    # Limpiar archivos de uploads antiguos (opcional)
    delete_old_uploads(days=30)
    logger.info("Aplicación iniciada en modo debug - Puerto 5000")
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)
