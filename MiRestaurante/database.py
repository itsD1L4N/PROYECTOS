
# -*- coding: utf-8 -*-
"""
================================================================================
BASE DE DATOS PROFESIONAL PARA RESTAURANTE - VERSIÓN 3.0 (SOLO ESTRUCTURA)
================================================================================
Solo crea tablas, columnas e índices. No inserta ningún dato de ejemplo.
================================================================================
"""

import sqlite3
from datetime import datetime

DB_NAME = 'restaurante.db'

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def verificar_y_agregar_columna(conn, tabla, columna, tipo_defecto):
    """Agrega una columna a una tabla si no existe."""
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({tabla})")
    columnas = [col[1] for col in cursor.fetchall()]
    if columna not in columnas:
        cursor.execute(f"ALTER TABLE {tabla} ADD COLUMN {columna} {tipo_defecto}")
        conn.commit()
        print(f"✅ Columna '{columna}' añadida a {tabla}")

def init_db():
    conn = get_db()
    cursor = conn.cursor()

    # ==================== TABLA: usuarios ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            rol TEXT NOT NULL,
            nombre TEXT
        )
    ''')
    verificar_y_agregar_columna(conn, 'usuarios', 'email', 'TEXT')
    verificar_y_agregar_columna(conn, 'usuarios', 'telefono', 'TEXT')
    verificar_y_agregar_columna(conn, 'usuarios', 'activo', 'BOOLEAN DEFAULT 1')
    verificar_y_agregar_columna(conn, 'usuarios', 'fecha_creacion', 'TIMESTAMP DEFAULT CURRENT_TIMESTAMP')

    # ==================== TABLA: mesas ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS mesas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero INTEGER UNIQUE NOT NULL,
            estado TEXT DEFAULT 'libre'
        )
    ''')
    verificar_y_agregar_columna(conn, 'mesas', 'capacidad', 'INTEGER DEFAULT 4')
    verificar_y_agregar_columna(conn, 'mesas', 'ubicacion', 'TEXT')

    # ==================== TABLA: menu_dia (almuerzos) ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS menu_dia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT UNIQUE NOT NULL,
            entrada TEXT,
            sopa TEXT,
            segundo1_nombre TEXT,
            segundo1_stock INTEGER DEFAULT 0,
            segundo2_nombre TEXT,
            segundo2_stock INTEGER DEFAULT 0,
            salida TEXT
        )
    ''')
    verificar_y_agregar_columna(conn, 'menu_dia', 'precio_almuerzo_completo', 'REAL DEFAULT 10')
    verificar_y_agregar_columna(conn, 'menu_dia', 'precio_segundo_suelto1', 'REAL DEFAULT 6')
    verificar_y_agregar_columna(conn, 'menu_dia', 'precio_segundo_suelto2', 'REAL DEFAULT 6')
    verificar_y_agregar_columna(conn, 'menu_dia', 'segundo_suelto_disponible', 'BOOLEAN DEFAULT 1')
    verificar_y_agregar_columna(conn, 'menu_dia', 'activo', 'BOOLEAN DEFAULT 1')

    # ==================== TABLA: productos_cafeteria ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS productos_cafeteria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            precio REAL DEFAULT 0,
            stock INTEGER DEFAULT 999,
            activo BOOLEAN DEFAULT 1
        )
    ''')
    verificar_y_agregar_columna(conn, 'productos_cafeteria', 'imagen', 'TEXT')
    verificar_y_agregar_columna(conn, 'productos_cafeteria', 'descripcion', 'TEXT')
    verificar_y_agregar_columna(conn, 'productos_cafeteria', 'categoria', 'TEXT')
    verificar_y_agregar_columna(conn, 'productos_cafeteria', 'orden', 'INTEGER DEFAULT 0')

    # ==================== TABLA: producto_variantes ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS producto_variantes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            producto_id INTEGER NOT NULL,
            nombre_variante TEXT,
            precio REAL NOT NULL,
            stock INTEGER DEFAULT 999,
            activo BOOLEAN DEFAULT 1,
            FOREIGN KEY (producto_id) REFERENCES productos_cafeteria(id) ON DELETE CASCADE
        )
    ''')
    verificar_y_agregar_columna(conn, 'producto_variantes', 'imagen', 'TEXT')

    # ==================== TABLA: pedidos_almuerzo ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pedidos_almuerzo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mesa_id INTEGER,
            segundo_elegido TEXT,
            cantidad INTEGER DEFAULT 1,
            tipo_pedido TEXT DEFAULT 'completo',
            precio_unitario REAL DEFAULT 0,
            para_llevar BOOLEAN DEFAULT 0,
            fecha TEXT,
            hora TEXT,
            estado TEXT DEFAULT 'pendiente',
            usuario_id INTEGER
        )
    ''')
    verificar_y_agregar_columna(conn, 'pedidos_almuerzo', 'notas', 'TEXT')

    # ==================== TABLA: pedidos_cafeteria ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pedidos_cafeteria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mesa_id INTEGER,
            items TEXT,
            total REAL,
            fecha TEXT,
            hora TEXT,
            estado TEXT DEFAULT 'pendiente',
            usuario_id INTEGER
        )
    ''')
    verificar_y_agregar_columna(conn, 'pedidos_cafeteria', 'notas', 'TEXT')

    # ==================== TABLA: caja_diaria ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS caja_diaria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT UNIQUE NOT NULL,
            apertura REAL DEFAULT 0,
            cierre REAL DEFAULT NULL,
            estado TEXT DEFAULT 'abierta',
            retiros TEXT,
            ingresos_extra TEXT
        )
    ''')
    verificar_y_agregar_columna(conn, 'caja_diaria', 'notas', 'TEXT')

    # ==================== TABLA: settings ====================
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')

    # ==================== ÍNDICES (optimización) ====================
    indices = [
        "CREATE INDEX IF NOT EXISTS idx_menu_dia_fecha ON menu_dia(fecha)",
        "CREATE INDEX IF NOT EXISTS idx_pedidos_almuerzo_fecha ON pedidos_almuerzo(fecha)",
        "CREATE INDEX IF NOT EXISTS idx_pedidos_cafeteria_fecha ON pedidos_cafeteria(fecha)",
        "CREATE INDEX IF NOT EXISTS idx_productos_cafeteria_nombre ON productos_cafeteria(nombre)",
        "CREATE INDEX IF NOT EXISTS idx_producto_variantes_producto_id ON producto_variantes(producto_id)",
        "CREATE INDEX IF NOT EXISTS idx_usuarios_username ON usuarios(username)",
        "CREATE INDEX IF NOT EXISTS idx_mesas_numero ON mesas(numero)"
    ]
    for idx in indices:
        cursor.execute(idx)

    conn.commit()
    conn.close()
    print("✅ Base de datos creada/actualizada (estructura vacía, sin datos de ejemplo)")

if __name__ == '__main__':
    init_db()