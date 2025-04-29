import logging
import sqlite3
import pandas as pd
import time
from datetime import datetime
from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, ConversationHandler,
    ContextTypes, filters
)

# --- Configuración ---
BOT_TOKEN = "7693778500:AAGAG0N4c0MCFuivbRl48w3OF_eaWksyhtA"
EXCEL_FILE = "RegistroBolsas.xlsx"
DB_FILE = "entregas.db"
ALLOWED_USER_IDS = [677369649, 87654321]

# Estados de conversación
(ACTION, QUERY, CONFIRM_DELIVERY,
 BATCH, ADD_CED, ADD_NAME, ADD_NOMINA,
 DELETE_CED, CONFIRM_DELETE) = range(9)

# --- Base de datos ---
def init_db():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS entregas(
            ID INTEGER PRIMARY KEY,
            Cedula TEXT,
            Nombre TEXT,
            Unidad TEXT,
            Direcciones TEXT,
            Sede TEXT,
            Nomina TEXT,
            Chequeo INTEGER DEFAULT 0,
            Numero_Asignado INTEGER,
            Autorizado TEXT,
            Fecha_Entrega TEXT,
            Entregado_Por TEXT
        )
    ''')
    conn.commit()
    return conn

# Cargar datos de Excel si tabla vacía
def load_excel(conn):
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        logging.warning("No se encontró el archivo Excel: %s", EXCEL_FILE)
        return
    if conn.execute("SELECT COUNT(*) FROM entregas").fetchone()[0] == 0:
        df.to_sql('entregas', conn, if_exists='append', index=False)

# Guardar SQLite de vuelta a Excel con nombre dinámico
def save_excel(conn, filename=None):
    df = pd.read_sql_query(
        "SELECT ID, Cedula, Nombre, Unidad, Direcciones, Sede, Nomina, Chequeo, Numero_Asignado, Autorizado, Fecha_Entrega, Entregado_Por FROM entregas", conn)
    if not filename:
        df.to_excel(EXCEL_FILE, index=False)
    else:
        df.to_excel(filename, index=False)

# Obtener siguiente número por nómina
def next_numero(conn, nomina):
    r = conn.execute(
        "SELECT MAX(Numero_Asignado) FROM entregas WHERE Nomina=?", (nomina,)
    ).fetchone()[0]
    return (r or 0) + 1

# --- Handlers ---
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if uid not in ALLOWED_USER_IDS:
        await update.message.reply_text("No autorizado.")
        return ConversationHandler.END
    kb = [
        ['Entregar Bolsa', 'Carga Lote'],
        ['Consultar', 'Resumen'],
        ['Resumen por Direcciones'],
        ['Agregar Persona', 'Eliminar Persona'],
        ['Salir']
    ]
    await update.message.reply_text(
        "Seleccione una opción:",
        reply_markup=ReplyKeyboardMarkup(kb, one_time_keyboard=True)
    )
    return ACTION

async def action_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text
    context.user_data['action'] = text
    conn = context.application.bot_data['conn']
    if text == 'Entregar Bolsa':
        await update.message.reply_text('Ingrese cédula o nombre:')
        return QUERY
    if text == 'Carga Lote':
        await update.message.reply_text('Ingrese cédulas separadas por coma:')
        return BATCH
    if text == 'Consultar':
        await update.message.reply_text('Ingrese cédula para consultar:')
        return QUERY
    if text == 'Resumen':
        for nom in ['Activo', 'Jubilado']:
            ent = conn.execute("SELECT COUNT(*) FROM entregas WHERE Nomina=? AND Chequeo=1", (nom,)).fetchone()[0]
            tot = conn.execute("SELECT COUNT(*) FROM entregas WHERE Nomina=?", (nom,)).fetchone()[0]
            await update.message.reply_text(f"{nom}: Entregadas {ent}, Pendientes {tot - ent}")
        return await start(update, context)
    if text == 'Resumen por Direcciones':
        rows = conn.execute(
            "SELECT Direcciones, COUNT(*) FROM entregas WHERE Nomina='Activo' AND Chequeo=0 GROUP BY Direcciones"
        ).fetchall()
        msg = '\n'.join(f"{addr}: {cnt}" for addr, cnt in rows) or 'No hay activos pendientes'
        await update.message.reply_text(msg)
        return await start(update, context)
    if text == 'Agregar Persona':
        await update.message.reply_text('Nueva persona - Ingresa cédula:')
        return ADD_CED
    if text == 'Eliminar Persona':
        await update.message.reply_text('Eliminar - Ingresa cédula:')
        return DELETE_CED
    if text == 'Salir':
        # Al salir, generar Excel con fecha de jornada
        fecha = datetime.now().strftime('%Y%m%d')
        filename = f"entrega_bolsa_{fecha}.xlsx"
        save_excel(context.application.bot_data['conn'], filename)
        await update.message.reply_text(f"Jornada finalizada. Archivo guardado como {filename}", reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    return ACTION

async def query_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
#    q = update.message.text.strip()
#    conn = context.application.bot_data['conn']
#    row = conn.execute(
#        "SELECT ID, Cedula, Nombre, Unidad, Direcciones, Sede, Nomina, Chequeo, Numero_Asignado, Autorizado, Fecha_Entrega, Entregado_Por FROM entregas WHERE Cedula=? OR Nombre LIKE ?",
#        (q, f"%{q}%")
#    ).fetchone()
#    if not row:
#        await update.message.reply_text('Persona no encontrada.')
#        return await start(update, context)
#    (ID, Cedula, Nombre, Unidad, Direcciones, Sede, Nomina,
#     Chequeo, Numero, Autorizado, Fecha_Entrega, Entregado_Por) = row
#    action = context.user_data['action']
#    if action == 'Consultar':
#        await update.message.reply_text(f"{Nombre} - Entregado: {'Sí' if Chequeo else 'No'}")
#        return await start(update, context)
#    if action == 'Entregar Bolsa':
#        if Chequeo:
#            await update.message.reply_text('Ya entregó su bolsa.')
#            return await start(update, context)
#        num = next_numero(conn, Nomina)
#        fecha = datetime.now().strftime('%Y-%m-%d %H:%M')
#        quien = update.effective_user.full_name
#        conn.execute(
#            "UPDATE entregas SET Chequeo=1, Numero_Asignado=?, Fecha_Entrega=?, Entregado_Por=? WHERE ID=?",
#            (num, fecha, quien, ID)
#        )
#        conn.commit()
#        save_excel(conn)
#        await update.message.reply_text(f"Asignado número {num} a {Nombre}")
#        return await start(update, context)
#    return await start(update, context)

##    return CONFIRM_DELIVERY

    q = update.message.text.strip()
    conn = context.application.bot_data['conn']
    c = conn.cursor()
    # Seleccionar columnas explícitas para evitar valores extras
    c.execute(
        "SELECT ID, Cedula, Nombre, Unidad, Direcciones, Sede, Nomina, Chequeo, Numero_Asignado, Autorizado, Cedula_Autorizado, Fecha_Entrega, Entregado_Por "
        "FROM entregas WHERE Cedula=? OR Nombre LIKE ?",
        (q, f"%{q}%")
    )
    row = c.fetchone()
    if not row:
        await update.message.reply_text('No encontrado.')
        return await start(update, context)
    ID, ced, nombre, unidad, direc, sede, nomina, chequeo, num, auth, auth_ced, fecha, por = row
    if chequeo:
        await update.message.reply_text('Ya entregado.')
        return await start(update, context)
    context.user_data['ID'] = ID
    context.user_data['nomina'] = nomina
    info = (
        f"ID: {ID}\nCédula: {ced}\nNombre: {nombre}"
        f"\nUnidad: {unidad}\nSede: {sede}\nNómina: {nomina}"
    )
    await update.message.reply_text(
        info + "\nConfirmar entrega? si/no",
        reply_markup=ReplyKeyboardMarkup([['si','no']], one_time_keyboard=True)
    )
    return CONFIRM_DELIVERY

async def confirm_delivery(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.lower() != 'si':
        await update.message.reply_text('Cancelado', reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
    conn = context.application.bot_data['conn']
    ID = context.user_data['ID']
    nomina = context.user_data['nomina']
    num = next_numero(conn, nomina)
    fecha = datetime.now().strftime('%Y-%m-%d')
    quien = update.effective_user.full_name
    conn.execute(
        "UPDATE entregas SET Chequeo=1, Numero_Asignado=?, Fecha_Entrega=?, Entregado_Por=? WHERE ID=?",
        (num, fecha, quien, ID)
    )
    conn.commit()
    df = pd.read_sql_query("SELECT * FROM entregas", conn)
    df.to_excel(EXCEL_FILE, index=False)
    await update.message.reply_text(f"Entregado # {num} (Nómina {nomina})", reply_markup=ReplyKeyboardRemove())
    return await start(update, context)

async def batch_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    ceds = [c.strip() for c in update.message.text.split(',')]
    conn = context.application.bot_data['conn']
    results = []
    for q in ceds:
        row = conn.execute("SELECT ID, Nombre, Nomina, Chequeo FROM entregas WHERE Cedula=?", (q,)).fetchone()
        if not row:
            results.append(f"{q}: no encontrado")
        elif row[3]:
            results.append(f"{row[1]}: ya entregado")
        else:
            num = next_numero(conn, row[2])
            fecha = datetime.now().strftime('%Y-%m-%d %H:%M')
            quien = update.effective_user.full_name
            conn.execute(
                "UPDATE entregas SET Chequeo=1, Numero_Asignado=?, Fecha_Entrega=?, Entregado_Por=? WHERE ID=?",
                (num, fecha, quien, row[0])
            )
            results.append(f"{row[1]}: #{num}")
    conn.commit()
    save_excel(conn)
    await update.message.reply_text("\n".join(results))
    return await start(update, context)

async def add_ced(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['new_ced'] = update.message.text.strip()
    await update.message.reply_text('Ingresa Nombre:')
    return ADD_NAME

async def add_name(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['new_name'] = update.message.text.strip()
    await update.message.reply_text('Ingresa Dirección y Nómina (ej. Oficina de Gestion Humana Activo):')
    return ADD_NOMINA

async def add_nom(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    parts = update.message.text.strip().rsplit(' ', 1)
    if len(parts) != 2:
        await update.message.reply_text('Formato inválido. Usa: <Dirección> <Nomina>')
        return ADD_NOMINA
    direccion, nom = parts
    conn = context.application.bot_data['conn']
    new_id = (conn.execute("SELECT MAX(ID) FROM entregas").fetchone()[0] or 0) + 1
    conn.execute(
        "INSERT INTO entregas(ID, Cedula, Nombre, Unidad, Direcciones, Sede, Nomina) VALUES(?,?,?,?,?,?,?)",
        (new_id, context.user_data['new_ced'], context.user_data['new_name'], '', direccion, '', nom)
    )
    conn.commit()
    save_excel(conn)
    await update.message.reply_text('Persona agregada correctamente.')
    return await start(update, context)

async def delete_ced(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    ced = update.message.text.strip()
    context.user_data['del_ced'] = ced
    conn = context.application.bot_data['conn']
    row = conn.execute("SELECT Nombre FROM entregas WHERE Cedula=?", (ced,)).fetchone()
    if not row:
        await update.message.reply_text('No existe esa cédula.')
        return await start(update, context)
    await update.message.reply_text(f"Confirma eliminar a {row[0]}? (si/no)")
    return CONFIRM_DELETE

async def confirm_delete(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if update.message.text.lower() != 'si':
        return await start(update, context)
    conn = context.application.bot_data['conn']
    conn.execute("DELETE FROM entregas WHERE Cedula=?", (context.user_data['del_ced'],))
    conn.commit()
    save_excel(conn)
    await update.message.reply_text('Persona eliminada.')
    return await start(update, context)

# --- Main ---

def main():
    logging.basicConfig(level=logging.INFO)
    conn = init_db()
    load_excel(conn)
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.bot_data['conn'] = conn
    conv = ConversationHandler(
        entry_points=[CommandHandler('start', start)],
        states={
            ACTION: [MessageHandler(filters.TEXT & ~filters.COMMAND, action_handler)],
            QUERY: [MessageHandler(filters.TEXT & ~filters.COMMAND, query_handler)],
            CONFIRM_DELIVERY: [MessageHandler(filters.Regex('^(si|no)$'), confirm_delivery)],
            BATCH: [MessageHandler(filters.TEXT & ~filters.COMMAND, batch_handler)],
            ADD_CED: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_ced)],
            ADD_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_name)],
            ADD_NOMINA: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_nom)],
            DELETE_CED: [MessageHandler(filters.TEXT & ~filters.COMMAND, delete_ced)],
            CONFIRM_DELETE: [MessageHandler(filters.Regex('^(si|no)$'), confirm_delete)]
        },
        fallbacks=[CommandHandler('start', start)]
    )
    app.add_handler(conv)
    while True:
        try:
            app.run_polling()
            break
        except Exception as e:
            logging.error(f"Error: {e}, reintentando en 15s...")
            time.sleep(15)

if __name__ == '__main__':
    main()
