import re
import os
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler

TOKEN = "8770484638:AAFYQjU3SdO9O4c-Klgs4mwdY9xUxuSyGiA"

INVENTARIO_PATH = Path(__file__).parent / 'cod_prod_Tipicos.xlsx'
SALIDA_PATH = Path('G:/Mi unidad/Tipicos del Valle/TipdelValle_Reportes/2026/Reportes_Bot')
SALIDA_PATH.mkdir(parents=True, exist_ok=True)

df_inventario = pd.read_excel(INVENTARIO_PATH)
df_inventario['nombre_lower'] = df_inventario['nombre'].str.lower().str.strip()

SINONIMOS = {
    'pechugas': 'Pechuga Asada',
    'pechuga': 'Pechuga Asada',
    'caldos': 'Caldo con Arepa',
    'caldo': 'Caldo con Arepa',
    'rancho': 'Huevos Rancheros',
    'rancheros': 'Huevos Rancheros',
    'ranchero': 'Huevos Rancheros',
    'tinto': 'Tinto',
    'empanadas': 'Empanada',
    'empanada': 'Empanada',
    'queso': 'Queso Costeño x libra',
    'aguas': 'Botella de Agua Cristal',
    'agua': 'Botella de Agua Cristal',
    'jugos': 'Jugo Hit',
    'hit': 'Jugo Hit',
    'jugo': 'Jugo Hit',
    'geitore': 'Gatorade',
    'gatorade': 'Gatorade',
    'desmechada': 'Carne Desmechada',
    'desmech': 'Carne Desmechada',
    'te': 'Bebida Te',
    'tehatsu': 'Te Hatsu',
    'cocacola': 'Gaseosa CocaCola',
    'coca': 'Gaseosa CocaCola',
    'chicle': 'Chiclets Traident 4',
    'chicles': 'Chiclets Traident 4',
    'carne': 'Carne Asada',
    'gaseosa': 'Gaseosa CocaCola',
    'ho2': 'Agua H2o',
    'h2o': 'Agua H2o',
    'agua ho2': 'Agua H2o',
}

PENDIENTE_CONFIRMACION = {}

CONFIRMAR_SINDERO, CORREGIR_ITEM = range(2)

def normalizar_palabra(palabra):
    palabra = palabra.lower().strip()
    palabra = palabra.replace('ón', 'on').replace('á', 'a').replace('é', 'e')
    palabra = palabra.replace('í', 'i').replace('ú', 'u').replace('ñ', 'n')
    palabra = palabra.replace('.', ' ')
    palabra = re.sub(r'\s+', ' ', palabra)
    if palabra.endswith('nes'):
        palabra = palabra[:-3]
    elif palabra.endswith('es'):
        palabra = palabra[:-2]
    elif palabra.endswith('s'):
        palabra = palabra[:-1]
    return palabra

def buscar_producto(nombre_producto):
    nombre_input = nombre_producto.lower().strip()
    
    if nombre_input in SINONIMOS:
        nombre_buscado = SINONIMOS[nombre_input]
        for _, row in df_inventario.iterrows():
            if row['nombre_lower'] == nombre_buscado.lower():
                return row, True
    
    nombre_normalizado = normalizar_palabra(nombre_producto)
    mejores = []
    for _, row in df_inventario.iterrows():
        nombre_inv = normalizar_palabra(row['nombre'])
        score = SequenceMatcher(None, nombre_normalizado, nombre_inv).ratio()
        if score >= 0.5:
            mejores.append((row, score))
    
    if mejores:
        mejores.sort(key=lambda x: x[1], reverse=True)
        mejor_match, mejor_score = mejores[0]
        if len(mejores) > 1 and mejores[1][1] >= mejor_score - 0.1:
            return None, 'ambiguo'
        return mejor_match, True
    
    for _, row in df_inventario.iterrows():
        nombre_inv = row['nombre_lower']
        if nombre_normalizado in nombre_inv or nombre_inv in nombre_normalizado:
            return row, True
    
    if nombre_input in SINONIMOS:
        nombre_buscado = SINONIMOS[nombre_input]
        for _, row in df_inventario.iterrows():
            if row['nombre_lower'] == nombre_buscado.lower():
                return row, True
    
    return None, 'no_encontrado'

def parsear_mensaje(texto):
    lineas = texto.strip().split('\n')
    productos = []
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    indice = 0
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
        
        match = re.match(r'^\[\d+/\d+/\d+,\s+\d+:\d+:\d+\]\s', linea)
        if match:
            linea = linea[match.end():].strip()
            if ':' in linea:
                linea = linea.split(':', 1)[1].strip()
        
        match = re.match(r'^(\d+)\s+(.+?)\s+\$?([\d,]+)$', linea, re.IGNORECASE)
        if match:
            indice += 1
            cantidad = int(match.group(1))
            nombre = match.group(2).strip()
            valor_total = int(match.group(3).replace(',', ''))
            
            producto, encontrado = buscar_producto(nombre)
            if encontrado == True and producto is not None and not producto.empty:
                codigo = int(producto['codigo'])
                precio_unitario = int(producto['precio'])
                esperado = cantidad * precio_unitario
                
                precio = precio_unitario
                validado = True
                if abs(valor_total - esperado) >= 1000:
                    precio = valor_total // cantidad if cantidad > 0 else 0
                    validado = False
            else:
                codigo = 0
                precio_unitario = 0
                precio = 0
                esperado = 0
                validado = False
            
            productos.append({
                'indice': indice,
                'cantidad': cantidad,
                'nombre': nombre,
                'nombre_inventario': producto['nombre'] if producto is not None else nombre,
                'codigo': codigo,
                'precio': precio,
                'precio_unitario': precio_unitario,
                'total': valor_total,
                'esperado': esperado,
                'validado': validado,
                'encontrado': encontrado,
                'fecha': fecha_hoy
            })
    
    return productos

def obtener_ultimo_numero():
    ultimo = 4628
    try:
        archivos = list(SALIDA_PATH.glob('*.xlsx'))
        for archivo in archivos:
            try:
                df = pd.read_excel(archivo)
                if 'NUMERO FACTURA' in df.columns:
                    max_num = df['NUMERO FACTURA'].max()
                    if max_num > ultimo:
                        ultimo = int(max_num)
            except:
                pass
    except:
        pass
    return ultimo

def obtener_contador_fecha(fecha):
    contador = 1
    try:
        fecha_str = datetime.strptime(fecha, '%Y-%m-%d').strftime('%m%d')
        archivos = list(SALIDA_PATH.glob(f'2026{fecha_str}-??-LO-101.xlsx'))
        if archivos:
            contador = len(archivos) + 1
    except:
        pass
    return contador

def generar_excel(productos, local='LO-101', airport='SKVP'):
    if not productos:
        return None, None
    
    fecha = productos[0]['fecha']
    fecha_str = datetime.strptime(fecha, '%Y-%m-%d').strftime('%m%d')
    
    numero_base = obtener_ultimo_numero()
    contador_fecha = obtener_contador_fecha(fecha)
    datos = []
    
    for i, p in enumerate(productos):
        datos.append({
            'AEROPUERTO': airport,
            'LOCAL': local,
            'NUMERO FACTURA': numero_base + i + 1,
            'FECHA DE EMISIÓN': fecha,
            'CÓDIGO PRODUCTO': p['codigo'],
            'DESCRIPCION': p.get('nombre_inventario', p['nombre']),
            'CANTIDAD': p['cantidad'],
            'VALOR NETO': p['precio'],
            'IMPUESTOS': None,
            'TOTAL': p['total']
        })
    
    df = pd.DataFrame(datos)
    
    nombre = f'2026{fecha_str}-{contador_fecha:02d}-{local}.xlsx'
    ruta = SALIDA_PATH / nombre
    df.to_excel(ruta, index=False)
    
    return ruta, df

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🤖 *Reportes Bot*\n\n"
        "Envía el reporte en formato:\n"
        "```\n"
        "cantidad producto valor\n"
        "```\n"
        "Ejemplo:\n"
        "```\n"
        "3 Empanada 15000\n"
        "5 Tinto 15000\n"
        "2 Chicharon 60000\n"
        "```",
        parse_mode='Markdown'
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📋 *Cómo usar:*\n\n"
        "1. Copia el reporte de WhatsApp\n"
        "2. Pega aquí en este chat\n"
        "3. Yo genero el Excel automáticamente\n\n"
        "Cuando haya productos ambiguos, te pediré confirmación.\n\n"
        "*Para corregir productos:*\n"
        "Usa: `/corregir 3:Queso` (corrige item 3 a Queso)\n\n"
        "Usa /cancelar para cancelar una operación pendiente.",
        parse_mode='Markdown'
    )

async def confirmar_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in PENDIENTE_CONFIRMACION:
        await update.message.reply_text("No hay nada pendiente de confirmación.")
        return
    
    data = PENDIENTE_CONFIRMACION[user_id]
    productos = data['productos']
    
    for p in data['no_encontrados']:
        p['codigo'] = 0
        p['encontrado'] = True
    
    for p in data['ambiguedades']:
        p['encontrado'] = True
    
    ruta, df = generar_excel(productos)
    
    if ruta:
        total_general = sum(p['total'] for p in productos)
        resumen = f"✅ *Reporte confirmado*\n\n"
        resumen += f"📦 *{len(productos)} productos*\n"
        resumen += f"💰 Total: ${total_general:,}\n"
        
        if data['alertas']:
            resumen += f"\n⚠️ *Precios ajustados:*\n"
            for alerta in data['alertas']:
                resumen += f"{alerta}\n"
        
        await update.message.reply_text(resumen, parse_mode='Markdown')
        
        await update.message.reply_document(document=open(ruta, 'rb'))
    else:
        await update.message.reply_text("❌ Error generando Excel")
    
    del PENDIENTE_CONFIRMACION[user_id]

async def cancelar_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id in PENDIENTE_CONFIRMACION:
        del PENDIENTE_CONFIRMACION[user_id]
        await update.message.reply_text("❌ Operación cancelada.")
    else:
        await update.message.reply_text("No hay nada pendiente.")

async def corregir_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    
    if user_id not in PENDIENTE_CONFIRMACION:
        await update.message.reply_text("No hay productos pendientes de confirmar. Envía un reporte primero.")
        return
    
    args = context.args
    if not args or len(args) < 2:
        await update.message.reply_text(
            "Usa: `/corregir 3:Queso`\n"
            "Ejemplo: `/corregir 5:Agua` (cambia el item 5 a Agua)",
            parse_mode='Markdown'
        )
        return
    
    try:
        parte = args[0].split(':')
        indice = int(parte[0])
        nuevo_producto = parte[1] if len(parte) > 1 else args[1]
        
        for p in args[1:]:
            nuevo_producto += ' ' + p
    except (ValueError, IndexError):
        await update.message.reply_text(
            "❌ Formato inválido. Usa: `/corregir 3:Queso`",
            parse_mode='Markdown'
        )
        return
    
    data = PENDIENTE_CONFIRMACION[user_id]
    productos = data['productos']
    
    item_encontrado = None
    for p in productos:
        if p['indice'] == indice:
            item_encontrado = p
            break
    
    if not item_encontrado:
        await update.message.reply_text(f"❌ No hay item con índice {indice}")
        return
    
    producto_en_inventario, encontrado = buscar_producto(nuevo_producto)
    
    if encontrado == True and producto_en_inventario is not None:
        item_encontrado['nombre'] = nuevo_producto
        item_encontrado['nombre_inventario'] = producto_en_inventario['nombre']
        item_encontrado['codigo'] = int(producto_en_inventario['codigo'])
        item_encontrado['precio_unitario'] = int(producto_en_inventario['precio'])
        item_encontrado['precio'] = int(producto_en_inventario['precio'])
        item_encontrado['total'] = item_encontrado['cantidad'] * item_encontrado['precio']
        item_encontrado['encontrado'] = True
        item_encontrado['validado'] = True
        
        if str(indice) in data.get('no_encontrados_indices', []):
            data['no_encontrados_indices'].remove(str(indice))
        
        await update.message.reply_text(
            f"✅ *Item {indice} corregido:*\n"
            f"- Ahora: {item_encontrado['cantidad']} x {item_encontrado['nombre_inventario']} = ${item_encontrado['total']:,}",
            parse_mode='Markdown'
        )
    else:
        await update.message.reply_text(f"❌ No encontré '{nuevo_producto}' en el inventario")
        return
    
    PENDIENTE_CONFIRMACION[user_id] = data
    
    no_encontrados = [p for p in productos if p.get('codigo', 0) == 0 or p.get('encontrado') == 'no_encontrado']
    ambiguedades = [p for p in productos if p.get('encontrado') == 'ambiguo']
    
    if not no_encontrados and not ambiguedades:
        await update.message.reply_text("✅ Todos los productos están correctos. Usa /confirmar para generar el Excel.")
    else:
        await update.message.reply_text("⬆️Algunos items aún necesitan corrección. Usa /corregir o /confirmar.")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    mensaje = update.message.text
    
    productos = parsear_mensaje(mensaje)
    
    if not productos:
        await update.message.reply_text(
            "❌ No entendí el formato.\n\n"
            "Usa: `cantidad producto valor`\n"
            "Ej: `3 Empanada 15000`",
            parse_mode='Markdown'
        )
        return
    
    ambiguedades = []
    no_encontrados = []
    no_encontrados_indices = []
    alertas = []
    
    for p in productos:
        if p['codigo'] == 0:
            no_encontrados.append(p)
        elif p['encontrado'] == 'no_encontrado':
            p['codigo'] = 0
            no_encontrados.append(p)
        elif p['encontrado'] == 'ambiguo':
            ambiguedades.append(p)
        elif not p['validado']:
            precio_real = p['precio_unitario']
            precio_digitado = p['precio']
            alertas.append(f"⚠️ Item {p['indice']}: {p['nombre']} - precio unitario ${precio_real:,} vs digitado ${precio_digitado:,}")
    
    for p in no_encontrados:
        no_encontrados_indices.append(str(p['indice']))
    
    PENDIENTE_CONFIRMACION[user_id] = {
        'productos': productos,
        'ambiguedades': ambiguedades,
        'no_encontrados': no_encontrados,
        'no_encontrados_indices': no_encontrados_indices,
        'alertas': alertas
    }
    
    if no_encontrados or ambiguedades:
        texto = "❓ *Confirma los siguientes items:*\n\n"
        for i, p in enumerate(no_encontrados):
            texto += f"*{p['indice']}.* {p['cantidad']} x {p['nombre']} = ${p['total']:,}\n"
            texto += f"   ⬆️ ¿Qué producto es?\n"
        for p in ambiguedades:
            texto += f"*{p['indice']}.* {p['cantidad']} x {p['nombre']} = ${p['total']:,}\n"
            texto += f"   ⬆️ ¿Es esto correcto?\n"
        
        botones = []
        if no_encontrados or ambiguedades:
            botones.append("✅ Confirmar todo")
            botones.append("❌ Cancelar")
        
        teclado = [[btn] for btn in botones]
        
        await update.message.reply_text(texto, parse_mode='Markdown')
        return
    
    ruta, df = generar_excel(productos)
    
    if ruta:
        total_general = sum(p['total'] for p in productos)
        resumen = f"✅ *Reporte procesado*\n\n"
        resumen += f"📦 *{len(productos)} productos*\n"
        resumen += f"💰 Total: ${total_general:,}\n"
        
        await update.message.reply_text(resumen, parse_mode='Markdown')
        
        await update.message.reply_document(document=open(ruta, 'rb'))
    else:
        await update.message.reply_text("❌ Error generando Excel")

async def error_handle(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f"Update {update} caused error {context.error}")

def main():
    application = Application.builder().token(TOKEN).build()
    
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("confirmar", confirmar_command))
    application.add_handler(CommandHandler("cancelar", cancelar_command))
    application.add_handler(CommandHandler("corregir", corregir_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error_handle)
    
    print("🤖 Bot Started. Enviame un mensaje!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()