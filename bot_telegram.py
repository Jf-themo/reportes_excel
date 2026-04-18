import re
import os
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher
import json

import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

TOKEN = "8770484638:AAFYQjU3SdO9O4c-Klgs4mwdY9xUxuSyGiA"

INVENTARIO_PATH = Path(__file__).parent / 'cod_prod_Tipicos.xlsx'
SALIDA_PATH = Path('G:/Mi unidad/Tipicos del Valle/TipdelValle_Reportes/2026/Reportes_Bot')
SALIDA_PATH.mkdir(parents=True, exist_ok=True)

df_inventario = pd.read_excel(INVENTARIO_PATH)
df_inventario['nombre_lower'] = df_inventario['nombre'].str.lower().str.strip()

def normalizar_palabra(palabra):
    palabra = palabra.lower().strip()
    palabra = palabra.replace('ón', 'on').replace('á', 'a').replace('é', 'e')
    palabra = palabra.replace('í', 'i').replace('ú', 'u').replace('ñ', 'n')
    if palabra.endswith('nes'):
        palabra = palabra[:-3]
    elif palabra.endswith('es'):
        palabra = palabra[:-2]
    elif palabra.endswith('s'):
        palabra = palabra[:-1]
    return palabra

def buscar_producto(nombre_producto):
    nombre_normalizado = normalizar_palabra(nombre_producto)
    mejor_match = None
    mejor_score = 0
    
    for _, row in df_inventario.iterrows():
        nombre_inv = normalizar_palabra(row['nombre'])
        score = SequenceMatcher(None, nombre_normalizado, nombre_inv).ratio()
        if score > mejor_score and score >= 0.4:
            mejor_score = score
            mejor_match = row
    
    if mejor_match is None:
        for _, row in df_inventario.iterrows():
            nombre_inv = row['nombre_lower']
            if nombre_normalizado in nombre_inv or nombre_inv in nombre_normalizado:
                return row
    
    return mejor_match

def parsear_mensaje(texto):
    lineas = texto.strip().split('\n')
    productos = []
    fecha_hoy = datetime.now().strftime('%Y-%m-%d')
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
        
        match = re.match(r'^(\d+)\s+([a-zA-ZáéíóúñÑ\s]+?)\s+\$?([\d,]+)', linea, re.IGNORECASE)
        if match:
            cantidad = int(match.group(1))
            nombre = match.group(2).strip()
            valor_total = int(match.group(3).replace(',', ''))
            
            producto = buscar_producto(nombre)
            if producto is not None and not producto.empty:
                codigo = int(producto['codigo'])
                precio_unitario = int(producto['precio'])
                esperado = cantidad * precio_unitario
                
                if abs(valor_total - esperado) < 1000:
                    precio = precio_unitario
                    validado = True
                else:
                    precio = valor_total // cantidad if cantidad > 0 else 0
                    validado = False
            else:
                codigo = 0
                precio_unitario = 0
                precio = 0
                esperado = 0
                validado = False
            
            productos.append({
                'cantidad': cantidad,
                'nombre': nombre,
                'codigo': codigo,
                'precio': precio,
                'precio_unitario': precio_unitario,
                'total': valor_total,
                'esperado': esperado,
                'validado': validado,
                'fecha': fecha_hoy
            })
    
    return productos

def generar_excel(productos, local='LO-101', airport='SKVP'):
    if not productos:
        return None, None
    
    fecha = productos[0]['fecha']
    fecha_str = datetime.strptime(fecha, '%Y-%m-%d').strftime('%m%d')
    
    numero_base = 4628
    datos = []
    
    for i, p in enumerate(productos):
        datos.append({
            'AEROPUERTO': airport,
            'LOCAL': local,
            'NUMERO FACTURA': numero_base + i + 1,
            'FECHA DE EMISIÓN': fecha,
            'CÓDIGO PRODUCTO': p['codigo'],
            'DESCRIPCION': p['nombre'],
            'CANTIDAD': p['cantidad'],
            'VALOR NETO': p['precio'],
            'IMPUESTOS': None,
            'TOTAL': p['total']
        })
    
    df = pd.DataFrame(datos)
    
    numero_reporte = obtener_numero_reporte(fecha)
    nombre = f'2026{fecha_str}-{numero_reporte:02d}-{local}.xlsx'
    print(f"💾 Guardando: {nombre}")
    ruta = SALIDA_PATH / nombre
    df.to_excel(ruta, index=False)
    
    return ruta, df

def obtener_numero_reporte(fecha):
    fecha_obj = datetime.strptime(fecha, '%Y-%m-%d')
    fecha_str = fecha_obj.strftime('%Y%m%d')
    
    archivos = list(SALIDA_PATH.glob(f'*{fecha_str}*-LO-101.xlsx'))
    return len(archivos) + 1

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
        "Formarto: `cantidad producto valor`",
        parse_mode='Markdown'
    )

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
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
    
    alerta = ""
    no_encontrado = ""
    for p in productos:
        if p['codigo'] == 0:
            no_encontrado += f"❌ {p['nombre']}: no está en inventario\n"
        elif not p['validado']:
            precio_real = p['precio_unitario']
            precio_digitado = p['precio']
            alerta += f"⚠️ {p['nombre']}: precio unitario ${precio_real} → usado ${precio_digitado}\n"
    
    if no_encontrado:
        await update.message.reply_text(no_encontrado, parse_mode='Markdown')
    
    ruta, df = generar_excel(productos)
    
    if ruta:
        total_general = sum(p['total'] for p in productos)
        resumen = f"✅ *Reporte procesado*\n\n"
        resumen += f"📦 *{len(productos)} productos*\n"
        resumen += f"💰 Total: ${total_general:,}\n"
        if alerta:
            resumen += f"\n{alerta}"
        
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
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.add_error_handler(error_handle)
    
    print("🤖 Bot Started. Enviame un mensaje!")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()