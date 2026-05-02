import re
import os
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
from PIL import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

INVENTARIO_PATH = Path(__file__).parent / 'cod_prod_Tipicos.xlsx'
SALIDA_PATH = Path(__file__).parent / 'Reportes'

class ProcesadorReportes:
    def __init__(self, inventario_path=INVENTARIO_PATH):
        self.df_inventario = pd.read_excel(inventario_path)
        self.df_inventario['nombre_lower'] = self.df_inventario['nombre'].str.lower().str.strip()
        self.salida_path = SALIDA_PATH
        self.salida_path.mkdir(exist_ok=True)

    def OCR_imagen(self, imagen_path):
        img = Image.open(imagen_path)
        texto = pytesseract.image_to_string(img, lang='spa')
        return texto

    def _encontrar_producto(self, nombre_producto):
        nombre_lower = nombre_producto.lower().strip()
        mejor_match = None
        mejor_score = 0

        for _, row in self.df_inventario.iterrows():
            score = SequenceMatcher(None, nombre_lower, row['nombre_lower']).ratio()
            if score > mejor_score and score >= 0.6:
                mejor_score = score
                mejor_match = row

        return mejor_match

    def _parsear_mensaje(self, texto):
        lineas = texto.strip().split('\n')
        productos = []
        current_fecha = datetime.now().strftime('%Y-%m-%d')

        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue

            match = re.match(r'^(\d+)\s+(.+?)\s+\$?([\d,]+)\s*$', linea, re.IGNORECASE)
            if match:
                cantidad = int(match.group(1))
                nombre = match.group(2).strip()
                valor_str = match.group(3).replace(',', '')
                valor_total = int(valor_str)

                producto_encontrado = self._encontrar_producto(nombre)
                if producto_encontrado:
                    codigo = producto_encontrado['codigo']
                    precio_unitario = producto_encontrado['precio']
                    valor_esperado = cantidad * precio_unitario
                   	validado = abs(valor_total - valor_esperado) < 1000
                else:
                    codigo = 'NO ENCONTRADO'
                    precio_unitario = 0
                    valor_esperado = 0
                    validado = False

                productos.append({
                    'cantidad': cantidad,
                    'nombre_original': nombre,
                    'codigo_encontrado': codigo,
                    'precio_unitario': precio_unitario,
                    'valor_total': valor_total,
                    'valor_esperado': valor_esperado,
                    'validado': validado,
                    'fecha': current_fecha
                })

        return productos

    def generar_excel(self, productos, numero_local='LO-101', airport='SKVP'):
        if not productos:
            print("No hay productos para procesar")
            return None

        fecha = productos[0]['fecha']
        fecha_str = datetime.strptime(fecha, '%Y-%m-%d').strftime('%Y%m%d')

        ultimo_numero = 4628
        datos = []

        for i, prod in enumerate(productos):
            datos.append({
                'AEROPUERTO': airport,
                'LOCAL': numero_local,
                'NUMERO FACTURA': ultimo_numero + i + 1,
                'FECHA DE EMISIÓN': fecha,
                'CÓDIGO PRODUCTO': prod['codigo_encontrado'],
                'DESCRIPCION': prod['nombre_original'],
                'CANTIDAD': prod['cantidad'],
                'VALOR NETO': prod['precio_unitario'],
                'IMPUESTOS': None,
                'TOTAL': prod['valor_total']
            })

        df = pd.DataFrame(datos)
        nombre_archivo = f'2026{fecha_str}-{numero_local}.xlsx'
        ruta_salida = self.salida_path / nombre_archivo
        df.to_excel(ruta_salida, index=False)

        return ruta_salida, df

    def procesar_imagen(self, ruta_imagen):
        print(f"Procesando: {ruta_imagen}")
        texto = self.OCR_imagen(ruta_imagen)
        print("--- Texto extraído ---")
        print(texto)
        print("---------------------")

        productos = self._parsear_mensaje(texto)
        print(f"\nProductos encontrados: {len(productos)}")

        for p in productos:
            print(f"  - {p['cantidad']} x {p['nombre_original']}: ${p['valor_total']} (Validado: {p['validado']})")

        if productos:
            ruta, df = self.generar_excel(productos)
            print(f"\nExcel generado: {ruta}")
            return ruta, df
        return None, None


def main():
    import sys
    if len(sys.argv) > 1:
        imagen = sys.argv[1]
    else:
        imagen = '13_abril_imagen.png'

    procesador = ProcesadorReportes()
    ruta, df = procesador.procesar_imagen(imagen)

    if df is not None:
        print("\n=== RESUMEN ===")
        print(df[['CANTIDAD', 'DESCRIPCION', 'VALOR NETO', 'TOTAL']].to_string())


if __name__ == '__main__':
    main()