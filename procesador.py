import re
import os
import sys
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher
import json
import pandas as pd

INVENTARIO_PATH = Path(__file__).parent / 'cod_prod_Tipicos.xlsx'
SALIDA_PATH = Path(__file__).parent / 'Reportes'

class ProcesadorReportes:
    def __init__(self):
        self.df_inventario = pd.read_excel(INVENTARIO_PATH)
        self.df_inventario['nombre_lower'] = self.df_inventario['nombre'].str.lower().str.strip()
        self.salida_path = SALIDA_PATH
        self.salida_path.mkdir(exist_ok=True)

    def buscar_producto(self, nombre_producto):
        nombre_lower = nombre_producto.lower().strip()
        mejor_match = None
        mejor_score = 0

        for _, row in self.df_inventario.iterrows():
            score = SequenceMatcher(None, nombre_lower, row['nombre_lower']).ratio()
            if score > mejor_score and score >= 0.5:
                mejor_score = score
                mejor_match = row

        return mejor_match

    def parsear_mensaje(self, texto):
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

                producto = self.buscar_producto(nombre)
                if producto is not None and not producto.empty:
                    codigo = int(producto['codigo'])
                    precio = int(producto['precio'])
                    esperado = cantidad * precio
                    validado = abs(valor_total - esperado) < 1000
                else:
                    codigo = 0
                    precio = 0
                    esperado = 0
                    validado = False

                productos.append({
                    'cantidad': cantidad,
                    'nombre': nombre,
                    'codigo': codigo,
                    'precio': precio,
                    'total': valor_total,
                    'esperado': esperado,
                    'validado': validado,
                    'fecha': fecha_hoy
                })

        return productos

    def generar_excel(self, productos, local='LO-101', airport='SKVP'):
        if not productos:
            return None

        fecha = productos[0]['fecha']
        fecha_str = datetime.strptime(fecha, '%Y-%m-%d').strftime('%Y%m%d')

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
        nombre = f'2026{fecha_str}-{local}.xlsx'
        ruta = self.salida_path / nombre
        df.to_excel(ruta, index=False)

        return ruta, df

    def procesar(self, mensaje_texto):
        productos = self.parsear_mensaje(mensaje_texto)
        
        if not productos:
            return {'error': 'No se encontraron productos en el mensaje'}
        
        resultado = {'productos': productos, 'total_items': len(productos)}
        
        for p in productos:
            if not p['validado']:
                resultado['alerta'] = f"Producto no validado: {p['nombre']} (esperado: {p['esperado']}, dado: {p['total']})"
        
        ruta, df = self.generar_excel(productos)
        
        if ruta:
            resultado['excel'] = str(ruta)
            resultado['resumen'] = df[['CANTIDAD', 'DESCRIPCION', 'VALOR NETO', 'TOTAL']].to_dict('records')
        
        return resultado


def main():
    if len(sys.argv) > 1:
        if sys.argv[1] == '--json':
            data = json.loads(sys.argv[2])
            mensaje = data.get('text', data.get('message', ''))
        else:
            mensaje = sys.argv[1]
    else:
        mensaje = input("Ingresa el mensaje: ")

    proc = ProcesadorReportes()
    resultado = proc.procesar(mensaje)
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()