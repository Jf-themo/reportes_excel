# Plan de Ejecución: Automatización Reportes de Ventas WhatsApp

## 1. Arquitectura

```
[Reporte WhatsApp] → [Reenviar a grupo con Bot N8N]
                                        ↓
                              [N8N procesa texto]
                                        ↓
                              [Genera Excel LO-101]
                                        ↓
                              [Guarda en carpeta]
```

## 2. Flujo Completo

### Paso 1: Llega el mensaje
Ejemplo formato que recibes:
```
3 Empanada 15000
5 Tinto 15000
2 Chicharon 60000
```

### Paso 2: Reenvías al grupo
El grupo tiene el bot de N8N que escucha mensajes nuevos.

### Paso 3: N8N procesa
- Extrae líneas: cantidad + producto + valor
- Busca código en inventario
- Valida: cantidad × precio = valor
- Genera Excel con formato LO-101
- Guarda en `Reportes/`

### Paso 4: Listo
Excel generado con nombre `2026MMDD-01-LO-101.xlsx`

## 3. Componentes

### A. Script Python (`procesador.py`)
- Parsea texto: "3 Empanada 15000"
- Busca código de producto
- Valida contra inventario
- Genera Excel LO-101

### B. Flujo N8N
- Trigger: Cuando llegue mensaje al grupo
- Ejecuta script Python
- Guarda Excel

## 4. Requisitos

- N8N cloud o local con WhatsApp Business API
- Carpeta `Reportes/` para Excel generados
- Archivo `cod_prod_Tipicos.xlsx` (ya lo tienes)

## 5. Próximo paso

Necesito confirmar:
1. ¿Ya tienes N8N configurado con WhatsApp Business?
2. ¿El grupo ya tiene el bot o necesitas crearlo?

Si tienes N8N funcionando, te passo el workflow completo para importar.