# Instrucciones: Ejecutar Bot con Windows

## Opción 1: Tarea Programada (recomendado)

### Crear la tarea:

1. Presiona `Win + R`
2. Escribe `taskschd.msc` y presiona Enter
3. Clic en **"Crear tarea básica"**

4. Configura:
   - **Nombre**: `ReportesBot`
   - **Desencadenar**: `Al iniciar sesión`
   - **Acción**: `Iniciar un programa`

5. Programa:
   - **Programa/script**: `C:\Users\hh\Documents\Reportes_Excel\iniciar_bot.bat`
   - **Iniciar en**: `C:\Users\hh\Documents\Reportes_Excel`

6. Finaliza

---

## Opción 2: Inicio automático simple

1. Presiona `Win + R`
2. Escribe `shell:startup` y presiona Enter
3. Arrastra `iniciar_bot.bat` a esa carpeta

---

## Opción 3: Batch con reinicio automático

Si quieres que se reinicie solo si se cierra:

```batch
@echo off
cd /d "C:\Users\hh\Documents\Reportes_Excel"
:bucle
python bot_telegram.py
echo Bot cerrado, reiniciando...
timeout /t 5
goto :bucle
```

Guarda como `iniciar_bot_loop.bat` y úsalo en la tarea programada.

---

## Verificar que funciona

1. Reinicia tu PC
2. Busca en Telegram tu bot
3. Envíale un mensaje de prueba:
   ```
   3 Empanada 15000
   ```
4. Debe responder con el Excel