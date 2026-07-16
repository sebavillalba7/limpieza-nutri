# Limpieza Nutrición Inferiores

App gratuita (Streamlit) para que los nutricionistas del club carguen la planilla de Excel de cada categoría y extraigan los datos, ya limpios y ordenados, al formato de la base de datos `NUTRI_LONG`.

## Cómo correrla

1. Instalar dependencias (una sola vez):
   ```
   pip install -r requirements.txt
   ```
2. Iniciar la app:
   ```
   streamlit run app.py
   ```
3. Se abre sola en el navegador (`http://localhost:8501`). Es 100% local y gratis, no necesita internet ni licencias.

## Escudo del club

No llegó ningún archivo de imagen adjunto, así que la app arranca con un ícono genérico 🔴⚫. Para que aparezca el escudo real de Unión de Santa Fe: guardá el archivo como

```
assets/escudo.png
```

(en la misma carpeta que `app.py`) y va a mostrarse automáticamente arriba de todo, sin tocar el código.

## Qué hace la app

1. **Cargar planilla**: subís el Excel del club (una hoja por categoría, ej. `4ta (2006)`, `5ta (2008)`, etc.).
2. **Filtros**: elegís una o varias categorías (hojas) y una o varias fechas de evaluación (`FECHEVAL`), con botones para seleccionar todas o ninguna.
3. **Vista previa editable**: aparece una tabla con todos los jugadores encontrados, cada uno con un check de "Incluir". Podés desmarcar los que no querés exportar, y también editar a mano el nombre si el ajuste automático se equivocó (por ejemplo, algún doble apellido poco común o un error de carga en el Excel original).
4. **Exportar**: descargás un CSV con los jugadores seleccionados en formato `NUTRI_LONG`. Si además subís tu base de datos actual (CSV o Excel), la app la fusiona con los nuevos jugadores y te da la base completa actualizada (CSV y Excel), sacando duplicados exactos.

## Mapeo de columnas

| Campo NUTRI_LONG | Columna en el Excel del club |
|---|---|
| JUGADOR | APELLIDO (ajustado a formato "Apellido N") |
| CAT | Columna `CAT` si existe; si no, se completa con el nombre de la hoja (ej. "4ta") |
| FECHA NAC | FECHNAC |
| EDAD | EDAD |
| POS | Columna `POS` si existe; si no, queda vacía |
| PESO | PESO |
| TALLA | TALLA |
| S6PLIEGUES | S6PLIEG |
| IMO | I M/O |

## Formato de nombre de jugador

Se ajusta automáticamente a `Apellido N` (apellido completo + inicial del nombre), sin comas, tildes ni caracteres raros, respetando apellidos compuestos (ej. "Ruiz Díaz Benjamín" → "Ruiz Diaz B"). Los casos ambiguos quedan disponibles para corregir a mano en la vista previa antes de exportar.
