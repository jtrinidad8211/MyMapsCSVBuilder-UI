# MyMaps CSV Builder UI

Interfaz React + TypeScript + Vite para convertir un Excel en un paquete CSV + KML + KMZ preparado para Google My Maps.

## Desarrollo

Inicia primero el API en `http://localhost:5237` y luego ejecuta:

```powershell
npm install
npm run dev
```

Vite publica la UI en `http://localhost:5173` y redirige `/api` al backend local.

## Flujo

1. Carga un libro `.xlsx` o `.xlsm`.
2. Indica la fila que contiene los encabezados y selecciona la hoja.
3. Elige las columnas del CSV e identifica el código, el nombre del cliente y si las coordenadas están juntas, separadas o no existen en el Excel.
4. Completa solo coordenadas faltantes o vuelve a consultar todas.
5. Descarga un ZIP con el CSV, el KML y un KMZ. Importa el KML o KMZ en My Maps y mantén la capa en **Estilos individuales** para aplicar los colores del Excel; la columna `Color` del CSV por sí sola no controla el estilo. La opción adicional **Generar la carpeta capas-por-leyenda** crea un KML por elemento, numerado según el orden configurado, para importarlos como capas separadas; puede dejarse desactivada si no se necesita. My Maps no permite cambiar el orden que aplica dentro de una sola capa agrupada por `Leyenda`.

El color del marcador se toma del relleno de la celda que contiene el código del cliente. Las filas sin color usan azul de forma predeterminada. El nombre del cliente se guarda como descripción del marcador. Las filas que no tengan coordenadas válidas se entregan aparte en `*-sin-coordenadas.csv` y no se mezclan con el CSV que se importará en My Maps.

## Verificación

```powershell
npm run lint
npm run build
```
