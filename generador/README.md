# Generador de datos de Consulta CECOs

Este generador convierte los Excel de Nisira en los archivos livianos que utiliza la web `Consulta_Cecos`.

## Archivos de entrada

Coloque dentro de la carpeta `entradas`:

- `Cecos_ASA.xlsx`
- `Cecos_IA.xlsx`
- `CECOS_Historico.xlsx`
- `ASA_DataCostos.xlsx`
- `IA_DataCostos.xlsx`

También puede colocar una copia del `cecos.json` actual para conservar la lista `uso_planillas`. Si el generador se encuentra dentro del repositorio, intentará encontrar automáticamente el `cecos.json` de la raíz.

## Uso

1. Instale Python una sola vez desde <https://www.python.org/downloads/>. Durante la instalación marque `Add Python to PATH`.
2. Copie los cinco Excel dentro de `entradas`.
3. Haga doble clic en `Actualizar_CECOs.bat`.
4. Espere el mensaje `GENERACION TERMINADA`.
5. Revise la carpeta `salida`.

## Archivos generados

- `salida/cecos.json`
- `salida/historico.json`
- `salida/consumos/index.json`
- `salida/consumos/ASA_###.json`
- `salida/consumos/INTERANDINA_###.json`
- `salida/reporte_generacion.json`

El generador no publica importes. Solamente conserva fecha, detalle, cuenta, cantidad, labor, partida, rubro y número de movimientos.

Los archivos originales no se modifican.
