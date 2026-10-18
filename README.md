# Tenor de nuevo lote compuesto

## Instalación
    python -m venv venv
    venv\Scripts\activate        # Windows  (Linux/Mac: source venv/bin/activate)
    pip install -r requirements.txt

## Ejecución
    streamlit run app.py
Se abre en http://localhost:8501.

## Uso
1. Carga "Comercialización Concentrados Alto Tenor.xlsx" (solo se lee la hoja OFICIAL).
2. Elige cómo se compone el nuevo lote:
   - **Un solo rango o lista**: `601-607` o `601, 605, 610` + un peso total (se reparte en partes iguales entre sus lotes).
   - **Varios rangos**: cada rango con su propia cantidad de material (ej. `601-612` → 800 y `615-649` → 1200). Con ➕ se agregan más rangos.
3. Elige si los números son el valor de la columna LOTE o la fila de Excel.
4. Calcular → descargar el .xlsx → súbelo a Google Drive y ábrelo con Hojas de cálculo de Google.

## Fórmula
Por rango: peso_por_lote = peso_rango / n_rango ; aporte_i = peso_por_lote × tenor_i
Nuevo lote: tenor = Σ aportes (de todos los rangos) / Σ pesos de los rangos
(con un solo rango es la fórmula original)

## Pruebas rápidas
    python test_calculo.py
