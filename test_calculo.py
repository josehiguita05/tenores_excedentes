"""Pruebas con un Excel sintético."""
import io
import pandas as pd
import calculo as C

filas = [["LOTE", "Tenor Au Quintana (g/t)"]] + [[600 + i, t] for i, t in
         enumerate([10, 20, "30,5", None, "abc", 40, 50, 60], start=1)]
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    pd.DataFrame(filas).to_excel(w, sheet_name="OFICIAL", header=False, index=False)
    pd.DataFrame([[1]]).to_excel(w, sheet_name="OTRA", header=False, index=False)
buf.seek(0)
df = C.leer_hoja_oficial(buf)

# --- un solo rango (comportamiento original)
r = C.procesar(df, [("601-603", "1500,5")])
assert abs(r.tenor_nuevo - (10 + 20 + 30.5) / 3) < 1e-9
r = C.procesar(df, [("601, 606", "100")])
assert abs(r.tenor_nuevo - 25) < 1e-9
r = C.procesar(df, [("2-3", "100")], modo="fila")        # filas 2 y 3 de Excel = lotes 601 y 602
assert r.grupos[0].lotes == ["601", "602"]

# --- varios rangos, cada uno con su peso
# R1: 601-602, 100 -> aportes 50*10 + 50*20 = 1500 ; R2: 606-608, 300 -> 100*(40+50+60) = 15000
r = C.procesar(df, [("601-602", "100"), ("606-608", "300")])
assert abs(r.tenor_nuevo - 16500 / 400) < 1e-9, r.tenor_nuevo
assert r.cantidad == 5 and r.peso_total == 400
assert abs(r.grupos[0].tenor_grupo - 15) < 1e-9 and abs(r.grupos[1].tenor_grupo - 50) < 1e-9
# rangos con pesos iguales y mismo n => promedio simple de los tenores de los grupos
assert abs(C.procesar(df, [("601", "10"), ("602", "10")]).tenor_nuevo - 15) < 1e-9

# --- errores
casos = [
    [("601-604", "100")],                         # tenor vacío
    [("601-605", "100")],                         # vacío + texto
    [("999", "100")],                             # lote inexistente
    [("601", "-5")], [("601", "abc")], [("", "10")], [("9-1", "10")],
    [("601-602", "10"), ("602", "10")],           # lote en dos rangos
    [("601-602", "10"), ("606-607", "")],         # peso faltante en rango 2
    [("601-602", "10"), ("999", "5")],            # lote inexistente en rango 2
]
for c in casos:
    try:
        C.procesar(df, c); raise SystemExit(f"Debió fallar: {c}")
    except C.ErrorDatos as e:
        print("OK error ->", e)

open("/tmp/prueba.xlsx", "wb").write(
    C.generar_documento(C.procesar(df, [("601-602", "100"), ("606-608", "300")]), "NL-TEST", "kg"))
print("Todas las pruebas pasaron")
