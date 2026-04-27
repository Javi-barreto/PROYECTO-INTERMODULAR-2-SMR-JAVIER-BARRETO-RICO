import pandas as pd
import os

archivo_entrada = "ANALISIS-DE-FIREWALL.xlsx"
archivo_salida = "resultado_regla_3.xlsx"

columnas_necesarias = [
    "src_ip",
    "dns_src_IP",
    "dest_ip",
    "dest_port",
    "app",
    "action",
    "count",
    "total_bytes"
]

origen_objetivo = "10.216.64.4"

extension = os.path.splitext(archivo_entrada)[1].lower()

if extension == ".xlsx":
    df = pd.read_excel(archivo_entrada)
elif extension == ".csv":
    df = pd.read_csv(archivo_entrada)
else:
    raise ValueError("Formato no soportado. Usa .xlsx o .csv")

df.columns = (
    df.columns.astype(str)
    .str.strip()
    .str.replace("\n", "", regex=False)
    .str.replace("\r", "", regex=False)
)

faltan = [col for col in columnas_necesarias if col not in df.columns]
if faltan:
    raise ValueError(f"Faltan estas columnas en el archivo: {faltan}")

df = df[columnas_necesarias].copy()

df["src_ip"] = df["src_ip"].astype(str).str.strip()
df["dns_src_IP"] = df["dns_src_IP"].astype(str).str.strip()
df["dest_ip"] = df["dest_ip"].astype(str).str.strip()
df["app"] = df["app"].astype(str).str.strip().str.lower()
df["action"] = df["action"].astype(str).str.strip().str.lower()

df["dest_port"] = pd.to_numeric(df["dest_port"], errors="coerce")
df["count"] = pd.to_numeric(df["count"], errors="coerce")
df["total_bytes"] = pd.to_numeric(df["total_bytes"], errors="coerce")

filtro = df["src_ip"] == origen_objetivo

resultado = df[filtro].copy()
resultado = resultado.sort_values(by=["total_bytes", "count"], ascending=False)

destinos_unicos = (
    resultado[["dest_ip"]]
    .drop_duplicates()
    .sort_values(by="dest_ip")
    .reset_index(drop=True)
)

resumen = pd.DataFrame({
    "metrica": [
        "total_filas_analizadas",
        "total_coincidencias",
        "origen_filtrado",
        "total_destinos_unicos"
    ],
    "valor": [
        len(df),
        len(resultado),
        origen_objetivo,
        destinos_unicos["dest_ip"].nunique()
    ]
})

if archivo_salida.endswith(".xlsx"):
    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        resultado.to_excel(writer, sheet_name="coincidencias", index=False)
        destinos_unicos.to_excel(writer, sheet_name="destinos_unicos", index=False)
        resumen.to_excel(writer, sheet_name="resumen", index=False)
else:
    resultado.to_csv(archivo_salida, index=False)

print(f"Filas analizadas: {len(df)}")
print(f"Coincidencias encontradas: {len(resultado)}")
print(f"Destinos únicos: {destinos_unicos['dest_ip'].nunique()}")
print(f"Archivo generado: {archivo_salida}")