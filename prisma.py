import pandas as pd

# 1. Cargar el archivo 
archivo_entrada = 'prueba.csv' 
try:
    df = pd.read_csv(archivo_entrada, quotechar='"', skipinitialspace=True, on_bad_lines='warn')
    df.columns = df.columns.str.strip()
    print(f"Cargados {len(df)} registros.")
except Exception as e:
    print(f"Error al cargar: {e}")

# 2. Definir las Keywords de tu RSL
keywords = ['retail', 'sales', 'anomaly', 'outlier', 'unsupervised', 'isolation forest', 'lof']

# 3. Aplicar el filtrado
def filtrar_prisma(row):
    texto = (str(row.get('Title', '')) + " " + 
             str(row.get('Abstract', '')) + " " + 
             str(row.get('Author Keywords', ''))).lower()
    
    if any(k in texto for k in keywords):
        return 'Pre_INCLUIDO'
    return 'Pre_EXCLUIDO'

df['Estado_PRISMA'] = df.apply(filtrar_prisma, axis=1)

# 4. Crear los DataFrames para cada hoja
df_incluidos = df[df['Estado_PRISMA'] == 'Pre_INCLUIDO'].copy()
df_excluidos = df[df['Estado_PRISMA'] == 'Pre_EXCLUIDO'].copy()

# 5. Guardar en un archivo Excel con múltiples hojas
nombre_excel = 'Resultados_PRISMA2_Antoni.xlsx'

with pd.ExcelWriter(nombre_excel, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Data_Total', index=False)
    df_incluidos.to_excel(writer, sheet_name='Pre_Incluidos', index=False)
    df_excluidos.to_excel(writer, sheet_name='Pre_Excluidos', index=False)

print(f"\n¡Éxito! Se ha creado el archivo: {nombre_excel}")
print(f"Hoja 'Pre_Incluidos': {len(df_incluidos)} registros.")
print(f"Hoja 'Pre_Excluidos': {len(df_excluidos)} registros.")
