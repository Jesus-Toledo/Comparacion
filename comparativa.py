import pandas as pd

# Cargar archivos
agosto = pd.read_excel("Alarmas_Agosto.xlsx")
septiembre = pd.read_excel("Alarmas_Septiembre.xlsx")

# Asegurarse de que los IDs son texto y están limpios
agosto['Alarm_ID'] = agosto['Alarm_ID'].astype(str).str.strip()
septiembre['Alarm_ID'] = septiembre['Alarm_ID'].astype(str).str.strip()

# Extraer sets de IDs
agosto_ids = set(agosto['Alarm_ID'])
septiembre_ids = set(septiembre['Alarm_ID'])

# Comparaciones
repetidos = sorted(list(agosto_ids & septiembre_ids))
cerrados = sorted(list(agosto_ids - septiembre_ids))
nuevos = sorted(list(septiembre_ids - agosto_ids))

# Igualar longitudes de las listas para que puedan ir en columnas
max_len = max(len(repetidos), len(cerrados), len(nuevos))
repetidos += [None] * (max_len - len(repetidos))
cerrados += [None] * (max_len - len(cerrados))
nuevos += [None] * (max_len - len(nuevos))

# Crear DataFrame con 3 columnas
df_resultados = pd.DataFrame({
    'Repetidos': repetidos,
    'Cerrados': cerrados,
    'Nuevos': nuevos
})

# Guardar en Excel
df_resultados.to_excel("comparacion_tickets_final.xlsx", index=False)