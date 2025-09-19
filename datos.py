import pandas as pd

# Cargar archivos
agosto = pd.read_excel("Alarmas_agosto.xlsx")
septiembre = pd.read_excel("Alarmas_septiembre.xlsx")

# Asegurarse de que los IDs son texto y están limpios
agosto['Alarm_ID'] = agosto['Alarm_ID'].astype(str).str.strip()
septiembre['Alarm_ID'] = septiembre['Alarm_ID'].astype(str).str.strip()

# Extraer sets de IDs
agosto_ids = set(agosto['Alarm_ID'])
septiembre_ids = set(septiembre['Alarm_ID'])

# Comparaciones de IDs
repetidos_ids = agosto_ids & septiembre_ids
cerrados_ids = agosto_ids - septiembre_ids
nuevos_ids = septiembre_ids - agosto_ids

# Filtrar las filas completas según los IDs
repetidos = pd.concat([
    agosto[agosto['Alarm_ID'].isin(repetidos_ids)],
    septiembre[septiembre['Alarm_ID'].isin(repetidos_ids)]
], axis=0, ignore_index=True)

nuevos = septiembre[septiembre['Alarm_ID'].isin(nuevos_ids)]

cerrados = agosto[agosto['Alarm_ID'].isin(cerrados_ids)]

# Guardar en Excel con 3 hojas
with pd.ExcelWriter("comparacion_tickets_final.xlsx", engine="openpyxl") as writer:
    repetidos.to_excel(writer, sheet_name="Repetidos", index=False)
    nuevos.to_excel(writer, sheet_name="Nuevos", index=False)
    cerrados.to_excel(writer, sheet_name="Cerrados", index=False)
