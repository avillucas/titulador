import json

# Si ya tienes ambos archivos en tu carpeta de trabajo:
with open("certificaciones_validez_nacional.json", "r", encoding="utf-8") as f:
    circ_items = json.load(f)["certificaciones_validez_nacional"]

with open("CATALOGO_2025_FP_agrupado.json", "r", encoding="utf-8") as f:
    catalogo = json.load(f)

# Generación del índice por resolución/nombre
circ_dict = {
    c["resolucion_jurisdiccional"].strip().lower(): c 
    for c in circ_items
}

resultado = []
for item in catalogo:
    res_j = item.get("Res Jurisdiccional", "").strip().lower()
    match = circ_dict.get(res_j)
    
    if match:
        cert = f"{match['trayecto']} según Resolución Nro. {match['resolucion_validez_nacional']}"
        cohorte = f"{match['cohortes_desde']} a {match['cohortes_hasta']}"
        res_vn = match["resolucion_validez_nacional"]
        res_j_val = match["resolucion_jurisdiccional"]
        titulo = match["trayecto"]
    else:
        cert = item.get("Certificación", "")
        cohorte = item.get("Cohortes asociadas a Res de V. N.", "")
        res_vn = item.get("Res Validez Nacional", "")
        res_j_val = item.get("Res Jurisdiccional", "")
        titulo = item.get("Nombre del Trayecto", "")

    resultado.append({
        "codigo": item.get("Código", ""),
        "trayecto_titulo": titulo,
        "certificacion": cert,
        "cohorte": cohorte,
        "resolucion_jurisdiccional": res_j_val,
        "resolucion_validez_nacional": res_vn,
        "modulos": item.get("Módulos Acreditables", [])
    })

with open("catalogo_completo_integrado.json", "w", encoding="utf-8") as out:
    json.dump(resultado, out, ensure_ascii=False, indent=2)

print(f"Generado con éxito: {len(resultado)} registros.")