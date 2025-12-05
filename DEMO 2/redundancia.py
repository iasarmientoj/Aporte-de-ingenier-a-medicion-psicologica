import pandas as pd
from sentence_transformers import SentenceTransformer, util

def auditar_banco_items():
    archivo_excel = "Banco de items.xlsx"
    umbral_alerta = 0.70  # 65% de similitud. Ajusta esto según qué tan estricto quieras ser.
    
    print("1. Cargando modelo de lenguaje (esto puede tardar un poco la primera vez)...")
    # Usamos un modelo multilingüe optimizado para paráfrasis
    model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')
    
    print(f"2. Leyendo preguntas de {archivo_excel}...")
    try:
        df = pd.read_excel(archivo_excel, sheet_name='items')
        # Asegurarnos de que leemos las columnas correctas
        ids = df['N'].tolist()
        preguntas = df['PREGUNTA'].tolist()
    except Exception as e:
        print(f"Error leyendo el archivo: {e}")
        return

    print("3. Generando embeddings (vectores semánticos)...")
    # Convertimos cada pregunta en un vector numérico
    embeddings = model.encode(preguntas, convert_to_tensor=True)

    print("4. Calculando matriz de similitud (Coseno)...")
    # Calcula la similitud de todos contra todos
    cosine_scores = util.cos_sim(embeddings, embeddings)

    # Lista para guardar las alertas encontradas
    alertas = []

    # Iteramos sobre la matriz. 
    # Usamos range(len) y el segundo bucle empieza en i+1 para:
    # A) No comparar una pregunta consigo misma.
    # B) No repetir pares (ej: si comparo 1 con 5, no comparar luego 5 con 1).
    for i in range(len(preguntas)):
        for j in range(i + 1, len(preguntas)):
            score = cosine_scores[i][j].item()
            
            if score > umbral_alerta:
                alertas.append({
                    'Item_A': ids[i],
                    'Item_B': ids[j],
                    'Similitud': score,
                    'Pregunta_A': preguntas[i],
                    'Pregunta_B': preguntas[j]
                })

    # --- REPORTE DE RESULTADOS ---
    print("\n" + "="*60)
    print(f"REPORTE DE AUDITORÍA DE ÍTEMS (Umbral > {umbral_alerta*100}%)")
    print("="*60)

    if not alertas:
        print("✅ No se encontraron preguntas redundantes o muy similares.")
    else:
        # Ordenamos las alertas de mayor similitud a menor
        alertas.sort(key=lambda x: x['Similitud'], reverse=True)
        
        for alerta in alertas:
            porcentaje = alerta['Similitud'] * 100
            print(f"⚠️  ALERTA DE REDUNDANCIA DETECTADA: {porcentaje:.2f}% de similitud")
            print(f"   • Ítem {alerta['Item_A']}: {alerta['Pregunta_A']}")
            print(f"   • Ítem {alerta['Item_B']}: {alerta['Pregunta_B']}")
            print("-" * 60)

if __name__ == "__main__":
    auditar_banco_items()
    input('Presione Enter papra salir.')