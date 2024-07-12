import sqlite3

# Configuración de las bases de datos
origen_db = 'matrices_tarifarias_org.db'
destino_db = 'matrices_tarifarias.db'

# Conexión a la base de datos de origen
origen_conn = sqlite3.connect(origen_db)
origen_cursor = origen_conn.cursor()

# Conexión a la base de datos de destino
destino_conn = sqlite3.connect(destino_db)
destino_cursor = destino_conn.cursor()

# Consulta para obtener los datos de los campos primer_transbordo y segundo_transbordo de la tabla matriz_tarifaria_transbordos
query_select = """
SELECT numero_de_servicio, origen, destino, primer_transbordo, segundo_transbordo
FROM matriz_tarifaria_transbordos
"""

# Consulta para actualizar los campos primer_transbordo y segundo_transbordo en la base de datos de destino
query_update = """
UPDATE matriz_tarifaria_transbordos
SET primer_transbordo = ?, segundo_transbordo = ?
WHERE numero_de_servicio = ? AND origen = ? AND destino = ?
"""

try:
    # Ejecutar la consulta en la base de datos de origen
    origen_cursor.execute(query_select)
    rows = origen_cursor.fetchall()
    
    # Actualizar los datos en la base de datos de destino
    for row in rows:
        numero_de_servicio, origen, destino, primer_transbordo, segundo_transbordo = row
        print("Haciendo :", numero_de_servicio, " ", origen, " ", destino)
        destino_cursor.execute(query_update, (primer_transbordo, segundo_transbordo, numero_de_servicio, origen, destino))
    
    # Confirmar los cambios en la base de datos de destino
    destino_conn.commit()

except Exception as e:
    print(f"Error: {e}")
    destino_conn.rollback()

finally:
    # Cerrar los cursores y conexiones
    origen_cursor.close()
    origen_conn.close()
    destino_cursor.close()
    destino_conn.close()