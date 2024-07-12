import pandas as pd
import sqlite3
import re
import pyodbc

def get_db_connection():
    # Establecer la cadena de conexión a tu base de datos SQL Server
    server_name = "172.16.0.171"
    database_name = "AlttusTI"
    username = "Socket.Innova"
    password = "ppXtbpuF8AKR2i7P"


    # Crear la cadena de conexión
    conn_str = f"DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};UID={username};PWD={{{password}}};"

    try:
        # Conectar a la base de datos
        conn = pyodbc.connect(conn_str)
        return conn
    except pyodbc.Error as e:
        print(f"Error al conectar a la base de datos: {e}")
        return None


def sql_execute(query, params=None):
    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
        except pyodbc.Error as e:
            print(f"Error al ejecutar la consulta: {e}")
        finally:
            cursor.close()
            conn.close()
def sql_query(query, params=None):
    conn = get_db_connection()
    if conn:
        try:
            cursor = conn.cursor()
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            results = cursor.fetchall()
            return results
        except pyodbc.Error as e:
            print(f"Error al ejecutar la consulta: {e}")
            return None
        finally:
            cursor.close()
            conn.close()
    else:
        return None

def extract_number(s):
    match = re.search(r'\d+', s)
    return int(match.group()) if match else None

def crear_tablas(cursor):
    # Crear la tabla si no existe
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS matriz_tarifaria_servicios (
        matriz_t_s_id INTEGER,
        origen TEXT,
        destino TEXT,
        precio_normal REAL,
        precio_preferente REAL,
        numero_de_servicio INTEGER
    )
    ''')

    cursor.execute('''
    CREATE TABLE IF NOT EXISTS matriz_tarifaria_transbordos (
        matriz_t_t_id INTEGER,
        origen TEXT,
        destino TEXT,
        precio_normal REAL,
        precio_preferente REAL,
        numero_de_servicio INTEGER,
        primer_transbordo TEXT,
        segundo_transbordo TEXT
    )
    ''')

def obtener_ultimo_registro_m_t_s(cursor):
    cursor.execute('''
        SELECT * FROM matriz_tarifaria_servicios
        ORDER BY matriz_t_s_id DESC
        LIMIT 1
    ''')

    # Obtener el resultado
    ultimo_registro = cursor.fetchone()

    # Mostrar el último registro
    return ultimo_registro

def obtener_ultimo_registro_m_t_t(cursor):
    cursor.execute('''
        SELECT * FROM matriz_tarifaria_transbordos
        ORDER BY matriz_t_t_id DESC
        LIMIT 1
    ''')

    # Obtener el resultado
    ultimo_registro = cursor.fetchone()

    # Mostrar el último registro
    return ultimo_registro



# Ruta del archivo Excel
file_path = 'C:/Users/minec/Desktop/ROGER.xlsx'

# Leer el archivo Excel desde la quinta fila (índice 4)
df = pd.read_excel(file_path, sheet_name='V.ESP BOL', header=6)

# Mostrar los nombres de las columnas para verificar
#print("Columnas:", df.columns.tolist())

# Renombrar la segunda columna a 'origen'
df.rename(columns={df.columns[1]: 'origen'}, inplace=True)

numero_de_servicio = extract_number(str(df.columns[0]))

print("El servicio: ", numero_de_servicio)

if pd.notna(numero_de_servicio):
    try:
        numero_de_servicio = int(numero_de_servicio)
    except ValueError:
        print(f"Error al convertir el numero_de_servicio: {numero_de_servicio}")
else:
    numero_de_servicio = None  # Si el número de servicio es nulo, se deja como None

# Verificar algunas filas del DataFrame
#print("Primeras filas del DataFrame:\n", df.head())

# Conectar a la base de datos SQLite3
conn = sqlite3.connect('matrices_tarifarias.db')
cursor = conn.cursor()

crear_tablas(cursor)

# Obtener los nombres de los orígenes para la tabla de transbordos
destinos_transbordo = input("Ingrese los nombres de los orígenes que deben ir en la tabla de transbordos (separados por coma): ").split(',')
destino_final = input("Ingresa el Destino Final de la Matriz Tarifaria: ")

print("Los destinos transbordos: ", destinos_transbordo)
print("El destino final: ", destino_final)

# Procesar y añadir los datos a la base de datos
for index, row in df.iterrows():
    origen = row['origen']
    
    if origen == "Punto Anterior":
        print("Llego al final...")
        break
    
    if pd.notna(origen) and origen != 'origen':  # Evitar filas con 'origen' nulo y el encabezado duplicado
            # Insertar en la tabla de transbordos
            for col in df.columns[2:]:
                
                destino = col
                
                print("Origen:", origen, " - Destino: ", destino)
                precio_normal = row[col]
                #print("Precio normal: ", precio_normal)
                
                # Verificar si el índice siguiente está dentro del rango del DataFrame
                if index + 1 < len(df):
                    precio_preferente = df.at[index + 1, col]
                else:
                    precio_preferente = None
                    
                # if type(precio_normal) == str:
                #     precio_normal = int(precio_normal.replace("$",""))
                # if type(precio_preferente) == str:
                #     precio_preferente = int(precio_preferente.replace("$",""))
                    
                #print("Precio_preferente: ", precio_preferente)
                
                if pd.notna(precio_normal):
                    try:
                        precio_normal = float(precio_normal)
                    except ValueError:
                        print(f"Error al convertir el precio_normal: {precio_normal}")
                        continue
                else:
                    precio_normal = None
                
                if pd.notna(precio_preferente):
                    try:
                        precio_preferente = float(precio_preferente)
                    except ValueError:
                        print(f"Error al convertir el precio_preferente: {precio_preferente}")
                        continue
                else:
                    precio_preferente = None
                    
                #print("Precio normal: ", precio_normal)
                #print("Precio_preferente: ", precio_preferente)

                if precio_normal is None and precio_preferente is None:
                    if destino_final in destino:
                        break
                    continue
                
                if destino.strip() in destinos_transbordo:  # Comprobar si el origen debe ir en la tabla de transbordos
                    print("Transbordo")
                    primer_transbordo = segundo_transbordo = "NE"  # Valores predeterminados para transbordos

                    ultimo_registro = obtener_ultimo_registro_m_t_t(cursor)
                    
                    if ultimo_registro != None:
                        ultimo_registro = int(ultimo_registro[0]) + 1
                    else:
                        ultimo_registro = 20001
                    
                    cursor.execute('''
                    INSERT INTO matriz_tarifaria_transbordos (matriz_t_t_id, origen, destino, precio_normal, precio_preferente, numero_de_servicio, primer_transbordo, segundo_transbordo)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (ultimo_registro, origen, destino, precio_normal, precio_preferente, numero_de_servicio, primer_transbordo, segundo_transbordo))
                    
                    
                    sqlSelect = """SELECT TOP 1 M.Id_Geocerca_Origen, M.Id_Geocerca_Destino 
                    FROM alttusti.matriz_tarifaria_transbordos M 
                    INNER JOIN alttusti.geocerca G ON M.Id_Geocerca_Origen = G.Id_Geocerca 
                    INNER JOIN alttusti.geocerca GD ON M.Id_Geocerca_Destino = GD.Id_Geocerca 
                    WHERE G.Nombre = ?
                    AND GD.Nombre = ?
                    ORDER BY G.Id_geocerca DESC"""
                    results = sql_query(sqlSelect, (origen, destino))
                    if results:
                        for row_2 in results:
                            origen_2, destino_2 = row_2
                            

                            sqlConsulta = """
                            INSERT INTO alttusti.matriz_tarifaria_transbordos_copy (Id_Tarifa_Transbordos,Id_Geocerca_Origen,Id_Geocerca_Destino,Costo_Normal,Costo_Preferente,Id_Servicio,	Primer_Transbordo,Segundo_Transbordo)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """
                            sql_execute(sqlConsulta, (ultimo_registro, origen_2, destino_2, precio_normal, precio_preferente, numero_de_servicio, 0, 0))

                    
                    if destino_final in destino:
                        break
                else:
                    print("Directo")
                    
                    ultimo_registro = obtener_ultimo_registro_m_t_s(cursor)
                    
                    if ultimo_registro != None:
                        ultimo_registro = int(ultimo_registro[0]) + 1
                    else:
                        ultimo_registro = 20001
                        
                    cursor.execute('''
                    INSERT INTO matriz_tarifaria_servicios (matriz_t_s_id, origen, destino, precio_normal, precio_preferente, numero_de_servicio)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ''', (ultimo_registro, origen, destino, precio_normal, precio_preferente, numero_de_servicio))



                    sqlSelect = """SELECT TOP 1 M.Id_Geocerca_Origen, M.Id_Geocerca_Destino 
                    FROM alttusti.matriz_tarifaria_servicios M 
                    INNER JOIN alttusti.geocerca G ON M.Id_Geocerca_Origen = G.Id_Geocerca 
                    INNER JOIN alttusti.geocerca GD ON M.Id_Geocerca_Destino = GD.Id_Geocerca 
                    WHERE G.Nombre = ?
                    AND GD.Nombre = ?
                    ORDER BY G.Id_geocerca DESC"""
                    results = sql_query(sqlSelect, (origen, destino))
                    if results:
                        for row_2 in results:
                            origen_2, destino_2 = row_2
                                
                            sqlConsulta = """
                            INSERT INTO alttusti.matriz_tarifaria_servicios_copy (Id_Tarifa_Servicio,Id_Geocerca_Origen,Id_Geocerca_Destino,Costo_Normal,Costo_Preferente,Id_Servicio)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """
                            sql_execute(sqlConsulta, (ultimo_registro, origen_2, destino_2, precio_normal, precio_preferente, numero_de_servicio))
                    
                    if destino_final in destino:
                        break

# Confirmar los cambios y cerrar la conexión
conn.commit()
conn.close()

print("Datos insertados correctamente.")