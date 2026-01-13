import mysql.connector
from datetime import datetime
import sys
import os
import io
import mmap

def parse_date(date_str):
    if date_str and date_str.strip():
        try:
            return datetime.strptime(date_str, '%d/%m/%Y').strftime('%Y-%m-%d')
        except:
            return None
    return None

def create_table(conn):
    try:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS persona (
                dni VARCHAR(8),
                apellido_paterno VARCHAR(100),
                apellido_materno VARCHAR(100),
                nombres VARCHAR(100),
                fecha_nacimiento DATE,
                fecha_inscripcion DATE,
                fecha_emision DATE, 
                fecha_caducidad DATE,
                ubigeo_nacimiento VARCHAR(6),
                ubigeo_direccion VARCHAR(6),
                direccion TEXT,
                sexo CHAR(1),
                estado_civil VARCHAR(20),
                digito_verificador VARCHAR(1),
                madre VARCHAR(200),
                padre VARCHAR(200),
                PRIMARY KEY (dni)
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
        """)
        conn.commit()
        print("Tabla creada o ya existente")
    except Exception as e:
        print(f"Error creando la tabla: {str(e)}")
        sys.exit(1)

def process_file_in_batches(filename, batch_size=1000, buffer_size=8192):
    # Configuración de la conexión
    db_config = {
        'host': '192.168.8.217',
        'user': 'xd',
        'password': 'xddd',
        'database': 'dbsistemamunicipal',
        'charset': 'utf8mb4',
        'connect_timeout': 28800,  # 8 horas
        'pool_size': 5
    }

    # Conectar a la base de datos
    try:
        conn = mysql.connector.connect(**db_config)
        print("Conexión exitosa a la base de datos")
        
        # Configurar timeouts más largos
        cursor = conn.cursor()
        cursor.execute("SET SESSION wait_timeout=28800")  # 8 horas
        cursor.execute("SET SESSION interactive_timeout=28800")  # 8 horas
        conn.commit()
    except Exception as e:
        print(f"Error conectando a la base de datos: {str(e)}")
        return

    # Crear la tabla si no existe
    create_table(conn)

    cursor = conn.cursor()

    # Optimizaciones de MySQL
    cursor.execute("SET FOREIGN_KEY_CHECKS=0")
    cursor.execute("SET UNIQUE_CHECKS=0")
    cursor.execute("SET AUTOCOMMIT=0")
    cursor.execute("SET NET_WRITE_TIMEOUT=7200")
    cursor.execute("SET NET_READ_TIMEOUT=7200")

    batch = []
    total_processed = 0
    errors = 0
    partial_line = ''

    try:
        print(f"Iniciando procesamiento del archivo: {filename}")
        
        # Abrir el archivo en modo binario
        with open(filename, 'rb') as file:
            # Crear un objeto mmap
            with mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                # Decodificar como UTF-8
                text = mm.read().decode('utf-8', errors='replace')
                lines = text.split('\n')
                
                for line_number, line in enumerate(lines, 1):
                    try:
                        # Ignorar líneas vacías o que solo contienen |
                        if not line.strip() or line.strip() == '|':
                            continue

                        # Dividir la línea por el separador |
                        fields = [f.strip() for f in line.split('|')]
                        
                        # Si tenemos los campos esperados
                        if len(fields) >= 16:
                            record = (
                                fields[0],  # dni
                                fields[1],  # apellido_paterno
                                fields[2],  # apellido_materno
                                fields[3],  # nombres
                                parse_date(fields[4]),  # fecha_nacimiento
                                parse_date(fields[5]),  # fecha_inscripcion
                                parse_date(fields[6]),  # fecha_emision
                                parse_date(fields[7]),  # fecha_caducidad
                                fields[8],  # ubigeo_nacimiento
                                fields[9],  # ubigeo_direccion
                                fields[10], # direccion
                                fields[11], # sexo
                                fields[12], # estado_civil
                                fields[13], # digito_verificador
                                fields[14], # madre
                                fields[15]  # padre
                            )
                            batch.append(record)

                        # Cuando el lote alcanza el tamaño especificado, insertamos
                        if len(batch) >= batch_size:
                            try:
                                cursor.executemany("""
                                    INSERT IGNORE INTO persona (
                                        dni, apellido_paterno, apellido_materno, nombres,
                                        fecha_nacimiento, fecha_inscripcion, fecha_emision,
                                        fecha_caducidad, ubigeo_nacimiento, ubigeo_direccion,
                                        direccion, sexo, estado_civil, digito_verificador,
                                        madre, padre
                                    ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                                """, batch)
                                conn.commit()
                                total_processed += len(batch)
                                print(f"Procesados {total_processed} registros")
                                batch = []
                            except mysql.connector.Error as e:
                                if e.errno == 2006:  # MySQL server has gone away
                                    print("Reconectando a la base de datos...")
                                    conn = mysql.connector.connect(**db_config)
                                    cursor = conn.cursor()
                                    continue
                                print(f"Error en el lote cerca de la línea {line_number}: {str(e)}")
                                errors += 1
                                batch = []
                                conn.rollback()

                    except Exception as e:
                        print(f"Error procesando línea {line_number}: {str(e)}")
                        errors += 1
                        continue

            # Procesar la última línea parcial si existe
            if partial_line:
                # Procesar partial_line igual que las líneas anteriores
                pass

            # Procesar el último lote si existe
            if batch:
                try:
                    cursor.executemany("""
                        INSERT IGNORE INTO persona (
                            dni, apellido_paterno, apellido_materno, nombres,
                            fecha_nacimiento, fecha_inscripcion, fecha_emision,
                            fecha_caducidad, ubigeo_nacimiento, ubigeo_direccion,
                            direccion, sexo, estado_civil, digito_verificador,
                            madre, padre
                        ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    """, batch)
                    conn.commit()
                    total_processed += len(batch)
                    print(f"Procesados {total_processed} registros")
                except Exception as e:
                    print(f"Error en el último lote: {str(e)}")
                    errors += 1
                    conn.rollback()

    except Exception as e:
        print(f"Error general: {str(e)}")
    finally:
        try:
            # Restaurar configuración de MySQL
            cursor.execute("SET FOREIGN_KEY_CHECKS=1")
            cursor.execute("SET UNIQUE_CHECKS=1")
            cursor.execute("SET AUTOCOMMIT=1")
            
            cursor.close()
            conn.close()
        except:
            pass
        
        print("\nResumen del proceso:")
        print(f"Total de registros procesados: {total_processed}")
        print(f"Total de errores encontrados: {errors}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python script.py <ruta_del_archivo>")
        sys.exit(1)
    
    filename = sys.argv[1]
    print(f"Iniciando procesamiento del archivo: {filename}")
    process_file_in_batches(filename)
