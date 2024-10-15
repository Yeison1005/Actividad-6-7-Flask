# Importando Libreria mysql.conector para conectar Python con MYSQL
import mysql.connector


def connectionBD():
    try:
        # connection = mysql.conector.connect
        connection = mysql.conector.connect(
            host="localhost",
            user="root",
            passwd="",
            database="app_empresa_bd",
            charset='utf8mb4',
            collation='utf8mb4_unicode_ci',
            raise_on_warnings=True
        )
        if connection.is_connected():
            # print("conexion exitosa a la BD")
            return connection
        
    except mysql.connector.Error as error:
        print(f"no se pudo conectar:{error}")