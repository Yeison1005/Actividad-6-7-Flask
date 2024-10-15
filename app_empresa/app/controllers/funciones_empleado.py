#para subir archivos tipo foto al servidor
from werkzeug.utils import secure_filename
import uuid # modulo de python para crear un string

from config import connectionBD #conexion a BD

import datetime
import re
import os

from os import remove #modulo para remover archivo
from os import path #modulo para obtener la ruta a directorio

import openpyxl #para generar el excel
#biblioteca o modulo send_file para forzar la descarga
from flask import send_file

#1 FUNCION: funcion que recibe del formulario nuevo empleado
def procesar_form_empleado(dataForm, foto_perfil):
    #formateando salario
    salario_sin_puntos= re.sub('[^0-9]+', '', dataForm['salario'])
    #convertir salario a INT
    salario_entero = int(salario_sin_puntos)
    
    result_foto_perfil = procesar_imagen_perfil(foto_perfil)
    try:
        with connectionBD() as conexion_MySQLdb:
            with conexion_MySQLdb.cursor(dictionary=True) as cursor:
                
                sql = "INSERT INTO empleados (nombre_empleado, apellidos_empleado, tipo_identidad, n_identidad, fecha_nacimiento, sexo, grupo_rh, email, telefono, profesion, salario, foto_perfil) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                
                #creando una tupla con los valores del INSERT
                valores = (dataForm['nombre_empleado'], dataForm['apellidos_empleado'], dataForm['tipo_identidad'],
                           dataForm['n_identidad'], dataForm['fecha_nacimiento'], dataForm['sexo'],
                           dataForm['grupo_rh'], dataForm['email'], dataForm['telefono'], dataForm['profesion'], salario_entero, result_foto_perfil)
                cursor.execute(sql, valores)
                
                conexion_MySQLdb.commit()
                resultado_insert = cursor.rowcount
                return resultado_insert
            
    except Exception as e:
        return f'se produjo un error en procesar_form_empleado: {str(e)}'
    
# 2 FUNCION : funcion que guarda la imagen de perfil
def procesar_imagen_perfil(foto):
    try:
        #Nombre original del archivo
        filename = secure_filename(foto.filename)
        extension = os.path.splitext(filename)[1]
        
        #creando un string de 50 caracteres
        nuevoNameFile = (uuid.uuid4().hex + uuid.uuid4().hex)[:100]
        nombreFile = nuevoNameFile + extension
        
        #construir la ruta completa de subida del archivo
        basepath = os.path.abspach(os.path.dirname(__file__))
        upload_dir = os.path.join(basepath, f'../static/img/foto_perfil/')
        
        #validar si exite la ruta y crearla si no existe
        if not os.path.exists(upload_dir):
            os.makedirs(upload_dir)
            #Dando permiso a la carpeta
            os.chmod(upload_dir, 0o755)
            
        #construir la ruta completa de subida del archivo
        upload_path = os.path.join(upload_dir, nombreFile)
        foto.save(upload_path)
        
        return nombreFile
    except Exception as e:
        print("Error al procesar archivo", e)
        return []
    
# 3 FUNCION : lista de Empleados
def sql_lista_empleadosBD():
    try:
        with connectionBD() as conexion_MySQLdb:
            with conexion_MySQLdb.cursor(dictionary=True) as cursor:
                querySQL = (f"""
                        SELECT
                            e.id_empleado,
                            e.nombre_empleado,
                            e.apellidos_empleado,
                            e.salario,
                            e.foto_perfil,
                            CASE
                                WHEN e.sexo = 1 THEN 'Masculino'
                                ELSE 'Femenino'
                            END AS SEXO
                        FROM empleados AS e
                        ORDER BY e.id_empleado DESC
                        """)
                cursor.execute(querySQL,)
                empleadosBD = cursor.fetchall()
        return empleadosBD
    except Exception as e:
        print(
            f"Error en la funcion sql_lista_empleadosBD: {e}")
        return None
    
# 4 FUNCION : Detalles del Empleado
def sql_detalles_empleadosBD(idEmpleado):
    try:
        with connectionBD() as conexion_MySQLdb:
            with conexion_MySQLdb.cursor(dictionary=True) as cursor:
                querySQL = ("""
                        SELECT
                            e.id_empleado,
                            e.nombre_empleado,
                            e.apellidos_empleado,
                            e.tipo_identidad,
                            e.n_identidad,
                            e.fecha_nacimiento,
                            CASE
                                WHEN e.sexo = 1 THEN 'Masculino'
                                ELSE 'Femenino'
                            END AS sexo,
                            e.grupo_rh,
                            e.email,
                            e.telefono,
                            e.profesion,
                            e.salario,
                            e.foto_perfil,
                            DATE_FORMAT(e.fecha_registro, '%Y-%m_-%d %h:%i %p') AS fecha_registro
                        FROM empleados AS e
                        WHERE id_empleado =%s
                        ORDER BY e.id_empleado DESC
                        """)
                cursor.execute(querySQL, (idEmpleado,))
                empleadosBD = cursor.fetchone()
        return empleadosBD
    except Exception as e:
        print(
            f"Error en la funcion sql_detalles_empleadosBD: {e}")
        return None
    
# 5 FUNCION : funcion empleados informe (Reporte)
def empleadosReporte():
    try:
        with connectionBD() as conexion_MySQLdb:
            with conexion_MySQLdb.cursor(dictionary=True) as cursor:
                querySQL = ("""
                        SELECT
                            e.id_empleado,
                            e.nombre_empleado,
                            e.apellidos_empleado,
                            e.tipo_identidad,
                            e.n_identidad,
                            e.fecha_nacimiento,
                            e.grupo_rh,
                            e.email,
                            e.telefono,
                            e.profesion,
                            e.salario,
                            DATE_FORMAT(e.fecha_registro, '%d de %b %Y %h:%i %p') AS fecha_registro,
                            CASE
                                WHEN e.sexo = 1 THEN 'Masculino'
                                ELSE 'Femenino'
                            END AS sexo,
                        FROM empleados AS e
                        ORDER BY e.id_empleado DESC
                        """)
                cursor.execute(querySQL,)
                empleadosBD = cursor.fetchall()
        return empleadosBD
    except Exception as e:
        print(
            f"Error en la funcion empleadosReporte: {e}")
        return None
    
# 6 FUNCION : funcion que exporta un excel con empleados 
def generarReporteExcel():
    dataEmpleados = empleadosReporte()
    wb = openpyxl.workbook()
    hoja = wb.active
    
    #Agregar los registros a la hoja
    for registro in dataEmpleados:
        nombre_empleado = registro['nombre_empleado']
        apellidos_empleado = registro['apellidos_empleado']
        tipo_identidad = registro['tipo_identidad']
        n_identidad = registro['n_identidad']
        fecha_nacimiento = registro['fecha_nacimiento']
        sexo = registro['sexo']
        grupo_rh = registro['grupo_rh']
        email = registro['email']
        telefono = registro['telefono']
        profesion = registro['profesion']
        salario = registro['salario']
        fecha_registro = registro['fecha_registro']
        
        #Agregar los valores a la hoja
        hoja.append((nombre_empleado, apellidos_empleado, tipo_identidad, n_identidad, fecha_nacimiento, sexo , grupo_rh, email, telefono, profesion, salario,fecha_registro))
        
        #Itera a traves de las filas y aplica el formato a la column G
        for fila_num
