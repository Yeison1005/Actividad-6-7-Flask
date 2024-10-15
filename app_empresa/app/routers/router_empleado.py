from app import app
from flask import render_template, request, flash, redirect, url_for, session, jsonify
from mysql.connector.errors import Error

#   Importando conexión a BD
from controllers.funciones_empleado import *


PATH_URL = "empleados" # carpeta templates/empleados

@app.route('/registrar-empleado', methods=['GET'])
def viewFormEmpleado():
    if 'conectado' in session:
        return render_template(f'{PATH_URL}/form_empleado.html')
    else:
        flash('Primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))

@app.route('/form-registrar-empleado', methods=['POST'])
def formEmpleado():
    if 'conectado' in session:
        if 'foto_perfil' in request.files:
            foto_perfil = request.files['foto_perfil']
            resultado = procesar_form_empleado(request.form, foto_perfil)
            if resultado:
                return redirect(url_for('lista_empleados'))
            else:
                flash('El empleado NO fue registrado.', 'error')
                return render_template(f'{PATH_URL}/form_empleado.html')
        else:
            flash('Primero debes iniciar sesión.', 'error')
            return redirect(url_for('inicio'))

@app.route('/lista-de-empleados', methods=['GET'])
def lista_empleados():
    if 'conectado' in session:
        return render_template(f'{PATH_URL}/lista_empleados.html', empleados=sql_lista_empleadosBD())
    else:
        flash('Primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))

@app.route('/detalles-empleado/', methods=['GET'])
@app.route('/detalles-empleado/<int:idEmpleado>', methods=['GET'])
def detalleEmpleado(idEmpleado=None):
    if 'conectado' in session:
        # Verificamos si el parámetro idEmpleado es None o no está presente en la URL
        if idEmpleado is None:
            return redirect(url_for('inicio'))
        else:
            detalle_empleado = sql_detalles_empleadosBD(idEmpleado) or []
            return render_template(f'{PATH_URL}/detalles_empleado.html', detalle_empleado=detalle_empleado)
    else:
        flash('Primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))

# Buscando en empleados
@app.route('/buscando-empleado', methods=['POST'])
def viewBuscarEmpleadoBD():
    if 'conectado' in session:
        resultadoBusqueda = buscarEmpleadoUnico(request.json['busqueda'])
        if resultadoBusqueda:
            return render_template(f'{PATH_URL}/resultado_busqueda_empleado.html', dataBusqueda=resultadoBusqueda)
        else:
            return jsonify({'fin': 0})

@app.route('/editar-empleado/<int:id>', methods=['GET'])
def viewEditarEmpleado(id):
    if 'conectado' in session:
        respuestaEmpleado = buscarEmpleadoUnico(id)
        if respuestaEmpleado:
            return render_template(f'{PATH_URL}/form_empleado_update.html', respuestaEmpleado=respuestaEmpleado)
        else:
            flash('El empleado no existe.', 'error')
            return redirect(url_for('inicio'))
    else:
        flash('Primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))

# Recibir formulario para actualizar información de empleado
@app.route('/actualizar-empleado', methods=['POST'])
def actualizarEmpleado():
    if 'conectado' in session:
        resultData = procesar_Actualizacion_form(request)
        if resultData:
            return redirect(url_for('lista_empleados'))

@app.route('/lista-de-usuarios', methods=['GET'])
def usuarios():
    if 'conectado' in session:
        resp_usuariosBD = lista_usuariosBD()
        return render_template('usuarios/lista_usuarios.html', resp_usuariosBD=resp_usuariosBD)
    else:
        return redirect(url_for('inicioCPanel'))

@app.route('/borrar-usuario/<string:id>', methods=['GET'])
def borrarUsuario(id):
    resp = eliminarUsuario(id)
    if resp:
        flash('El Usuario fue eliminado correctamente.', 'success')
        return redirect(url_for('usuarios'))

@app.route('/borrar-empleado/<string:id_empleado>/<string:foto_perfil>', methods=['GET'])
def borrarEmpleado(id_empleado, foto_perfil):
    resp = eliminarEmpleado(id_empleado, foto_perfil)
    if resp:
        flash('El Empleado fue eliminado correctamente.', 'success')
        return redirect(url_for('lista_empleados'))

@app.route('/descargar-informe-empleados/', methods=['GET'])
def reporteBD():
    if 'conectado' in session:
        return generarReportesExcel()
    else:
        flash('Primero debes iniciar sesión.', 'error')
        return redirect(url_for('inicio'))