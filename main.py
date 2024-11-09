from flask import Flask, render_template, request, redirect, session, url_for
import pymysql
from functools import wraps

app = Flask(__name__)
app.secret_key = 'tu_clave_secreta'

def conexion():
    return pymysql.connect(host="127.0.0.1",
                           user="root",
                           passwd="",
                           db="empresareparto"
                           )

def requiere_login(rol):
    def decorador(func):
        @wraps(func)
        def envoltura(*args, **kwargs):
            if 'usuario' not in session or session['rol'] != rol:
                return redirect("/login")
            return func(*args, **kwargs)
        return envoltura
    return decorador

@app.route("/")
def index():
    return render_template("index.html")  

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = conexion()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    SELECT id, username, password, tipo, empresa_id, trabajador_id 
                    FROM usuarios 
                    WHERE username = %s AND password = %s
                """, (username, password))
                usuario = cursor.fetchone()
                
                if usuario:
                    session['usuario'] = usuario[1] 
                    session['rol'] = usuario[3]     
                    if usuario[4]:               
                        session['empresa_id'] = usuario[4]
                    if usuario[5]:               
                        session['trabajador_id'] = usuario[5]
                    return redirect(url_for('index'))
                else:
                    return render_template("login.html", error="Usuario o contraseña incorrectos")
        except Exception as e:
            print(e)
            return render_template("login.html", error="Error al iniciar sesión")
        finally:
            conn.close()

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route("/empresa/registrar", methods=["GET", "POST"])
def empresa_registrar():
    if request.method == "POST":
        nombre = request.form["nombre"]
        direccion = request.form["direccion"]
        telefono = request.form["telefono"]
        email = request.form["email"]
        username = request.form["username"]
        password = request.form["password"]
        confirm_password = request.form["confirm_password"]
        if password != confirm_password:
            return "Las contraseñas no coinciden"
        
        conn = conexion()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT * FROM usuarios WHERE username = %s", (username,))
                if cursor.fetchone():
                    return "El nombre de usuario ya existe"
                cursor.execute("""
                    INSERT INTO empresas(nombre, direccion, telefono, email) 
                    VALUES(%s, %s, %s, %s)
                """, (nombre, direccion, telefono, email))
                
                empresa_id = cursor.lastrowid
                cursor.execute("""
                    INSERT INTO usuarios(username, password, tipo, empresa_id) 
                    VALUES(%s, %s, 'cliente', %s)
                """, (username, password, empresa_id))
                
            conn.commit()
            return redirect(url_for('login'))
        except Exception as e:
            print(e)
            return "Error al registrar la empresa"
        finally:
            conn.close()
    
    return render_template("registroEmpresa.html")

@app.route("/empresa/mostrar")
@requiere_login('administrador')
def empresa_mostrar():
    conn = conexion()
    empresas = []
    with conn.cursor() as cursor:
        cursor.execute("SELECT * FROM empresas")
        empresas = cursor.fetchall()
    conn.close()
    return render_template("mostrarEmpresas.html", empresas=empresas)

@app.route("/trabajador/registrar", methods=["GET", "POST"])
def trabajador_registrar():
    if request.method == "POST":
        nombre = request.form["nombre"]
        tipo = request.form["tipo"]
        username = request.form["username"]
        password = request.form["password"]
        confirm_password = request.form["confirm_password"]
        if password != confirm_password:
            return "Las contraseñas no coinciden"
        
        conn = conexion()
        try:
            with conn.cursor() as cursor:
                cursor.execute("SELECT * FROM usuarios WHERE username = %s", (username,))
                if cursor.fetchone():
                    return "El nombre de usuario ya existe"
                cursor.execute("""
                    INSERT INTO trabajadores(nombre, tipo, estado) 
                    VALUES(%s, %s, 'disponible')
                """, (nombre, tipo))
                
                trabajador_id = cursor.lastrowid
                cursor.execute("""
                    INSERT INTO usuarios(username, password, tipo, trabajador_id) 
                    VALUES(%s, %s, 'trabajador', %s)
                """, (username, password, trabajador_id))
                
            conn.commit()
            return redirect(url_for('login'))
        except Exception as e:
            print(e)
            return "Error al registrar el trabajador"
        finally:
            conn.close()
    
    return render_template("registroTrabajador.html")

@app.route("/trabajador/mostrar")
@requiere_login('administrador')
def trabajador_mostrar():
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("SELECT * FROM trabajadores")
        trabajadores = cursor.fetchall()
    conn.close()
    return render_template("mostrarTrabajadores.html", trabajadores=trabajadores)

@app.route("/solicitud/nueva", methods=["GET", "POST"])
@requiere_login('cliente')
def solicitud_nueva():
    if request.method == "POST":
        empresa_id = session['empresa_id']
        
        conn = conexion()
        with conn.cursor() as cursor:
            cursor.execute("INSERT INTO solicitudes(empresa_id, estado) VALUES(%s, 'pendiente')",
                         (empresa_id,))
            solicitud_id = cursor.lastrowid
            
            descripcion = request.form["descripcion"]
            cantidad = request.form["cantidad"]
            precio = request.form["precio"]
            direccion = request.form["direccion"]
            fecha_entrega = request.form["fecha_entrega"]
            
            cursor.execute("""
                INSERT INTO pedidos(solicitud_id, descripcion, cantidad, precio, 
                                  direccion_entrega, fecha_entrega, estado)
                VALUES(%s, %s, %s, %s, %s, %s, 'pendiente')
            """, (solicitud_id, descripcion, cantidad, precio, direccion, fecha_entrega))
            
        conn.commit()
        conn.close()
        return redirect(url_for('solicitud_mostrar'))
    
    return render_template("nuevaSolicitud.html")

@app.route("/solicitud/mostrar")
@requiere_login('cliente')
def solicitud_mostrar():
    empresa_id = session['empresa_id']
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT s.*, p.* 
            FROM solicitudes s 
            JOIN pedidos p ON s.id = p.solicitud_id 
            WHERE s.empresa_id = %s
        """, (empresa_id,))
        solicitudes = cursor.fetchall()
    conn.close()
    return render_template("mostrarSolicitudes.html", solicitudes=solicitudes)

@app.route("/pedidos/pendientes")
@requiere_login('administrador')
def pedidos_pendientes():
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT p.*, e.nombre as empresa_nombre 
            FROM pedidos p 
            JOIN solicitudes s ON p.solicitud_id = s.id
            JOIN empresas e ON s.empresa_id = e.id
            WHERE p.estado = 'pendiente'
        """)
        pedidos = cursor.fetchall()
        
        cursor.execute("SELECT * FROM trabajadores WHERE estado = 'disponible'")
        trabajadores = cursor.fetchall()
    conn.close()
    return render_template("pedidosPendientes.html", pedidos=pedidos, trabajadores=trabajadores)

@app.route("/pedido/asignar", methods=["POST"])
@requiere_login('administrador')
def pedido_asignar():
    pedido_id = request.form["pedido_id"]
    trabajador_id = request.form["trabajador_id"]
    
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            INSERT INTO asignaciones(pedido_id, trabajador_id, estado_entrega)
            VALUES(%s, %s, 'en camino')
        """, (pedido_id, trabajador_id))
        
        cursor.execute("UPDATE pedidos SET estado = 'en camino' WHERE id = %s", 
                      (pedido_id,))
        
        cursor.execute("UPDATE trabajadores SET estado = 'ocupado' WHERE id = %s", 
                      (trabajador_id,))
    conn.commit()
    conn.close()
    return redirect(url_for('pedidos_pendientes'))

# Ruta del dashboard
@app.route("/dashboard")
@requiere_login('administrador')
def dashboard():
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT estado, COUNT(*) as total 
            FROM pedidos 
            GROUP BY estado
        """)
        pedidos_por_estado = cursor.fetchall()
        
        cursor.execute("""
            SELECT tipo, COUNT(*) as total 
            FROM trabajadores 
            GROUP BY tipo
        """)
        trabajadores_por_tipo = cursor.fetchall()
        
        cursor.execute("""
            SELECT DATE(fecha_entrega) as fecha, COUNT(*) as total 
            FROM pedidos 
            WHERE estado = 'entregado' 
            AND fecha_entrega >= DATE_SUB(NOW(), INTERVAL 7 DAY)
            GROUP BY DATE(fecha_entrega)
        """)
        pedidos_por_dia = cursor.fetchall()
        
    conn.close()
    return render_template("dashboard.html", 
                         pedidos_por_estado=pedidos_por_estado,
                         trabajadores_por_tipo=trabajadores_por_tipo,
                         pedidos_por_dia=pedidos_por_dia)
# Agregar esta nueva ruta
@app.route("/registro/seleccionar")
def seleccionar_registro():
    return render_template("seleccionarRegistro.html")

@app.route("/admin/registrar", methods=["GET", "POST"])
def admin_registrar():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        
        conn = conexion()
        try:
            with conn.cursor() as cursor:
                cursor.execute("""
                    INSERT INTO usuarios(username, password, tipo) 
                    VALUES(%s, %s, 'administrador')
                """, (username, password))
                
            conn.commit()
            return redirect(url_for('login'))
        except Exception as e:
            print(e)
            return "Error al registrar el administrador"
        finally:
            conn.close()
    
    return render_template("registroAdmin.html")
@app.route("/pedidos/asignados")
@requiere_login('trabajador')
def pedidos_asignados():
    trabajador_id = session.get('trabajador_id')
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT p.*, e.nombre as empresa_nombre, a.estado_entrega
            FROM pedidos p 
            JOIN asignaciones a ON p.id = a.pedido_id
            JOIN solicitudes s ON p.solicitud_id = s.id
            JOIN empresas e ON s.empresa_id = e.id
            WHERE a.trabajador_id = %s AND a.estado_entrega = 'en camino'
        """, (trabajador_id,))
        pedidos = cursor.fetchall()
    conn.close()
    return render_template("pedidosAsignados.html", pedidos=pedidos)
@app.route("/pedido/actualizar/<int:pedido_id>", methods=["POST"])
@requiere_login('trabajador')
def pedido_actualizar(pedido_id):
    nuevo_estado = request.form["estado"]
    trabajador_id = session.get('trabajador_id')
    
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            UPDATE asignaciones 
            SET estado_entrega = %s 
            WHERE pedido_id = %s AND trabajador_id = %s
        """, (nuevo_estado, pedido_id, trabajador_id))
        
        if nuevo_estado == 'entregado':
            cursor.execute("""
                UPDATE pedidos SET estado = 'entregado' 
                WHERE id = %s
            """, (pedido_id,))
            
            cursor.execute("""
                UPDATE trabajadores SET estado = 'disponible' 
                WHERE id = %s
            """, (trabajador_id,))
            
    conn.commit()
    conn.close()
    return redirect(url_for('pedidos_asignados'))
@app.route("/historial/entregas")
@requiere_login('trabajador')
def historial_entregas():
    trabajador_id = session.get('trabajador_id')
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT 
                p.id as pedido_id,
                p.descripcion,
                p.cantidad,
                p.precio,
                p.direccion_entrega,
                p.fecha_entrega,
                p.estado as pedido_estado,
                e.nombre as empresa_nombre,
                a.estado_entrega,
                a.fecha_asignacion
            FROM pedidos p 
            JOIN asignaciones a ON p.id = a.pedido_id
            JOIN solicitudes s ON p.solicitud_id = s.id
            JOIN empresas e ON s.empresa_id = e.id
            WHERE a.trabajador_id = %s 
            ORDER BY a.fecha_asignacion DESC
        """, (trabajador_id,))
        entregas = cursor.fetchall()
    conn.close()
    return render_template("historialEntregas.html", entregas=entregas)
@app.route("/admin/historial")
@requiere_login('administrador')
def historial_admin():
    conn = conexion()
    with conn.cursor() as cursor:
        cursor.execute("""
            SELECT 
                p.id as pedido_id,
                e.nombre as empresa_nombre,
                p.descripcion,
                p.cantidad,
                p.precio,
                p.direccion_entrega,
                p.fecha_entrega,
                p.estado as pedido_estado,
                s.estado as solicitud_estado,
                s.fecha_solicitud,
                COALESCE(t.nombre, 'Sin asignar') as trabajador_nombre,
                COALESCE(a.estado_entrega, 'pendiente') as estado_entrega,
                COALESCE(a.fecha_asignacion, NULL) as fecha_asignacion,
                t.id as trabajador_id,
                COALESCE(t.tipo, 'sin asignar') as tipo_trabajo  # Modificado aquí
            FROM pedidos p 
            JOIN solicitudes s ON p.solicitud_id = s.id
            JOIN empresas e ON s.empresa_id = e.id
            LEFT JOIN asignaciones a ON p.id = a.pedido_id
            LEFT JOIN trabajadores t ON a.trabajador_id = t.id
            ORDER BY s.fecha_solicitud DESC
        """)
        pedidos = cursor.fetchall()
    conn.close()
    return render_template("historialAdmin.html", pedidos=pedidos)
if __name__ == "__main__":
    app.run(debug=True)