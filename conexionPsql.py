import psycopg2     #---> se importa la libreria

conexion = psycopg2.connect(host = "localhost",
                                database = "zapateria",
                                user = "odoo",
                                password = "odoo")  #----> se crea la conexion

cur = conexion.cursor()   # ----> se crea un cursor

cur.execute('SELECT * FROM shoes')  #-----> query hacia una tabla

rows = cur.fetchall()  #----> devuelve todas las filas de la consulta

for row in rows:
    print(row)     #----> se itera para mostrar los registros

cur.close()    # ----> se cierra el cursor
conexion.close()   # ---> se cierra la conexion