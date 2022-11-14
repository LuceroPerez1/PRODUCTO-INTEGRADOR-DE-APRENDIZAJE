import pandas as pd
import re
import sys
import sqlite3
from sqlite3 import Error
import datetime
import time
import os
import openpyxl

BLACK = '\033[30m'
RED = '\033[31m'
GREEN = '\033[32m'
YELLOW = '\033[33m'
BLUE = '\033[34m'
RESET = '\033[35m'

print(os.getcwd())

if not os.path.exists("PIA1.db"):
    print(RED + "No existen tablas")
    try:
        with sqlite3.connect("PIA1.db") as conn:
            mi_cursor = conn.cursor()
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS clientes(id INTEGER PRIMARY KEY, nombre_cliente TEXT NOT NULL);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS salas(clave INTEGER PRIMARY KEY, nombre_sala TEXT NOT NULL, cupo INTEGER);")
            mi_cursor.execute("CREATE TABLE IF NOT EXISTS reservaciones(folio INTEGER PRIMARY KEY, numero_cliente INTEGER, sala INTEGER, fecha_reservacion timestamp, turno TEXT NOT NULL, nombre_del_evento TEXT NOT NULL, FOREIGN KEY (numero_cliente) REFERENCES clientes(id), FOREIGN KEY (sala) REFERENCES salas(clave));")
            print(RED + "Tablas creadas" + BLACK)
    except Error as e:
        print(e)
else:
    while True:
        print(BLACK + """
        |--------------------------|
        |     Menú de opciones     |
        |--------------------------|
        |(1) Reservaciones         |
        |(2) Reportes              |
        |(3) Registrar una sala    |
        |(4) Registrar un cliente  |
        |(5) Salir                 |
        |--------------------------|\n""")
        opcion_menu_principal = input(BLACK + "Elige una opción del menú principal:  ")
        if opcion_menu_principal == "":
            print(RED + "No puede dejar el campo en blanco, favor de seleccionar una opción del menú. \n")
            continue
        elif opcion_menu_principal.isspace():
            print(RED + "No puede dejar el campo en blanco, favor de seleccionar una opción del menú. \n")
            continue
        elif not opcion_menu_principal in "12345":
            print(RED + "Debe ingresar un número del 1 al 5. \n")
            continue
        elif opcion_menu_principal == "1":
            while True:    
                print(BLACK + """
                |----------------Menú de reservaciones---------------|
                |(1) Registrar nueva reservación                     |
                |(2) Modificar descripción de una reservación        |
                |(3) Consultar disponibilidad de salas para una fecha|
                |(4) Eliminar una reservación                        |
                |(5) Salir                                           |
                |----------------------------------------------------|""")
                opcion_menu_reservaciones = input(BLACK + "Elige una opción del menú de reservaciones:  ")
                if opcion_menu_reservaciones == "":
                    print(RED + "No puede dejar el campo en blanco, favor de seleccionar una opción del menú. \n")
                    continue
                elif opcion_menu_reservaciones.isspace():
                    print(RED + "No puede dejar el campo en blanco, favor de seleccionar una opción del menú. \n")
                    continue
                elif not opcion_menu_reservaciones in ("12345"):
                    print(RED + "Debe ingresar un número del 1 al 5. \n")
                    continue
                elif opcion_menu_reservaciones == "1":
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT * FROM clientes;")
                            datos_clientes = mi_cursor.fetchall()
                            print(YELLOW + "---------------------------------------")
                            print(RED + "               CLIENTES                ")
                            print(YELLOW + "---------------------------------------")
                            print(BLACK + "ID CLIENTE       NOMBRE DE CLIENTE")
                            for clave, nombre_cliente in datos_clientes:
                                print(clave,BLACK + "\t\t",nombre_cliente, BLACK + "")
                            print(YELLOW + "---------------------------------------")
                    except Error as e:
                            print(e)
                            break
                    while True:
                        try:
                            numero_cliente = int(input(BLACK + "\nIngrese su ID Usuario:  "))
                        except:
                            print(RED + "Solamente se permiten números, favor de intentarlo de nuevo. \n")
                            continue
                        try:
                            with sqlite3.connect("PIA1.db") as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("SELECT * FROM clientes WHERE id=(?)", [numero_cliente])
                                folios = mi_cursor.fetchall()
                        except Error as e:
                            print(e)
                            break
                        if folios:
                            try:
                                with sqlite3.connect("PIA1.db") as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute("SELECT * FROM salas")
                                    salas = mi_cursor.fetchall()
                                    print(YELLOW + "---------------------------------------")
                                    print(RED + "                SALAS                  ")
                                    print(YELLOW + "---------------------------------------")
                                    print(BLACK + "ID SALA       NOMBRE DE SALA     CUPO")
                                    for sala, nombre_sala, cupo in salas:
                                        print(sala, BLACK + "\t\t",nombre_sala, BLACK + "\t",cupo, BLACK + "")
                                    print(YELLOW + "---------------------------------------")
                            except Error as e:
                                print(e)
                                break
                        else:
                            print(RED + "El número que ingresó no coincide con ningún ID cliente, favor de intentarlo de nuevo.")
                            continue
                        while True:
                            try:
                                sala = int(input(BLACK + "Seleccione un ID sala:  "))
                            except:
                                print(RED + "Solamente se permiten números, favor de intenarlo de nuevo. \n")
                                continue
                            try:
                                with sqlite3.connect("PIA1.db") as conn:
                                    mi_cursor = conn.cursor()
                                    mi_cursor.execute("SELECT * FROM salas WHERE clave=(?)", [sala])
                                    sala = mi_cursor.fetchall()
                            except Error as e:
                                print(e)
                                break
                            if not sala:
                                print(RED + "El número que ingresó no coincide con ningún ID sala, favor de intentarlo de nuevo. \n")
                                continue
                            else:
                                try:
                                    fecha_reservacion = input("Ingrese la fecha deseada (dd/mm/aaa): \n")
                                    fecha_reservacion = datetime.datetime.strptime(fecha_reservacion,"%d/%m/%Y").date()
                                    fecha_actual = (datetime.date.today())
                                    limite_fecha = (fecha_reservacion - fecha_actual).days
                                    if limite_fecha <=1:
                                        print(RED + "Debe de reservar con más de 2 días de anticipación. \n")
                                        continue
                                except:
                                    print(RED + "Debe ingresar una fecha en el formato señalado. \n")
                                    continue
                            while True:
                                print(BLACK + """
                                |-----Turnos-----|
                                |(M) Matutino    |
                                |(V) Vespertino  |
                                |(N) Nocturno    |
                                |----------------|""")
                                turno = input(BLACK + "Ingrese un turno:  ").upper()
                                if turno == "":
                                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un turno.\n")
                                    continue
                                elif turno.isspace():
                                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un turno. \n")
                                    continue
                                elif turno not in "MVN":
                                    print(RED + "Debe de ingresar una de las opciones del menú de turnos (M, V o N). \n")
                                    continue
                                elif turno == "M":
                                    turno = "Matutino"
                                elif turno == "V":
                                    turno = "Vespertino"
                                elif turno == "N":
                                    turno = "Nocturno"
                                while True:
                                    nombre_del_evento = input(BLACK + "Ingrese el nombre del evento:  ").title()
                                    if nombre_del_evento == "":
                                        print(RED + "Debe ingresar un nombre. \n")
                                        continue
                                    try:
                                        with sqlite3.connect("PIA1.db") as conn:
                                            mi_cursor = conn.cursor()
                                            valores = (numero_cliente, sala[0][0], fecha_reservacion, turno,  nombre_del_evento)
                                            mi_cursor.execute("INSERT INTO reservaciones(numero_cliente, sala, fecha_reservacion, turno,  nombre_del_evento) VALUES(?,?,?,?,?)", [numero_cliente, sala[0][0], fecha_reservacion, turno, nombre_del_evento])
                                            print(BLUE + "Registro agregado exitosamente")
                                            mi_cursor.execute("SELECT MAX(folio) FROM reservaciones;")
                                            id_reservacion = mi_cursor.fetchall()
                                            print(YELLOW + "------------------------------------------------------------------------------------------------------")
                                            print(BLACK + "ID Reserva        Fecha reservación       Turno       Nombre sala     Nombre cliente      Nombre del evento")
                                            print("------------------------------------------------------------------------------------------------------")
                                            print(id_reservacion[0][0], BLACK + "               ",fecha_reservacion, BLACK + "        ",turno, BLACK + "       ",sala[0][1], BLACK + "          ",folios[0][1], BLACK + "         ",nombre_del_evento)
                                            print(YELLOW + "------------------------------------------------------------------------------------------------------")
                                            break
                                    except Error as e:
                                        print(e)
                                        break
                                break
                            break
                        break
                    break
                elif opcion_menu_reservaciones == "2":
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave")
                            intento = mi_cursor.fetchall()
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            print(RED + "                             RESERVACIONES                                               ")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            print(BLACK + "ID Reserva | Fecha reservación | Turno | Nombre Cliente | Nombre Sala | Nombre del evento")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            for folio, fecha_reservacion, turno, nombre_cliente, nombre_sala, nombre_del_evento in intento:
                                print(folio, BLACK +"      ",fecha_reservacion, BLACK +"         ",turno, BLACK +"      ", nombre_cliente, BLACK + "     ", nombre_sala, BLACK +"    ", nombre_del_evento, BLACK +"")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                    except Error as e:
                        print(e)
                        break
                    while True:
                        try:
                            numero_reserva = int(input(BLACK + "Ingresa el ID Reserva:  "))
                        except:
                            print(RED + "Solamente se permiten números, favor de intentarlo de nuevo. \n")
                            continue
                        break
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT folio FROM reservaciones WHERE folio=(?)", [numero_reserva])
                            id = mi_cursor.fetchall()
                    except Error as e:
                        print(e)
                        break
                    if id:
                        try:
                            with sqlite3.connect("PIA1.db") as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave WHERE folio=(?)", [numero_reserva])
                                identificador = mi_cursor.fetchall()
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                print(RED + "                             RESERVACIONES                                               ")
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                print(BLACK + "ID Reserva | Fecha reservación | Turno | Nombre Cliente | Nombre Sala | Nombre del evento")
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                print(identificador[0][0], BLACK +"          ", identificador[0][1], BLACK +"      ", identificador[0][2], BLACK + "     ", identificador[0][3], BLACK + "       ", identificador[0][4], BLACK +"      ", identificador[0][5])
                        except Error as e:
                            print(e)
                            break
                        while True:
                                nombre_nuevo = input(BLACK + "Ingrese el nuevo nombre del evento:  ").title()
                                if nombre_nuevo == "":
                                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un nombre. \n")
                                    continue
                                nombres =  nombre_nuevo, numero_reserva
                                try:
                                    with sqlite3.connect("PIA1.db") as conn:
                                        mi_cursor = conn.cursor()
                                        mi_cursor.execute("UPDATE reservaciones SET nombre_del_evento=(?) WHERE folio=(?);", (nombre_nuevo, numero_reserva))
                                        edicion = mi_cursor.fetchall()
                                        mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave WHERE folio=(?)", [numero_reserva])
                                        reservacion = mi_cursor.fetchall()
                                        print(BLUE + "Se ha editado el nombre del evento")
                                        print(YELLOW + "------------------------------------------------------------------------------------------------------")
                                        print(BLACK + "ID Reserva     Fecha reservación     Turno    Nombre cliente   Nombre sala     Nombre del evento")
                                        print(YELLOW + "------------------------------------------------------------------------------------------------------")
                                        print(reservacion[0][0], BLACK +"              ",reservacion[0][1], BLACK +"     ",reservacion[0][2], BLACK +"    ",reservacion[0][3], BLACK +"          ",reservacion[0][4], BLACK +"       ",reservacion[0][5])
                                        print(YELLOW + "------------------------------------------------------------------------------------------------------")
                                        break
                                except Exception as E:
                                    print(E)
                                    continue
                    else:
                        print(RED + "No se encontró ningun ID reserva con ese número. \n")
                        continue

                elif opcion_menu_reservaciones == "3":
                    while True:
                        try:
                            fecha_disponible = input(BLACK + "Ingrese la fecha que desea consultar (dd/mm/aaa): \n")
                            fecha_disponible = datetime.datetime.strptime(fecha_disponible,"%d/%m/%Y").date()
                            break
                        except:
                            print(RED + "Debe ingresar una fecha en el formato señalado. \n")
                            continue 
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT clave FROM salas")
                            claves_salas = mi_cursor.fetchall()
                            mi_cursor.execute("SELECT sala, turno FROM reservaciones WHERE fecha_reservacion=(?)", [fecha_disponible])
                            ocupados = mi_cursor.fetchall()
                            posibles = []
                            turnos_existentes = ["Matutino", "Vespertino", "Nocturno"]
                            if claves_salas:
                                for sala in claves_salas:
                                    for turno in turnos_existentes:
                                        posibles.append(((sala[0]),turno))
                                posibles = set(posibles)
                                ocupados = set(ocupados)
                                turnos_disponibles = posibles - ocupados
                                turnos_disponibles = set(turnos_disponibles)
                                print(YELLOW + "----------------------")
                                print(RED + "   SALAS DISPONIBLES  ")
                                print(YELLOW + "----------------------")
                                print(BLACK + "ID SALA     TURNOS    ")
                                print(YELLOW + "----------------------")
                                for clave, turno in turnos_disponibles:   
                                    print(clave,BLACK + "        ", turno, BLACK + "     ")
                                print(YELLOW + "----------------------")
                            else:
                                print(RED + "No existe ninguna sala registrada con esa fecha. \n")
                    except Error as e:
                            print(e)
                            break
                elif opcion_menu_reservaciones == "4":
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave")
                            borrados = mi_cursor.fetchall()
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            print(RED + "                             RESERVACIONES                                               ")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            print(BLACK + "ID Reserva | Fecha reservación | Turno | Nombre Cliente | Nombre Sala | Nombre del evento")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                            for folio, fecha_reservacion, turno, nombre_cliente, nombre_sala, nombre_del_evento in borrados:
                                print(folio,BLACK + "           ",fecha_reservacion, BLACK + "     ", BLACK + turno,"     ", nombre_cliente, BLACK + "   ", nombre_sala, BLACK + "        ", nombre_del_evento, BLACK + "")
                            print(YELLOW + "-----------------------------------------------------------------------------------------")
                    except Error as e:
                        print(e)
                        break
                    while True:
                        try:
                            numero_reserva = int(input(BLACK + "Ingresa el ID Reserva:  "))
                        except:
                            print(RED + "Solamente se permiten números, favor de intentarlo de nuevo. \n")
                            continue
                        try:
                            with sqlite3.connect("PIA1.db") as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("SELECT * FROM reservaciones WHERE folio=(?)", [numero_reserva])
                                numero_existe = mi_cursor.fetchall()
                        except Error as e:
                            print(e)
                            break
                        if numero_existe:
                            while True:
                                try:
                                    fecha_eliminada = input(BLACK + "Ingrese la fecha que desea eliminar (dd/mm/aaa): \n")
                                    fecha_eliminada = datetime.datetime.strptime(fecha_eliminada,"%d/%m/%Y").date()
                                    break
                                except:
                                    print(RED + "Debe ingresar una fecha en el formato señalado. \n")
                                    continue 
                            fecha_actual = (datetime.date.today())
                            limite_fecha = (fecha_eliminada - fecha_actual).days
                            if limite_fecha <=2:
                                print(RED + "Debe de eliminar la reservación, al menos con 3 días de anticipación. \n")
                                break
                            else:
                                try:
                                    with sqlite3.connect("PIA1.db") as conn:
                                        mi_cursor = conn.cursor()
                                        mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave")
                                        reserva = mi_cursor.fetchall()
                                        print(YELLOW + "-----------------------------------------------------------------------------------------")
                                        print(RED + "RESERVACION DEL ID: ", (numero_reserva))
                                        print(YELLOW + "-----------------------------------------------------------------------------------------")
                                        print(BLACK + "ID RESERVA | FECHA RESERVACION | TURNO | NOMBRE CLIENTE | NOMBRE SALA | NOMBRE DEL EVENTO")
                                        print(YELLOW + "-----------------------------------------------------------------------------------------")
                                        print(reserva[0][0],BLACK + "                  ",reserva[0][1], BLACK + "  ",reserva[0][2], BLACK + "  ",reserva[0][3], BLACK + "     ",reserva[0][4], BLACK + "         ",reserva[0][5], BLACK + "")
                                        print(YELLOW + "-----------------------------------------------------------------------------------------")
                                except Error as e:
                                    print(e)
                                    break
                                print(RED + """\nEliminar reservacion:
                                        (1) SI
                                        (2) NO \n""")
                                confirmacion = input(BLACK + "Esta seguro que desea eliminar un elemento?  ")
                                if confirmacion == "":
                                    print(RED + "No puede dejar el campo en blanco, favor de seleccionar una opción. \n")
                                    continue
                                if not confirmacion in ("12"):
                                    print(RED + "Debe de ingresar una de las opciones (1 ó 2). \n")
                                    continue
                                if confirmacion == "1":
                                    try:
                                        with sqlite3.connect("PIA1.db") as conn:
                                            mi_cursor = conn.cursor()
                                            mi_cursor.execute("DELETE FROM reservaciones WHERE folio=(?)", [numero_reserva])
                                            print(BLUE + "Se ha eliminado la reservación. \n")
                                    except Error as e:
                                        print(e)
                                        break
                            if confirmacion == "2":
                                break
                            break
                        else:
                            print(RED + "No se encontró ningun folio en la base de datos con ese número. \n")
                            break
                elif opcion_menu_reservaciones == "5":
                    break
        elif opcion_menu_principal == "2":
            while True:
                print(BLACK + """
                |----------------Menú de reportes-----------------------|
                |(1) Reporte en pantalla de reservaciones para una fecha|
                |(2) Exportar reporte tabular en Excel                  |
                |(3) Salir                                              |
                |-------------------------------------------------------|""")
                opcion_menu_reportes = input(BLACK + "Elige una opción del menú:  ")
                if opcion_menu_reportes == "":
                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un turno. \n")
                    continue
                elif opcion_menu_reportes.isspace():
                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un turno. \n")
                    continue
                elif not opcion_menu_reportes in ("123"):
                    print(RED + "Debe de ingresar una de las opciones del menú de reportes (1, 2 ó 3). \n")
                    continue
                elif opcion_menu_reportes == "1":
                    while True:
                        try:
                            fecha_consulta = input(BLACK + "Ingrese la fecha que desee consultar (dd/mm/aaa): \n")
                            fecha_consulta = datetime.datetime.strptime(fecha_consulta,"%d/%m/%Y").date()
                            break
                        except:
                            print(RED + "Debe ingresar una fecha en el formato señalado. \n")
                            continue
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT fecha_reservacion FROM reservaciones WHERE fecha_reservacion = (?)", [fecha_consulta])
                            fecha_existe = mi_cursor.fetchall()
                    except Error as e:
                        print(e)
                        break
                    if fecha_existe:
                        try:
                            with sqlite3.connect("PIA1.db") as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave WHERE fecha_reservacion=(?)", [fecha_consulta])
                                intento = mi_cursor.fetchall()
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                print(RED + "                             RESERVACIONES                                               ")
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                print(BLACK + "ID Reserva | Fecha reservación | Turno | Nombre Cliente | Nombre Sala | Nombre del evento")
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                                for folio, fecha_reservacion, turno, nombre_cliente, nombre_sala, nombre_del_evento in intento:
                                    print(folio, BLACK + "           ",fecha_reservacion, BLACK +"       ",turno, BLACK + "    ", nombre_cliente, BLACK + "     ", nombre_sala, BLACK +"      ", nombre_del_evento, BLACK +"")
                                print(YELLOW + "-----------------------------------------------------------------------------------------")
                        except Error as e:
                            print(e)
                            break
                    else:
                        print(RED + "No se encontró ninguna reserva con la fecha que ingresó. ")
                        continue

                elif opcion_menu_reportes == "2":
                    while True:
                        try:
                            fecha_consulta = input(BLACK + "Ingrese la fecha que desee consultar (dd/mm/aaa): \n")
                            fecha_consulta = datetime.datetime.strptime(fecha_consulta,"%d/%m/%Y").date()
                            break
                        except:
                            print(RED + "Debe ingresar una fecha en el formato señalado. \n")
                            continue
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            mi_cursor.execute("SELECT fecha_reservacion FROM reservaciones WHERE fecha_reservacion = (?)", [fecha_consulta])
                            fecha_existe = mi_cursor.fetchall()
                    except Error as e:
                        print(e)
                        break
                    if fecha_existe:
                        try:
                            with sqlite3.connect("PIA1.db") as conn:
                                mi_cursor = conn.cursor()
                                mi_cursor.execute("SELECT reservaciones.folio, reservaciones.fecha_reservacion, reservaciones.turno , clientes.nombre_cliente, salas.nombre_sala, reservaciones.nombre_del_evento FROM reservaciones INNER JOIN clientes ON reservaciones.numero_cliente = clientes.id INNER JOIN salas ON reservaciones.sala = salas.clave WHERE fecha_reservacion=(?)", [fecha_consulta])
                                datos_excel = mi_cursor.fetchall()
                                libro_excel = openpyxl.Workbook()  
                                hoja_excel = libro_excel.create_sheet("Hoja")
                                hoja = libro_excel.active
                                hoja.append(('ID Reserva', 'Fecha reservación', 'Turno', 'Nombre Cliente', 'Nombre Sala', 'Nombre del evento'))
                                for datos_reserva in datos_excel:
                                    hoja.append(datos_reserva)
                                    libro_excel.save("PIA1.xlsx")
                                print(BLUE + "Libro exitosamente creado")
                        except Error as e:
                            print(e)
                            break
                    else:
                        print(RED + "No se encontró ninguna reserva con la fecha que ingresó. ")
                        continue
                elif opcion_menu_reportes == "3":
                    break
        elif opcion_menu_principal == "3":
            while True:
                nombre_sala = input(BLACK + "Ingrese el nombre de la sala:  ").title()
                if nombre_sala == "":
                    print(RED + "No puede dejar el campo en blanco, favor de ingresar un nombre de la sala")
                    continue
                if nombre_sala.isspace():
                    print(RED + "Debe de ingresar el nombre de la sala, favor de intentarlo de nuevo. \n")
                    continue
                while True:
                    try:
                        cupo = int(input(BLACK + "Ingrese el cupo de la sala:  "))
                    except:
                        print(RED + "Favor de ingresar una cantidad. \n")
                        continue
                    if cupo <=1:
                        print(RED + "El cupo tiene que ser mayor a cero. \n")
                        continue
                    try:
                        with sqlite3.connect("PIA1.db") as conn:
                            mi_cursor = conn.cursor()
                            valores = (nombre_sala, cupo)
                            mi_cursor.execute("INSERT INTO salas(nombre_sala, cupo) VALUES(?,?)", [nombre_sala, cupo])
                            print(BLACK + "Registro agregado exitosamente")
                            mi_cursor.execute("SELECT MAX(clave) FROM salas;")
                            id_sala = mi_cursor.fetchall()
                            print(YELLOW + "-------------------------")
                            print(BLUE + f"Tu ID Sala es: {id_sala[0][0]}")
                            print(YELLOW + "-------------------------")
                            break
                    except Error as e:
                        print(e)
                        break
                break
        elif opcion_menu_principal == "4":
            while True:
                nombre_cliente = input(BLACK + "Ingrese su nombre:  ").title()
                if nombre_cliente == "":
                    print(RED + "No puede dejar el espacio en blanco, favor de ingresar su nombre")
                    continue
                elif nombre_cliente.isspace():
                    print(RED + "Debe de ingresar su nombre, favor de intentarlo de nuevo. \n")
                    continue
                elif (not re.match("^[a-zA-Z_ ]*$", nombre_cliente)):
                    print(RED + "Solamente se permiten letras, favor de intentarlo de nuevo. \n")
                    continue
                try:
                    with sqlite3.connect("PIA1.db") as conn:
                        mi_cursor = conn.cursor()
                        valores = (nombre_cliente)
                        mi_cursor.execute("INSERT INTO clientes(nombre_cliente) VALUES(?)", [nombre_cliente])
                        print(BLACK + "Registro agregado exitosamente")
                        mi_cursor.execute("SELECT MAX(id) FROM clientes;")
                        id_client3 = mi_cursor.fetchall()
                        print(YELLOW + "-------------------------")
                        print(BLUE + f"Tu ID Usuario es: {id_client3[0][0]}")
                        print(YELLOW + "-------------------------")
                except Error as e:
                    print(e)
                    break
                break
        elif opcion_menu_principal == "5":
            break