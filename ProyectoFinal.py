# pandas
# xlrd
# openpyxl

import pandas as pd  # Importación de una libreria
from openpyxl import load_workbook  # Importación de una libreria


def EscribirArchivo(prestamos):  # Guarda en un archivo excel los prestamos
    # definimos el nombre del archivo excel en el que vamos a guardar datos
    writer = pd.ExcelWriter('prestamos.xlsx', engine='openpyxl')
    wb = writer.book
    df = pd.DataFrame(prestamos)  # escribimos en el archivo la lista
    df.to_excel(writer, index=False)
    wb.save('prestamos.xlsx')  # Guarda el archivo


def Lineas():  # Imprime asterizcos
    print("***********************************")


def CargarPersonas():  # Funcion que lee el archivo de de personas en formato excel y retorna los datos leidos
    # Definimos el nombre del archivo excel a leer y lo asignamos en una variable la cual vamos retornar en la funcion
    rd = pd.read_excel('personasv2.xlsx', engine='openpyxl')
    return rd  # Retornamos los datos que se encontraron en el excel


def CargarPrestamos():  # Funcion que lee el archivo de prestamos en formato excel y retorna los datos leidos
    # Definimos el nombre del archivo excel a leer y lo asignamos en una variable la cual vamos retornar en la funcion
    rd = pd.read_excel('prestamos.xlsx', engine='openpyxl')
    return rd  # Retornamos los datos que se encontraron en el excel


def CargarLibros():  # Funcion que lee el archivo de libros en formato excel y retorna los datos leidos
    # Definimos el nombre del archivo excel a leer y lo asignamos en una variable la cual vamos retornar en la funcion
    rd = pd.read_excel('librosv2.xlsx', engine='openpyxl')
    return rd  # Retornamos los datos que se encontraron en el excel


def Menu():  # Funcion que muestra las opciones de menú
    Lineas()  # Llamada a la funcion Lineas
    print("a - Ver lista de personas")  # Impresión en consola de un mensaje
    # Impresión en consola de un mensaje
    print("b - Ordenar lista de personas")
    # Impresión en consola de un mensaje
    print("c - Imprimir registro de lista de persona")
    print("d - Ver lista de libros")  # Impresión en consola de un mensaje
    print("e - Buscar libro")  # Impresión en consola de un mensaje
    print("f - Prestar libro")  # Impresión en consola de un mensaje
    print("g - Ver prestamo de libros")  # Impresión en consola de un mensaje
    # Impresión en consola de un mensaje
    print("h - Opción para salir del programa ")


personas = [[]]  # Inicializamos lista en blanco
libros = [[]]  # Inicializamos lista en blanco
prestamos = [[]]  # Inicializamos la lista en blanco

codigoLibro = ""  # Inicializamos la variable string
codigoPersona = ""  # Inicializamos la variable string

encontroLibro = False  # Inicializamos la variable booleana
encontroPersona = False  # Inicializamos la variable booleana

# Se llama a la funcion CargarPersonas la cual busca en el archivo, carga los datos y los guarda en la variable
personas = CargarPersonas()
# Se llama a la funcion CargarLibros la cual busca en el archivo, carga los datos y los guarda en la variable
libros = CargarLibros()
# Se llama a la funcion CargarPrestamos la cual busca en el archivo, carga los datos y los guarda en la variable
prestamos = CargarPrestamos()

# Convierte la columna de nombre a string para que no falle a la hora de ordenar
personas['Nombre'] = personas['Nombre'].astype(str)
# Variable la cual nos indica cuantas veces se va a continuar en el menú ( True = sigue, False=detiene el bucle y sale del menú)
ciclo = True

CargarPersonas()
while ciclo:
    Menu()  # Llama a la funcion del menú para que despliegue las opciones disponibles
    # Leemos la opción de menú a ejecutar
    opcion = input("Selecione una opción del menú: ")

    if(opcion == 'a'):  # Validamos si la opción selecionada es igual a la predefinida
        Lineas()  # Llamada a la funcion Lineas
        print("LISTA DE PERSONAS")
        print(personas)  # Imprimimos en consola la lista de personas

    elif(opcion == 'b'):  # Validamos si la opción selecionada es igual a la predefinida
        # se usa el sort_values para poder indicar por cual columna vamos a ordenar
        personas = personas.sort_values(by='Nombre')

    elif(opcion == 'c'):  # Validamos si la opción selecionada es igual a la predefinida
        # Se lee el codigo de persona a buscar
        codigo = input("Ingrese código de persona a buscar : ")
        # Ciclo for el cual le indicamos 0 que es de donde inicia hasta el tamaño de la lista para recorrerla
        for x in range(0, len(personas)):
            # se busca el código de la persona en la lista para comparar con el codigo ingresado
            if(str(personas['Codigo'][x]) == codigo):
                Lineas()  # Llamada a la funcion Lineas
                # Impresión en consola de un mensaje
                print("Se encontró la persona:")
                # Impresión en consola de un mensaje
                print("Código: ", personas['Codigo'][x])
                # Impresión en consola de un mensaje
                print("Nombre: ", personas['Nombre'][x])
                # Impresión en consola de un mensaje
                print("Correo; ", personas['Correo'][x])

    elif(opcion == 'd'):  # Validamos si la opción selecionada es igual a la predefinida
        Lineas()  # Llamada a la funcion Lineas
        print("LISTA DE LIBROS")  # Impresión en consola de un mensaje
        print(libros)  # Imprimimos en consola la lista de libros

    elif(opcion == 'e'):  # Validamos si la opción selecionada es igual a la predefinida
        # Se lee el codigo de libro a buscar
        codigo = input("Ingrese código de libro a buscar : ")
        # Ciclo for el cual le indicamos 0 que es de donde inicia hasta el tamaño de la lista para recorrerla
        for x in range(0, len(libros)):
            # se busca el código del libro en la lista para comparar con el codigo ingresado
            if(str(libros['idLibro'][x]) == codigo):
                Lineas()  # Llamada a la funcion Lineas
                # Impresión en consola de un mensaje
                print("Se encontró el libro : ")
                # Impresión en consola de un mensaje
                print("ID Libro: ", libros['idLibro'][x])
                # Impresión en consola de un mensaje
                print("Nombre: ", libros['nombre'][x])
                # Impresión en consola de un mensaje
                print("Genero; ", libros['Genero'][x])
                # Impresión en consola de un mensaje
                print("Autor; ", libros['Autor'][x])

    elif(opcion == 'f'):  # Validamos si la opción selecionada es igual a la predefinida
        # Se solicita el codigo del libo a prestar
        codigoLibro = input("Ingrese código de libro a prestar : ")
        # Se solicita el codigo de la persona a prestar el libro
        codigoPersona = input(
            "Ingrese código de persona a la que se le va a prestar el libro : ")
        # Ciclo for el cual le indicamos 0 que es de donde inicia hasta el tamaño de la lista para recorrerla
        for x in range(0, len(personas)):
            # se busca el código de la persona en la lista para comparar con el codigo
            if(str(personas['Codigo'][x]) == codigoPersona):
                # Asignamos el codigo de la persona encotrada a una variable para guardarla en la lista
                codEncPersona = personas['Codigo'][x]
                # Asignamos el nombre de la persona encotrada a una variable para guardarla en la lista
                nomEncPersona = personas['Nombre'][x]
                # Asignamos true a la variable lo cual indica que si encontró la persona
                encontroLibro = True

        # Ciclo for el cual le indicamos 0 que es de donde inicia hasta el tamaño de la lista para recorrerla
        for x in range(0, len(libros)):
            # se busca el código del libro en la lista para comparar con el codigo
            if(str(libros['idLibro'][x]) == codigoLibro):
                # Asignamos el codigo del libro encotrado a una variable para guardarla en la lista
                codEncLibro = libros['idLibro'][x]
                # Asignamos el nombre del libro encotrado a una variable para guardarla en la lista
                nomEncLibro = libros['nombre'][x]
                # Asignamos true a la variable lo cual indica que si encontró el libro
                encontroLibro = True

        if encontroLibro and encontroLibro:  # Validamos que se encontró el libro y la persona en la lista, eso con las dos variables de tipo boleanas
            prestamos = prestamos.append({'CodigoPersona': codEncPersona,
                                          'NombrePersona': nomEncPersona,
                                          'IDLibro': codEncLibro,
                                          'NombreLibro': nomEncLibro}, ignore_index=True)  # agregamos una nueva linea de prestamo a la lista
            # Mandamos a llamar la funcion que escribe en el archivo de excel para guardar los prestamos
            EscribirArchivo(prestamos)
        else:
            print("Codigo de libro o persona no encontrado")
    elif(opcion == 'g'):  # Validamos si la opción selecionada es igual a la predefinida
        Lineas()  # Llamada a la funcion Lineas
        print("LIBROS PRESTADOS")  # Impresión en consola de un mensaje
        print(prestamos)  # Imprimimos en consola la lista de prestamoa

    elif(opcion == 'h'):  # Validamos si la opción selecionada es igual a la predefinida
        ciclo = False  # Asignamos false a la variable para que salga del menú
