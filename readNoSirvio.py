""" import pandas as pd

ruta_archivo = "Empleados-OPERATIVO.xlsx"
nombre_hoja = "Lista Empleados"

# Saltar la primera fila y seleccionar solo las columnas 2 y 3
datos_excel = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, skiprows=1, usecols="B:C")

print(datos_excel.head()) """

import pandas as pd
import cairo
import os
from pathlib import Path
import shutil
import subprocess

ruta_archivo = "Empleados-OPERATIVO.xlsx"
nombre_hoja = "Lista Empleados"

# Leer el archivo de Excel y seleccionar las columnas 1 y 2
datos_excel = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, usecols=[1, 2])

# Extraer el nombre de la columna 1 y el texto de la columna 2
nombre_completo = datos_excel.iloc[:, 0].str.split(" ", expand=True)
nombre = nombre_completo[2]
primer_apellido = nombre_completo[0]
puesto = datos_excel.iloc[:, 1]

# Variables globales para almacenar los textos de cada campo
name = ""
position = ""
email = ""
cellphone = ""
address_part1 = "Blvd. Teniente Azueta #130 int. 210"
address_part2 = "Recinto Portuario, Ensenada B.C. C.P. 22800"
telefone = "(646) 175.7732"
ext = ""


#print(nombre + " " + primer_apellido + " " + puesto)
# Imprimir los valores extraídos junto con el texto de la columna 2


def mostrar_vista_previa():
    global surface
    # Cargar la imagen a modificar
    surface = cairo.ImageSurface.create_from_png("images/FrenteVacio.png")
    # Se crea un contexteo
    context = cairo.Context(surface)

    # Configurar la fuente y el texto
    font_path = "Fonts/microsoft_jhenghei.ttf"
    context.select_font_face(font_path)
    context.set_source_rgb(255, 255, 255)  # Color del texto

    # Dibujar el texto con fondo transparente
    #x_name = 705
    #y_name = 40
    #x_position = 672
    y_position = 62
    x_telefone = 680
    y_telefone = 95
    x_ext = 810
    y_ext = 95
    #x_email = 650
    
    #x_cellphone = 680
    
    if cellphone == "":
        x_address_part1 = 622
        y_address_part1 = 145
        #x_address_part2 = 592
        y_address_part2 = 165
    else:
        #global x_cellphone
        
        x_address_part1 = 622
        y_address_part1 = 165
        #x_address_part2 = 592
        y_address_part2 = 185
    
    ##### NAME AND POSITION ########################
    cuadro_x = 685
    cuadro_y = 16
    cuadro_ancho = 250
    cuadro_alto = 50
    context.set_font_size(19)
    name_extends = context.text_extents(name)
    position_extends = context.text_extents(position)
    

    margen_vertical = (cuadro_alto - (name_extends.height + position_extends.height)) / 2

    name_x = cuadro_x + (cuadro_ancho - name_extends.width) / 2
    name_y = cuadro_y + margen_vertical + name_extends.height

    position_x = cuadro_x + (cuadro_ancho - position_extends.width) / 2
    position_y = cuadro_y + margen_vertical + name_extends.height + 20

    # Dibujar el cuadro de texto transparente
    context.rectangle(cuadro_x, cuadro_y, cuadro_ancho, cuadro_alto)
    context.set_source_rgba(0, 0, 0, 0)
    #context.fill()
    #dibujar textos
    context.move_to(name_x, name_y)
    context.show_text(name)

    context.move_to(position_x, position_y)
    context.show_text(position)
    

    #POSITION-----------------------------------------
    """ # Mide el ancho del texto previo
    extents = context.text_extents(name)
    ancho_name = extents.width
    # Mide el ancho del texto a centrar
    extents = context.text_extents(position)
    ancho_position = extents.width
    # Calcula la posición X para centrar el texto
    x_position = (x_name + ancho_name/2) - (ancho_position/2)
    # Dibujar
    context.set_font_size(18)
    context.move_to(x_position, y_position)
    context.show_text(position) """
    #-------------------------------------------------

    #TELEFONE----------------------------------------
    context.move_to(x_telefone, y_telefone)
    context.show_text(telefone)
    #------------------------------------------------

    #EXT----------------------------------------
    context.move_to(x_telefone+130, y_telefone)
    context.show_text(ext)
    #------------------------------------------------

    #CELLPHONE--------------------------------------
    if telefone != "" and cellphone != "":
        y_cellphone = 115
        # Mide el ancho del texto previo
        extents = context.text_extents(telefone)
        extents2 = context.text_extents(ext)
        ancho_telefone = extents.width + extents2.width
        # Mide el ancho del texto a centrar
        extents = context.text_extents(cellphone)
        ancho_cellphone = extents.width
        # Calcula la posición X para centrar el texto
        x_cellphone = (x_telefone + ancho_telefone/2) - (ancho_cellphone/2)
        
        #Dibujar
        context.move_to(x_cellphone, y_cellphone)
        context.show_text(cellphone)
        #icono
        icono = cairo.ImageSurface.create_from_png("images/WhatsappIcon.png")
        icono_x = x_cellphone - 25
        icono_y = y_cellphone - 13
       # Dibujar el icono en el contexto como máscara
        context.save()  # Guardar el estado actual del contexto
        context.set_source_surface(icono, icono_x, icono_y)
        context.paint_with_alpha(1)  # Utilizar la imagen como máscara con opacidad completa
        context.restore()  # Restaurar el estado del contexto
    #-------------------------------------------------

    #EMAIL-------------------------------------------
    if cellphone != "":
        y_email = 135
        # Mide el ancho del texto previo
        extents = context.text_extents(cellphone)
        ancho_cellphone = extents.width
        # Mide el ancho del texto a centrar
        extents = context.text_extents(email)
        ancho_email = extents.width
        # Calcula la posición X para centrar el texto
        x_email = (x_cellphone + ancho_cellphone/2) - (ancho_email/2)
        context.move_to(x_email, y_email)
    else:
        y_email = 110
        # Mide el ancho del texto previo
        extents = context.text_extents(telefone) 
        extents2 = context.text_extents(ext) 
        ancho_telefone = extents.width + extents2.width
        # Mide el ancho del texto a centrar
        extents = context.text_extents(email)
        ancho_email = extents.width
        # Calcula la posición X para centrar el texto
        x_email = (x_telefone + ancho_telefone/2) - (ancho_email/2)
        context.move_to(x_email, y_email)

    #Dibujar
    
    context.show_text(email)
    #-----------------------------------------------

    #ADDRESS-------------------------------------------
    #address_part1
    context.set_font_size(17)
    context.move_to(x_address_part1, y_address_part1)
    context.show_text(address_part1)
    # Mide el ancho del texto previo
    extents = context.text_extents(address_part1)
    ancho_address_part1 = extents.width
    # Mide el ancho del texto a centrar
    extents = context.text_extents(address_part2)
    ancho_address_part2 = extents.width
    # Calcula la posición X para centrar el texto
    x_address_part2 = (x_address_part1 + ancho_address_part1/2) - (ancho_address_part2/2)
    #Dibujar
    #address_part2
    context.move_to(x_address_part2, y_address_part2)
    context.show_text(address_part2)
    #-------------------------------------------------

    # Guardar la imagen modificada
    surface.write_to_png("imagen_modificada.png")

def custom_title(s):
    no_capitalize = ["y", "e", "o", "u", "de"]
    words = s.split(' ')
    for i in range(len(words)):
        if words[i].lower() not in no_capitalize or i == 0:  # Convertimos la palabra a minúsculas antes de verificar
            words[i] = words[i].capitalize()
        else:
            words[i] = words[i].lower()  # Convertimos la palabra a minúsculas si está en la lista
    return ' '.join(words)

user_homefolder = str(Path.home())   
def save_result():
    global surface
    global name
    first_name_text = name
    if first_name_text != "":
        export_firma_folder = (user_homefolder +
                            '/Downloads/' + first_name_text)
        export_file_name = first_name_text + ".png"
        if os.path.exists(export_firma_folder):
            shutil.rmtree(export_firma_folder)
        os.makedirs(export_firma_folder)
        export_front = (export_firma_folder + '/' +
                        'Frente_' + export_file_name)

        surface.write_to_png(export_front)

        #formatted_path = os.path.normpath(export_firma_folder)
        #subprocess.Popen(r'explorer /open,"{}"'.format(formatted_path))

for nom, ape, pue in zip(nombre, primer_apellido, puesto):
    name = nom + " " + ape
    name = custom_title(name)
    position = pue
    mostrar_vista_previa()
    save_result()
    #print(f"Nombre: {nom}, Apellido: {ape}, Puesto: {pue}") 


# Crear un nuevo DataFrame con los nombres y apellidos separados
""" nuevo_dataframe = pd.DataFrame({
    "Nombre": nombre,
    "Apellido": primer_apellido,
    "Puesto": datos_excel2
    
})

print(nuevo_dataframe.head())
 """