import cairo
from PIL import Image

surface = cairo.ImageSurface.create_from_png("images/FrenteVacio.png")
# Se crea un contexto
context = cairo.Context(surface)

# Calcular la posición de los textos
texto1 = "Gabriel Luna"
texto2 = "Gerente Operativo"
telefone = "(646) 175.7732 ext. 125"
email = "g.luna@grupoaceves.com"
#email = ""
#cellphone = "(646) 125.4512"
cellphone = ""
address_part1 = "Blvd. Teniente Azueta #130 int. 210"
address_part2 = "Recinto Portuario, Ensenada B.C. C.P. 22800"
#address_part1 = "Av. Alejandro Von Humboldt 17618-Int. B,"
#address_part2 = "Garita de Otay, 22430 Tijuana, B.C."

# Dimensiones Cuadro de texto nombre y posicion
cuadro_x = 685
cuadro_y = 16
cuadro_ancho = 250
cuadro_alto = 50

# Dimensiones Cuadro de texto telefono, correo, celular y address
if cellphone == "":
    cuadro_tel_x = 645
    cuadro_tel_y = 70
    cuadro_tel_ancho = 290
    cuadro_tel_alto = 50
    #cuadro address
    cuadro_addr_x = 600
    cuadro_addr_y = 125
    cuadro_addr_ancho = 330
    cuadro_addr_alto = 50
else:
    cuadro_tel_x = 645
    cuadro_tel_y = 70
    cuadro_tel_ancho = 290
    cuadro_tel_alto = 60
    #cuadro address
    cuadro_addr_x = 580
    cuadro_addr_y = 135
    cuadro_addr_ancho = 330
    cuadro_addr_alto = 50



context.set_font_size(19)
###name
texto1_extents = context.text_extents(texto1)
texto2_extents = context.text_extents(texto2)

margen_vertical = (cuadro_alto - (texto1_extents.height + texto2_extents.height)) / 2

texto1_x = cuadro_x + (cuadro_ancho - texto1_extents.width) / 2
texto1_y = cuadro_y + margen_vertical + texto1_extents.height

texto2_x = cuadro_x + (cuadro_ancho - texto2_extents.width) / 2
texto2_y = cuadro_y + margen_vertical + texto1_extents.height + 20



#####telefone

telefone_extents = context.text_extents(telefone)
email_extents = context.text_extents(email)
cellphone_extents = context.text_extents(cellphone)

margen_vertical_tel = (cuadro_tel_alto - (telefone_extents.height + email_extents.height + cellphone_extents.height)) / 2

telefone_x = cuadro_tel_x + (cuadro_tel_ancho - telefone_extents.width) / 2
telefone_y = cuadro_tel_y + margen_vertical_tel + telefone_extents.height

if cellphone == "":
    #####email
    email_x = cuadro_tel_x + (cuadro_tel_ancho - email_extents.width) / 2
    email_y = cuadro_tel_y + margen_vertical_tel + telefone_extents.height + 20
else:
    ####cellphone
    cellphone_x = cuadro_tel_x + (cuadro_tel_ancho - cellphone_extents.width) / 2
    cellphone_y = cuadro_tel_y + margen_vertical_tel + telefone_extents.height + 20

    #####email
    email_x = cuadro_tel_x + (cuadro_tel_ancho - email_extents.width) / 2
    email_y = cuadro_tel_y + margen_vertical_tel + cellphone_extents.height + 40


#### address ##################
address_part1_extents = context.text_extents(address_part1)
address_part2_extents = context.text_extents(address_part2)

margen_vertical_address = (cuadro_addr_alto - (address_part1_extents.height + address_part2_extents.height)) / 2

address_part1_x = cuadro_addr_x + (cuadro_addr_ancho - address_part1_extents.width) / 2
address_part1_y = cuadro_addr_y + margen_vertical_address+ address_part1_extents.height

address_part2_x = cuadro_addr_x + (cuadro_addr_ancho - address_part2_extents.width) / 2
address_part2_y = cuadro_addr_y + margen_vertical_address+ address_part1_extents.height + 20

""" texto1_x = cuadro_x + (cuadro_ancho - texto1_extents.width) / 2
texto1_y = cuadro_y + margen_vertical + texto1_extents.height

texto2_x = cuadro_x + (cuadro_ancho - texto2_extents.width) / 2
texto2_y = cuadro_y + margen_vertical + texto1_extents.height + 20 """


# Dibujar el cuadro de texto transparente

context.rectangle(cuadro_x, cuadro_y, cuadro_ancho, cuadro_alto)
context.rectangle(cuadro_tel_x, cuadro_tel_y, cuadro_tel_ancho, cuadro_tel_alto)
context.rectangle(cuadro_addr_x, cuadro_addr_y, cuadro_addr_ancho, cuadro_addr_alto)
context.set_source_rgba(0, 0, 0, 0)
context.fill()

# Dibujar los textos centrados dentro del cuadro
context.set_source_rgb(255, 255, 255)  # Establecer color del texto
context.set_font_size(19)

context.move_to(texto1_x, texto1_y)
context.show_text(texto1)

context.move_to(texto2_x, texto2_y)
context.show_text(texto2)
context.set_font_size(18)

context.move_to(telefone_x, telefone_y)
context.show_text(telefone)
if cellphone != "":
    context.move_to(cellphone_x, cellphone_y)
context.show_text(cellphone)

context.move_to(email_x, email_y)
context.show_text(email)

context.move_to(address_part1_x, address_part1_y)
context.show_text(address_part1)

context.move_to(address_part2_x, address_part2_y)
context.show_text(address_part2)

# Guardar la imagen resultante
surface.write_to_png("imagen_modificada.png")  # Especifica el nombre y la ruta de la imagen resultante

# Mostrar la imagen resultante
imagen_modificada = Image.open("imagen_modificada.png")  # Cargar la imagen resultante
imagen_modificada.show()
