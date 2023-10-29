import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import os  # Agregamos la importación de la biblioteca os
import subprocess  # Agregamos la importación de la biblioteca subprocess

# Funciones para mostrar mensajes en el área de texto
def mostrar_mensaje1(valor1, valor2):
    texto_area.insert("end", f"Los valores ingresados son {valor1} y {valor2}\n")

def mostrar_mensaje2():
    texto_area.insert("end", "Presionaste el Botón 2\n")

# Funciones para abrir una imagen y mostrarla en una ventana
def abrir_imagen():
    # Abre un cuadro de diálogo para seleccionar la imagen
    ruta_imagen = filedialog.askopenfilename(filetypes=[("Archivos de imagen", "*.png;*.jpg;*.jpeg;*.gif;*.bmp")])

    if ruta_imagen:
        # Carga la imagen seleccionada con Pillow
        imagen = Image.open(ruta_imagen)

        # Crea una instancia de la imagen tkinter
        imagen_tk = ImageTk.PhotoImage(imagen)

        # Crea una ventana para mostrar la imagen
        ventana_imagen = tk.Toplevel()
        ventana_imagen.title("Imagen")
        etiqueta = tk.Label(ventana_imagen, image=imagen_tk)
        etiqueta.pack()

        # Evita que la imagen sea eliminada por la recolección de basura
        etiqueta.imagen_tk = imagen_tk

# Función para mostrar una imagen en la parte superior del frame de la imagen
def mostrar_imagen(ruta_imagen):
    # Carga la imagen proporcionada con Pillow
    imagen = Image.open(ruta_imagen)

    # Crea una instancia de la imagen tkinter
    imagen_tk = ImageTk.PhotoImage(imagen)

    # Crea una ventana para mostrar la imagen
    ventana_imagen = tk.Toplevel()
    ventana_imagen.title("Imagen")
    etiqueta = tk.Label(ventana_imagen, image=imagen_tk)
    etiqueta.pack()

    # Evita que la imagen sea eliminada por la recolección de basura
    etiqueta.imagen_tk = imagen_tk

    # Evita que la imagen sea eliminada por la recolección de basura
    etiqueta.imagen_tk = imagen_tk

# Función para abrir un archivo Excel
# Función para abrir un archivo Excel
def abrir_excel(nombre_archivo):
    if nombre_archivo:
        try:
            # Verifica si el archivo existe en el directorio actual
            if os.path.exists(nombre_archivo):
                # Utiliza la biblioteca subprocess para abrir el archivo Excel con la aplicación predeterminada
                subprocess.Popen(["libreoffice","--calc", nombre_archivo], shell=True)
            else:
                texto_area.insert("end", f"El archivo {nombre_archivo} no existe en el directorio actual.\n")
        except Exception as e:
            texto_area.insert("end", f"Error al abrir el archivo Excel: {str(e)}\n")

# Variables globales para almacenar los valores ingresados
valor1_global = "Sin valor ingresado"
valor2_global = "Sin valor ingresado"

# Función para abrir la ventana de entrada de valores
def abrir_ventana_valores():
    global valor1_global, valor2_global
    
    # Crea una nueva ventana emergente
    ventana_valores = tk.Toplevel(ventana)
    ventana_valores.title("Ingresar Valores")
    
    # Configura el tamaño de la ventana emergente
    ventana_valores.geometry("400x200")  # Ajusta las dimensiones según tu preferencia

    # Crea un frame para los botones
    frame_botones = tk.Frame(ventana_valores)
    frame_botones.pack(side=tk.BOTTOM, pady=10)  # Ajusta el espaciado vertical según tu preferencia

    # Crea campos de texto para ingresar valores
    etiqueta_valor1 = tk.Label(ventana_valores, text="Valor 1:")
    etiqueta_valor1.pack()
    valor1 = tk.Entry(ventana_valores)
    valor1.pack()

    etiqueta_valor2 = tk.Label(ventana_valores, text="Valor 2:")
    etiqueta_valor2.pack()
    valor2 = tk.Entry(ventana_valores)
    valor2.pack()

    # Función para guardar los valores ingresados y cerrar la ventana
    def guardar_valores():
        global valor1_global, valor2_global
        valor1_global = valor1.get()
        valor2_global = valor2.get()
        # Puedes realizar acciones adicionales aquí si es necesario
        ventana_valores.destroy()  # Cierra la ventana de entrada de valores

    # Botón para guardar los valores
    boton_guardar = tk.Button(frame_botones, text="Guardar", command=guardar_valores, height=2, width=20)
    boton_guardar.pack(side=tk.LEFT, padx=10)  # Ajusta el espaciado horizontal según tu preferencia

    # Botón para salir de la ventana de entrada de valores
    boton_salir = tk.Button(frame_botones, text="Salir", command=ventana_valores.destroy, height=2, width=20)
    boton_salir.pack(side=tk.RIGHT, padx=10)  # Ajusta el espaciado horizontal según tu preferencia

    

# Función para cerrar la ventana principal y salir del programa
def salir_programa():
    ventana.quit()
    ventana.destroy()

# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Biodial Software")

# Configurar el tamaño de la ventana
ventana.geometry("800x800")  # Cambia las dimensiones según tu preferencia

# Crear un frame para la imagen en la parte superior
frame_imagen = tk.Frame(ventana)
frame_imagen.pack(fill=tk.BOTH, expand=True)

# Cargar una imagen inicial en el frame de la imagen
imagen_fondo = Image.open("biodial.jpeg")  # Reemplaza con la ruta de tu imagen de fondo
imagen_fondo = ImageTk.PhotoImage(imagen_fondo)
label_imagen = tk.Label(frame_imagen, image=imagen_fondo)
label_imagen.pack(fill=tk.BOTH, expand=True)

# Crear un frame para el texto en la parte media
frame_texto = tk.Frame(ventana)
frame_texto.pack(fill=tk.BOTH, expand=True)

# Crear un área de texto para mostrar los mensajes en el frame de texto
texto_area = tk.Text(frame_texto, wrap=tk.WORD, font=("Helvetica", 12), height=4)  # Cambia el tamaño de fuente aquí
texto_area.pack(fill=tk.BOTH, expand=True)

# Crear un frame para los botones en la parte inferior
frame_botones = tk.Frame(ventana)
frame_botones.pack(fill=tk.BOTH, expand=True)

# Crear los botones con tamaño personalizado en el frame de botones
boton_resultado = tk.Button(frame_botones, text="Mostrar resultados", command=lambda:mostrar_mensaje1(valor1_global, valor2_global), height=2, width=100)
boton_excel = tk.Button(frame_botones, text="Abrir Excel", command=lambda:abrir_excel("Clasificacion y factor de Consumo Biodial.xlsx"), height=2, width=100)
boton_abrir = tk.Button(frame_botones, text="Abrir Imagen", command=abrir_imagen, height=2, width=100)
boton_abrir_con_ruta = tk.Button(frame_botones, text="Abrir Imagen con Ruta", command=lambda: mostrar_imagen("pronostico.PNG"), height=2, width=100)
boton_valores = tk.Button(frame_botones, text="Ingresar Valores", command=abrir_ventana_valores, height=2, width=100)
boton_salir = tk.Button(frame_botones, text="Salir", command=salir_programa, height=2, width=100)

# Configurar la geometría de los botones en el frame de botones
boton_resultado.pack(expand=True)
boton_excel.pack(expand=True)
boton_abrir.pack(expand=True)
boton_abrir_con_ruta.pack(expand=True)
boton_valores.pack(expand=True)
boton_salir.pack(expand=True)

# Iniciar el bucle de la aplicación
ventana.mainloop()
