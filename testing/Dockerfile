# Usa una imagen base de Python
FROM python:3.9

# Establece el directorio de trabajo en /app
WORKDIR /app

# Copia los archivos locales al contenedor
COPY . .

# Instala las dependencias de sistema necesarias para tkinter
RUN apt-get update && apt-get install -y python3 python3-tk libreoffice

RUN pip install Pillow

# Configura la variable de entorno para la GUI
ENV DISPLAY=:0

# Ejecuta tu script cuando el contenedor se inicie
CMD ["python", "main.py"]
