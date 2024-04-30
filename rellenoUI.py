import openpyxl
import tkinter as tk
from tkinter import messagebox
import pyautogui
import time

def procesar_datos():
    try:
        # Procesar los datos del Excel y realizar las acciones necesarias
        # Aquí va el código para procesar los datos del Excel, similar a la versión original
        # ...
        print('alexis romero')
        
        messagebox.showinfo("Éxito", "¡Los datos se procesaron con éxito!")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

def mostrar_interfaz_grafica():
    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.title("Procesador de Datos")

    # Crear etiquetas y botón
    etiqueta = tk.Label(ventana, text="Haga clic en el botón para procesar los datos del Excel:")
    boton = tk.Button(ventana, text="Procesar Datos", command=procesar_datos)

    # Colocar widgets en la ventana
    etiqueta.pack()
    boton.pack()

    # Iniciar el bucle principal de la ventana
    ventana.mainloop()

if __name__ == "__main__":
    # Mostrar la interfaz gráfica al ejecutar el programa
    mostrar_interfaz_grafica()
