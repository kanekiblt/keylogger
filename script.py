from pynput import keyboard
import pandas as pd
import threading
import time
import string
import os

contador = {letra: 0 for letra in string.ascii_lowercase}

archivo_excel = "pulsaciones.xlsx"
archivo_temp = "pulsaciones_temp.xlsx"
archivo_log = "teclas_log.txt"

lock = threading.Lock()

def guardar_excel():
    while True:
        try:
            time.sleep(10)

            with lock:
                datos = [(letra.upper(), contador[letra]) for letra in string.ascii_lowercase]

            df = pd.DataFrame(datos, columns=["Letra", "Cantidad"])
            df.to_excel(archivo_temp, index=False)

            try:
                os.replace(archivo_temp, archivo_excel)
                print("Excel actualizado")
            except PermissionError:
                print("Archivo abierto, no se pudo actualizar")

        except Exception as e:
            print("Error guardando:", e)


def al_presionar(tecla):
    try:
        letra = tecla.char.lower()

        with lock:
            if letra in contador:
                contador[letra] += 1

        with open(archivo_log, "a", encoding="utf-8") as f:
            f.write(letra)

    except:
        pass  


thread_guardado = threading.Thread(target=guardar_excel, daemon=True)
thread_guardado.start()

print("Contador de teclas activo (en segundo plano)")

with keyboard.Listener(on_press=al_presionar) as listener:
    listener.join()