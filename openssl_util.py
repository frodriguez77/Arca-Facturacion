"""
Utilidad para localizar el ejecutable de OpenSSL en Windows.
Busca en las rutas más comunes y también en el PATH del sistema.
"""
import os
import shutil

_RUTAS_CONOCIDAS = [
    r"C:\Program Files\OpenSSL-Win64\bin\openssl.exe",
    r"C:\Program Files\OpenSSL\bin\openssl.exe",
    r"C:\Program Files (x86)\OpenSSL-Win32\bin\openssl.exe",
    r"C:\Program Files\Git\mingw64\bin\openssl.exe",
    r"C:\Program Files\Git\usr\bin\openssl.exe",
    r"C:\Git\mingw64\bin\openssl.exe",
    r"C:\tools\OpenSSL\bin\openssl.exe",
]

def encontrar_openssl():
    """
    Retorna la ruta completa al ejecutable de OpenSSL.
    Lanza Exception si no lo encuentra.
    """
    # 1. Buscar en rutas conocidas
    for ruta in _RUTAS_CONOCIDAS:
        if os.path.isfile(ruta):
            return ruta

    # 2. Buscar en el PATH del sistema
    en_path = shutil.which('openssl')
    if en_path:
        return en_path

    raise Exception(
        "No se encontró OpenSSL en el sistema.\n"
        "Instalalo desde https://slproweb.com/products/Win32OpenSSL.html\n"
        "o asegurate de tener Git for Windows instalado."
    )
