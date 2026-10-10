"""
Taller Excel + Python
Lee el archivo personas.xlsx y permite consultarlo desde un menú interactivo.

Retos resueltos:
    1. Mostrar únicamente Name y Email.
    2. Mostrar cuántos registros tiene el archivo.
    3. Filtrar personas cuya empresa contenga una palabra (por defecto "Group").
    4. Pedir un ID con input() y mostrar solo esa persona.
    5. Menú interactivo.
"""

from pathlib import Path

import openpyxl
from tabulate import tabulate

# El Excel se busca junto a este script, sin importar desde dónde se ejecute
ARCHIVO = Path(__file__).parent / "personas.xlsx"

ENCABEZADOS = ["#", "ID", "Name", "Company", "Email", "MAC Address"]


# ---------------------------------------------------------------- Lectura ---

def cargar_personas(ruta):
    """Lee el Excel y devuelve una lista de diccionarios, uno por persona."""
    libro = openpyxl.load_workbook(ruta, read_only=True)
    hoja = libro.active

    personas = []
    # min_row=2 salta la fila de encabezados
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        # Se ignoran las filas totalmente vacías
        if not any(fila):
            continue
        personas.append({
            "numero": len(personas) + 1,
            "id": str(fila[0] or ""),
            "name": str(fila[1] or ""),
            "company": str(fila[2] or ""),
            "email": str(fila[3] or ""),
            "mac": str(fila[4] or ""),
        })

    libro.close()
    return personas


# ----------------------------------------------------------- Presentación ---

def mostrar_tabla(filas, encabezados):
    """Imprime una tabla, o un aviso si no hay filas."""
    if not filas:
        print("\n  No hay registros para mostrar.")
        return
    print()
    print(tabulate(filas, headers=encabezados, tablefmt="fancy_grid",
                   colalign=("center",) * len(encabezados)))


def mostrar_personas(personas):
    """Muestra todas las columnas de las personas recibidas."""
    filas = [[p["numero"], p["id"], p["name"], p["company"], p["email"], p["mac"]]
             for p in personas]
    mostrar_tabla(filas, ENCABEZADOS)


def pausar():
    input("\nPresiona Enter para volver al menú...")


# ------------------------------------------------------------------ Retos ---

def ver_todos(personas):
    """Opción 1: muestra todos los registros."""
    mostrar_personas(personas)
    print(f"\n  {len(personas)} registros mostrados.")


def ver_nombre_y_email(personas):
    """Reto 1: muestra únicamente Name y Email."""
    filas = [[p["name"], p["email"]] for p in personas]
    mostrar_tabla(filas, ["Name", "Email"])


def contar_registros(personas):
    """Reto 2: muestra cuántos registros tiene el archivo."""
    print(f"\n  El archivo tiene {len(personas)} registros.")


def filtrar_por_empresa(personas):
    """Reto 3: filtra las personas cuya empresa contenga una palabra."""
    palabra = input("\nPalabra a buscar en la empresa [Group]: ").strip()
    # Si el usuario no escribe nada se usa la palabra del reto
    if not palabra:
        palabra = "Group"

    # La comparación no distingue mayúsculas de minúsculas
    encontradas = [p for p in personas if palabra.lower() in p["company"].lower()]

    if not encontradas:
        print(f"\n  Ninguna empresa contiene '{palabra}'.")
        return
    mostrar_personas(encontradas)
    print(f"\n  {len(encontradas)} persona(s) con '{palabra}' en la empresa.")


def buscar_por_id(personas):
    """Reto 4: pide un ID y muestra solo esa persona.

    Acepta el número de la columna # o el ID del archivo (completo o su inicio).
    """
    texto = input("\nIngresa el # o el ID de la persona: ").strip().lower()

    # Validación: no se permite una búsqueda vacía
    if not texto:
        print("\n  Debes ingresar un valor.")
        return

    if texto.isdigit() and 1 <= int(texto) <= len(personas):
        encontradas = [personas[int(texto) - 1]]
    else:
        encontradas = [p for p in personas if p["id"].lower().startswith(texto)]

    if not encontradas:
        print(f"\n  No se encontró ninguna persona con '{texto}'.")
    elif len(encontradas) > 1:
        # El inicio del ID coincide con varias personas: se pide más precisión
        mostrar_personas(encontradas)
        print(f"\n  {len(encontradas)} personas coinciden; escribe más caracteres del ID.")
    else:
        mostrar_personas(encontradas)


# ------------------------------------------------------------------- Menú ---

# Cada opción guarda el texto que se muestra y la función que ejecuta
OPCIONES = {
    "1": ("Ver todos", ver_todos),
    "2": ("Ver solo Name y Email", ver_nombre_y_email),
    "3": ("Buscar por ID", buscar_por_id),
    "4": ("Filtrar empresa", filtrar_por_empresa),
    "5": ("Contar registros", contar_registros),
    "0": ("Salir", None),
}


def mostrar_menu():
    print("\n" + "=" * 34)
    print("   TALLER EXCEL + PYTHON")
    print("=" * 34)
    for clave, (texto, _) in OPCIONES.items():
        print(f"  {clave}. {texto}")
    print("=" * 34)


def menu(personas):
    """Reto 5: menú interactivo. Se repite hasta que el usuario elige Salir."""
    while True:
        mostrar_menu()
        opcion = input("Elige una opción: ").strip()

        # Validación: solo se aceptan las opciones del menú
        if opcion not in OPCIONES:
            print("\n  Opción no válida. Escribe uno de los números del menú.")
            continue

        accion = OPCIONES[opcion][1]
        if accion is None:
            print("\n¡Hasta luego!")
            break

        accion(personas)
        pausar()


def main():
    # Validaciones de lectura: el archivo debe existir, abrir y tener datos
    try:
        personas = cargar_personas(ARCHIVO)
    except FileNotFoundError:
        print(f"No se encontró el archivo: {ARCHIVO}")
        return
    except PermissionError:
        print(f"No se pudo abrir {ARCHIVO.name}. Ciérralo en Excel e intenta de nuevo.")
        return
    except Exception as error:
        print(f"No se pudo leer {ARCHIVO.name}: {error}")
        return

    if not personas:
        print(f"El archivo {ARCHIVO.name} no tiene registros.")
        return

    try:
        menu(personas)
    except (KeyboardInterrupt, EOFError):
        # Ctrl+C cierra el programa sin mostrar un error
        print("\n\nPrograma interrumpido. ¡Hasta luego!")


if __name__ == "__main__":
    main()
