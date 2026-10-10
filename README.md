# Taller Excel + Python

Aplicación de consola para consultar registros de un archivo Excel.

## Requisitos

- Python 3
- Dependencias de `requirements.txt`
- Un archivo `personas.xlsx` junto a `lectura.py`

Instala las dependencias con:

```powershell
python -m pip install -r requirements.txt
```

Coloca tu archivo `personas.xlsx` en la carpeta del proyecto y ejecuta:

```powershell
python lectura.py
```

El archivo Excel se excluye del repositorio porque puede contener datos personales. La primera fila debe contener encabezados y los registros deben incluir ID, nombre, empresa, correo y dirección MAC en ese orden.

## Menú

- Ver todos los registros
- Mostrar únicamente nombre y correo
- Buscar una persona por número o ID
- Filtrar por empresa
- Contar registros
