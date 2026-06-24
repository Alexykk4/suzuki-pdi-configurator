# Suzuki PDI Configurator

Este repositorio contiene la herramienta para la preparación pre-entrega (PDI) de vehículos Suzuki.

## Contenido del Repositorio

- `FORMATO PDI.xlsx`: Plantilla en formato Excel que contiene las hojas de trabajo para los diferentes modelos de vehículos.
- `modificar_pdi.py`: Script interactivo en Python para configurar de forma rápida los datos del PDI (Modelo, Línea, Transmisión, Color, Número de Llave y VIN) y mandar a imprimir la hoja de forma automática en Windows.

## Instrucciones para Windows

### 1. Requisitos Previos

Necesitas tener instalado Python 3 y Microsoft Excel. Instala las librerías necesarias ejecutando el siguiente comando en la terminal (PowerShell o CMD):

```bash
pip install openpyxl pywin32
```

### 2. Ejecución

Ejecuta el configurador interactivo:

```bash
python modificar_pdi.py
```

Sigue las instrucciones en pantalla para seleccionar el modelo, línea, transmisión, color, número de llave y terminación de VIN. Una vez guardados los cambios, el script mandará a imprimir automáticamente la hoja modificada a tu impresora predeterminada.
