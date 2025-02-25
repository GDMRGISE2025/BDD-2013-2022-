# -*- coding: utf-8 -*-
"""
Created on Fri Jan 24 17:06:12 2025

@author: masol
"""

import pandas as pd
import glob
import os

# Ruta donde se encuentran los archivos de Excel
excel_folder_path = "C:Insertar carpeta con las matrices de cada año" 
output_file = "promedio_exportaciones_filtrado.xlsx"

# Obtener una lista de todos los archivos Excel en la carpeta
excel_files = glob.glob(f"{excel_folder_path}/*.xlsx")

# Filtrar y descartar archivos temporales (~$)
excel_files = [file for file in excel_files if not os.path.basename(file).startswith("~$")]

# Lista para almacenar las matrices leídas y los conjuntos de productos
matrices = []
product_sets = []

# Leer cada archivo Excel y obtener el conjunto de productos
for file in excel_files:
    df = pd.read_excel(file, index_col=0)  # Leer archivo ignorando la primera columna como índice
    matrices.append(df)
    product_sets.append(set(df.columns))

# Identificar los productos comunes (los productos del archivo con menos columnas)
common_products = set.intersection(*product_sets)

# Filtrar todas las matrices para que contengan solo los productos comunes
filtered_matrices = [df[common_products] for df in matrices]

# Asegurar que todas las matrices tengan las mismas dimensiones
if not all(m.shape == filtered_matrices[0].shape for m in filtered_matrices):
    raise ValueError("Las matrices no se pudieron igualar correctamente en dimensiones.")

# Calcular el promedio de las matrices
average_matrix = sum(filtered_matrices) / len(filtered_matrices)

# Restaurar la primera fila y columna para la salida
average_matrix.index.name = "country_id"

# Exportar la matriz promedio a un archivo Excel
average_matrix.to_excel(output_file)

print(f"Matriz promedio filtrada guardada en: {output_file}")
