#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DIAGNOSTICO Y LIMPIEZA - Cuestionario Nordico
Tesis de Salud Pública
"""

import pandas as pd
import numpy as np
import os
import glob
import warnings
warnings.filterwarnings('ignore')

print("=" * 70)
print("DIAGNOSTICO DE ARCHIVOS")
print("=" * 70)

# =============================================================================
# PASO 1: CONFIGURACIÓN - EL USUARIO DEBE ESPECIFICAR SU RUTA
# =============================================================================

# IMPORTANTE: Cambiar esta ruta por la ubicación de tus datos
# Ejemplo: ruta_base = '/ruta/a/tus/datos/'
ruta_base = input("Ingresa la ruta donde están tus archivos Excel: ")

print(f"\n1. Verificando ruta base:")
print(f"   {ruta_base}")

# Verificar si la carpeta existe
if os.path.exists(ruta_base):
    print("   ✓ La carpeta EXISTE")
    
    # Listar todos los archivos en la carpeta
    print(f"\n2. Archivos encontrados en la carpeta:")
    try:
        archivos = os.listdir(ruta_base)
        for i, archivo in enumerate(archivos, 1):
            print(f"   {i}. {archivo}")
    except Exception as e:
        print(f"   Error al listar: {e}")
    
    # Buscar específicamente archivos Excel
    print(f"\n3. Archivos Excel (.xlsx) encontrados:")
    try:
        archivos_xlsx = glob.glob(os.path.join(ruta_base, '*.xlsx'))
        if archivos_xlsx:
            for archivo in archivos_xlsx:
                print(f"   - {os.path.basename(archivo)}")
        else:
            print("   No se encontraron archivos .xlsx")
    except Exception as e:
        print(f"   Error: {e}")
        
else:
    print("   ✗ La carpeta NO EXISTE")
    print("   Verifica la ruta manualmente")

# =============================================================================
# PASO 2: INTENTAR CARGAR ARCHIVO
# =============================================================================

print(f"\n4. Intentando cargar archivo...")

# Lista de posibles nombres (sin incluir rutas absolutas)
posibles_nombres = [
    'Anexo I. Datos fuente del Cuestionario Nórdico.xlsx',
    'Anexo I. Datos fuente del Cuestionario Nordico.xlsx',
    'Anexo I. Datos fuente.xlsx',
    'Datos fuente del Cuestionario Nordico.xlsx',
    'Anexo I.xlsx',
]

archivo_cargado = False

for nombre in posibles_nombres:
    ruta_completa = os.path.join(ruta_base, nombre)
    print(f"\n   Intentando: {nombre}")
    
    if os.path.exists(ruta_completa):
        print(f"   ✓ Archivo encontrado!")
        try:
            df = pd.read_excel(ruta_completa, sheet_name='Formularsvar 1')
            print(f"   ✓ Cargado exitosamente: {len(df)} registros")
            archivo_cargado = True
            break
        except Exception as e:
            print(f"   ✗ Error al cargar: {e}")
    else:
        print(f"   ✗ No existe")

if not archivo_cargado:
    print(f"\n{'='*70}")
    print("SOLUCIÓN:")
    print(f"{'='*70}")
    print("1. Verifica el nombre EXACTO del archivo Excel")
    print("2. Asegúrate que la hoja se llame 'Formularsvar 1'")
    print("3. Modifica la lista 'posibles_nombres' con tu nombre exacto")
    print(f"{'='*70}")
    exit()

# =============================================================================
# PASO 3: PROCESAMIENTO DE DATOS
# =============================================================================

print(f"\n{'='*70}")
print("PROCESANDO DATOS")
print(f"{'='*70}")

# Limpiar nombres de columnas
df.columns = df.columns.str.strip()

print(f"\nColumnas cargadas: {len(df.columns)}")
print(f"Primeras columnas: {list(df.columns[:3])}...")

# RENOMBRAR COLUMNAS CLAVE
nuevos_nombres = {
    'ID': 'id',
    'En qué puerto trabajas?': 'puerto',
    ' ¿Cuál es tu género? ': 'genero',
    ' ¿Tu edad? ': 'edad',
    '¿Cuánto mides sin zapatos en centímetros?': 'altura_cm',
    '¿Cuánto pesa, en kilogramos?': 'peso_kg',
    '¿Seleccione su puesto de trabajo en este momento?': 'puesto_actual',
    '¿Cuántos años en este tipo de trabajo?': 'antiguedad_actual',
    '¿Cómo está tu salud en general? (marque solo uno)': 'estado_salud',
    '¿Cuánto fuma (marque solo uno)': 'tabaquismo'
}
df.rename(columns=nuevos_nombres, inplace=True)

# LIMPIEZA BASICA
df['puerto'] = df['puerto'].str.strip()
df['genero'] = df['genero'].str.strip()
df['edad'] = pd.to_numeric(df['edad'], errors='coerce')

def corregir_altura(valor):
    if pd.isna(valor):
        return np.nan
    try:
        num = float(str(valor).replace(',', '.'))
        if num < 2.5:
            return num * 100
        elif num < 100:
            return num * 100
        else:
            return num
    except:
        return np.nan

df['altura_cm'] = df['altura_cm'].apply(corregir_altura)
df['peso_kg'] = pd.to_numeric(df['peso_kg'], errors='coerce')
df['imc'] = df['peso_kg'] / ((df['altura_cm']/100) ** 2)

def clasificar_imc(imc):
    if pd.isna(imc):
        return 'No calculable'
    elif imc < 18.5:
        return 'Bajo peso'
    elif imc < 25:
        return 'Normal'
    elif imc < 30:
        return 'Sobrepeso'
    else:
        return 'Obesidad'

df['categoria_imc'] = df['imc'].apply(clasificar_imc)

# PUESTO DE TRABAJO
df['puesto_actual'] = df['puesto_actual'].astype(str).str.strip()

def clasificar_puesto(puesto):
    if puesto == 'nan':
        return 'No especificado'
    puesto = puesto.lower()
    if any(x in puesto for x in ['operaciones', 'operador', 'conductor', 'capataz', 'controlador']):
        return 'Operaciones'
    elif any(x in puesto for x in ['mantenimiento', 'mecanico', 'electricista']):
        return 'Mantenimiento'
    elif any(x in puesto for x in ['administracion', 'almacenista']):
        return 'Administrativo'
    else:
        return 'Otros'

df['categoria_puesto'] = df['puesto_actual'].apply(clasificar_puesto)

# Mostrar resultados
print(f"\n{'='*70}")
print("RESUMEN DE DATOS PROCESADOS")
print(f"{'='*70}")
print(f"Total de registros: {len(df)}")
print(f"Variables procesadas: {len(df.columns)}")
print("\nVariables principales:")
print(f"- Edad: media {df['edad'].mean():.1f} años")
print(f"- IMC: media {df['imc'].mean():.1f}")
print(f"- Géneros: {df['genero'].unique()}")
print(f"- Puertos: {df['puerto'].unique()}")

print(f"\n{'='*70}")
print("PROCESO COMPLETADO")
print(f"{'='*70}")
