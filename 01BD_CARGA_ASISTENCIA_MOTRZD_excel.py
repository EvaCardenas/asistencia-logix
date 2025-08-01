#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from pathlib import Path
from datetime import datetime  # ✅ Asegúrate de incluir esto

# Crear carpeta de salida si no existe
Path("salida").mkdir(exist_ok=True)

# Leer los archivos desde la carpeta 'data'
archivo_excel_1 =pd.ExcelFile( "data/Reporte_Almacén_Asistencia 31-07.xlsx")
archivo_excel_2 = pd.ExcelFile("data/Reporte_Citas_20250731_1153.xlsx")

hoja_nombre = 'REPORTE CITAS'.strip()

if hoja_nombre in archivo_excel_2.sheet_names:
    df_2 = archivo_excel_2.parse(hoja_nombre)

    # Convertir fechas
    df_2['FECHA_ENTREGA'] = pd.to_datetime(df_2['FECHA_ENTREGA'], unit='D', origin='1900-01-01')
    df_2['FECHA_ENTREGA'] = df_2['FECHA_ENTREGA'].dt.strftime('%d/%m/%Y')

    cambios = {'REPORTE CITAS': {'renombradas': [], 'agregadas': []}}

    columnas_renombrar = {
        'USUARIO_ICC': 'USUARIO_ULTIMO_MOVIMIENTO',
        'FECHA_PRIMER_ESTADO_MOTORIZADO_ULT_VISITA': 'FECHA_ULTIMO_MOVIMIENTO'
    }

    for original, nuevo in columnas_renombrar.items():
        if original in df_2.columns:
            df_2.rename(columns={original: nuevo}, inplace=True)
            cambios[hoja_nombre]['renombradas'].append(f'{original} → {nuevo}')

    columnas_a_agregar = ['MOTIVO', 'PUNTO_ENCUENTRO_COMENTARIO', 'SUB_AREA', 'ID_HUELLA']
    for col in columnas_a_agregar:
        if col not in df_2.columns:
            df_2[col] = ''
            cambios[hoja_nombre]['agregadas'].append(col)

    columnas_deseadas = [
        'CAMPANA', 'CITA', 'PEDIDO', 'TIPO_CITA', 'MOTIVO_REPRO', 'TIPO_VENTA', 'CODIGO_VENTA',
        'TIPO_DOCUMENTO', 'NUMERO_DOCUMENTO', 'CLIENTE', 'TERCERO_TIPODOCUMENTO', 'TERCERO_NUMERODOCUMENTO',
        'TERCERO_NOMBRE', 'FECHAONECLICK', 'FECHAPICKING', 'FECHA_CREACION_PEDIDO_BPO',
        'SLA_TIPO_ENTREGA', 'FECHA_PACTADA', 'SEDE', 'DEPARTAMENTO_ENTREGA', 'PROVINCIA_ENTREGA',
        'DISTRITO_ENTREGA', 'DIRECCION_ENTREGA', 'DIRECCION_ORIGINAL', 'REFERENCIAS', 'TELEFONO_1',
        'TELEFONO_2', 'PRODUCTO', 'IMEI', 'OPERADOR', 'TELEFONO_PORTA', 'PLAN', 'ESTADO_DOCUMENTACION',
        'MOTIVO_DOCUMENTACION', 'OBSERVACIONES', 'FECHA_ENTREGA', 'MOTORIZADO', 'POS_MOTORIZADO',
        'MONTO_A_COBRAR', 'NUMERO_HUELLAS', 'CODIGO_RESPUESTA_RENIEC', 'DETALLE_RESPUESTA_RENIEC',
        'FECHA_CONSULTA_RENIEC', 'COURIER', 'SOCIO', 'VISITAS', 'PUNTOVENTA', 'CATEGORIA',
        'OBSERVACION_MOTORIZADO', 'PUNTO_ENCUENTRO', 'PUNTO_ENCUENTRO_MOTIVO', 'PUNTO_ENCUENTRO_COMENTARIO',
        'CODIGO_MEDIO_DE_PAGO', 'ID_VOUCHER_1', 'ID_VOUCHER_2', 'ID_VOUCHER_3', 'ID_VOUCHER_4',
        'ID_VOUCHER_5', 'MONTO_MEDIO_PAGO_1', 'MONTO_MEDIO_PAGO_2', 'MONTO_MEDIO_PAGO_3',
        'MONTO_MEDIO_PAGO_4', 'MONTO_MEDIO_PAGO_5', 'ORDEN_AVANZADA_TDE', 'DETALLE', 'FECHA_HORA_XTORE',
        'ID_HUELLA', 'MODALIDAD_PROVENIENCIA', 'TIENDA_PUNTO_DE_VENTA', 'VEP', 'TIPO_DELIVERY_1',
        'CORTE_SE', 'RANGO_VISITA_TIPO_DELIVERY_1', 'TIPO_DELIVERY_2', 'RANGO_VISITA_TIPO_DELIVERY_2',
        'DIA_DESPACHO', 'CUMPL_FECHA_PACTADA', 'SUB_AREA', 'MOTIVO', 'TIPO_PRODUCTO',
        'FECHA_ULTIMO_MOVIMIENTO', 'USUARIO_ULTIMO_MOVIMIENTO', 'FECHA_PRIMER_ESTADO_MOTORIZADO',
        'FECHA_ESTADO_MOTORIZADO', 'HORA_ESTADO_MOTORIZADO'
    ]

    columnas_existentes = [col for col in columnas_deseadas if col in df_2.columns]
    columnas_faltantes = list(set(columnas_deseadas) - set(df_2.columns))
    df_2 = df_2[columnas_existentes]

    print(f"\n✅ Columnas conservadas: {len(columnas_existentes)}")
    if columnas_faltantes:
        print(f"❌ Columnas faltantes: {columnas_faltantes}")

    if cambios[hoja_nombre]['renombradas']:
        print("🔁 Renombradas:")
        for c in cambios[hoja_nombre]['renombradas']:
            print(f" - {c}")
    if cambios[hoja_nombre]['agregadas']:
        print("➕ Agregadas:")
        for c in cambios[hoja_nombre]['agregadas']:
            print(f" - {c}")
else:
    print(f"❌ La hoja '{hoja_nombre}' no existe.")

# Leer df_1 y hacer merge con df_2
df_1 = archivo_excel_1.parse(archivo_excel_1.sheet_names[0])
df_1.columns = df_1.columns.str.strip()
df_2.columns = df_2.columns.str.strip()

print("Primeras filas de MOTORIZADO en df_1:", df_1['MOTORIZADO'].head())
print("Primeras filas de MOTORIZADO en df_2:", df_2['MOTORIZADO'].head())

df_merged = pd.merge(df_2, df_1[['MOTORIZADO']], on='MOTORIZADO', how='inner')

print(f"\n✅ El DataFrame resultante tiene {df_merged.shape[0]} filas y {df_merged.shape[1]} columnas.")
print(df_merged.head())

# Generar nombre con fecha actual
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
output_filename = f"salida/df_merged_motorizado_{fecha_hoy}.xlsx"

# Guardar archivo Excel
df_merged.to_excel(output_filename, index=False)
print(f"\n📁 Archivo guardado como: {output_filename}")

