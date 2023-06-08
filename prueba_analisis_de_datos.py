import pandas as pd
import openpyxl
import glob
import seaborn as sns
import re
from unidecode import unidecode 
import matplotlib.pyplot as plt

def encontrar_fila_inicio(archivo):
    # Cargar el archivo de Excel utilizando openpyxl
    wb = openpyxl.load_workbook(archivo)
    sheet = wb.active

    # Buscar la fila de inicio recorriendo las celdas de cada columna
    for row in sheet.iter_rows():
        if row[3].value is not None:
            return row[3].row

    return None  # Si no se encuentra una fila de inicio válida

def remove_accents(x):
    if isinstance(x, str):
        return unidecode(x)
    else:
        return x

#Lectura de archivos de excel descargados
archivos_excel = glob.glob('*.xlsx')
lista_dataframes = []

for archivo in archivos_excel:
    startIndex = encontrar_fila_inicio(archivo)
    df = pd.read_excel(archivo, skiprows  = startIndex - 1)
    df.columns = df.columns.str.replace('\n', ' ')
    df.columns = df.columns.str.replace('  ', ' ')
    lista_dataframes.append(df)

#union de todos los archivos en un solo dataframe
df_concatenado = pd.concat(lista_dataframes)


#Columnas no reelevantes
lista_columnas_a_eliminar = ['Código de la Institución',
                             'IES PADRE',
                             'ID Sector IES',
                             'ID Caracter',
                             'Código del departamento (IES)',
                             'Código del Municipio (IES)',
                             'Municipio de domicilio de la IES',
                             'Código SNIES del programa',
                             'ID Nivel Académico',
                             'ID Nivel de Formación',
                             'ID Metodología',
                             'ID Área',
                             'Id_Nucleo',
                             'Núcleo Básico del Conocimiento (NBC)',
                            'Código del Departamento (Programa)',
                             'Departamento de oferta del programa',
                             'Código del Municipio (Programa)',
                             'Municipio de oferta del programa',
                             'ID Sexo']

#columnas categóricas a utilizar
cols_cat = ['Principal o Seccional',
            'Sector IES',
           'Caracter IES',
           'Departamento de domicilio de la IES',
           'Nivel Académico',
           'Nivel de Formación',
           'Metodología',
           'Área de Conocimiento',
            'Sexo'           
           ]

#Columnas para aplicar formato
cols_format = ['Institución de Educación Superior (IES)',
               'Principal o Seccional',
                'Sector IES',
               'Caracter IES',
               'Departamento de domicilio de la IES',
               'Programa Académico',
               'Nivel Académico',
               'Nivel de Formación',
               'Metodología',
               'Área de Conocimiento',
                'Sexo'           
               ]

#Formateo de valores en columnas categóricas
for column in df_concatenado.columns:
    if column in cols_format:
        df_concatenado[column] = df_concatenado[column].apply(remove_accents)
        df_concatenado[column] = df_concatenado[column].str.lower()


#Impresion de graficas para analizar valores categóricos
fig, ax = plt.subplots(nrows=len(cols_cat), ncols=1, figsize=(10,40))
fig.subplots_adjust(hspace=5)

for i, col in enumerate(cols_cat):
    sns.countplot(x=col,data=df_concatenado,ax=ax[i])
    ax[i].set_title(col)
    ax[i].set_xticklabels(ax[i].get_xticklabels(),rotation=30)

#Correcion de valores categóricos
df_concatenado['Caracter IES'] = df_concatenado['Caracter IES'].str.replace('institucion universitaria/escuela tecnologica','institucion tecnologica',regex=False)

df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'bogota.*','bogota d.c.',regex=True)
df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'cundinam.*','cundinamarca',regex=True)
df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'guajira','la guajira',regex=False)
df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'norte de sa.*','norte de santander',regex=True)
df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'valle del c.*','valle del cauca',regex=True)
df_concatenado['Departamento de domicilio de la IES'] = df_concatenado['Departamento de domicilio de la IES'].str.replace(r'archipiela.*','san andres y providencia',regex=True)

df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace('tecnologico','tecnologica',regex=False)
df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace('universitario','universitaria',regex=False)
df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace(r'especializacion medico quirur.*','especializacion medico quirurgica',regex=True)
df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace(r'especializacion tecnico profe.*','especializacion tecnico profesional',regex=True)
df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace('no aplica','sin programa especifico',regex=False)
df_concatenado['Nivel de Formación'] = df_concatenado['Nivel de Formación'].str.replace(r'especializacion.*','especializacion',regex=True)

df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace(r'a distancia.*','virtual',regex=True)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace(r'distancia.*','virtual',regex=True)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace('presencial-virtual','hibrido',regex=False)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace('virtual-dual','hibrido',regex=False)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace('dual','hibrido',regex=False)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace('presencial-hibrido','hibrido',regex=False)
df_concatenado['Metodología'] = df_concatenado['Metodología'].str.replace('no aplica','sin programa especifico',regex=False)

df_concatenado['Área de Conocimiento'] = df_concatenado['Área de Conocimiento'].str.replace('ingenieria arquitectura urbanismo y afines','ingenieria, arquitectura, urbanismo y afines',regex=True)
df_concatenado['Área de Conocimiento'] = df_concatenado['Área de Conocimiento'].str.replace('economia administracion contaduria y afines','economia, administracion, contaduria y afines',regex=True)
df_concatenado['Área de Conocimiento'] = df_concatenado['Área de Conocimiento'].str.replace('agronomia veterinaria y afines','agronomia, veterinaria y afines',regex=True)
df_concatenado['Área de Conocimiento'] = df_concatenado['Área de Conocimiento'].str.replace('sin programa especifico','sin clasificar',regex=True)
df_concatenado['Área de Conocimiento'] = df_concatenado['Área de Conocimiento'].str.replace('no aplica','sin clasificar',regex=True)

df_concatenado['Sexo'] = df_concatenado['Sexo'].str.replace('femenino','mujer',regex=True)
df_concatenado['Sexo'] = df_concatenado['Sexo'].str.replace('masculino','hombre',regex=True)

#Eliminar columnas no necesarias
df_concatenado = df_concatenado.drop(lista_columnas_a_eliminar, axis=1)

#eliminar datos faltantes
df_concatenado.dropna(inplace=True)

#Eliminando registros repetidos

df_concatenado.drop_duplicates(inplace=True)

#Exportación a archivo excel
df_concatenado.to_excel('datos_limpios.xlsx',index=False)