import pandas as pd
def load_data(file_name):
    return pd.read_excel("Libro4.xlsx")

def mask_dataframe_single_date(dataframe,date):
    mask = (dataframe['Fecha Mvto'] == date) 
    return dataframe.loc[mask]

def cantidad_de_repartos(df):
    repartos = df['Número']
    return len(repartos.drop_duplicates())

def dataframe_cleaning(df):
    df_less_cols=df[['Fecha Mvto','Artículo','Descripción Artículo','Bultos','Unids']]
    articulos = df_less_cols[['Artículo','Descripción Artículo','Fecha Mvto']]
    articulos_unicos = articulos.drop_duplicates(subset="Artículo")
    return df_less_cols,articulos_unicos

def suma_y_merge(df_datos,df_articulos):
    df_suma = df_datos.groupby('Artículo').sum()
    df_final = df_suma.merge(df_articulos, on='Artículo')
    return df_final[['Fecha Mvto','Artículo','Descripción Artículo','Bultos','Unids']]

def dataframe_repartos(df):
    repartos = df[['Número', 'Descripción Transporte']]
    return repartos.drop_duplicates(subset='Número')