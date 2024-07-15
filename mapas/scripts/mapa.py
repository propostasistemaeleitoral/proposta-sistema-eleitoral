import pandas as pd
import geopandas as gpd
from shapely.geometry import mapping, shape, Polygon, MultiPolygon
import folium
import datetime
import os

#########################################################################################################################################################
# Painel de controlo

# Escolher um dos cenários (100, 113 ou 126)

# numero_circulos_uninominais=100

# numero_circulos_uninominais=113

numero_circulos_uninominais=126

#########################################################################################################################################################

caminho_ficheiro_input_excel = "../input/eleitores_2022.xlsx"
caminho_ficheiro_shapefile = "../2023_agregado_shapefile/2023_agregado.shp"

#########################################################################################################################################################

# Ajustar a precisão das geometrias
precisao = 6 

#########################################################################################################################################################

# Função para ajustar a precisão das geometrias
def ajustar_precisao(geometria, precisao):
    def arredondar_coordenadas(coords, precisao):
        return tuple(round(coord, precisao) for coord in coords)
    
    def processar_geometria(geom, precisao):
        if isinstance(geom, Polygon):
            exterior = [arredondar_coordenadas(ponto, precisao) for ponto in geom.exterior.coords]
            interiors = [[arredondar_coordenadas(ponto, precisao) for ponto in interior.coords] for interior in geom.interiors]
            return Polygon(exterior, interiors)
        elif isinstance(geom, MultiPolygon):
            return MultiPolygon([processar_geometria(parte, precisao) for parte in geom.geoms])
        return geom
    
    return processar_geometria(geometria, precisao)


#########################################################################################################################################################

df = pd.read_excel(caminho_ficheiro_input_excel, sheet_name=f"{numero_circulos_uninominais}")
df['dico'] = df['dicofre'].str[:4]
df['dico'] = df['dico'].astype(str)

#########################################################################################################################################################

df_concelhos_agregados = df.groupby('dico').agg({
    'eleitores': 'sum', 
    'distrito_ilha': 'first', 
    'concelho': 'first', 
    'freguesia': 'first'
}).reset_index()

gdf = gpd.read_file(caminho_ficheiro_shapefile)
gdf['dico'] = gdf['DICOFRE'].str.slice(0, 4)
gdf['dico'] = gdf['dico'].astype(str)
gdf['dicofre'] = gdf['DICOFRE'].astype(str)

# Combinar o DataFrame com o GeoDataFrame
merged_gdf = gdf.merge(df_concelhos_agregados, on='dico')
merged_gdf['geometry'] = merged_gdf['geometry'].apply(lambda geom: ajustar_precisao(geom, precisao))

# Disolver por concelhos 
dissolved_gdf_concelho = merged_gdf.dissolve(by='dico', as_index=False)
# Converter para coordenadas
dissolved_gdf_wgs84_concelho = dissolved_gdf_concelho.to_crs(epsg=4326)

#########################################################################################################################################################

# Combinar o DataFrame com o GeoDataFrame
merged_gdf = gdf.merge(df, on='dicofre')
merged_gdf['geometry'] = merged_gdf['geometry'].apply(lambda geom: ajustar_precisao(geom, precisao))

#  Disolver por circulos atuais 
dissolved_gdf_circulo = merged_gdf.dissolve(by='circulo_atual', as_index=False)

# Converter para coordenadas
dissolved_gdf_wgs84_circulo = dissolved_gdf_circulo.to_crs(epsg=4326)

#########################################################################################################################################################

df_freguesias = pd.read_excel(caminho_ficheiro_input_excel, sheet_name=f"{numero_circulos_uninominais}")

if numero_circulos_uninominais == 100:

    lista_concelho_dividir=[
    '0109 Santa Maria da Feira',
    # '0302 Barcelos',
    '0303 Braga',
    '0308 Guimarães',
    # '0312 Vila Nova de Famalicão',
    '0603 Coimbra',
    # '1009 Leiria',
    '1105 Cascais',
    '1106 Lisboa',
    '1107 Loures',
    '1110 Oeiras',
    '1111 Sintra',
    # '1114 Vila Franca de Xira',
    '1115 Amadora',
    '1116 Odivelas',
    '1304 Gondomar',
    # '1306 Maia',
    '1308 Matosinhos',
    '1312 Porto',
    '1317 Vila Nova de Gaia',
    '1503 Almada',
    '1510 Seixal',
    # '1512 Setúbal',
    # '3103 Funchal'
    ]

elif numero_circulos_uninominais == 113:

    lista_concelho_dividir=[
    '0109 Santa Maria da Feira',
    # '0302 Barcelos',
    '0303 Braga',
    '0308 Guimarães',
    '0312 Vila Nova de Famalicão',
    '0603 Coimbra',
    '1009 Leiria',
    '1105 Cascais',
    '1106 Lisboa',
    '1107 Loures',
    '1110 Oeiras',
    '1111 Sintra',
    '1114 Vila Franca de Xira',
    '1115 Amadora',
    '1116 Odivelas',
    '1304 Gondomar',
    '1306 Maia',
    '1308 Matosinhos',
    '1312 Porto',
    '1317 Vila Nova de Gaia',
    '1503 Almada',
    '1510 Seixal',
    # '1512 Setúbal',
    # '3103 Funchal'
    ]

elif numero_circulos_uninominais == 126:

    lista_concelho_dividir=[
    '0109 Santa Maria da Feira',
    '0302 Barcelos',
    '0303 Braga',
    '0308 Guimarães',
    '0312 Vila Nova de Famalicão',
    '0603 Coimbra',
    '1009 Leiria',
    '1105 Cascais',
    '1106 Lisboa',
    '1107 Loures',
    '1110 Oeiras',
    '1111 Sintra',
    '1114 Vila Franca de Xira',
    '1115 Amadora',
    '1116 Odivelas',
    '1304 Gondomar',
    '1306 Maia',
    '1308 Matosinhos',
    '1312 Porto',
    '1317 Vila Nova de Gaia',
    '1503 Almada',
    '1510 Seixal',
    '1512 Setúbal',
    '3103 Funchal'
    ]

# Filtrar o  DataFrame
df_freguesias = df_freguesias[df_freguesias['concelho'].isin(lista_concelho_dividir)]
# Combinar o DataFrame com o  GeoDataFrame, para freguesias filtradas
merged_gdf_freguesias = gdf.merge(df_freguesias, on='dicofre')
# Converter para coordenadas
merged_gdf_wgs84_freguesias = merged_gdf_freguesias.to_crs(epsg=4326)

#########################################################################################################################################################

df_circulo_uninominal = pd.read_excel(caminho_ficheiro_input_excel, sheet_name=f"{numero_circulos_uninominais}")

gdf_circulo_uninominal = gpd.read_file(caminho_ficheiro_shapefile)
gdf_circulo_uninominal['dicofre'] = gdf_circulo_uninominal['DICOFRE']

# Combinar o DataFrame com o  GeoDataFrame
merged_gdf_circulo_uninominal = gdf_circulo_uninominal.merge(df_circulo_uninominal, on='dicofre')
merged_gdf_circulo_uninominal['geometry'] = merged_gdf_circulo_uninominal['geometry'].apply(lambda geom: ajustar_precisao(geom, precisao))

#  Disolver por circulos uninominais e somar os eleitores
dissolved_gdf_circulo_uninominal = merged_gdf_circulo_uninominal.dissolve(
    by='circulo_uninominal',
    as_index=False,
    aggfunc={'eleitores': 'sum'}
)

# Converter para coordenadas
dissolved_gdf_wgs84_circulo_uninominal = dissolved_gdf_circulo_uninominal.to_crs(epsg=4326)

#########################################################################################################################################################

map_center = dissolved_gdf_wgs84_concelho.geometry.centroid.iloc[0].coords[:][0]
m = folium.Map(location=[map_center[1], map_center[0]], zoom_start=7)

#########################################################################################################################################################

# Para layer da distribuição, por circulo uninominal
choropleth = folium.Choropleth(
    geo_data=dissolved_gdf_wgs84_circulo_uninominal.to_json(),
    data=dissolved_gdf_wgs84_circulo_uninominal,
    columns=['circulo_uninominal', 'eleitores'],
    aliases=['Círculo uninominal', 'Eleitores:'],
    key_on='feature.properties.circulo_uninominal',
    fill_color='YlOrRd',
    fill_opacity=0.7,
    line_opacity=0, # já tenho a layer  (dos círculos uninominais) para isto
    legend_name='Eleitores',
    name='Número e densidade de eleitores, por círculo uninominal'
).add_to(m)

choropleth.geojson.add_child(
    folium.features.GeoJsonTooltip(['circulo_uninominal', 'eleitores'], labels=True)
)

#########################################################################################################################################################

# Para layer da distribuição, por concelho
geojson_layer = folium.GeoJson(
    data=dissolved_gdf_wgs84_concelho.to_json(),
    name='Número de eleitores e limites, por concelho',  
    show=False,
    style_function=lambda x: {
        'fillColor': 'transparent',  
        'color': 'blue', 
        'weight': 1, 
        'fillOpacity': 0,  
        'opacity': 2 
    },
    tooltip=folium.features.GeoJsonTooltip(
        fields=['concelho', 'eleitores'],
        aliases=['Concelho:', 'Eleitores:'],  
        localize=True,
        sticky=False, 
    )
)

geojson_layer.add_to(m)

#########################################################################################################################################################

# Para layer da distribuição, por freguesia
geojson_layer_fregusias = folium.GeoJson(
    data=merged_gdf_wgs84_freguesias.to_json(),
    name='Número de eleitores e limites, por freguesia (apenas para concelhos dividdos)', 
    show=False,
    style_function=lambda x: {
        'fillColor': 'transparent',  
        'color': 'white',  
        'weight': 1,  
        'fillOpacity': 0,  
        'opacity': 1  
    },
    tooltip=folium.features.GeoJsonTooltip(
        fields=['freguesia', 'eleitores'],
        aliases=['Freguesia:', 'Eleitores:'],  
        localize=True,
        sticky=False, 
    )
)

geojson_layer_fregusias.add_to(m)

#########################################################################################################################################################

# Limites dos círculos uninominais
boundary_group = folium.FeatureGroup(name='Limites dos círculos uninominais', show=True)

folium.GeoJson(
    dissolved_gdf_wgs84_circulo_uninominal,
    style_function=lambda x: {'color': 'green', 'weight': 5, 'fillOpacity': 0, 'interactive': False}
).add_to(boundary_group)

boundary_group.add_to(m)

#########################################################################################################################################################

# Limites dos círculos atuais
district_boundaries_group = folium.FeatureGroup(name='Limites dos círculos atuais', show=True)
folium.GeoJson(
    dissolved_gdf_wgs84_circulo,
    style_function=lambda x: {'color': 'black', 'weight': 10, 'fillOpacity': 0, 'interactive': False}
).add_to(district_boundaries_group)
district_boundaries_group.add_to(m)

#########################################################################################################################################################

#  Criar menu para adicionar/retirar camadas
folium.LayerControl().add_to(m)

#########################################################################################################################################################

momento_agora = datetime.datetime.now()
ficheiro_output = f"mapa_{numero_circulos_uninominais}_uninominais_{momento_agora.strftime('%Y')}_{momento_agora.strftime('%m')}_{momento_agora.strftime('%d')}_{momento_agora.strftime('%H')}_{momento_agora.strftime('%M')}_{momento_agora.strftime('%S')}.xlsx"

#########################################################################################################################################################

m.save(f'../output/{ficheiro_output}.html')

print(f"Ficheiro {ficheiro_output} guardado")