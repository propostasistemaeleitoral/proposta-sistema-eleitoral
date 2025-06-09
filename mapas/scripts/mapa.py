import pandas as pd
import geopandas as gpd
from shapely.geometry import mapping, shape, Polygon, MultiPolygon
import folium
import datetime
import os

caminho_ficheiro_input_excel = "../input/dados/eleitores_2022.xlsx"

caminho_ficheiro_shapefile_par = "../input/shapefiles/2022/1_par/par.shp"
caminho_ficheiro_shapefile_mun = "../input/shapefiles/2022/2_mun/mun.shp"
caminho_ficheiro_shapefile_dis = "../input/shapefiles/2022/3_dis/dis.shp"

lista_numero_circulos_uninominais =[100, 113, 126]

count = 0
for numero_circulos_uninominais in lista_numero_circulos_uninominais:

    #########################################################################################################################################################

    df_par = pd.read_excel(caminho_ficheiro_input_excel, sheet_name=f"{numero_circulos_uninominais}")
    df_par['dico'] = df_par['dicofre'].str[:4]
    df_par['dico'] = df_par['dico'].astype(str)


    gdf_par = gpd.read_file(caminho_ficheiro_shapefile_par)
    gdf_mun = gpd.read_file(caminho_ficheiro_shapefile_mun)
    gdf_dis = gpd.read_file(caminho_ficheiro_shapefile_dis)

    #########################################################################################################################################################
    # Simplificar Shapefile
    gdf_dis['geometry'] = gdf_dis['geometry'].simplify(tolerance=15, preserve_topology=True) 
    # Converter para coordenadas
    gdf_dis_coord = gdf_dis.to_crs(epsg=4326)
    #########################################################################################################################################################
    # Concelhos
    df_concelhos_agregados = df_par.groupby('dico').agg({
        'eleitores': 'sum', 
        'distrito_ilha': 'first', 
        'concelho': 'first', 
    }).reset_index()

    gdf_mun["c_mun_f"] = gdf_mun["c_mun"].astype(str)

    # Combinar o DataFrame com o GeoDataFrame
    merged_gdf_mun = gdf_mun.merge(df_concelhos_agregados, left_on='c_mun', right_on='dico')

    # Simplificar Shapefile
    merged_gdf_mun['geometry'] = merged_gdf_mun['geometry'].simplify(tolerance=15, preserve_topology=True)  
    # Converter para coordenadas
    merged_gdf_mun_coord = merged_gdf_mun.to_crs(epsg=4326)
    #########################################################################################################################################################
    # Freguesias
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
    df_par_filtrado = df_par[df_par['concelho'].isin(lista_concelho_dividir)]
    # Combinar o DataFrame com o  GeoDataFrame, para freguesias filtradas
    merged_gdf_freguesias_filtrado = gdf_par.merge(df_par_filtrado, left_on='c_par', right_on='dicofre')

    # Simplificar Shapefile
    merged_gdf_freguesias_filtrado['geometry'] = merged_gdf_freguesias_filtrado['geometry'].simplify(tolerance=15, preserve_topology=True)  
    # Converter para coordenadas
    merged_gdf_freguesias_filtrado_coord = merged_gdf_freguesias_filtrado.to_crs(epsg=4326)
    #########################################################################################################################################################
    # Círculos uninominais

    merged_gdf_circulo_uninominal = gdf_par.merge(df_par, left_on='c_par', right_on='dicofre')

    #  Disolver por circulos uninominais e somar os eleitores
    dissolved_gdf_circulo_uninominal = merged_gdf_circulo_uninominal.dissolve(
        by='circulo_uninominal',
        as_index=False,
        aggfunc={'eleitores': 'sum'}
    )

    # Simplificar Shapefile
    dissolved_gdf_circulo_uninominal['geometry'] = dissolved_gdf_circulo_uninominal['geometry'].simplify(tolerance=15, preserve_topology=True)  
    # Converter para coordenadas
    dissolved_gdf_circulo_uninominal_coord = dissolved_gdf_circulo_uninominal.to_crs(epsg=4326)

    #########################################################################################################################################################
########################################################################################################################################################
 
    filtered_gdf_concelho_f = merged_gdf_mun_coord[merged_gdf_mun_coord['concelho'] == '0510 Vila de Rei']

    map_center = filtered_gdf_concelho_f.geometry.centroid.iloc[0].coords[:][0]

    m = folium.Map(location=[map_center[1], map_center[0]], zoom_start=7)

    #########################################################################################################################################################

    # Para layer da distribuição, por circulo uninominal



    choropleth = folium.Choropleth(
        geo_data=dissolved_gdf_circulo_uninominal_coord.to_json(),
        data=dissolved_gdf_circulo_uninominal_coord,
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
    folium.features.GeoJsonTooltip(
        fields=['circulo_uninominal', 'eleitores'],
        aliases=['Círculo uninominal:', 'Eleitores:'],
        labels=True
    )
    )

    #########################################################################################################################################################

    # Para layer da distribuição, por concelho
    geojson_layer = folium.GeoJson(
        data=merged_gdf_mun_coord.to_json(),
        name='Número de eleitores e delimitação, por concelho',  
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
        data=merged_gdf_freguesias_filtrado_coord.to_json(),
        name='Número de eleitores e delimitação, por freguesia (apenas para concelhos divididos)', 
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
    boundary_group = folium.FeatureGroup(name='Delimitação dos círculos uninominais', show=True)

    folium.GeoJson(
        dissolved_gdf_circulo_uninominal_coord,
        style_function=lambda x: {'color': 'green', 'weight': 4, 'fillOpacity': 0, 'interactive': False}
    ).add_to(boundary_group)

    boundary_group.add_to(m)

    #########################################################################################################################################################

    # Limites dos círculos atuais
    district_boundaries_group = folium.FeatureGroup(name='Delimitação dos círculos atuais', show=True)
    folium.GeoJson(
        gdf_dis_coord,
        style_function=lambda x: {'color': 'black', 'weight': 4, 'fillOpacity': 0, 'interactive': False}
    ).add_to(district_boundaries_group)
    district_boundaries_group.add_to(m)

    #########################################################################################################################################################

    #  Criar menu para adicionar/retirar camadas
    folium.LayerControl().add_to(m)

    #########################################################################################################################################################

    momento_agora = datetime.datetime.now()
    ficheiro_output = f"mapa_{numero_circulos_uninominais}_uninominais_{momento_agora.strftime('%Y')}_{momento_agora.strftime('%m')}_{momento_agora.strftime('%d')}_{momento_agora.strftime('%H')}_{momento_agora.strftime('%M')}_{momento_agora.strftime('%S')}"

    #########################################################################################################################################################

    m.save(f'../output/{ficheiro_output}.html')

    print(f"Ficheiro {ficheiro_output} guardado")