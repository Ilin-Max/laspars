import numpy as np
import warnings
import lasio
import pandas as pd
from time import sleep

def linear_interpolation(x1, y1, x2, y2, x):
    y = y1 + ((y2 - y1) / (x2 - x1)) * (x - x1)
    return y

def find_value_GIS_in_diskret_depth(kern_md, gis_md, gis_value, round = 2):
    gis_value_in_depth_kern = np.empty(len(kern_md))
    
    if np.max(kern_md) > np.max(gis_md) or np.min(kern_md) < np.min(gis_md):
        warnings.warn(f"Глубины отбора керна выходят за границы глубин, проведения ГИС", UserWarning)
    
    for i in range(len(kern_md)):
        if np.min(gis_md) <= kern_md[i] <= np.max(gis_md):
            
            if kern_md[i] < np.max(gis_md):
                low_border_index = np.where(gis_md <= kern_md[i])[0][-1]
                top_border_index = low_border_index + 1
                x1 = gis_md[low_border_index]
                x2 = gis_md[top_border_index]
                y1 = gis_value[low_border_index]
                y2 = gis_value[top_border_index]
                x = kern_md[i]
            else:
                top_border_index = np.where(gis_md >= kern_md[i])[0][0]
                low_border_index = top_border_index - 1
                x1 = gis_md[low_border_index]
                x2 = gis_md[top_border_index]
                y1 = gis_value[low_border_index]
                y2 = gis_value[top_border_index]
                x = kern_md[i]
            
            gis_value_in_depth_kern[i] = np.round(linear_interpolation(x1, y1, x2, y2, x), round)
        
        else:
            gis_value_in_depth_kern[i] = np.nan
        
    return gis_value_in_depth_kern

def find_index_low_top_border(depth_array, depth_to_search):
    low_border_index = np.where(depth_array <= depth_to_search)[0][-1]
    top_border_index = low_border_index + 1
    return low_border_index, top_border_index

def make_DataFrame_with_GIS_value_in_diskret_depth(exel_file, las_file):
    df = pd.read_excel(exel_file)
    depth_to_search_array = np.array(df[df.keys()[0]])

    las = lasio.LASFile(las_file)
    depth_curves_array = np.array(las.index)
    curves_name = las.curves.keys()
    data = np.ones((len(curves_name[1:]), len(depth_to_search_array)))

    for i in range(0, len(depth_to_search_array)):
        low_border_index, top_border_index = find_index_low_top_border(depth_curves_array, depth_to_search_array[i])
        for j in range(1, len(las.curves)):
            x1 = depth_curves_array[low_border_index]
            x2 = depth_curves_array[top_border_index]
            y1 = las.curves[j].data[low_border_index]
            y2 = las.curves[j].data[top_border_index]
            x = depth_to_search_array[i]
            data[j - 1, i] = linear_interpolation(x1, y1, x2, y2, x)

    output_data = {df.keys()[0]: depth_to_search_array}

    for i in range(len(data)):
        output_data[las.curves[i+1].mnemonic] = data[i]

    return output_data

def make_exel_with_GIS_value_in_diskret_depth(exel_file, las_file, name_output_file):
    output_data = make_DataFrame_with_GIS_value_in_diskret_depth(exel_file, las_file)
    output_df = pd.DataFrame(output_data)
    if len(name_output_file) == 0:
        name_output_file = "output_file"
    output_df.to_excel(f"{name_output_file}.xlsx", index=False)
    

exel_file_path = input("Введи адрес файла exel с глубинами керна:\n")
las_file_path = input("Введи адрес файла las из которого необходимо вытащить значения кривых:\n")
name_output_file = input("Введи название нового файла exel в котором все сохранится:\n")

make_exel_with_GIS_value_in_diskret_depth(exel_file_path, las_file_path, name_output_file)
print(f"Файл {name_output_file}.xlsx сохранен")

sleep(5)






