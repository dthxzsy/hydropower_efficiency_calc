# interpolation.py
import numpy as np
import pandas as pd

def interpolate_efficiency(water_level, data_list):
    if pd.isna(water_level):
        return np.nan

    data_list = sorted(data_list, key=lambda x: x[0])
    levels = np.array([x[0] for x in data_list])
    effs = np.array([x[1] for x in data_list])

    if len(levels) == 0 or water_level < levels.min() or water_level > levels.max():
        print(f" 有效水位 {water_level} 超出效率表范围")
        return np.nan

    idx = np.where(np.isclose(levels, water_level))[0]
    if len(idx) > 0:
        return effs[idx[0]]

    for i in range(len(levels) - 1):
        if levels[i] <= water_level <= levels[i + 1]:
            x0, x1 = levels[i], levels[i + 1]
            y0, y1 = effs[i], effs[i + 1]
            return y0 + (water_level - x0) * (y1 - y0) / (x1 - x0)

    return np.nan

