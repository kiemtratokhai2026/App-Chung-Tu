import pandas as pd
import os

_mapping_cache = None

def load_mapping():
    global _mapping_cache
    if _mapping_cache is None:
        if os.path.exists('MAPPING.xlsx'):
            df = pd.read_excel('MAPPING.xlsx', dtype=str)
            # key = (norm_model, brand)
            _mapping_cache = {}
            for _, row in df.iterrows():
                brand = str(row['BRAND']).strip().upper()
                model = str(row['MODEL']).strip().upper()
                key = f"{model}|{brand}"
                _mapping_cache[key] = {
                    'internal_code': str(row['INTERNAL CODE']).strip() if pd.notna(row['INTERNAL CODE']) else "",
                    'size_cm': float(row['SIZE (CM)']) if pd.notna(row['SIZE (CM)']) else 0.0,
                    'hs_code': str(row['HS CODE']).strip().replace('.0', '') if pd.notna(row['HS CODE']) else "",
                    'ten_hang': str(row['TEN HANG (VN)']).strip() if pd.notna(row['TEN HANG (VN)']) else ""
                }
        else:
            _mapping_cache = {}
    return _mapping_cache

def lookup_mapping(model, brand):
    mapping = load_mapping()
    key = f"{model.upper()}|{brand.upper()}"
    return mapping.get(key, None)

if __name__ == '__main__':
    res = lookup_mapping("SC52", "MICHELIN")
    print("SC52 MICHELIN:", res)
    res2 = lookup_mapping("RECMF-8300", "TAKA PRO")
    print("RECMF-8300 TAKA PRO:", res2)
