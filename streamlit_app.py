
import pandas as pd
import numpy as np

# 1. Dosyaları Yükleme (Kendi dosya yollarınıza göre düzenleyin)
sell_in_df = pd.read_excel("sell_in_verisi.xlsx")
alloc_df1 = pd.read_excel("agustos_1_alokasyon.xlsx")
alloc_df2 = pd.read_excel("agustos_2_alokasyon.xlsx")
master_df = pd.read_excel("customer_master.xlsx")

print("Dosyalar başarıyla yüklendi.")

# 2. Kolon adı bulma yardımcı fonksiyonu
def find_col(df, keywords):
    for col in df.columns:
        col_lower = str(col).lower()
        if any(kw in col_lower for kw in keywords):
            return col
    return df.columns[0]

# Sell-in kolonları
s_cust = find_col(sell_in_df, ['customer', 'musteri', 'unvan', 'company'])
s_ean = find_col(sell_in_df, ['ean', 'barkod', 'material', 'sku'])
s_desc = find_col(sell_in_df, ['desc', 'tanim', 'text'])
s_inv = find_col(sell_in_df, ['invcd', 'fatura', 'qty'])
s_del = find_col(sell_in_df, ['delnot', 'sevk', 'delivery'])

# Sell-in miktarını toplama (Invcd + del.not.in)
sell_in_df['Sell_In_Qty'] = pd.to_numeric(sell_in_df[s_inv], errors='coerce').fillna(0) + \
                            pd.to_numeric(sell_in_df[s_del], errors='coerce').fillna(0)

# 3. 1. ve 2. Alokasyon Dosyalarını Birleştirme
alloc_combined = pd.concat([alloc_df1, alloc_df2], ignore_index=True)

# İstenmeyen kolonları eleme
ignored_keywords = ['brand', 'description', 'file', 'unnamed', 'urun', 'kod', 'ean', 'desc', 'total', 'toplam', 'not', 'ofis', 'sip', 'pr']
valid_alloc_cols = [col for col in alloc_combined.columns if not any(kw in str(col).lower() for kw in ignored_keywords)]

# 4. Master Hiyerarşi Kuralları
# Master dosyasındaki kolonları tespit et
m_cust = find_col(master_df, ['customer', 'musteri'])
m_hie2_val = find_col(master_df, ['custhier2lev', 'hie2'])
m_hie2_name = find_col(master_df, ['namecusthier2lev', 'hier2'])
m_hie3_name = find_col(master_df, ['namecusthier3lev', 'hier3'])
m_hie4_name = find_col(master_df, ['namecusthier4lev', 'hier4'])

def get_hierarchy_name(row):
    row_str = " ".join(str(val).lower() for val in row.values)
    hie2_val = str(row.get(m_hie2_val, '')).lower()
    
    if 'distributor' in hie2_val:
        return row.get(m_hie2_name, row.get(m_cust))
    elif 'cash' in row_str or 'rka' in row_str:
        return row.get(m_hie3_name, row.get(m_cust))
    else:
        return row.get(m_hie4_name, row.get(m_cust))

master_df['Matched_Hier'] = master_df.apply(get_hierarchy_name, axis=1)

# Müşteri eşleme sözlüğü oluşturma
mapping_dict = {}
for col in valid_alloc_cols:
    col_norm = str(col).strip().lower()
    match = master_df[master_df[m_cust].astype(str).str.lower().str.contains(col_norm, na=False)]
    if not match.empty:
        mapping_dict[col] = match['Matched_Hier'].values[0]
    else:
        mapping_dict[col] = col # Bulunamazsa kendi adını koru

# 5. Ana Rapor Oluşturma ve Hesaplamalar
report_rows = []

for _, row in sell_in_df.iterrows():
    cust = row[s_cust]
    ean = row[s_ean]
    desc = row[s_desc]
    sell_in_qty = row['Sell_In_Qty']
    
    # Alokasyon miktarını bul
    alok_qty = 0
    for col, hier_name in mapping_dict.items():
        if col in alloc_combined.columns:
            # EAN eşleşmesine göre filtrele
            filtered = alloc_combined[alloc_combined.apply(lambda r: str(ean) in str(r.values), axis=1)] if s_ean in alloc_combined.columns else alloc_combined
            alok_qty += pd.to_numeric(filtered[col], errors='coerce').fillna(0).sum()

    if alok_qty == 0:
        alok_qty = 0 # Sıfır kontrolü

    abs_ga = abs(alokQty - sell_in_qty) if 'alokQty' in locals() else abs(alok_qty - sell_in_qty)
    
    # Realizasyon ve Accuracy
    realization_ratio = (sell_in_qty / alok_qty * 100) if alok_qty > 0 else 0
    accuracy = max(0, 1 - (abs_ga / sell_in_qty)) if sell_in_qty > 0 else 0
    
    # Yorum / Durum
    if realization_ratio > 100:
        yorum = "On Top Adet verilmiş"
    elif realization_ratio == 100:
        yorum = "100%"
    elif realization_ratio == 0:
        yorum = "0%"
    elif realization_ratio >= 30:
        yorum = "%30-%100"
    else:
        yorum = "<%30"

    report_rows.append({
        "Unique": f"{mapping_dict.get(cust, cust)}{ean}",
        "EAN": ean,
        "DESC": desc,
        "CUSTOMER": mapping_dict.get(cust, cust),
        "SELL IN": sell_in_qty,
        "ALOK": alok_qty,
        "ABS GA": abs_ga,
        "% REALIZATION": f"{round(realization_ratio)}%",
        "ACCURACY": f"{round(accuracy * 100, 1)}%",
        "YORUM": yorum
    })

report_df = pd.DataFrame(report_rows)

# 6. Excel Olarak Çıktı Alma
output_filename = "Alokasyon_Accuracy_Raporu_Python.xlsx"
report_df.to_excel(output_filename, index=False)
print(f"Rapor başarıyla oluşturuldu ve kaydedildi: {output_filename}")
