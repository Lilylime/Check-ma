import pandas as pd
from fuzzywuzzy import fuzz

# Đọc file Excel
sheets = pd.read_excel(
    r'C:\Users\Acer Nitro 5\Desktop\check ma.xlsx',  sheet_name=None, engine='openpyxl')

# Kiểm tra danh sách sheet
print("Các sheet trong file:", sheets.keys())

# Truy cập từng sheet theo tên:
df1 = pd.DataFrame(sheets['All'])  # thay 'Sheet1' bằng tên thực tế

# Đảm bảo các cột tồn tại sẵn và đặt đúng tên
for col in ['Fix Company name', 'Fix Code']:
    if col not in df1.columns:
        df1[col] = ""
    df1[col] = df1[col].astype(str)  # ép kiểu rõ ràng sang string

if 'Same' not in df1.columns:
    df1['Same'] = 0.0  # cột này là số, để kiểu float

assigned = set()

for i in range(len(df1)):
    if i in assigned:
        continue

    name_i = str(df1.loc[i, 'Ten']).strip()
    df1.at[i, 'Fix Company name'] = name_i
    df1.at[i, 'Same'] = 1  # chính nó
    df1.at[i, 'Fix Code'] = str(df1.at[i, 'Code'])

    for j in range(i + 1, len(df1)):
        if j in assigned:
            continue

        name_j = str(df1.loc[j, 'Ten']).strip()
        score = fuzz.partial_ratio(name_i.lower(), name_j.lower())

        if score >= 70:
            df1.at[j, 'Fix Company name'] = name_i
            df1.at[j, 'Same'] = round(score / 100, 2)
            df1.at[j, 'Fix Code'] = str(df1.at[i, 'Code'])
            assigned.add(j)

    assigned.add(i)

# Ghi kết quả vào sheet mới
with pd.ExcelWriter(r'C:\Users\Acer Nitro 5\Desktop\check ma.xlsx',
                    engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df1.to_excel(writer, sheet_name='Checking ', index=False)
# thêm mới
