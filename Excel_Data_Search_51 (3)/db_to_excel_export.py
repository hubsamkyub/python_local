import sqlite3
import pandas as pd

db_path = r"C:\Users\Admin\Desktop\번역\TW\TW_STRING_DB\unique_texts_250611_.db"
table_name = "unique_texts"   # 내보낼 테이블명
excel_path = 'output.xlsx'

conn = sqlite3.connect(db_path)
df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
conn.close()

df.to_excel(excel_path, index=False)
print(f"엑셀 파일로 저장 완료: {excel_path}")
