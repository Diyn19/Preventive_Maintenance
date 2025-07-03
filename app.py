import pandas as pd
import requests
from io import BytesIO
from flask import Flask, render_template, request, abort
import os

app = Flask(__name__)

GITHUB_XLSX_URL = 'https://raw.githubusercontent.com/Diyn19/data/master/data.xlsx'

def load_excel_from_github(url):
    try:
        response = requests.get(url, timeout=5)
        content_type = response.headers.get('Content-Type', '')
        if response.status_code == 200 and ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type or url.endswith('.xlsx')):
            return pd.ExcelFile(BytesIO(response.content), engine='openpyxl')
        else:
            print(f"❌ Excel 下載失敗：{response.status_code} - {content_type}")
    except Exception as e:
        print(f"❌ 錯誤下載 Excel: {e}")
    abort(500, description="⚠️ 無法從 GitHub 載入 Excel 檔案")

def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    return df.fillna('')

@app.route('/')
def index():
    xls = load_excel_from_github(GITHUB_XLSX_URL)

    df_department = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:F", skiprows=4, nrows=1))
    df_seasons = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:D", skiprows=8, nrows=2))
    df_project1 = clean_df(pd.read_excel(xls, sheet_name='首頁', usecols="A:E", skiprows=12, nrows=3))
    df = clean_df(pd.read_excel(xls, sheet_name=0, header=13, nrows=250, usecols="A:O"))
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df.empty

    return render_template(
        'index.html',
        tables=df.to_dict(orient='records'),
        keyword=keyword,
        store_id='',
        repair_item='',
        personal_page=False,
        report_page=False,
        department_table=df_department.to_dict(orient='records'),
        seasons_table=df_seasons.to_dict(orient='records'),
        project1_table=df_project1.to_dict(orient='records'),
        no_data_found=no_data_found,
    )

@app.route('/<name>')
def personal(name):
    sheet_map = {
        '吳宗鴻': '吳宗鴻',
        '湯家瑋': '湯家瑋',
        '狄澤洋': '狄澤洋'
    }
    sheet_name = sheet_map.get(name)
    if not sheet_name:
        return f"找不到{name}的分頁", 404

    xls = load_excel_from_github(GITHUB_XLSX_URL)

    df_top = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:G", nrows=4))
    df_top = df_top.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_project = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="H:L", nrows=3))
    df_project = df_project.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_bottom = clean_df(pd.read_excel(xls, sheet_name=sheet_name, usecols="A:J", skiprows=5))

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df_bottom = df_bottom[df_bottom.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df_bottom.empty

    return render_template(
        'index.html',
        tables_top=df_top.to_dict(orient='records'),
        tables_project=df_project.to_dict(orient='records'),
        tables_bottom=df_bottom.to_dict(orient='records'),
        keyword=keyword,
        store_id='',
        repair_item='',
        personal_page=True,
        report_page=False,
        no_data_found=no_data_found,
        show_top=True,
        show_project=True
    )

@app.route('/report')
def report():
    keyword = request.args.get('keyword', '').strip()
    store_id = request.args.get('store_id', '').strip()
    repair_item = request.args.get('repair_item', '').strip()
    no_data_found = False
    tables = []

    if keyword or store_id or repair_item:
        xls = load_excel_from_github(GITHUB_XLSX_URL)
        df = clean_df(pd.read_excel(xls, sheet_name='IM'))
        df = df[['案件類別', '門店編號', '門店名稱', '報修時間', '報修類別', '報修項目', '報修說明', '設備號碼', '服務人員', '工作內容']]

        if keyword:
            df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

        if store_id:
            df = df[df['門店編號'].astype(str).str.contains(store_id, case=False)]

        if repair_item:
            df = df[df['報修類別'].astype(str).str.strip() == repair_item.strip()]

        if df.empty:
            no_data_found = True
        else:
            tables = df.to_dict(orient='records')

    return render_template(
        'index.html',
        tables=tables,
        keyword='',
        store_id=store_id,
        repair_item=repair_item,
        personal_page=False,
        report_page=True,
        no_data_found=no_data_found
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
