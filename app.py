import pandas as pd
from flask import Flask, render_template, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

app = Flask(__name__)

# Google Sheets 認證與讀取函式
def get_sheet(sheet_name, header=0, skiprows=None, usecols=None, nrows=None):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # 從環境變數讀金鑰JSON字串
    key_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
    if not key_json:
        raise Exception("環境變數 GOOGLE_CREDENTIALS_JSON 未設定或為空")

    creds_dict = json.loads(key_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

    client = gspread.authorize(creds)

    # 你原本的這行有誤，worksheet() 裡面要傳字串 sheet_name 而不是 data
    sheet = client.open_by_key('13BpLAGfdFyiXno69hFTkhHxkqyIE_8PiFuek8M4UnzU').worksheet(sheet_name)

    data = sheet.get_all_values()
    df = pd.DataFrame(data)

    if header is not None:
        df.columns = df.iloc[header]
        df = df.drop(index=list(range(0, header + 1)))
    if skiprows:
        df = df.iloc[skiprows:]
    if usecols:
        df = df.iloc[:, pd.Index(usecols)]
    if nrows:
        df = df.head(nrows)
    df = df.fillna('')
    return df.reset_index(drop=True)
    
# 清理資料
def clean_df(df):
    df.columns = df.columns.astype(str).str.replace('\n', '', regex=False)
    df = df.fillna('')
    return df

@app.route('/')
def index():
    df_department = get_sheet('首頁', header=4, nrows=1)
    df_department = clean_df(df_department)

    df_seasons = get_sheet('首頁', header=8, nrows=2)
    df_seasons = clean_df(df_seasons)

    df_project1 = get_sheet('首頁', header=12, nrows=3)
    df_project1 = clean_df(df_project1)

    df = get_sheet('首頁', header=13, nrows=250)
    df = clean_df(df)
    df = df[['門市編號', '門市名稱', 'PMQ_檢核', '專案檢核', 'HUB', '完工檢核']]

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    tables = df.to_dict(orient='records')

    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        if df.empty:
            no_data_found = True
        tables = df.to_dict(orient='records')

    return render_template(
        'index.html',
        tables=tables,
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

    df_top = get_sheet(sheet_name, header=0, nrows=4)
    df_top = clean_df(df_top)
    df_top = df_top.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_project = get_sheet(sheet_name, header=0, usecols=range(7, 12), nrows=3)
    df_project = clean_df(df_project)
    df_project = df_project.applymap(lambda x: int(x) if isinstance(x, (int, float)) and x == int(x) else x)

    df_bottom = get_sheet(sheet_name, header=5)
    df_bottom = clean_df(df_bottom)

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    if keyword:
        df_bottom = df_bottom[df_bottom.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        no_data_found = df_bottom.empty

    tables_bottom = df_bottom.to_dict(orient='records')

    return render_template(
        'index.html',
        tables_top=df_top.to_dict(orient='records'),
        tables_project=df_project.to_dict(orient='records'),
        tables_bottom=tables_bottom,
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
        df = get_sheet('IM')
        df = clean_df(df)
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
    app.run(debug=True)
