import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# 清理資料
def clean_df(df):
    df.columns = df.columns.str.replace('\n', '', regex=False)
    df = df.fillna('')
    return df

@app.route('/')
def index():
    df_department = pd.read_excel('data.xlsx', sheet_name='首頁', usecols="A:F", nrows=1)
    df_department = clean_df(df_department)

    df_workstats = pd.read_excel('data.xlsx', sheet_name='首頁', usecols="A:G", skiprows=2, nrows=2)
    df_workstats = clean_df(df_workstats)

    df = pd.read_excel('data.xlsx', sheet_name=0, header=13, nrows=250, usecols="A:O")
    df = clean_df(df)
    df = df[['門市編號', '門市名稱', 'PMQ2檢核', '專案檢核', 'HUB', '完工檢核']]

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
        workstats_table=df_workstats.to_dict(orient='records'),
        no_data_found=no_data_found
    )

@app.route('/<name>')
def personal(name):
    sheet_map = {
        '吳宗鴻': '吳宗鴻',
        '湯家瑋': '湯家瑋',
        '狄澤洋': '狄澤洋'
    }
    sheet_name = sheet_map.get(name, None)
    if not sheet_name:
        return f"找不到{name}的分頁", 404

    df_top = pd.read_excel('data.xlsx', sheet_name=sheet_name, usecols="A:J", nrows=4)
    df_top = clean_df(df_top)
    for column in df_top.columns:
        df_top[column] = df_top[column].apply(lambda x: int(x) if isinstance(x, (int, float)) else x)

    df_bottom = pd.read_excel('data.xlsx', sheet_name=sheet_name, usecols="A:J", skiprows=5)
    df_bottom = clean_df(df_bottom)

    keyword = request.args.get('keyword', '').strip()
    no_data_found = False
    tables_bottom = df_bottom.to_dict(orient='records')

    if keyword:
        df_bottom = df_bottom[df_bottom.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
        if df_bottom.empty:
            no_data_found = True
        tables_bottom = df_bottom.to_dict(orient='records')

    return render_template(
        'index.html',
        tables_top=df_top.to_dict(orient='records'),
        tables_bottom=tables_bottom,
        keyword=keyword,
        store_id='',
        repair_item='',
        personal_page=True,
        report_page=False,
        no_data_found=no_data_found
    )

@app.route('/report')
def report():
    keyword = request.args.get('keyword', '').strip()
    store_id = request.args.get('store_id', '').strip()
    repair_item = request.args.get('repair_item', '').strip()
    no_data_found = False
    tables = []

    if keyword or store_id or repair_item:
        df = pd.read_excel('data.xlsx', sheet_name='IM')
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
