# ライブラリの読み込み
import datetime
import xlrd
import re
import csv


# 対象ファイルの読み込み
target_file_path = "/Users/okadatetsuhei/Documents/workspace/Project/teraoka/csvconverter/data-sheet.xlsx"
wb = xlrd.open_workbook(target_file_path)
sheet = wb.sheet_by_name('XD')

### 商品基礎情報の抜き出し
syouhinmei = sheet.cell_value(3,6)
sanchi = sheet.cell_value(4,6)
kikaku = sheet.cell_value(4,10)
maker = sheet.cell_value(5,6)
rott = sheet.cell_value(5,11)
torihikisaki = sheet.cell_value(6,6)
cyakune = sheet.cell_value(6,10)
buturyuu = sheet.cell_value(7,6)
genka = sheet.cell_value(7,10)
torihikisaki_cd = sheet.cell_value(8,6)
hontai = sheet.cell_value(8,10)
haccyuu_cd = sheet.cell_value(9,6)
sougaku = sheet.cell_value(9,10)
jan_cd = sheet.cell_value(10,5)
neireritu = sheet.cell_value(10,10)

productinfo_df = {
    "商品名":syouhinmei,
    "産地":sanchi,
    "規格":kikaku,
    "メーカー":maker,
    "ロット":rott,
    "取引先/帳合":torihikisaki,
    "着値":int(cyakune),
    "物流":buturyuu,
    "原価":int(genka),
    "取引先コード":torihikisaki_cd,
    "本体":int(hontai),
    "発注コード":haccyuu_cd,
    "総額":int(sougaku),
    "JAN":jan_cd,
    "値入率":int(neireritu)*100
}

# print(productinfo_df["JAN"])

###納品情報の抜き出し
def get_list_2d(sheet, start_row, end_row, start_col, end_col):
    return [sheet.row_values(row, start_col, end_col + 1) for row in range(start_row, end_row + 1)]

l_2d = get_list_2d(sheet, 12, 168, 2, 10)



###納品日をyyyyMMddの形式にする
new_year_flg = False
#実行日の年を取得
year = datetime.datetime.now(datetime.timezone.utc).year
for i, dt in enumerate(l_2d[0]):
    if i < 2:
        continue
    # 納品情報から納品月日を取得
    month = dt.split('/')[0]
    day = dt.split('/')[1]
    # 12/31の場合は年を１つカウントアップ
    if month == "12" and day == "31":
        new_year_flg= True
    elif new_year_flg:
        year = year + 1
    # Date型に変換
    new_date = datetime.datetime.strptime(str(year)+"/"+month+"/"+day, "%Y/%m/%d")
    # 納品日をYYYYMMDDの形式に変更
    new_date = new_date.strftime("%Y%m%d")
    l_2d[0][i] = new_date


def making_output_row(row, suryou, target_date, product_info):
    return {
            "処理区分":1,
            "納品日":target_date,
            "便コード":11,
            "商品コード":productinfo_df["JAN"],
            "店舗コード":row[0],
            "チェーンコード":"",
            "本体売単価":productinfo_df["本体"],
            "総額売単価":productinfo_df["総額"],
            "原単価":productinfo_df["原価"],
            "二重売価区分":"",
            "二重売価":"",
            "価格丸め区分":"",
            "総額丸め区分":"",
            "税額丸め区分":"",
            "売価入力区分":"",
            "バーコード内価格区分":"",
            "軽量区分":"",
            "特売区分":"",
            "規格量":re.compile('\d+').findall(productinfo_df["規格"])[0],
            "規格量単位コード":"",
            "上限重量":"",
            "下限重量":"",
            "定単価区分":"",
            "定単価／定重量":"",
            "内容量":re.compile('\d+').findall(productinfo_df["規格"])[0],
            "内容量単位コード":"",
            "加工日":"",
            "広告文コード":"",
            "-":"",
            "受注日":"",
            "作業日":"",
            "受注数":suryou,
            "フリー１":"",
            "フリー２":"",
            "フリー３":"",
            "第１第２ラベラー使用区分":"",
            "売価印字区分":""
        }


### 出力ようcsvデータに整形する
out_data = []
out_data_row = []

for i, row in enumerate(l_2d):
    if i == 0:
        #ヘッダ回避
        continue
    elif not row[0]:
        #店舗コードのないデータはスキップ
        continue
    else:
        # 納品日と納品数を取得する
        for x in list(range(2, 9)):
            if not row[x] or int(row[x] == 0): #指定がないまたは納品数０の場合はスキップ
                continue
            else:
                #納品日
                target_date = l_2d[0][x]
                #納品数
                nouhin_num = int(row[x])
                #出力用配列に格納
                out_data_row.append(making_output_row(row, nouhin_num, target_date, productinfo_df))
    

def making_csvfile(data):
    with open('/Users/okadatetsuhei/Documents/workspace/Project/teraoka/csvconverter/sample_dictwriter.DAT',
                 'w', 
                 newline='', 
                 encoding='shift_jis') as f:
        writer = csv.DictWriter(f, data[0].keys())
        writer.writerows(data)
            

# CSV ファイルの書き出し
making_csvfile(out_data_row)
print(out_data_row[0])
    