import openpyxl
import pandas as pd
import datetime
import re
from openpyxl.chart import Reference
from openpyxl.chart.axis import DateAxis
import matplotlib.pyplot as plt

#日程を取得
date = datetime.datetime.now()

#excelのデータを読み込み、値を取得
Excel_File = pd.ExcelFile(f'株価リスト_{date.month}月.xlsx')
dfs = []
#複数シートから
for sheetname in Excel_File.sheet_names:
    df = Excel_File.parse(sheetname)
    #列名を「株価」から変える
    df = df.rename(columns={'株価':sheetname})
    #会社名をインデックスにする
    df = df.set_index('会社名')
    df[sheetname] = df[sheetname].str.replace(r'[^\d.]', '', regex=True).astype(float)
    #株価推移を削除する
    df = df.drop('株価推移', axis=1)
    dfs.append(df)
#表を結合させる
conbine_df = pd.concat(dfs, axis=1)
#EXCELに出力
conbine_df.to_excel('test.xlsx', index=True)
#取得した値をもとにグラフ化

#日本語対応させる
plt.rcParams['font.family'] = "MS Gothic"

#行と列を転置
conbine_df_T = conbine_df.transpose()

#企業ごとに表を抽出する
for index, row in conbine_df.iterrows():
    plt.plot(row, label=f'行{index}')
    print(index)
    #グラフの設定
    plt.title("株価比較")
    plt.xlabel("日付")
    plt.ylabel("株価")

    plt.savefig(f'株価推移グラフ_{index}_{date.month}月.png')
    plt.show()

#グラフをPDF化

