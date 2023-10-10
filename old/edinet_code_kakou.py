import pandas as pd

FILECODE = 'EdinetcodeDlInfo.csv'


if __name__ == '__main__':

    df = pd.read_csv(FILECODE, encoding="cp932", usecols=[0,6,11],
                    names=('name_code', 'name', 'syoken_code'), dtype={"syoken_code": str}, skiprows=2)
    #値がNullになっているものを削除
    df = df.dropna(how='any', axis=0)
    df_ex = df.copy()
    #0を消去する
    df_ex['syoken_code'] = df['syoken_code'].str[:4]
    #csvを出力
    df_ex.to_csv("output.csv", encoding="cp932" ,index=False)


