import pandas as pd


def get_jp_list(excel_name: str):
    df = pd.read_excel(excel_name)
    return df['Local Code'].astype(str).to_list()[:-2907]


if __name__ == '__main__':
    print(get_jp_list('data_jp.xls'))
