import pandas as pd
import os
import glob
import time
import xlrd
import warnings
from datetime import datetime
from tqdm import tqdm


def MoveColumn(name_df, name_column, new_column_num):
    """name_dfのデータフレームのname_columnの列をnew_column_num番目に移動して
    データフレームを返す。"""

    data_column = name_df.pop(name_column)
    name_df.insert(new_column_num, name_column, data_column)
    return name_df


def MakeDenpyo(directory=r'', save=True):
    """カレントディレクトリ内にある作業伝票.xlsxをかき集めて縦に積層し、
    作業伝票まとめ.csvを同ディレクトリ内に出力する。"""

    print('---------------------Depyo---------------------')
    print('Searching files: ')

    # time系は全て実行時時間測定のため
    time_0 = time.time()

    # カレントディレクトリ内にある作業伝票.xlsxを検索し、リスト化してall_filesへ
    all_files = glob.glob(directory + r'*\**\*作業伝票*.xlsx', recursive=True)

    time_1 = time.time()
    print('Done, ', round(time_1 - time_0, 2), 'sec \n')

    print('Checking files : ')

    # all_filesリストからファイルサイズ、編集日時を取得してリストinリストをpairsへ
    # pairsからdf_fileを作成してファイルサイズ10kB以上のファイルのみ抽出
    pairs = []
    for file in all_files:
        size = os.path.getsize(file)
        m_time = os.path.getmtime(file)
        pairs.append([size, m_time, file])
    df_file = pd.DataFrame(pairs, columns=['size', 'm_time', 'file'])
    df_file = df_file[df_file['size'] > 10000]

    time_2 = time.time()
    print('Done, ', round(time_2 - time_1, 2), 'sec \n')

    print('Reading files : ')

    # df_fileにあるファイル名からエクセルファイルを個別に読み込んでいく
    li = []
    err_file_list = []
    for filename in tqdm(df_file['file']):
        # まずはA列のみ読み込んでheader行をしてデータをc_indexへ
        df_index = pd.read_excel(filename, sheet_name=0, usecols='A:A')
        c_index = df_index[df_index.iloc[:, 0] == 'header'].index.values

        # c_indexを使って列名を特定してエクセルデータ読み込んでdfへ変換
        # ロット、S/N無い行を削除
        df = pd.read_excel(filename, sheet_name=0, header=c_index + 1) \
            .dropna(subset=['ロット', 'S/N'])

        # S/Nに数字が入っている場合、floatタイプになって問題になるのでIntに変換
        if df['S/N'].dtype == "float64":
            df['S/N'] = df['S/N'].astype('Int64')

        # S/Nをstr(文字)に変換
        df['S/N'] = df['S/N'].astype(str)

        # ファイル列追加。filenameとなっているがデータ中身はファイルのフルパス
        df['file'] = filename

        # ファイルフルパスからフォルダーの日時部分を抽出して、ちゃんと日時に変換できる場合のみ
        # dfをそのままリスト化(li)。変換できなかった場合はファイルフルパスをerr_file_listに追加
        try:
            df['date'] = datetime.strptime(
                str(os.path.split(os.path.split(filename)[0])[1])[0:8],
                '%Y%m%d').strftime('%Y-%m-%d')
            li.append(df)
        except IndexError:
            err_file_list.append(filename)
        except ValueError:
            err_file_list.append(filename)

    time_3 = time.time()
    print('Done, ', round(time_3 - time_2, 2), 'sec \n')

    print('Processing : ')

    # liから改めてdf化(df_denpyo)
    df_denpyo = pd.concat(li, axis=0, ignore_index=True)

    # 不要列削除
    df_denpyo = df_denpyo.drop(columns=['header', '↓'])

    # TESECファイル名データに.XLS表記無い場合は.XLSを追加
    df_denpyo.loc[
        ~df_denpyo['イニシャル測定_ファイル名'].str.contains('.XLS', na=False),
        'イニシャル測定_ファイル名'] = \
        df_denpyo['イニシャル測定_ファイル名'] + '.XLS'
    df_denpyo.loc[
        ~df_denpyo['アフター測定_ファイル名'].str.contains('.XLS', na=False),
        'アフター測定_ファイル名'] = \
        df_denpyo['アフター測定_ファイル名'] + '.XLS'

    # 下記の新しい列及びデータ追加。
    # ロット_S/N、ロット_WA_No-水準_No、WA_No、水準_No、Chip_No、ロット_WA_NO
    df_denpyo['ロット_S/N'] = df_denpyo['ロット'] + '_' + df_denpyo['S/N']
    df_denpyo['ロット_WA_No-水準_No'] = \
        df_denpyo['ロット_S/N'].str.rsplit('-', n=1, expand=True)[0]
    df_denpyo_SN = df_denpyo['S/N'].str.split(pat="-", expand=True)
    df_denpyo_SN.columns = ['WA_No', '水準_No', 'Chip_No']
    df_denpyo_SN = df_denpyo_SN.assign(
        WA_No=lambda x: x['WA_No'].str.extract(r'(\d+)'))
    df_denpyo_SN = df_denpyo_SN.assign(
        水準_No=lambda x: x['水準_No'].str.extract(r'(\d+)'))
    df_denpyo_SN = df_denpyo_SN.assign(
        Chip_No=lambda x: x['Chip_No'].str.extract(r'(\d+)'))
    df_denpyo_SN['水準_No_origin'] = df_denpyo_SN['水準_No']
    df_denpyo_SN['Chip_No_origin'] = df_denpyo_SN['Chip_No']

    df_denpyo_SN['Chip_No'][df_denpyo_SN['Chip_No_origin'].isnull()] = \
        df_denpyo_SN['水準_No']
    df_denpyo_SN['水準_No'][df_denpyo_SN['Chip_No_origin'].isnull()] = \
        df_denpyo_SN['Chip_No_origin']

    df_denpyo_SN['Chip_No'][
        df_denpyo_SN['Chip_No_origin'].isnull() &
        df_denpyo_SN['水準_No_origin'].isnull()] = df_denpyo_SN['WA_No']
    df_denpyo_SN['WA_No'][
        df_denpyo_SN['Chip_No_origin'].isnull() &
        df_denpyo_SN['水準_No_origin'].isnull()] = df_denpyo_SN['Chip_No_origin']

    df_denpyo_SN = df_denpyo_SN.drop(
        columns=['Chip_No_origin', '水準_No_origin'])

    df_denpyo = df_denpyo.merge(
        df_denpyo_SN, left_index=True, right_index=True)
    df_denpyo['ロット_WA_NO'] = df_denpyo['ロット'] + '_' + df_denpyo['WA_No']

    # S\Nデータの先頭に'を追加(エクセル誤変換対策)
    df_denpyo['S/N'] = '\'' + df_denpyo['S/N']

    # 'date', 'ロット_S/N', '投入日時'が被っている行を削除。2重ファイル対策
    df_denpyo = df_denpyo.drop_duplicates(subset=['date', 'ロット_S/N', '投入日時'],
                                          keep='first')
    # 'Duty (%)', '投入回数', 'Pulse_time (h)', 'DC_time (min)'などの列追加
    df_denpyo['Duty (%)'] = \
        df_denpyo['パルス幅(us)_実測'] * df_denpyo['周波数(kHz)_実測'] * 0.001
    df_denpyo['DCタイマー時間(min)_実際'] = \
        df_denpyo['タイマー時間(h)_実際'] * df_denpyo['Duty (%)'] * 60
    df_denpyo['投入回数'] = df_denpyo.groupby('ロット_S/N').cumcount() + 1
    df_denpyo['Pulse_time (h)'] = \
        df_denpyo.groupby('ロット_S/N')['タイマー時間(h)_実際'].cumsum()
    df_denpyo['DC_time (min)'] = \
        df_denpyo.groupby('ロット_S/N')['DCタイマー時間(min)_実際'].cumsum()

    # ファイルタイプ変更
    df_denpyo = df_denpyo.astype({
        'WA_No': 'float64',
        '水準_No': 'float64',
        'Chip_No': 'float64'
    }
    )
    df_denpyo = df_denpyo.astype({
        'WA_No': 'Int16',
        '水準_No': 'Int16',
        'Chip_No': 'Int16'
    }
    )

    # 整列
    df_denpyo = MoveColumn(df_denpyo, 'file', 0)
    df_denpyo = MoveColumn(df_denpyo, 'date', 1)
    df_denpyo = MoveColumn(df_denpyo, 'ロット_S/N', 8)
    df_denpyo = MoveColumn(df_denpyo, 'ロット_WA_No-水準_No', 9)
    df_denpyo = MoveColumn(df_denpyo, 'ロット_WA_NO', 10)
    df_denpyo = MoveColumn(df_denpyo, 'WA_No', 11)
    df_denpyo = MoveColumn(df_denpyo, '水準_No', 12)
    df_denpyo = MoveColumn(df_denpyo, 'Chip_No', 13)
    df_denpyo = MoveColumn(df_denpyo, '投入回数', 16)
    df_denpyo = MoveColumn(df_denpyo, 'Pulse_time (h)', 17)
    df_denpyo = MoveColumn(df_denpyo, 'DC_time (min)', 18)
    df_denpyo = MoveColumn(df_denpyo, 'Duty (%)', 28)

    time_4 = time.time()
    print('Done, ', round(time_4 - time_3, 2), 'sec \n')

    # saveがTrueだったらcsvへ保存
    if save:
        print('Saving : ')
        df_denpyo.to_csv(directory + r'作業伝票まとめ.csv', encoding='utf_8_sig')
        time_5 = time.time()
        print('Done, ', round(time_5 - time_4, 2), 'sec \n')

    print('-----------------------------------------------')

    # 引数としてdf_denpyoとerr_file_listを返す
    return df_denpyo, err_file_list


def MakeTESEC(directory=r'', save=True):
    """カレントディレクトリ内にある.XLSファイル（TESECファイル）をかき集めて
    測定項目及び各ファイルを縦に積層し、TESECまとめ.csvを同ディレクトリ内に出力する。"""

    print('---------------------TESEC---------------------')
    print('Searching files: ')
    time_0 = time.time()
    # カレントディレクトリ内にある.XLSを検索し、リスト化してall_filesへ
    all_files = glob.glob(directory + r'*\**\*.XLS', recursive=True)

    time_1 = time.time()
    print('Done, ', round(time_1 - time_0, 2), 'sec \n')

    print('Checking files : ')

    # all_filesリストからファイルサイズ、編集日時、ファイル名を取得してリストinリストをpairsへ
    # pairsからdf_fileを作成して被っているファイル名がある行を削除
    pairs = []
    for file in all_files:
        size = os.path.getsize(file)
        m_time = os.path.getmtime(file)
        fname = os.path.split(file)[1]
        pairs.append([size, m_time, file, fname])
    df_file = pd.DataFrame(pairs, columns=['size', 'm_time', 'file', 'fname'])
    df_file = df_file.drop_duplicates(subset='fname', keep='first')
    time_2 = time.time()
    print('Done, ', round(time_2 - time_1, 2), 'sec \n')

    print('Reading files : ')
    # df_fileにあるファイル名からTESECファイルを個別に読み込んでいく
    li = []
    for filename in tqdm(df_file['file']):
        # レシピ名読み取り
        wb = xlrd.open_workbook(filename, logfile=open(os.devnull, 'w'))
        recipe_name = pd.read_excel(
            wb, sheet_name='Result', skiprows=1, nrows=1,
            usecols="B:B", header=None, engine='xlrd').values.flat[0]

        # dfへ変換
        df = pd.read_excel(wb, sheet_name='Result', header=7, engine='xlrd')
        df = df.set_index('MEASURE #:')
        df = df.drop(columns=['Unnamed: 1'])

        # dfの'ITEM NAME:'行にあるデータをリスト化してDELAYがある列インデックスを取得して
        # その列を削除する
        list_ITEM_NAME = list(df.loc['ITEM NAME:'])
        list_ITEM_NAME.insert(0, 'ZERO_INDEX')
        C_list_DELAY = [i for i, x in enumerate(list_ITEM_NAME)
                        if x == 'DELAY']
        df = df.drop(columns=C_list_DELAY)

        # 縦横変換(df_t)
        df_t = df.transpose().reset_index()
        df_t = df_t.drop(columns=['S/NO'])

        # 測定順毎に横に並んでいるデータをさらに縦に積層(df_t_melt)
        df_t_melt = pd.melt(df_t, id_vars=['index',
                                           'ITEM NAME:',
                                           'MIN LIMIT:',
                                           'MAX LIMIT:',
                                           'BIAS 1:',
                                           'BIAS 2:',
                                           'BIAS 3:',
                                           'BIAS 4:'],
                            var_name='測定順'
                            )

        # 'ファイル名', 'レシピ'の列追加
        df_t_melt['ファイル名'] = os.path.basename(filename)
        df_t_melt['レシピ'] = recipe_name

        # 測定順がIntに変換できればdf_t_meltをそのままリスト化(li)
        # 変換できなかった場合（編集された測定順データの場合）はエラーメッセージ出力
        try:
            df_t_melt = df_t_melt.astype({'測定順': 'Int16'})
            li.append(df_t_melt)
        except TypeError:
            print(filename, 'のS/NOデータは数値ではありません。')

    time_3 = time.time()
    print('Done, ', round(time_3 - time_2, 2), 'sec \n')

    print('Processing : ')

    # liから改めてdf化(df_tesec)
    df_tesec = pd.concat(li, axis=0, ignore_index=True)

    # 列名変更
    df_tesec = df_tesec.rename(
        columns={'index': 'MEASURE #:'})

    # 'value'列を数値化。untestはnullになる。
    df_tesec['value'] = pd.to_numeric(df_tesec['value'], errors='coerce')

    # 整列
    df_tesec = MoveColumn(df_tesec, 'ファイル名', 0)
    df_tesec = MoveColumn(df_tesec, 'レシピ', 1)
    df_tesec = MoveColumn(df_tesec, '測定順', 2)
    df_tesec = MoveColumn(df_tesec, 'value', 5)

    time_4 = time.time()
    print('Done, ', round(time_4 - time_3, 2), 'sec \n')

    # saveがTrueだったらcsvへ保存
    if save:
        print('Saving : ')
        df_tesec.to_csv(directory + r'TESECまとめ.csv', encoding='utf_8_sig')
        time_5 = time.time()
        print('Done, ', round(time_5 - time_4, 2), 'sec \n')

    print('-----------------------------------------------')

    # 引数としてdf_tesecを返す
    return df_tesec


def MakeSummary(df_denpyo, df_tesec, directory=r'', save=True):
    """まずTESECまとめ.csvとTESEC測定項目_抽出.xlsxから必要データを抽出する。
    次に作業伝票まとめ.csvを投入回数で積層して抽出したTESECデータとマージする。
    出来がったデータフレームをpHTFBまとめ.csvとして出力し、
    再度投入回数毎に横に並べたデータフレームをpHTFBまとめ_for_excel.csvとして出力する。"""

    print('---------------------Summary---------------------')
    print('Processing: ')
    time_0 = time.time()
    df_tesec.loc[
        (df_tesec['ITEM NAME:'] == 'IGSS') |
        (df_tesec['ITEM NAME:'] == 'IDSS'), 'value'] = abs(df_tesec['value'])
    uni_recipe_tesec = set(df_tesec['レシピ'].unique())
    uni_file_tesec = set(df_tesec['ファイル名'].unique())
    df_tesec_ext = pd.read_excel(directory + 'TESEC測定項目_抽出.xlsx')
    uni_recipe_tesec_ext = set(df_tesec_ext['レシピ'].unique())
    new_recipe_list = uni_recipe_tesec - uni_recipe_tesec_ext
    df_tesec_ext = df_tesec_ext.astype({'測定番号': 'Int16'})
    df_tesec_ext = df_tesec_ext[~df_tesec_ext['測定番号'].isnull()]

    df_tesec = pd.merge(
        df_tesec, df_tesec_ext,
        how='inner',
        left_on=['レシピ', 'MEASURE #:', 'ITEM'
                                      ' NAME:'],
        right_on=['レシピ', '測定番号', '測定項目'])

    df_tesec['condition'] = df_tesec[[
        'MIN LIMIT:',
        'MAX LIMIT:',
        'BIAS 1:',
        'BIAS 2:',
        'BIAS 3:',
        'BIAS 4:']].values.tolist()
    try:
        df_tesec_ustc_val = df_tesec.set_index(
            ['ファイル名', '測定順', '測定項目_抽出'])['value'].unstack().reset_index()
    except ValueError:
        err_tesec_file = df_tesec['ファイル名'][
            df_tesec.duplicated(subset=['ファイル名', '測定順', '測定項目_抽出'])
        ].unique().tolist()
        print('\n以下のTESECファイルでデータが重複している可能性があるのでプログラムは'
              '停止しました。\npHTFBまとめ.csvファイルは保存されていません。'
              '\nファイルを確認後に再試行して下さい。')
        print('\t', err_tesec_file)
        input()
        raise

    df_tesec_ustc_val['Ron+'] = df_tesec_ustc_val['Ron+'] * 1000
    df_tesec_ustc_val['Ron-'] = df_tesec_ustc_val['Ron-'] * 1000

    for col_idss in [col for col in df_tesec_ustc_val.columns.values
                     if 'IDSS' in col]:
        df_tesec_ustc_val.loc[
            df_tesec_ustc_val[col_idss] < 3.04E-8, col_idss + '_MEAS'
        ] = 3.04E-8
        df_tesec_ustc_val.loc[
            df_tesec_ustc_val[col_idss] >= 3.04E-8, col_idss + '_MEAS'
        ] = df_tesec_ustc_val[col_idss]

    for col_igss in [col for col in df_tesec_ustc_val.columns.values
                     if 'IGSS' in col]:
        df_tesec_ustc_val.loc[
            df_tesec_ustc_val[col_igss] < 4.00E-8, col_igss + '_MEAS'
        ] = 4.00E-8
        df_tesec_ustc_val.loc[
            df_tesec_ustc_val[col_igss] >= 4.00E-8, col_igss + '_MEAS'
        ] = df_tesec_ustc_val[col_igss]

    df_tesec_ustc_cond = df_tesec.set_index(
        ['ファイル名', '測定順', '測定項目_抽出'])['condition'].unstack().reset_index()
    df_tesec_ustc_cond = df_tesec_ustc_cond.rename(
        columns={col: col + '_cond' for col in df_tesec_ustc_cond.columns
                 if col not in ['ファイル名', '測定順']})
    df_tesec_ustc = pd.merge(
        df_tesec_ustc_val, df_tesec_ustc_cond,
        how='inner',
        left_on=['ファイル名', '測定順'],
        right_on=['ファイル名', '測定順'])

    Meas_C_list = list(df_tesec['測定項目_抽出'].unique())
    Meas_C_list.sort()
    new_C_list = []
    for item in Meas_C_list:
        if ('IDSS' in item) | ('IGSS' in item):
            new_C_list.append(item + '_cond')
            new_C_list.append(item)
            new_C_list.append(item + '_MEAS')
        else:
            new_C_list.append(item + '_cond')
            new_C_list.append(item)

    new_C_list = ['ファイル名', '測定順'] + new_C_list
    df_tesec_ustc = df_tesec_ustc.reindex(columns=new_C_list)
    df_denpyo['Order'] = df_denpyo.index * 2 + 1
    df_denpyo = MoveColumn(df_denpyo, 'Order', 0)
    df_denpyo_init = df_denpyo[df_denpyo['投入回数'] == 1]
    index_init = list(df_denpyo_init.columns.values).index('イニシャル測定_測定順')
    df_denpyo_init = df_denpyo_init.iloc[:, :index_init + 1]
    df_denpyo_init = df_denpyo_init.rename(
        columns={'イニシャル測定_レシピ名': 'レシピ名',
                 'イニシャル測定_ファイル名': 'ファイル名',
                 'イニシャル測定_測定順': '測定順'})
    df_denpyo_init['投入回数'] = 0
    df_denpyo_init['Pulse_time (h)'] = 0
    df_denpyo_init['DC_time (min)'] = 0
    df_denpyo_init['Order'] = df_denpyo_init['Order'] - 1

    df_denpyo_after = df_denpyo.drop(columns=[
        'イニシャル測定_レシピ名',
        'イニシャル測定_ファイル名',
        'イニシャル測定_測定順'])
    df_denpyo_after = df_denpyo_after.rename(
        columns={'アフター測定_レシピ名': 'レシピ名',
                 'アフター測定_ファイル名': 'ファイル名',
                 'アフター測定_測定順': '測定順'})
    df_summary = pd.concat([df_denpyo_init, df_denpyo_after])
    df_summary = df_summary.sort_values('Order')
    uni_file_summary = set(df_summary['ファイル名'].value_counts().index)
    no_exist_file_list = uni_file_summary - uni_file_tesec - {''}
    df_summary = pd.merge(
        df_summary, df_tesec_ustc,
        how='left',
        left_on=['ファイル名', '測定順'],
        right_on=['ファイル名', '測定順'])

    C_list_df_sum = list(df_summary.columns)
    C_list_for_index = C_list_df_sum[
                       C_list_df_sum.index('機種'):
                       C_list_df_sum.index('投入回数')]
    C_list_for_values = C_list_df_sum[
                        0:C_list_df_sum.index('機種')
                        ] + C_list_df_sum[
                            C_list_df_sum.index('投入回数') + 1:]
    df_summary_pivot = df_summary.pivot(index=C_list_for_index,
                                        columns='投入回数',
                                        values=C_list_for_values
                                        ).reset_index()
    df_summary_pivot = df_summary_pivot.swaplevel(0, 1, 1)
    C_list_pivot = list(
        df_summary_pivot.columns.get_level_values(0).unique())
    df_summary_pivot = df_summary_pivot[C_list_pivot]
    df_summary_pivot.columns = [
        '_'.join(map(str, col)) for col in df_summary_pivot.columns]
    df_summary_pivot.columns = [
        col.replace('_', '', 1) if col[0] == '_' else col
        for col in df_summary_pivot.columns]
    df_summary_pivot = df_summary_pivot.sort_values("0_Order")
    C_list_df_sum_pivot = list(df_summary_pivot.columns)
    C_list_for_init_drop = C_list_df_sum_pivot[
                           C_list_df_sum_pivot.index('0_測定順') + 1:
                           C_list_df_sum_pivot.index(
                               '0_' + str(new_C_list[2]))]
    C_list_for_init_drop += [
        col for col in C_list_df_sum_pivot if 'Order' in col]
    df_summary_pivot = df_summary_pivot.drop(columns=C_list_for_init_drop)
    df_summary_pivot = df_summary_pivot.reset_index(drop=True)

    df_summary = df_summary.drop(columns=['Order'])

    time_1 = time.time()
    print('Done, ', round(time_1 - time_0, 2), 'sec \n')

    if save:
        print('Saving : ')
        df_summary.to_csv(directory + r'pHTFBまとめ.csv', encoding='utf_8_sig')
        df_summary_pivot.to_csv(
            directory + r'pHTFBまとめ_for_excel.csv', encoding='utf_8_sig')
        time_2 = time.time()
        print('Done, ', round(time_2 - time_1, 2), 'sec \n')

    print('-----------------------------------------------')

    return df_summary, new_recipe_list, no_exist_file_list


if __name__ == '__main__':
    warnings.simplefilter('ignore')
    time_init = time.time()
    # Dir = r'\\10.1.44.56\pd\8160_デバイス開発課2G\□共通ﾌｫﾙﾀﾞｰ\☆13.実験報告書\01_ローム\T1K6\T1K6-EVA\T1K6-EVA-000_pHTFB測定データ\\'
    # Dir = r'test_data_2\\'
    Dir = os.getcwd() + '\\'
    df_DENPYO, ERR_FILE_list = MakeDenpyo(directory=Dir, save=True)
    df_TESEC = MakeTESEC(directory=Dir, save=True)
    df_SUMMARY, NEW_RECIPE_list, NO_EXIST_FILE_list = MakeSummary(
        df_denpyo=df_DENPYO, df_tesec=df_TESEC, directory=Dir, save=True)
    ERR_FILE_list = [err_file for err_file in ERR_FILE_list
                     if 'テンプレート' not in err_file]

    time_finished = time.time()
    print('Everything finished successfully.\n'
          'Total time : ', round(time_finished - time_init, 2), 'sec')

    if ERR_FILE_list:
        print('\n以下の作業伝票は読み込まれませんでした。\n'
              '読み込ませるには適切な日時フォルダー内にファイルを移して下さい。')
        for err_file in ERR_FILE_list:
            print('\t', err_file)

    if NO_EXIST_FILE_list:
        print('\n以下のTESECファイルは見つかりませんでした。\n'
              'ファイル未保存か作業伝票のファイル名が異なっている可能性があります。')
        for no_exist_file in NO_EXIST_FILE_list:
            print('\t', no_exist_file)

    if NEW_RECIPE_list:
        print('\n下記のレシピで測定されたTESECデータは「pHTFBまとめ.csv」に\n'
              '反映されませんでした。\n'
              '反映させるには「TESEC測定項目_抽出.xlsx」にデータ追加して下さい。')
        for new_recipe in NEW_RECIPE_list:
            print('\t', new_recipe)
    print("Press Enter to exit ...")
    input()
pass
