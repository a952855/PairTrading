import pandas as pd
import matplotlib.pyplot as plt

from PairTrading import PairTrading

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False

def generate_pair_images(df, sheet, trad_date_list):

    a_name, b_name = sheet.split('+')

    fig = plt.figure(figsize=[15, 10], dpi=300)
    ax1 = fig.add_subplot(3,1,1)
    ax1.plot(df.date, df[df.columns[3]], label=f'{a_name[4:]}（{a_name[:4]}）標準化股價')
    ax1.plot(df.date, df[df.columns[4]], label=f'{b_name[4:]}（{b_name[:4]}）標準化股價')
    plt.xlabel('年分') # 設定x軸標題
    # plt.xticks(df.date, rotation='vertical') # 設定x軸label以及垂直顯示
    plt.title(f'{a_name[4:]}與{b_name[4:]}股票配對交易圖') # 設定圖表標題
    for date in trad_date_list:
        plt.axvspan(date[0], date[1], color='green', alpha=0.5)

    plt.legend(loc = 'upper right')

    ax2 = fig.add_subplot(3,1,2)
    ax2.plot(df.date, df[df.columns[5]], label=df.columns[5])
    plt.xlabel('年分') # 設定x軸標題
    # plt.xticks(df.date, rotation='vertical') # 設定x軸label以及垂直顯示
    plt.title('標準化股價價差') # 設定圖表標題
    for date in trad_date_list:
        plt.axvspan(date[0], date[1], color='green', alpha=0.5)

    ax2 = fig.add_subplot(3,1,3)
    ax2.plot(df.date, df[df.columns[5]], label=df.columns[5])
    plt.xlabel('年分') # 設定x軸標題
    # plt.xticks(df.date, rotation='vertical') # 設定x軸label以及垂直顯示
    plt.title('標準化股價價差') # 設定圖表標題

    plt.tight_layout() #隔開圖

    plt.savefig(f'./images/{sheet}.png') 

if __name__ == '__main__':
    
    path = r'./files/PairTradeFoodIndustry.xlsx'

    xls = pd.ExcelFile(path)
    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

    for sheet in xls.sheet_names[1:]:
        df = pd.read_excel(path, sheet_name=sheet)

        print(f'{sheet} 正在計算中，請稍等')
        '''
        group by data
        '''
        df.index = pd.to_datetime(df['date'], format='%y/%m/%d')
        grouped_df = df.groupby(by=[df.index.year])

        df_list = [group for name, group in grouped_df]

        # PairTrading
        sd = 0
        avg = 0


        data = []
        trad_date_list = []
        for i, df_per_year in enumerate(df_list):
            
            if i == 0:
                pt = PairTrading(df_per_year)
            else:
                pt = PairTrading(df_per_year, sd, avg)
                data.append(pt.get_result())
                trad_date_list += pt.trad_date_list

            sd = pt.get_standard_deviation()
            avg = pt.get_average()
            
        ans_df = pd.DataFrame(
            data,
            columns=['年分', '交易次數', '總報酬', '去年標準價平均', '去年標準差'])

        average = ans_df.mean()
        average['年分'] = ''
        ans_df.loc['平均'] = average

        ans_df.to_excel(writer, sheet_name=sheet)
        
        # image
        generate_pair_images(df, sheet, trad_date_list)

        workbook  = writer.book
        worksheet = writer.sheets[sheet]
        worksheet.insert_image('A15', f'./images/{sheet}.png')

        print(f'{sheet} 計算完畢')

    writer.save()
    input("程式結束，已生成output.xlsx. Press enter to exit.")