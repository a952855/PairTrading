import pandas as pd
import matplotlib.pyplot as plt

from io import BytesIO

from PairTrading import PairTrading

# 設定中文字體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False

if __name__ == '__main__':
    
    path = r'./files/PairTradeFoodIndustry.xlsx'

    # xls = pd.ExcelFile(path)
    # df = pd.read_excel(xls, '1210大成+1201味全')

    xls = pd.ExcelFile(path)
    writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

    for sheet in xls.sheet_names[1:]:
        df = pd.read_excel(path, sheet_name=sheet)
        # df_list.append(df)  ## This will append rows of one dataframe to another(just like your expected output)

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
        for i, df in enumerate(df_list):
            
            if i == 0:
                pt = PairTrading(df)
            else:
                pt = PairTrading(df, sd, avg)
                data.append(pt.get_result())
                trad_date_list += pt.trad_date_list

            sd = pt.get_standard_deviation()
            avg = pt.get_average()
            
        ans_df = pd.DataFrame(
            data,
            columns=['年分', '交易次數', '總報酬'])

        average = ans_df.mean()
        average['年分'] = ''
        ans_df.loc['平均'] = average

        ans_df.to_excel(writer, sheet_name=sheet)
        
        # image
        fig, ax = plt.subplots()
        print(df)

        plt.figure(figsize=[15, 4.8], dpi=300)
        plt.plot(df.date, df[df.columns[3]], label=df.columns[3])
        plt.plot(df.date, df[df.columns[4]], label=df.columns[4])
        plt.xlabel('年分') # 設定x軸標題
        # plt.xticks(df.date, rotation='vertical') # 設定x軸label以及垂直顯示
        plt.title(sheet) # 設定圖表標題
        # date = trad_date_list[0]
        for date in trad_date_list:
            plt.axvspan(date[0], date[1], color='green', alpha=0.5)

        plt.legend(loc = 'upper right')
        plt.savefig(f'{sheet}.png') 



        workbook  = writer.book
        worksheet = writer.sheets[sheet]



        # book=xls.Workbook('abc.xls')

        # sheet=book.add_worksheet('demo')

        worksheet.insert_image('A10', f'{sheet}.png') 

        print(f'{sheet} 計算完畢')

    writer.save()
    input("程式結束，已生成output.xlsx. Press enter to exit.")