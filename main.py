import pandas as pd

from PairTrading import PairTrading

if __name__ == '__main__':
    
    path = r'./files/PairTradeFoodIndustry.xlsx'

    xls = pd.ExcelFile(path)
    df = pd.read_excel(xls, '1210大成+1201味全')

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
        for i, df in enumerate(df_list):
            
            if i == 0:
                pt = PairTrading(df)
            else:
                pt = PairTrading(df, sd, avg)
                data.append(pt.get_result())
                
            sd = pt.get_standard_deviation()
            avg = pt.get_average()
            
        ans_df = pd.DataFrame(
            data,
            columns=['年分', '交易次數', '總報酬'])

        average = ans_df.mean()
        average['年分'] = ''
        ans_df.loc['平均'] = average

        ans_df.to_excel(writer, sheet_name=sheet)

        print(f'{sheet} 計算完畢')

    writer.save()
    input("程式結束，已生成output.xlsx. Press enter to exit.")