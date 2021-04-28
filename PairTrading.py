class PairTrading():
    
    def __init__(self, df, prev_sd = None, prev_avg = None, input_values = None):
        
        self.df = df
        self.prev_sd = prev_sd
        self.prev_avg = prev_avg
        self.input_values = input_values
        self.trad_date_list = []
        
    '''
    計算標準差
    '''
    def get_standard_deviation(self):

        num_list = list(self.df['diff'])
        avg = self.get_average()

        total = 0
        l = len(num_list)
        for num in num_list:
            total += (num - avg) ** 2

        return (total/(l-1)) ** 0.5

    '''
    計算平均
    '''
    def get_average(self):

        num_list = list(self.df['diff'])
        
        return  sum(num_list) / len(num_list)

    def get_result(self):
        
        is_open = False
        trad = 0
        prev_row = None
        
        gain = 0
        total = 0
        
        # print(f'{df["date"][0].year - 1}價差平均: {self.prev_avg}, 價差標準差: {self.prev_sd}')
        
        for index, row in self.df.iterrows():
            
            # 主要為XY之間標準化股價 價差大於 價差的平均加減1.5個價差的標準差 就開倉
            if (((self.prev_avg + self.prev_sd * self.input_values['input_sd']) < row['diff'] <= (self.prev_avg + self.prev_sd * self.input_values['input_sl'])) or 
                ((self.prev_avg - self.prev_sd * self.input_values['input_sl']) <= row['diff'] < (self.prev_avg - self.prev_sd * self.input_values['input_sd']))) and \
                is_open == False and \
                row[3] != row[4]:
                
                # print(f'{row["date"]} 開倉')
            
                is_open = True                  
                prev_row = row
                
            # XY之間標準化股價差收斂回價差的平均加減0.2個價差的標準差 就平倉，結算收益
            elif ((self.prev_avg - self.prev_sd * 0.2) <= row['diff'] <= (self.prev_avg + self.prev_sd * 0.2)) \
                and is_open == True:
                
                # print(f'{row["date"]} 平倉')
                is_open = False
                trad += 1
                
                total += prev_row[1] * prev_row[2] * 2
                self.trad_date_list.append((prev_row['date'], row['date'])) 
                
                if prev_row[3] > prev_row[4]:
                    
                    gain += (prev_row[1] - row[1]) * prev_row[2] \
                        + (row[2] - prev_row[2]) * prev_row[1]

                elif prev_row[3] < prev_row[4]:
                    
                    gain += (row[1] - prev_row[1]) * prev_row[2] \
                        + (prev_row[2] - row[2]) * prev_row[1]
                
            # 主要為XY之間標準化股價 價差大於 價差的平均加減5個價差的標準差 就強制平倉，就算收益
            elif ((row['diff'] > (self.prev_avg + self.prev_sd * self.input_values['input_sl']) or 
                 row['diff'] < (self.prev_avg - self.prev_sd * self.input_values['input_sl'])) and
                 is_open == True):
                
                # print(f'{row["date"]} 強制平倉')
                is_open = False
                trad += 1
                
                total += prev_row[1] * prev_row[2] * 2
                self.trad_date_list.append((prev_row['date'], row['date'])) 
                
                if prev_row[3] > prev_row[4]:
                    
                    gain += (prev_row[1] - row[1]) * prev_row[2] \
                        + (row[2] - prev_row[2]) * prev_row[1]
                    
                elif prev_row[3] < prev_row[4]:
                    
                    gain += (prev_row[2] - row[2]) * prev_row[1] \
                        + (row[1] - prev_row[1]) * prev_row[2]
        
        if total != 0:
            rate = gain / total * 100
        else:
            rate = 0
#         return f'總共交易 {trad} 次, 報酬: {gain}, 總交易金額: {total} \n'
        return [self.df['date'][0].year, trad, gain / 10, rate, self.prev_avg, self.prev_sd]