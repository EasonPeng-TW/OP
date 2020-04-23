import pandas as pd
import requests
from matplotlib.font_manager import findfont, FontProperties
import matplotlib.pyplot as plt
import datetime
import os
import csv

plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False
today = datetime.date.today() #今日日期
yesterday = str(today - datetime.timedelta(days=1))
today = str(today)
path = 'C:\\Users\\RF\\Desktop\\coding\\op\\'

#讀取昨日檔案, 可用pandas, 可用def
yesterday_data = []
def load_file(load_date, list_data, file_date):
    if os.path.isfile(path +load_date + '.csv'): # 檢查檔案在不在
        print('yes,找到', file_date , '檔案')
        with open(path + load_date +'.csv', encoding='utf-8') as f:
            rows = csv.reader(f)
            for line in rows:
                if 'Foreign' in line:
                    continue
                list_data.append(line)
        print(list_data)
    else:
        print('找不到', date, '檔案...')
    return
load_file(yesterday, yesterday_data, '昨日')


#爬取 op & fu 資料    
op = 'https://www.taifex.com.tw/cht/3/futAndOptDateExcel'
table = pd.read_html(requests.get(op, headers={'User-agent': 'Mozilla/5.0(Windows NT 6.1; Win64; x64)AppleWebKit/537.36(KHTML, like Gecko)Chrome/63.0.3239.132 Safari/537.36'}).text)
fu_url = 'https://www.taifex.com.tw/cht/3/futContractsDateExcel'
fu_table = pd.read_html(requests.get(fu_url, headers={'User-agent': 'Mozilla/5.0(Windows NT 6.1; Win64; x64)AppleWebKit/537.36(KHTML, like Gecko)Chrome/63.0.3239.132 Safari/537.36'}).text)
sm_fu_call = int(fu_table[1].iloc[11][9]) + int(fu_table[1].iloc[9][9])
sm_fu_put = int(fu_table[1].iloc[11][11]) + int(fu_table[1].iloc[9][11])
op_call_money = table[3].iloc[2, 4]
op_put_money = table[3].iloc[2, 8]
op_call_number = table[3].iloc[2, 2]
op_put_number = table[3].iloc[2, 6]
fu_call_money = table[3].iloc[2, 3]
fu_put_money = table[3].iloc[2, 7]
fu_call_number = table[3].iloc[2, 1]
fu_put_number = table[3].iloc[2, 5]



a = ['Op_call_money', 'Op_put_money', 'Op_call_number', 'Op_put_number', 'Fu_call_money', 'Fu_put_money', 'Fu_call_number', 'Fu_put_number', 'Sm_fu_call', 'Sm_fu_put']
b = [op_call_money, op_put_money, op_call_number,op_put_number, fu_call_money, fu_put_money, fu_call_number, fu_put_number, sm_fu_call, sm_fu_put ]
data = pd.DataFrame({'Foreign' : a,'Amount' : b})
data.to_csv(path + str(today) +'.csv', index = False, sep=str(','), encoding='utf-8')

#讀取今日檔案, 可用pandas
today_data = []
load_file(today, today_data, '今日')

#資料分析
yes_op_call_money = int(yesterday_data[0][1])
yes_op_put_money = int(yesterday_data[1][1])
yes_op_call_number = int(yesterday_data[2][1])
yes_op_put_number = int(yesterday_data[3][1])
yes_fu_call_money = int(yesterday_data[4][1])
yes_fu_put_money = int(yesterday_data[5][1])
yes_fu_call_number = int(yesterday_data[6][1])
yes_fu_put_number = int(yesterday_data[7][1])
yes_smfu_call_number = int(yesterday_data[8][1])
yes_smfu_put_number = int(yesterday_data[9][1])


today_op_call_money= int(today_data[0][1])
today_op_put_money = int(today_data[1][1])
today_op_call_number = int(today_data[2][1])
today_op_put_number = int(today_data[3][1])
today_fu_call_money = int(today_data[4][1])
today_fu_put_money = int(today_data[5][1])
today_fu_call_number = int(today_data[6][1])
today_fu_put_number = int(today_data[7][1])
today_smfu_call_number = int(today_data[8][1])
today_smfu_put_number = int(today_data[9][1])

dif_op_call_money = today_op_call_money - yes_op_call_money
dif_op_put_money = today_op_put_money - yes_op_put_money
dif_op_call_number = today_op_call_number - yes_op_call_number
dif_op_put_number = today_op_put_number - yes_op_put_number
dif_fu_call_money = today_fu_call_money - yes_fu_call_money
dif_fu_put_money = today_fu_put_number - yes_fu_put_money
dif_fu_call_number = today_fu_call_number - yes_fu_call_number
dif_fu_put_number = today_fu_put_number - yes_fu_put_number
dif_smfu_call = today_smfu_call_number - yes_smfu_call_number
dif_smfu_put = today_smfu_put_number - yes_smfu_put_number

yes_data_list = [yes_op_call_money, yes_op_put_money, yes_op_call_number, yes_op_put_number, yes_fu_call_money, yes_fu_put_money, yes_fu_call_number, yes_fu_put_number, yes_smfu_call_number, yes_smfu_put_number]
today_data_list = [today_op_call_money, today_op_put_money, today_op_call_number, today_op_put_number, today_fu_call_money, today_fu_put_money,today_fu_call_number, today_fu_put_number, today_smfu_call_number, today_smfu_put_number]

dif_list = [today_data_list[i] - yes_data_list[i] for i in range(len(today_data_list))]
print(dif_list)

# plt.bar(a[0:2], yes_data_list[0:2], label = 'yesterday', align = "edge", width = -0.35)
# plt.bar(a[0:2], today_data_list[0:2], label = 'today', align = "edge", width = 0.35)
# plt.legend()
# plt.savefig(path + str(today) + 'op_money_.png', dpi=300)
# plt.show()


def data_visualization(x_range, y_range, save_name):
    plt.bar(x_range, y_range, color=['firebrick','g'])
    plt.savefig(path + str(today) + save_name, dpi=200)
    plt.show()
    return

data_visualization(a[0:2], dif_list[0:2], '_dif_op_money.png')
print('外資多單op增加', str(dif_list[0]) , '金額')
print('外資空單op增加', str(dif_list[1]) , '金額')
data_visualization(a[2:4], dif_list[2:4], '_dif_op_number.png')
print('外資多單op增加', str(dif_list[2]) , '口')
print('外資空單op增加', str(dif_list[3]) , '口')
data_visualization(a[4:6], dif_list[4:6], '_dif_fu_money.png')
print('外資多單期貨增加', str(dif_list[4]) , '金額')
print('外資空單期貨增加', str(dif_list[5]) , '金額')
data_visualization(a[6:8], dif_list[6:8], '_dif_fu_number.png')
print('外資多單期貨增加', str(dif_list[6]) , '口')
print('外資空單期貨增加', str(dif_list[7]) , '口')
data_visualization(a[8:10], dif_list[8:10], '_dif_smfu_number.png')
print('外資小台多單期貨增加', str(dif_list[8]) , '口')
print('外資小台空單期貨增加', str(dif_list[9]) , '口')
#將圖畫在一張

plt.figure()
plt.subplot(1,2,1)
plt.bar(a[0:2], dif_list[0:2], color=['firebrick','g'])
plt.subplot(1,2,2)
plt.bar(a[2:4], dif_list[2:4], color=['firebrick','g'])
plt.savefig(path + str(today) + '_dif_op.png', dpi=200)
plt.show()

plt.figure()
plt.subplot(1,2,1)
plt.bar(a[4:6], dif_list[4:6], color=['firebrick','g'])
plt.subplot(1,2,2)
plt.bar(a[6:8], dif_list[6:8], color=['firebrick','g'])
plt.savefig(path + str(today) + '_dif_fu.png', dpi=200)
plt.show()


# import utils
# uploaded_image = utils.UploadToImgur(path + str(today) + 'op_money_.png', title='123')

# dpi=300
#畫圖表,rot旋轉X軸座標名稱
# datas = pd.Series(b, index=a)
# datas.loc[['Op_call_money', 'Op_put_money']].plot(kind='bar', rot=0) 
# plt.ylim(ymax= 2500000) #設定座標Y軸上限
