import pandas as pd
from matplotlib.font_manager import findfont, FontProperties
import matplotlib.pyplot as plt
import os, csv, time, datetime, requests


plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
plt.rcParams['axes.unicode_minus'] = False
today = datetime.date.today() #今日日期
yesterday = str(today - datetime.timedelta(days=1))
today = str(today)
path = 'C:\\Users\\RF\\Desktop\\coding\\op\\'
t = time.localtime()
t1 = time.asctime(t)

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
        print('找不到', file_date, '檔案...')
    
load_file(yesterday, yesterday_data, '昨日')


#爬取 op & fu 資料    
op = 'https://www.taifex.com.tw/cht/3/futAndOptDateExcel'
table = pd.read_html(requests.get(op, headers={'User-agent': 'Mozilla/5.0(Windows NT 6.1; Win64; x64)AppleWebKit/537.36(KHTML, like Gecko)Chrome/63.0.3239.132 Safari/537.36'}).text)
fu_url = 'https://www.taifex.com.tw/cht/3/futContractsDateExcel'
fu_table = pd.read_html(requests.get(fu_url, headers={'User-agent': 'Mozilla/5.0(Windows NT 6.1; Win64; x64)AppleWebKit/537.36(KHTML, like Gecko)Chrome/63.0.3239.132 Safari/537.36'}).text)
op_bc_bp = 'https://www.taifex.com.tw/cht/3/callsAndPutsDateExcel'
op_table = pd.read_html(requests.get(op_bc_bp, headers={'User-agent': 'Mozilla/5.0(Windows NT 6.1; Win64; x64)AppleWebKit/537.36(KHTML, like Gecko)Chrome/63.0.3239.132 Safari/537.36'}).text)

# sm_fu_call = int(fu_table[1].iloc[11][9]) + int(fu_table[1].iloc[9][9]) #外資和自營商小台
# sm_fu_put = int(fu_table[1].iloc[11][11]) + int(fu_table[1].iloc[9][11]) #外資和自營商小台
# fu_call_number_total = table[3].iloc[2, 1]
# fu_put_number_total = table[3].iloc[2, 5]

sm_fu_call = int(fu_table[1].iloc[11][9])
sm_fu_put = int(fu_table[1].iloc[11][11])
fu_call_number = int(fu_table[1].iloc[2][9])
fu_put_number = int(fu_table[1].iloc[2][11])
op_call_money = int(table[3].iloc[2][4])
op_put_money = int(table[3].iloc[2][8])
op_call_number = int(table[3].iloc[2][2])
op_put_number = int(table[3].iloc[2][6])
fu_call_money = int(table[3].iloc[2][3])
fu_put_money = int(table[3].iloc[2][7])
#BCBP_SCSP分計
# 外資
bc_nb = int(op_table[1].iloc[2][10])
sp_nb = int(op_table[1].iloc[5][12])
bp_nb = int(op_table[1].iloc[5][10])
sc_nb = int(op_table[1].iloc[2][12])
bc_mny = int(op_table[1].iloc[2][11])
sp_mny = int(op_table[1].iloc[5][13])
bp_mny = int(op_table[1].iloc[5][11])
sc_mny = int(op_table[1].iloc[2][13])
#自營商
bcsf_nb = int(op_table[1].iloc[0][10])
bpsf_nb = int(op_table[1].iloc[3][10])
scsf_nb = int(op_table[1].iloc[0][12])
spsf_nb = int(op_table[1].iloc[3][12])
bcsf_mny = int(op_table[1].iloc[0][11])
bpsf_mny = int(op_table[1].iloc[3][11])
scsf_mny = int(op_table[1].iloc[0][13])
spsf_mny = int(op_table[1].iloc[3][13])


a = ['Op_call_money', 'Op_put_money', 'Op_call_number', 'Op_put_number', 'Fu_call_money', 'Fu_put_money', 'Fu_call_number', 'Fu_put_number', 'Sm_fu_call', 'Sm_fu_put',
     'Bc_nb', 'Bp_nb', 'Bc_mny', 'Bp_mny', 'Sc_nb', 'Sp_nb', "Sc_mny", 'Sp_mny', 'bcsf_nb', 'bpsf_nb', 'bcsf_mny', 'bpsf_mny', 'scsf_nb', 'spsf_nb', 'scsf_mny', 'spsf_mny']
b = [op_call_money, op_put_money, op_call_number,op_put_number, fu_call_money, fu_put_money, fu_call_number, fu_put_number, sm_fu_call, sm_fu_put,
     bc_nb, bp_nb, bc_mny, bp_mny, sc_nb, sp_nb, sc_mny, sp_mny, bcsf_nb, bpsf_nb, bcsf_mny, bpsf_mny, scsf_nb, spsf_nb, scsf_mny, spsf_mny]
data = pd.DataFrame({'Foreign' : a,'Amount' : b})
data.to_csv(path + str(today) +'.csv', index = False, sep=str(','), encoding='utf-8')



#讀取今日檔案, 可用pandas
today_data = []
load_file(today, today_data, '今日')

#資料分析

yes_data_list = [int(yesterday_data[i][1]) for i in range(len(yesterday_data))]
today_data_list = [int(today_data[i][1]) for i in range(len(today_data))]
dif_list = [today_data_list[i] - yes_data_list[i] for i in range(len(today_data_list))]
print(dif_list)
#Normalize 
Normalize_list = [dif_list[i] / today_data_list[i] for i in range(len(today_data_list))]




def autolabel(rects):
    """Attach a text label above each bar in *rects*, displaying its height."""
    for rect in rects:
        height = rect.get_height()
        plt.annotate('{}'.format(height),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')

def data_visualization(x_range, y_range, save_name, title):
    rects = plt.bar(x_range, y_range, color=['firebrick','g'])
    plt.title(title)
    autolabel(rects) #remark content
    path = 'C:\\Users\\RF\\Desktop\\coding\\op\\' + today +'\\'
    if not os.path.isdir(path):
        os.mkdir(path)
    plt.savefig( path + today + save_name, dpi=200)
    plt.show()
    
def data_print(a=a, dif_list=dif_list):
    data_visualization(a[0:2], dif_list[0:2], '_外資總OP金額差異.png', '外資總OP金額差異')
    print('外資多單op增加', str(dif_list[0]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[0]))
    print('外資空單op增加', str(dif_list[1]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[1]))
    data_visualization(a[2:4], dif_list[2:4], '_外資總OP口數差異.png','外資總OP口數差異')
    print('外資多單總op增加', str(dif_list[2]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[2]))
    print('外資空單總op增加', str(dif_list[3]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[3]))
    data_visualization(a[4:6], dif_list[4:6], '_外資總期貨金額差異.png','外資總期貨金額差異')
    print('外資多單期貨增加', str(dif_list[4]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[4]))
    print('外資空單期貨增加', str(dif_list[5]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[5]))
    data_visualization(a[6:8], dif_list[6:8], '_外資大台期貨差異.png','外資大台期貨差異')
    print('外資多單期貨增加', str(dif_list[6]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[6]))
    print('外資空單期貨增加', str(dif_list[7]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[7]))
    data_visualization(a[8:10], dif_list[8:10], '_外資小台期貨差異.png','外資小台期貨差異')
    print('外資小台多單期貨增加', str(dif_list[8]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[8]))
    print('外資小台空單期貨增加', str(dif_list[9]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[9]))
    data_visualization(a[10:12], dif_list[10:12], '_外資BCBP口數差異.png','外資BCBP口數差異')
    print('外資BC增加', str(dif_list[10]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[10]))
    print('外資BP增加', str(dif_list[11]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[11]))
    data_visualization(a[12:14], dif_list[12:14], '_外資BCBP金額差異.png','外資BCBP金額差異')
    print('外資BC增加', str(dif_list[12]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[12]))
    print('外資BP增加', str(dif_list[13]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[13]))
    data_visualization(a[14:16], dif_list[14:16], '_外資SCSP口數差異.png','外資SCSP口數差異')
    print('外資SC增加', str(dif_list[14]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[14]))
    print('外資SP增加', str(dif_list[15]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[15]))
    data_visualization(a[16:18], dif_list[16:18], '_外資SCSP金額差異.png','外資SCSP金額差異')
    print('外資SC增加', str(dif_list[16]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[16]))
    print('外資SP增加', str(dif_list[17]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[17]))
    data_visualization(a[18:20], dif_list[18:20], '_自營BCBP口數差異.png','自營BCBP口數差異')
    print('自營BC增加', str(dif_list[18]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[18]))
    print('自營BP增加', str(dif_list[19]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[19]))
    data_visualization(a[20:22], dif_list[20:22], '_自營BCBP金額差異.png','自營BCBP金額差異')
    print('自營BC增加', str(dif_list[20]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[20]))
    print('自營BP增加', str(dif_list[21]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[21]))
    data_visualization(a[22:24], dif_list[22:24], '_自營SCSP口數差異.png','自營SCSP口數差異')
    print('自營SC增加', str(dif_list[22]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[22]))
    print('自營SP增加', str(dif_list[23]) , '口')
    print('percent: {:.2%}'.format(Normalize_list[23]))
    data_visualization(a[24:26], dif_list[24:26], '_自營SCS金額差異.png','自營SCSP金額差異')
    print('自營SC增加', str(dif_list[24]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[24]))
    print('自營SP增加', str(dif_list[25]) , '金額')
    print('percent: {:.2%}'.format(Normalize_list[25]))

    print(t1)

    #將圖畫在一張figure
    # path = 'C:\\Users\\RF\\Desktop\\coding\\op\\' + today +'\\'
    # plt.figure()
    # plt.subplot(1,2,1)
    # plt.bar(a[0:2], dif_list[0:2], color=['firebrick','g'])
    # plt.subplot(1,2,2)
    # plt.bar(a[2:4], dif_list[2:4], color=['firebrick','g'])
    # plt.savefig(path + str(today) + '_dif_op.png', dpi=200)
    # plt.show()

    # plt.figure()
    # plt.subplot(1,2,1)
    # plt.bar(a[4:6], dif_list[4:6], color=['firebrick','g'])
    # plt.subplot(1,2,2)
    # plt.bar(a[6:8], dif_list[6:8], color=['firebrick','g'])
    # plt.savefig(path + str(today) + '_dif_fu.png', dpi=200)
    # plt.show()

data_print()

# import utils
# uploaded_image = utils.UploadToImgur(path + str(today) + 'op_money_.png', title='123')

# dpi=300
#畫圖表,rot旋轉X軸座標名稱
# datas = pd.Series(b, index=a)
# datas.loc[['Op_call_money', 'Op_put_money']].plot(kind='bar', rot=0) 
# plt.ylim(ymax= 2500000) #設定座標Y軸上限