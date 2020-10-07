from selenium import webdriver
import pandas as pd
import numpy as np
from math import pi
import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
import time
import os

def first():
    chromeOptions = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : os.getcwd()}
    chromeOptions.add_experimental_option("prefs", prefs)
    
    # My computer uses a Chrome 85 version, so I use a chromedriver that supports it. 
    # Check the Chrome version used on your computer and install chromedriver to support it
    driver = webdriver.Chrome('chromedriver', chrome_options=chromeOptions)
    
    driver.implicitly_wait(3)
    driver.get("https://pay.tmoney.co.kr/ncs/pct/ugd/ReadTrcrStstDtl.dev?useYm=202005&rgtDtm=20200603114856")

    excel_download = driver.find_element_by_css_selector("#contents > div.view_box > div.view_bot.clfix > dl.view_file.clfix > dd > a")
    driver.implicitly_wait(3)

    excel_download.click()
    time.sleep(5)

    driver.close()
    print("파일 다운로드 실행...")

    # You need to install xlrd to convert xls -> csv
    return None
    
def second():
    # If you want to read file with csv, then you need to add "encoding='euc-kr'" option
    data = pd.read_excel('2020년 05월  교통카드 통계자료.xls', sheet_name='지하철 유무임별 이용현황', thousands=',')
    df = pd.DataFrame(data)

    df1 = pd.to_numeric(df['유임승차'], errors='coerce')
    df2 = pd.to_numeric(df['무임승차'], errors='coerce')

    df['유임승차비율'] = df1 / (df1 + df2)
    print(df.sort_values('유임승차비율', ascending=False).head(5)[['지하철역']])

def third():
    data = pd.read_excel('2020년 05월  교통카드 통계자료.xls', sheet_name='지하철 유무임별 이용현황', thousands=',')
    df = pd.DataFrame(data)

    df1 = df.sort_values('유임승차', ascending=False).head(1)[['유임승차', '유임하차', '무임승차', '무임하차']]
    df2 = df.sort_values('유임승차', ascending=True).head(1)[['유임승차', '유임하차', '무임승차', '무임하차']]
 
    val1 = df1.loc[df1.index[0], :].values.tolist()
    val1 += val1[:1]
    
    val2 = df2.loc[df2.index[0], :].values.tolist()
    val2 += val2[:1]

    N = 4
    categories = ['유임승차', '유임하차', '무임승차', '무임하차']    
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]

    font_path = "malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family = font_name)

    ax = plt.subplot(111, polar=True)
    plt.xticks(angles[:-1], categories, color='grey', size=10)
    ax.set_rlabel_position(45)
    
    plt.yticks([550000,1100000,1650000,2200000], ["550000","1100000","1650000","2200000"], color="red", size=7)
    plt.ylim(0, 2200000)

    ax.plot(angles, val1, linewidth=1, linestyle='solid', label='최대')
    ax.fill(angles, val1, 'orange', alpha=0.2)

    ax.plot(angles, val2, linewidth=1, linestyle='solid', label='최소')
    ax.fill(angles, val2, 'blue', alpha=0.2)

    plt.legend(loc='best', bbox_to_anchor=(0.05, 0.95))
    plt.show()
    
def fourth():
    data = pd.read_excel('2020년 05월  교통카드 통계자료.xls', sheet_name='지하철 시간대별 이용현황', thousands=',')
    df = pd.DataFrame(data)

    df1 = df['21:00:00~21:59:59']
    df2 = df['22:00:00~22:59:59']
    df3 = df['23:00:00~23:59:59']

    df1.drop(df1.index[0], inplace=True)
    df1 = df1.str.replace(",", "").astype(int)
    
    df2.drop(df2.index[0], inplace=True)
    df2 = df2.str.replace(",", "").astype(int)

    df3.drop(df3.index[0], inplace=True)
    df3 = df3.str.replace(",", "").astype(int)

    df['21~24시 탑승자수'] = df1 + df2 + df3

    df_result = df.sort_values('21~24시 탑승자수', ascending=False).head(5)
    
    font_path = "malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family = font_name)

    df_result.set_index('지하철역', inplace=True)

    df_result[['21~24시 탑승자수']].plot(kind='bar')
    plt.show()
    
def fifth():
    data = pd.read_excel('2020년 05월  교통카드 통계자료.xls', sheet_name='지하철 노선별 역별 이용현황', thousands=',')
    df = pd.DataFrame(data)

    df_temp = df.groupby(['호선명'], as_index = False).sum()
    df_result = df_temp[['승차승객수']]
    df_index = df_temp['호선명'][::-1]
    
    font_path = "malgun.ttf"
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family = font_name)

    # ax = sns.heatmap(df_result, cmap='YlGnBu')
    # Better to implement heatmap with above command
    
    plt.pcolor(df_result)
    plt.yticks(np.arange(0.5, len(df_index), 1), df_index)
    plt.title('호선별 승차 인원의 합', fontsize=20)
    plt.xlabel('heatmap', fontsize=14)
    plt.ylabel('호선명', fontsize=14)
    plt.colorbar()
                                  
    plt.show()
        
def main():
    first()
    second()
    third()
    fourth()
    fifth()

if __name__ == "__main__":
    main()
