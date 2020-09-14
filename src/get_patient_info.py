from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
import datetime
from selenium import webdriver
import re

def get_patient_info():
    Gu_dict = {'강남구': 'Gangnam-gu', '강동구':'Gangdong-gu', '강북구': 'Gangbuk-gu', '강서구': 'Gangseo-gu', '관악구':'Gwanak-gu', '광진구': 'Gwangjin-gu', '구로구':'Guro-gu', '금천구':'Geumcheon-gu',
               '노원구': 'Nowon-gu', '도봉구':'Dobong-gu', '동대문구':'Dongdaemun-gu', '동작구':'Dongjak-gu', '마포구':'Mapo-gu', '서대문구':'Seodaemun-gu', '서초구':'Seocho-gu', '성동구':'Seongdong-gu',
               '성북구':'Seongbuk-gu', '송파구':'Songpa-gu', '양천구':'Yangcheon-gu', '영등포구':'Yeongdeungpo-gu', '용산구':'Yongsan-gu', '은평구':'Eunpyeong-gu', '종로구':'Jongno-gu', '중구': 'Jung-gu', '중랑구':'Jungnang-gu'}
    First_col = ['patient_id', 'global_num', 'sex', 'birth_year' ,'age', 'country', 'province', 'city', 'disease', 'infection_case', 'infection_order', 'infected_by', 'contact_number', 'symptom_onset_date', 'confirmed_date',
                 'released_date', 'deceased_date', 'state']
    Infection_case = {'해외 접촉 추정': 'overseas inflow',
                      '확인중': ' ',
                      '확인 중': ' ',
                      '타시도 확진자 접촉': 'contact with patient',
                      '요양시설 관련': 'Day Care Center',
                      '성동구 아파트 관련': 'Seongdong-gu APT',
                      '은평구 병원 관련': "Eunpyeong St. Mary's Hospital",
                      '신천지 추정': 'Shincheonji Church',
                      '시청역 관련': 'Seoul City Hall Station safety worker',
                      '대자연코리아': 'Daezayeon Korea',
                      '의왕 물류센터 관련': 'Uiwang Logistics Center',
                      '리치웨이 관련': 'Richway',
                      '금천구 도정기 회사 관련': 'Geumcheon-gu rice milling machine manufacture',
                      '양천구 운동시설 관련': 'Yangcheon Table Tennis Club',
                      '대전 다단계 관련': 'Daejeon door-to-door sales',
                      '오렌지라이프 관련': 'Orange Life',
                      '수도권 개척교회 관련': 'SMR Newly Planted Churches Group',
                      '강남구 역삼동 모임': 'Gangnam Yeoksam-dong gathering',
                      '왕성교회 관련': 'Wangsung Church',
                      '''
                      '대전 꿈꾸는 교회': 'ㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌ',
                      '연아나뉴스클래스 관련': 'ㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌ',
                      '한국대학생선교회 관련':'ㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌ',
                      'kb 생명보험 관련': 'ㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌ',
                      '부천시 쿠팡 관련':'ㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌㅌ',
                      '''
                      '이태원 클럽 관련': 'Itaewon Clubs'}
    contact = re.compile('접촉')
    overseas = re.compile('해외')
    released = re.compile('퇴원')
    deceased = re.compile('사망')

    html = urlopen("https://www.seoul.go.kr/coronaV/coronaStatus.do")
    soup = BeautifulSoup(html, "html.parser")

    soul_text = soup.get_text()
    text_list = soul_text.splitlines()
    '''
    patient_info = soup.find_all(id = 'patient')
    for each in patient_info:
        print(each)
        day = each.select('tr')
        print(day)
        input()
    '''

    workbook = xlsxwriter.Workbook('Seoul_patient0914.xlsx')
    worksheet = workbook.add_worksheet()
    indx = 0
    for col in First_col:
        worksheet.write(0, indx, col)
        indx += 1
    row_indx = 1
    text_indx = 641

    while text_indx < len(text_list):
        if text_list[text_indx] == '':
            break
        try:
            worksheet.write(row_indx, 0, int(text_list[text_indx])+1000000000)
        except:
            worksheet.write(row_indx, 0, text_list[text_indx])
        try:
            worksheet.write(row_indx, 1, int(text_list[text_indx+1]))
        except:
            worksheet.write(row_indx, 1, text_list[text_indx+1])
        worksheet.write(row_indx, 6, 'Seoul')
        try:
            worksheet.write(row_indx, 7, Gu_dict[text_list[text_indx+3]])
        except:
            worksheet.write(row_indx, 7, 'etc')
        #infection_case 추가
        if overseas.search(text_list[text_indx+5]):
            worksheet.write(row_indx, 9, 'overseas inflow')
        elif contact.search(text_list[text_indx+5]):
            worksheet.write(row_indx, 9, 'contact with patient')
        else:
            try:
                worksheet.write(row_indx, 9, Infection_case[text_list[text_indx+5]])
            except:
                worksheet.write(row_indx, 9, text_list[text_indx + 5])
        if released.search(text_list[text_indx+6]):
            worksheet.write(row_indx, 17, 'released')
        elif deceased.search(text_list[text_indx+6]):
            worksheet.write(row_indx, 17, 'deceased')
        else:
            worksheet.write(row_indx, 17, 'isolated')
        date_time_list = text_list[text_indx+2].split('.')
        worksheet.write(row_indx, 14, '2020-'+date_time_list[0]+'-'+date_time_list[1])
        text_indx += 9
        row_indx +=1
    workbook.close()

get_patient_info()