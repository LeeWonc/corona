from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
import datetime
from selenium import webdriver
import re

def get_patient_info():
    Gu_dict = {'������': 'Gangnam-gu', '������':'Gangdong-gu', '���ϱ�': 'Gangbuk-gu', '������': 'Gangseo-gu', '���Ǳ�':'Gwanak-gu', '������': 'Gwangjin-gu', '���α�':'Guro-gu', '��õ��':'Geumcheon-gu',
               '�����': 'Nowon-gu', '������':'Dobong-gu', '���빮��':'Dongdaemun-gu', '���۱�':'Dongjak-gu', '������':'Mapo-gu', '���빮��':'Seodaemun-gu', '���ʱ�':'Seocho-gu', '������':'Seongdong-gu',
               '���ϱ�':'Seongbuk-gu', '���ı�':'Songpa-gu', '��õ��':'Yangcheon-gu', '��������':'Yeongdeungpo-gu', '��걸':'Yongsan-gu', '����':'Eunpyeong-gu', '���α�':'Jongno-gu', '�߱�': 'Jung-gu', '�߶���':'Jungnang-gu'}
    First_col = ['patient_id', 'global_num', 'sex', 'birth_year' ,'age', 'country', 'province', 'city', 'disease', 'infection_case', 'infection_order', 'infected_by', 'contact_number', 'symptom_onset_date', 'confirmed_date',
                 'released_date', 'deceased_date', 'state']
    Infection_case = {'�ؿ� ���� ����': 'overseas inflow',
                      'Ȯ����': ' ',
                      'Ȯ�� ��': ' ',
                      'Ÿ�õ� Ȯ���� ����': 'contact with patient',
                      '���ü� ����': 'Day Care Center',
                      '������ ����Ʈ ����': 'Seongdong-gu APT',
                      '���� ���� ����': "Eunpyeong St. Mary's Hospital",
                      '��õ�� ����': 'Shincheonji Church',
                      '��û�� ����': 'Seoul City Hall Station safety worker',
                      '���ڿ��ڸ���': 'Daezayeon Korea',
                      '�ǿ� �������� ����': 'Uiwang Logistics Center',
                      '��ġ���� ����': 'Richway',
                      '��õ�� ������ ȸ�� ����': 'Geumcheon-gu rice milling machine manufacture',
                      '��õ�� ��ü� ����': 'Yangcheon Table Tennis Club',
                      '���� �ٴܰ� ����': 'Daejeon door-to-door sales',
                      '������������ ����': 'Orange Life',
                      '������ ��ô��ȸ ����': 'SMR Newly Planted Churches Group',
                      '������ ���ﵿ ����': 'Gangnam Yeoksam-dong gathering',
                      '�ռ���ȸ ����': 'Wangsung Church',
                      '''
                      '���� �޲ٴ� ��ȸ': '����������������������',
                      '���Ƴ�����Ŭ���� ����': '��������������������',
                      '�ѱ����л�����ȸ ����':'����������������������',
                      'kb ������ ����': '������������������������������',
                      '��õ�� ���� ����':'��������������������������',
                      '''
                      '���¿� Ŭ�� ����': 'Itaewon Clubs'}
    contact = re.compile('����')
    overseas = re.compile('�ؿ�')
    released = re.compile('���')
    deceased = re.compile('���')

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
        #infection_case �߰�
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