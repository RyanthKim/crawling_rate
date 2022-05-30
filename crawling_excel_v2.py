import csv
import urllib.request
import requests
from bs4 import BeautifulSoup
import pyautogui
import openpyxl as xl

wb = xl.Workbook()
sheet1 = wb.active
sheet1.title = '전업권'
sheet1.append(["회사명","구분","공시일시","1등급","2등급","3등급","4등급","5등급","6등급","7등급","8등급","9등급","10등급","평균금리"])
sheet2 = wb.create_sheet("크롤링 설명")

print("안녕하세요 공시정보 crawling program 입니다. - by ryan")
print("version = 2.7 - 최종수정일 : 2020-05-22")
print("----------------------------------------")
year = input('년도를 입력해주세요 (4자리 예 : 2020) : ')
month = input('월을 입력해주세요 (2자리 예 : 01,02,03 ...) : ')
print("----------------------------------------")
combine = '{0}-{1}'.format(year,month)
btn = pyautogui.confirm('각 페이지의 금리 정보를 불러오시겠습니까?','공시정보 crawling')
if int(month) <= 12 and int(month) >= 10:
    month_new = "01"
    year_new = str(int(year) - 1)
if int(month) <= 6 and int(month) >= 4:
    month_new = "04"
    year_new = year
if int(month) <= 9 and int(month) >= 7:
    month_new = "07"
    year_new = year
if int(month) <= 12 and int(month) >= 10:
    month_new = "10"
    year_new = year

if btn == 'OK':
    # 은행연합회
    print(" -- 은행 시작")
    List_data = {
        "year" : "%s" % year,
        "month" : "%s" % month,
        "detail" : "0",
        'str' : 'KDB%BB%EA%BE%F7%C0%BA%C7%E0|NH%B3%F3%C7%F9%C0%BA%C7%E0|%BD%C5%C7%D1%C0%BA%C7%E0|%BF%EC%B8%AE%C0%BA%C7%E0|%BD%BA%C5%C4%B4%D9%B5%E5%C2%F7%C5%B8%B5%E5%C0%BA%C7%E0|%C7%CF%B3%AA%C0%BA%C7%E0|IBK%B1%E2%BE%F7%C0%BA%C7%E0|KB%B1%B9%B9%CE%C0%BA%C7%E0|%C7%D1%B1%B9%BE%BE%C6%BC%C0%BA%C7%E0|SH%BC%F6%C7%F9%C0%BA%C7%E0|DGB%B4%EB%B1%B8%C0%BA%C7%E0|BNK%BA%CE%BB%EA%C0%BA%C7%E0|%B1%A4%C1%D6%C0%BA%C7%E0|%C1%A6%C1%D6%C0%BA%C7%E0|%C0%FC%BA%CF%C0%BA%C7%E0|BNK%B0%E6%B3%B2%C0%BA%C7%E0|%C4%C9%C0%CC%B9%F0%C5%A9%C0%BA%C7%E0|%C7%D1%B1%B9%C4%AB%C4%AB%BF%C0%C0%BA%C7%E0'    
    }
    opt_1 = ["/주담대/분할","/주담대/일시","/신용/일반","/신용/마이너스"]
    number = 0
    for i in opt_1:
        number = number + 1
        List_data["opt_1"] = str(number)
        with requests.Session() as s:
            request_bank = s.post("https://portal.kfb.or.kr/compare/loan_household_search_result.php", data = List_data)

        soup = BeautifulSoup(request_bank.text,"html.parser")

        title = soup.find_all("table", class_="resultList_ty02")[0].find_all("th")
        tag = soup.find_all("table", class_="resultList_ty02")[0].find_all('tr')[2:]
        all_list = []
        for td in tag:
            temp = td.get_text().replace(" ","").replace("대출금리","").split()
            number_2 = 0
            for j in temp:
                try:
                    value = float(j)
                    temp.pop(number_2)
                    temp.insert(number_2,value)
                except ValueError:
                    True
                number_2 = number_2 + 1
            number_2 = 0
            for j in temp:
                if j == 0 or j == 0.0 or j == "-":
                    temp.pop(number_2)
                    temp.insert(number_2,"")
                number_2 = number_2 + 1
            temp_2 = temp
            temp.insert(6,temp_2[5])
            temp.insert(5,temp_2[4])
            temp.insert(4,temp_2[3])
            temp.insert(3,temp_2[2])
            temp.insert(2,temp_2[1])
            temp.insert(1,"{0}-{1}".format(year,month))
            temp.insert(1,"은행{0}".format(i))
            sheet1.append(temp[0:14])
    print(" -- 은행 완료")

    # 손해보험
    print(" -- 손해보험 시작")
    List_data = {
        'TYPE': 'C',
        'DETAIL': 'A',
        'P_CODE': "",
        'YEAR': '{}'.format(year),
        'MONTH': '{}'.format(month)
    }
    with requests.Session() as s:
        request = s.post("https://kpub.knia.or.kr/loan/NonPublicLoan.knia?detail", data = List_data)

    soup = BeautifulSoup(request.text,"html.parser")

    tag = soup.find_all("table", class_="etc_table1 group_tb")[0].find_all('tr')[2:]
    all_list = []
    for td in tag:
        temp = td.get_text().replace("대출금리","").split()
        number = 0
        for i in temp:
            try:
                value = float(i)
                temp.pop(number)
                temp.insert(number,value)
            except ValueError:
                False
            number = number + 1
        number = 0
        for i in temp:
            if i == 0 or i == 0.0:
                temp.pop(number)
                temp.insert(number,"")
            number = number + 1
            
        temp_2 = temp
        temp.insert(5,temp_2[5])
        temp.insert(5,temp_2[5])
        temp.insert(5,temp_2[5])
        temp.insert(1,temp_2[1])
        temp.insert(1,temp_2[1])
        temp.insert(1,"{0}-{1}".format(year,month))
        temp.insert(1,"손해보험")
        sheet1.append(temp[0:14])
    
    print(" -- 손해보험 완료")

    # 생명보험
    print(" -- 생명보험 시작")
    url = 'https://pub.insure.or.kr/loan/type/householdLoan/list.do?search_stdYm={0}-{1}&search_memberCd='.format(year,month)

    headers = {
        'Referer': 'https://pub.insure.or.kr/loan/type/householdLoan/list.do?search_stdYm=2019-07&search_memberCd=',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
    }


    data_list = {
    'credit_search_loanType': '02660002',
    'credit_search_view': '02680001',
    'credit_search_rateType': '1',
    'credit_search_stdYm': '{0}-{1}'.format(year,month)
    }
    with requests.session() as s:
        response = requests.post('https://pub.insure.or.kr/loan/type/householdLoan/creditLoanList.do',data = data_list)

    soup = BeautifulSoup(response.text,"html.parser")
    tag = soup.find_all("table", class_="table_t01 td_align_center b_line thead_pd01 data_table")[0].find_all('tr')[2:]
    all_list = []
    for td in tag:
        temp = td.get_text().split()
        temp.insert(1,"{0}-{1}".format(year,month))
        temp.insert(1,"생명보험")
        number = 0
        for i in temp:
            try:
                value = float(i)
                temp.pop(number)
                temp.insert(number,value)
            except ValueError:
                False
            number = number + 1
        number = 0
        for i in temp:
            if i == 0 or i == 0.0 or i == "-":
                temp.pop(number)
                temp.insert(number,"")
            number = number + 1
        del temp[3]
        temp_2 = temp
        temp.insert(7,temp_2[7])
        temp.insert(7,temp_2[7])
        temp.insert(7,temp_2[7])
        temp.insert(3,temp_2[3])
        temp.insert(3,temp_2[3])
        sheet1.append(temp)  
    
    print(" -- 생명보험 완료")
    
    # 카드신용
    print(" -- 카드신용 시작")
    url = 'https://gongsi.crefia.or.kr/portal/creditcard/creditcardDisclosureDetail25Ajax?cgcMode=25'
    headers = {
                'Referer': 'https://gongsi.crefia.or.kr/portal/creditcard/creditcardDisclosureDetail20?cgcMode=20',
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
    }
    response = requests.get(url, headers=headers)
    jsonObjs = response.json()
    dataList = jsonObjs['resultList']
    all_list = []
    for data in dataList :
        temp = []
        temp.append(data['mcCompany'])
        temp.append("카드신용")
        temp.append("{0}-{1}".format(year,month))
        if data.get('cgCardGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade1']))
        if data.get('cgCardGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade1']))
        if data.get('cgCardGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade1']))       
        if data.get('cgCardGrade4') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade4']))
        if data.get('cgCardGrade5') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade5']))
        if data.get('cgCardGrade6') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade6']))
        if data.get('cgCardGrade7') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade7']))
        if data.get('cgCardGrade7') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade7']))
        if data.get('cgCardGrade7') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade7']))
        if data.get('cgCardGrade7') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGrade7']))
        if data.get('cgCardGradeAvg') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['cgCardGradeAvg']))
        sheet1.append(temp)
    
    print(" -- 카드신용 완료")

    # 저축은행
    print(" -- 저축은행 시작")
    List_data = {
        '_JSON_': '%7B%22SORT_COLUMN%22%3A%22%22%2C%22SORT%22%3A%22%22%2C%22PRE_MONTH_MONEY%22%3A%22%22%2C%22SUBMIT_MONTH%22%3A%22{0}{1}%22%7D'.format(year,month)
    }
    with requests.Session() as s:
        request = s.post("https://www.fsb.or.kr/ratloanconf_0200.jct", data = List_data)

    jsonObjs = request.json()
    dataList = jsonObjs['REC']
    last_dic = {}
    all_list = []
    for last_data in dataList :
        temp = []
        temp.append(last_data['BANK_NAME'])
        temp.append("저축은행")
        temp.append("{0}-{1}".format(year,month))
        if last_data.get('A_RATE1') is not None:
            temp.append(float(last_data['A_RATE1']))
        else:
            temp.append("")
        if last_data.get('A_RATE2') is not None:
            temp.append(float(last_data['A_RATE2']))
        else:
            temp.append("")
        if last_data.get('A_RATE3') is not None:
            temp.append(float(last_data['A_RATE3']))
        else:
            temp.append("")
        if last_data.get('A_RATE4') is not None:
            temp.append(float(last_data['A_RATE4']))
        else:
            temp.append("")
        if last_data.get('A_RATE5') is not None:
            temp.append(float(last_data['A_RATE5']))
        else:
            temp.append("")
        if last_data.get('A_RATE6') is not None:
            temp.append(float(last_data['A_RATE6']))
        else:
            temp.append("")
        if last_data.get('A_RATE7') is not None:
            temp.append(float(last_data['A_RATE7']))
        else:
            temp.append("")
        if last_data.get('A_RATE8') is not None:
            temp.append(float(last_data['A_RATE8']))
        else:
            temp.append("")
        if last_data.get('A_RATE9') is not None:
            temp.append(float(last_data['A_RATE9']))
        else:
            temp.append("")
        if last_data.get('A_RATE10') is not None:
            temp.append(float(last_data['A_RATE10']))
        else:
            temp.append("")
        if last_data.get('A_RATE_AVE') is not None:
            temp.append(float(last_data['A_RATE_AVE']))
        else:
            temp.append("0")
        sheet1.append(temp)
    
    print(" -- 저축은행 완료")

    # 캐피탈
    print(" -- 캐피탈 시작")
    url = 'https://gongsi.crefia.or.kr/portal/creditloan/creditloanDisclosureDetail11/ajax?clgcMode=11&cardItem=13%2C12%2C502%2C619%2C103%2C580%2C158%2C44%2C39%2C154%2C40%2C130%2C134%2C41%2C25%2C156%2C6%2C55%2C32%2C58%2C52%2C59%2C61%2C57%2C62%2C'
    headers = {
                'Referer': 'https://gongsi.crefia.or.kr/portal/creditloan/creditloanDisclosureDetail11',
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
    }
    response = requests.get(url, headers=headers)
    jsonObjs = response.json()
    dataList = jsonObjs['resultList']
    all_list = []
    for data in dataList :
        temp = []
        temp.append(data['mcCompany'])
        temp.append("캐피탈")
        temp.append("{0}-{1}".format(year,month))
        if data.get('clgGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade1']))
        if data.get('clgGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade1']))
        if data.get('clgGrade1') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade1']))
        if data.get('clgGrade4') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade4']))
        if data.get('clgGrade5') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade5']))
        if data.get('clgGrade6') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade6']))
        if data.get('clgGrade10') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade10']))
        if data.get('clgGrade10') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade10']))
        if data.get('clgGrade10') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade10']))
        if data.get('clgGrade10') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgGrade10']))
        if data.get('clgInterestAvg') == "0.00":
            temp.append("")
        else:
            temp.append(float(data['clgInterestAvg']))
        sheet1.append(temp)
    
    print(" -- 캐피탈 완료")

    # 신차할부
    print(" -- 신차할부 시작")
    url = 'https://gongsi.crefia.or.kr/portal/quota/quotaFinancingDisclosureDetail1/detail1'
    data_list = {
    'ifgcMode' : '1',
    'ifgCompany' : '현대자동차',
    'ifgGoods' : '아반떼',
    'ifgPreRate' : '10',
    'ifgPeriod' : '36', 
    'ifgOrder' : 'ifgInterestRate'
    }
    headers = {
    'Cookie': '',
    'Referer': 'https://gongsi.crefia.or.kr/portal/quota/quotaFinancingDisclosureDetail1?ifgcMode=1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
    'X-CSRF-TOKEN': ''
    }

    with requests.session() as s:
        first_page = s.get("https://gongsi.crefia.or.kr/")
        html = first_page.text
        header = first_page.headers
        soup = BeautifulSoup(html,'html.parser')
        csrf = soup.find('input', {'name': '_csrf'})
        token = csrf['value']
        cookie = header['Set-Cookie'].split(';')[0]
        headers['Cookie'] = cookie
        headers['X-CSRF-TOKEN'] = token
        response = s.post(url, data=data_list,headers=headers)
        
    jsonObjs = response.json()
    dataList = jsonObjs['resultList']
    all_list = []
    for data in dataList :
        temp = []
        x = data['mcCompany']
        if data["ifgDiv"] == "2":
            x = x+" (D)*"
        value = float(round((data['ifgInterestRate1']+data['ifgInterestRate2'])/2,2))
        temp.append(x)
        temp.append("신차할부")
        temp.append("{0}-{1}".format(year,month))
        temp.append(value)
        temp.append("전분기 평균 실제금리 >>")
        temp.append(float(data['ifgResult']))
        sheet1.append(temp)
    
    print(" -- 신차할부 완료")

    # 중고차할부
    print(" -- 중고차할부 시작")
    grade = ['1등급','2등급','3등급','4등급','5등급','6등급','7등급','8등급','9등급','10등급']

    url = 'https://gongsi.crefia.or.kr/portal/quota/quotaFinancingDisclosureDetail2/detail2'
    headers = {
    'Cookie': '',
    'Referer': 'https://gongsi.crefia.or.kr/portal/quota/quotaFinancingDisclosureDetail1?ifgcMode=1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36',
    'X-CSRF-TOKEN': ''
    }
    company_list = {}
    all_list = []

    with requests.session() as s:
        first_page = s.get("https://gongsi.crefia.or.kr/")
        html = first_page.text
        header = first_page.headers
        soup = BeautifulSoup(html,'html.parser')
        csrf = soup.find('input', {'name': '_csrf'})
        token = csrf['value']
        cookie = header['Set-Cookie'].split(';')[0]
        headers['Cookie'] = cookie
        headers['X-CSRF-TOKEN'] = token
        number = 0
        for i in grade:
            number = number + 1
            data_list = {
                'ifgcMode': '2',
                'ifgCreditCompanyNm': 'NICE',
                'ifgCreditRatingNm': '%s' % i,
                'ifgPeriodNm': '36개월',
                'ifgCreditCompany': '1',
                'ifgCreditRating': '%d' % number,
                'ifgPeriod': '36',
                'ifgOrder': 'ifgInterestRate'
            }

            response = s.post(url, data=data_list,headers=headers)    

            jsonObjs = response.json()
            dataList = jsonObjs['resultList']
            for data in dataList :
                temp = {}
                x = data['mcCompany']
                if data["ifgDiv"] == "2":
                    x = x+" (D)*"
                company_list.update({x : {}})

        number = 0
        for i in grade:
            number = number + 1
            data_list = {
                'ifgcMode': '2',
                'ifgCreditCompanyNm': 'NICE',
                'ifgCreditRatingNm': '%s' % i,
                'ifgPeriodNm': '36개월',
                'ifgCreditCompany': '1',
                'ifgCreditRating': '%d' % number,
                'ifgPeriod': '36',
                'ifgOrder': 'ifgInterestRate'
            }
        
            response = s.post(url, data=data_list,headers=headers)    

            jsonObjs = response.json()
            dataList = jsonObjs['resultList']
            for data in dataList :
                temp = {}
                x = data['mcCompany']
                if data["ifgDiv"] == "2":
                    x = x+" (D)*"
                value = float(round((data['ifgInterestRate1']+data['ifgInterestRate2'])/2,2))
                temp = {i : value}
                
                company_list[x].update(temp)

    for i in company_list.keys():
        temp = []
        temp.append(i)
        temp.append("중고차할부")
        temp.append("{0}-{1}".format(year,month))
        for j in grade:
            if company_list[i].get(j) is not None:
                if company_list[i][j] == 0:
                    temp.append("")
                else:    
                    temp.append(company_list[i][j])
            else:
                temp.append("")
                
        sheet1.append(temp)


    print(" -- 중고차할부 완료")


    #사잇돌2
    print(" -- 사잇돌2 시작")
    List_data = {
        '_JSON_': '%7B%22SORT_COLUMN%22%3A%22%22%2C%22SORT%22%3A%22%22%2C%22SUBMIT_MONTH%22%3A%22202004%22%7D'.format(year_new,month_new)
    }
    with requests.Session() as s:
        request = s.post("https://www.fsb.or.kr/ratloanrage_0200.jct", data = List_data)

    jsonObjs = request.json()
    dataList = jsonObjs['REC']
    last_dic = {}
    all_list = []
    for last_data in dataList :
        temp = []
        temp.append(last_data['BANK_NAME']+" / "+last_data['PRODUCT_NAME'].replace("\n",""))
        temp.append("사잇돌2")
        temp.append("{0}-{1}".format(year,month))
        if last_data.get('CREDIT_1_3_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_1_3_AVE_RATE']))
        else:
            temp.append("")
        
        if last_data.get('CREDIT_1_3_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_1_3_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_1_3_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_1_3_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_4_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_4_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_5_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_5_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_6_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_6_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_7_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_7_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_8_10_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_8_10_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_8_10_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_8_10_AVE_RATE']))
        else:
            temp.append("")

        if last_data.get('CREDIT_8_10_AVE_RATE') is not None:
            temp.append(float(last_data['CREDIT_8_10_AVE_RATE']))
        else:
            temp.append("")    

        sheet1.append(temp)  
    print(" -- 사잇돌2 완료")
    print("------- 마무리중")
    sheet2.append(["",""])
    sheet2.append(["","*공시정보 설명*"])
    sheet2.append(["","- 신용카드는 장기카드대출을 가져옴"])
    sheet2.append(["","- 신차와 중고차 할부는 36개월 금리를 가져옴"])
    sheet2.append(["","- 신차할부는 각 업체별 최고금리 최저금리 평균값 가져옴"])
    sheet2.append(["","- 생명보험은 무증빙신용 금리만 가져옴"])
    yn = {"yn" : ""}
    yn["yn"] = input("작업이 종료되었습니다. 엑셀로 저장하시려면 엔터키를 입력해주세요.")
    if yn.get("yn") == "":
        wb.save('공시정보_crawling({0}-{1}).xlsx'.format(year,month))    
        pyautogui.alert('엑셀 파일 저장이 끝났습니다.')
    else:
        pyautogui.alert('작업이 취소되었습니다.')
else:
    pyautogui.alert('작업이 취소되었습니다.')
