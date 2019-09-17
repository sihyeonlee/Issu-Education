# -*- coding: utf-8 -*-
from urllib.parse import quote_plus, urlencode
from urllib.request import urlopen, Request
import json

auth_key = '****'

dic_location = {'서울특별시': 'B10', '부산광역시': 'C10', '대구광역시': 'D10', '인천광역시': 'E10',
                '광주광역시': 'F10', '대전광역시': 'G10', '울산광역시': 'H10', '세종특별자치시': 'I10',
                '경기도': 'J10', '강원도': 'K10', '충청북도': 'M10', '충청남도': 'N10', '전라북도': 'P10',
                '전라남도': 'Q10', '경상북도': 'R10', '경상남도': 'S10', '제주도': 'T10'}


def get_juso(location_code, sc_name, key):
    dic_result = {'순번': [], '학교명': [], '우편번호': [], '도로명주소':[], }

    url = 'http://open.neis.go.kr/hub/schoolInfo'
    queryParams = '?' + urlencode(
        {quote_plus('KEY'): key, quote_plus('Type'): 'json', quote_plus('pIndex'): '1', quote_plus('pSize'): '100',
         quote_plus('ATPT_OFCDC_SC_CODE'): location_code, quote_plus('SCHUL_NM'): sc_name})

    req = Request(url + queryParams)
    req.get_method = lambda: 'GET'
    response_body = urlopen(req).read()

    root_json = json.loads(response_body)

    print(root_json)

    result_cnt = root_json['schoolInfo'][0]['head'][0]['list_total_count']
    if result_cnt > 0:
        data_body = root_json['schoolInfo'][1]['row']
    else:
        print("검색 결과 없음")

        return -1

    print(result_cnt)

    for i in range(0, result_cnt):
        dic_result['순번'].append(i + 1)
        dic_result['학교명'].append(data_body[i]['SCHUL_NM'])
        dic_result['우편번호'].append(data_body[i]['ORG_RDNZC'])
        dic_result['도로명주소'].append(data_body[i]['ORG_RDNMA'])

    print(dic_result)


get_juso(dic_location['서울특별시'], '신림', auth_key)

