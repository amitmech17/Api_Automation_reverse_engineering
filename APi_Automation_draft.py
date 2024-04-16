import sys
import time

import requests
from solvers.svgcaptcha import solver
import subprocess
import logging
import json
from sheet2dict import Worksheet
import glob
import os
import datetime
import traceback

logging.basicConfig(level=logging.INFO, filename="Api_Automation_Draft.log", format="%(asctime)s %(levelname)s %(message)s")

relation_type_dict = {'SON OF':1, 'DAUGHTER OF':2, 'CARE OF':3, 'WIFE OF':4}
crop_dict = {'Finger Millet (Ragi/Mandika)': '041324700', 'African Sarson': '010900200', 'Bajra - Napier Hybrid': '010920900', 'Bajra Napier Grass': '010919200', 'Barley (Jau)': '010402100', 'Barnyard Millet (Kundiraivlli/Sawan': '011502200', 'Barseem': '010902300', 'Bengal Gram (Chana)': '011902900', 'Bengal Gram (Chana) - IR': '011902901', 'Bengal Gram (Chana) - RF': '011902902', 'Birdwood Grass': '010903400', 'Black Gram (Urad)': '011903700', 'Black Gram (Urad) - IR': '011903701', 'Black Gram (Urad) - RF': '011903702', 'Brown Sarson': '011604400', 'Buck Wheat (Kaspat)': '011504600', 'Buffel Grass (Anjan Grass)': '010904700', 'Castor (Rehri, Rendi, Arandi) - RF': '011605801', 'Common Millet (Panivaragu/Chena/Proso Millet/Hogm': '011507800', 'Cowpea': '011908100', 'Dhaincha': '011109000', 'Dharaf Grass': '010909100', 'Dinanath Grass': '010909300', 'Faba Bean (Horse Bean/Windsor Bean)': '011909900', 'Finger Millet (Ragi/Mandika)&nbsp;(Rabi Summer&nbsp;)': '041326000', 'Fingermillet (Ragi/Mandika)': '011510300', 'Fingermillet (Ragi/Mandika) - Hills': '011510301', 'Fingermillet (Ragi/Mandika) - IR': '011510302', 'Fingermillet (Ragi/Mandika) - RF': '011510304', 'Fingermillet (Ragi/Mandika) - Summer IR': '011510303', 'Fodder Maize': '010910500', 'Fodder Sorghum': '010910600', 'Golden Thimothy': '010911600', 'Green Gram (Moong)': '011912000', 'Green Gram (Moong) - IR': '011912001', 'Green Gram (Moong) - RF': '011912002', 'Groundnut (Pea Nut)': '011612100', 'Groundnut (Pea Nut) - IR': '011612101', 'Groundnut (Pea Nut) - RF': '011612102', 'Groundnut (Pea Nut) - Summer': '011612103', 'Groundnut (Pea Nut) - Summer IR': '011612104', 'Groundnut (Pea Nut) - Summer RF': '011612105', 'Grow bag': '041317900', 'Guar': '010912200', 'Guinea Grass': '010912400', 'Homestead Farming/Mixed cropping': '041318900', 'Horse Gram (Kulthi/Kultha)': '011913000', 'Horse Gram (Kulthi/Kultha) - RF': '011913001', 'Italian Millet (Thenai/Navane/Foxtail Millet/Kang)': '011513600', 'Italian Millet (Thenai/Navane/Foxtail Millet/Kang) - RF': '011513601', 'Italian Millet (Thenai/Navane/Foxtail Millet/Kang)(Rabi Summer)': '041326200', 'Italian Millet (Thenai/Navane/Foxtail Millet/Kang)(Rabi Winter&nbsp;)': '041324900', 'Jojoba': '011614200', 'Karan Rai': '011614500', 'Khesari (Chickling Vetch/ Grass Pea)': '011914600', 'Kodo Millet (Kodara/Varagu)': '011515000', 'Kolanchi(Tephrosia Purpurea)': '011115200', 'Korra': '011515400', 'Lentil (Masur)': '011916200', 'Lethyrus': '011916400', 'Linseed (Alsi)': '011616700', 'Linseed (Alsi) - RF': '011616701', 'Little Millet (Samai/Kutki/Kodo-Kutki)': '011516900', 'Little Millet (Samai/Kutki/Kodo-Kutki)(Rabi Summer)': '041326400', 'Little Millet (Samai/Kutki/Kodo-Kutki)(Rabi Winter)': '041325100', 'Lucerne (Alfalfa)': '010917300', 'Maize (Makka)': '011517400', 'Maize (Makka) - I': '011517405', 'Maize (Makka) - II': '011517406', 'Maize (Makka) - III': '011517407', 'Maize (Makka) - IR': '011517401', 'Maize (Makka) - Rabi': '011517403', 'Maize (Makka) - RF': '011517402', 'Maize (Makka) - Summer': '011517404', 'Marvel Grass': '010917900', 'Mehandi/Henna tree': '041319800', 'Mesta': '011118200', 'Mochai (Lab-Lab)': '011918500', 'Moth Bean (Kidney Bean/ Deww Gram)': '011918700', 'Mustard': '011619100', 'Niger (Ramtil)': '011619400', 'Oats': '010419600', 'Olive': '011619800', 'Other Vegetables': '041318800', 'Paddy - Aahu': '010420204', 'Paddy - Aman': '010420205', 'Paddy - Aus': '010420206', 'Paddy - Autumn': '010420207', 'Paddy - Boro': '010420208', 'Paddy - Hills': '010420215', 'Paddy - I': '010420203', 'Paddy - II': '010420212', 'Paddy - III': '010420214', 'Paddy - IR': '010420201', 'Paddy - RF (UnIR)': '010420202', 'Paddy - Sali': '010420209', 'Paddy - Summer': '010420210', 'Paddy - Summer IR': '010420213', 'Paddy (Dhan)': '010420200', 'Paddy High Yielding Variety': '010420211', 'Paddy(Local)': '041318000', 'Pea': '011920600', 'Pearl Millet (Bajra)': '011501900', 'Pearl Millet (Bajra) - IR': '011501901', 'Pearl Millet (Bajra) - RF': '011501902', 'Persian Clover': '010921400', 'Pigeon Pea (Red Gram/Arhar/Tur)': '011921700', 'Pigeon Pea (Red Gram/Arhar/Tur) - IR': '011921701', 'Pigeon Pea (Red Gram/Arhar/Tur) - RF': '011921702', 'Pillipesara': '011121800', 'Pulses': '011922600', 'Rajka Bajri': '010923100', 'Rajma (French Bean)': '011923200', 'Rajmash Bean': '011923300', 'Raya (Indian Mustard)': '011623400', 'Red Clover': '010923600', 'Rice': '010431500', 'Rice Fallow Black Gram': '011923800', 'Rice Fallow Cow Pea': '011924000', 'Rice Fallow Gingelly': '011624100', 'Rice Fallow Green Gram': '011924200', 'Rice Fallow Red Gram': '011924300', 'Ricebean': '010924400', 'Rocket Salad (Taramira)': '011624600', 'Ryegrass': '010925100', 'Safflower (Kusum/Kardi)': '011625700', 'Safflower (Kusum/Kardi) - RF': '011625701', 'Save - IR': '011526401', 'Save - RF': '011526402', 'Sen Grass': '010926500', 'Sesame (Gingelly/Til)/Sesamum': '011626700', 'Sesame (Gingelly/Til)/Sesamum - RF': '011626701', 'Setaria Grass': '010926800', 'Sorghum (Jowar/Great Millet)': '011527300', 'Sorghum (Jowar/Great Millet) - IR': '011527301', 'Sorghum (Jowar/Great Millet) - RF': '011527302', 'Sorghum (Jowar/Great Millet)(Rabi Summer&nbsp;)': '041326600', 'Sorghum (Jowar/Great Millet)(Rabi Winter&nbsp;)': '041325300', 'Soybean (Bhat)': '011627400', 'Soybean (Bhat) - IR': '011627401', 'Soybean (Bhat) - RF': '011627402', 'Stylosanthes': '010928000', 'Sudan Grass': '010928100', 'Sunflower (Suryamukhi)': '011628500', 'Sunflower (Suryamukhi) - IR': '011628501', 'Sunflower (Suryamukhi) - RF': '011628502', 'Sunflower (Suryamukhi) - Summer IR': '011628503', 'Sunnhemp (Patua)': '011128600', 'Tall Fescue Grass': '010928900', 'Teosinte': '010929300', 'Thysanolaena ( Broom Plant/ Grass)': '041326500', 'Toria': '011629600', 'Triticale': '010429700', 'Velimasal': '010930500', 'Wheat': '010430800', 'Wheat - Hills': '010430803', 'Wheat - IR': '010430801', 'Wheat - RF': '010430802', 'White Clover (Shaftal)': '010930900', 'Winged Bean': '011931200'}
season_code_dict = {"KHARIF": 1, "RABI": 2, "SUMMER/ZAID/OTHERS": 3}
land_type_dict = {"IRRIGATED": 1, "NON-IRRIGATED": 2}
username = "7988448331"
password = "Keshav@123"
year = "2022-2023"



def draft_section(year, account_number):
    # draft
    url = "https://fasalrin.gov.in:443/application/getLoanApplicationList?offset=0&limit=10&searchString={}&displayID=&pacID=&searchBy=accountNumber&financialYear={}&status=0".format(
        account_number, year)
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/applicationstatus",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def approve_section():
    # approved status=2
    # rejected status=3
    # approval
    url = "https://fasalrin.gov.in:443/application/getLoanApplicationList?offset=0&limit=10&searchString=&displayID=&pacID=&searchBy=accountNumber&financialYear=2022-2023&status=1"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/applicationstatus",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def all_data_loan(loan_id, year):
    url = "https://fasalrin.gov.in:443/application/getBeneficiaryDataById?loanFinancialDetailsID={}&financialYear={}".format(
        loan_id, year)
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/applicationstatus",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def get_state():
    # state
    url = "https://fasalrin.gov.in:443/locations/getStates"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/loanApplicationForm",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def get_district(state_id):
    # district
    url = "https://fasalrin.gov.in:443/locations/getDistrictFromState?stateID={}".format(state_id)
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/loanApplicationForm",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def get_sub_district(district_id):
    # sub district
    url = "https://fasalrin.gov.in:443/locations/getSubDistrictFromDistrict?districtID={}".format(district_id)
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/loanApplicationForm",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def get_village(state_id, district_id, sub_district_id):
    # get village
    url = "https://fasalrin.gov.in:443/locations/getVillageFromSubDistrict?subDistrictID={}&stateID={}&districtID={}".format(
        sub_district_id, state_id, district_id)
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
               "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/loanApplicationForm",
               "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
               "Priority": "u=1, i", "Connection": "close"}
    data = s.get(url, headers=headers).json()
    return dict(data)


def get_village_district_id(state, district, sub_district, village):
    all_data = {}
    state_api_response = get_state()
    state_id = None
    for state_api in state_api_response['data']:
        if state_api['stateName'].upper() == state.upper():
            state_id = state_api['stateID']
            break
    district_excel = district
    district_id = None
    district_api_response = get_district(state_id)
    for district_api in district_api_response['data']['hierarchy']['level3']:
        if district_api['level3Name'].upper() == district_excel.upper():
            district_id = district_api['level3ID']
            break
    if district_id is None:
        for district_api in district_api_response['data']['hierarchy']['level4']:
            if district_api['level4Name'].upper() == district_excel.upper():
                district_id = district_api['level4ID']
                break
    sub_district_excel = sub_district
    sub_district_id = None
    sub_district_api_response = get_sub_district(district_id)
    for sub_district_api in sub_district_api_response['data']['hierarchy']['level3']:
        if sub_district_api['level3Name'].upper() == sub_district_excel.upper():
            sub_district_id = sub_district_api['level3ID']
            break
    if sub_district_id is None:
        for sub_district_api in sub_district_api_response['data']['hierarchy']['level4']:
            if sub_district_api['level4Name'].upper() == sub_district_excel.upper():
                sub_district_id = sub_district_api['level4ID']
                break
    village_id = None
    village_excel = village
    village_api_response = get_village(state_id, district_id, sub_district_id)
    for village_api in village_api_response['data']['hierarchy']:
        if village_api['labels'].upper() == village_excel.upper():
            village_id = village_api['id']
            break
    all_data['state_id'] = state_id
    all_data['district_id'] = district_id
    all_data['sub_district_id'] = sub_district_id
    all_data['village_id'] = village_id
    return all_data


def age_calc(dob):
    years, month, day = map(int, dob.split("-"))
    today = datetime.date.today()
    age = today.year - years - ((today.month, today.day) < (month, day))
    return age


def first_update(all_data,excel_data, r_state, r_state_id, r_district, r_district_id, r_sub_district, r_sub_district_id, r_village, r_village_id):
    data = {}
    data['aadharNumber'] = all_data['data']['aadharNumber']
    data['dob'] = all_data['data']['dob']
    if all_data['data']['age'] == 0:
        data['age'] = age_calc(data['dob'])
    else:
        data['age'] = all_data['data']['age']
    data['beneficiaryID'] = all_data['data']['beneficiaryID']
    data['beneficiaryName'] = all_data['data']['beneficiaryName']
    data['beneficiaryPassbookName'] = all_data['data']['beneficiaryPassbookName']
    data['casteCategory'] = all_data['data']['casteCategory']
    data['farmerCategory'] = all_data['data']['farmerCategory']
    data['farmerType'] = all_data['data']['farmerType']
    data['financialDetailsID'] = all_data['data']['financialDetailsID']
    data['gender'] = all_data['data']['gender']
    data['idType'] = all_data['data']['idType']
    data['isMobileNoUpdated'] = 0
    if all_data['data']['mobile'] is None:
        data['mobile'] = "9100000000"
    else:
        data['mobile'] = all_data['data']['mobile']
    data['relation'] = relation_type_dict[excel_data['Relation Type']]
    data['relativeName'] = excel_data['Relative Name']
    data['resAddress'] = all_data['data']['resAddress']
    data['resDistrictId'] = r_district_id
    data['resDistrictName'] = r_district
    data['resPincode'] = all_data['data']['resPincode']
    data['resStateId'] = r_state_id
    data['resStateName'] = r_state
    data['resSubDistrictId'] = r_sub_district_id
    data['resSubDistrictName'] = r_sub_district
    data['resVillageId'] = r_village_id
    data['resVillageName'] = r_village
    data['socialCategory'] = all_data['data']['socialCategory']

    url = "https://fasalrin.gov.in:443/application/updateBeneficiaryDetails"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
               "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
               "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in", "Sec-Fetch-Site": "same-origin",
               "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
               "Referer": "https://fasalrin.gov.in/loanApplicationForm", "Accept-Encoding": "gzip, deflate, br",
               "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i", "Connection": "close"}

    # d_json = dict(json.dumps({"data": data}))
    d_json = {"data": {"aadharNumber": "{}".format(data['aadharNumber']), "age": int(data['age']),
                           "beneficiaryID": "{}".format(data['beneficiaryID']), "beneficiaryName": "{}".format(data['beneficiaryName']),
                           "beneficiaryPassbookName": "{}".format(data['beneficiaryPassbookName']), "casteCategory": int(data['casteCategory']), "dob": "{}".format(data['dob']),
                           "farmerCategory": int(data['farmerCategory']), "farmerType": int(data['farmerType']),
                           "financialDetailsID": "{}".format(data['financialDetailsID']), "gender": int(data['gender']), "idType": int(data['idType']),
                           "isMobileNoUpdated": int(data['isMobileNoUpdated']), "mobile": int(data['mobile']), "relation": int(data['relation']), "relativeName": "{}".format(data['relativeName']),
                           "resAddress": "{}".format(data['resAddress']),
                           "resDistrictId": "{}".format(data['resDistrictId']), "resDistrictName": "{}".format(data['resDistrictName']),
                           "resPincode": int(data['resPincode']), "resStateId": "{}".format(data['resStateId']),
                           "resStateName": "{}".format(data['resStateName']), "resSubDistrictId": "{}".format(data['resSubDistrictId']),
                           "resSubDistrictName": "{}".format(data['resSubDistrictName']), "resVillageId": "{}".format(data['resVillageId']),
                           "resVillageName": "{}".format(data['resVillageName']), "socialCategory": int(data['socialCategory'])}}
    data = s.put(url, headers=headers, json=d_json)
    response = dict(data.json())
    if data.status_code == 200 and response['status'] is True:
        return response
    else:
        logging.info("====== error in first update")
        print(response)
        sys.exit(1)

def second_update(all_data, r_state_id, r_district_id):
    data = {}
    data['stateId'] = r_state_id
    data['districtId'] = r_district_id
    data['bankId'] = all_data['data']['bankID']
    data['branchId'] = all_data['data']['fdBranchID']
    data['accountNumber'] = all_data['data']['accountNumber']
    data['accountHolder'] = all_data['data']['accountHolder']
    if data['accountHolder'] == 1:
        data['accountHolderCount'] = 0
    else:
        data['accountHolderCount'] = all_data['data']['accountHolderCount']
    data['beneficiaryId'] = all_data['data']['beneficiaryID']
    data['financialDetailsID'] = all_data['data']['financialDetailsID']
    data['isSearchByIfsc'] = 0
    data['isJointAccountHolderAadhaarMissingUpdated'] = 0

    url = "https://fasalrin.gov.in:443/application/updateAccountDetails"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                     "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                     "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in",
                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                     "Referer": "https://fasalrin.gov.in/loanApplicationForm", "Accept-Encoding": "gzip, deflate, br",
                     "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i"}
    d_json = {"data": {"accountHolder": int(data['accountHolder']), "accountHolderCount": data['accountHolderCount'], "accountNumber": "{}".format(data['accountNumber']),
                           "bankId": "{}".format(data['bankId']),
                           "beneficiaryId": "{}".format(data['beneficiaryId']),
                           "branchId": "{}".format(data['branchId']),
                           "districtId": "{}".format(data['districtId']),
                           "financialDetailsID": "{}".format(data['financialDetailsID']),
                           "isJointAccountHolderAadhaarMissingUpdated": int(data['isJointAccountHolderAadhaarMissingUpdated']), "isSearchByIfsc": int(data['isSearchByIfsc']),
                           "stateId": "{}".format(data['stateId'])}}
    data = s.put(url, headers=headers, json=d_json)
    response = dict(data.json())
    if data.status_code == 200 and response['status'] is True:
        return response
    else:
        logging.info("====== error in second update")
        print(response)
        sys.exit(1)


def third_update(all_data,excel_data, fin_year):
    data = {}
    if excel_data['KCC Loan Sanction amount'] == "":
        data['loanSanctionAmount'] = all_data['data']['loanSanctionedAmount']
    else:
        data['loanSanctionAmount'] = excel_data['KCC Loan Sanction amount']
    if excel_data['KCC loan sanctioned / KCC renewed on'] == "":
        data['sanctionedDate'] = all_data['data']['loanSanctionedDate']
    else:
        data['sanctionedDate'] = excel_data['KCC loan sanctioned / KCC renewed on'].replace("/", "-")
    data['overduePaymentDate'] = all_data['data']['overduePaymentDate']
    data['loanTenure'] = all_data['data']['loanTenure']
    data['drawingLimit'] = all_data['data']['drawingLimit']

    data['financialDetailsId'] = all_data['data']['financialDetailsID']
    data['loanFinancialDetailsID'] = all_data['data']['loanFinancialDetailsID']
    data['financialYear'] = fin_year

    url = "https://fasalrin.gov.in:443/application/updateLoanDetails"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                     "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                     "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in",
                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                     "Referer": "https://fasalrin.gov.in/loanApplicationForm", "Accept-Encoding": "gzip, deflate, br",
                     "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i", "Connection": "close"}
    d_json = {"data": {"drawingLimit": int(data['drawingLimit']), "financialDetailsId": "{}".format(data['financialDetailsId']),
                           "financialYear": "{}".format(data['financialYear']),
                           "loanFinancialDetailsID": "{}".format(data['loanFinancialDetailsID']),
                           "loanSanctionAmount": int(data['loanSanctionAmount']), "loanTenure": int(data['loanTenure']), "overduePaymentDate": "{}".format(data['overduePaymentDate']),
                           "sanctionedDate": "{}".format(data['sanctionedDate'])}}
    data = s.put(url, headers=headers, json=d_json)
    response = dict(data.json())
    if data.status_code == 200 and response['status'] is True:
        return response
    else:
        logging.info("====== error in third update")
        print(response)
        sys.exit(1)


def fourth_update(all_data,excel_data, fin_year, row_data):
    #activities agri = "94BC9CB5-DB21-4496-90F0-C628DD619571", horti = "CAE9DAB5-9184-4C2A-936C-68797519F38B" , animal = "7C529CA3-7278-4E73-8F6F-BE4E6941D409", fish = "DC4EE6F0-EB72-4B58-B2CC-4132DB0B72D5"

    data = {}
    data['financialYear'] = fin_year
    data['beneficiaryID'] = all_data['data']['beneficiaryID']
    data['loanFinancialDetailsID'] = all_data['data']['loanFinancialDetailsID']
    data['activity'] = []
    temp_data = {}
    temp_data['activityTypeID'] = "94BC9CB5-DB21-4496-90F0-C628DD619571"
    temp_data['activityData'] = {}
    temp_data['activityData']['loanSanctionedAmount'] = all_data['data']['drawingLimit']
    if excel_data['Sanction/Rollover Date'] == "":
        temp_data['activityData']['loanSanctionedDate'] = all_data['data']['loanSanctionedDate']
    else:
        temp_data['activityData']['loanSanctionedDate'] = excel_data['Sanction/Rollover Date'].replace("/", "-")
    temp_data['activityData']['rows'] = row_data
    data['activity'].append(temp_data)
    data['applicationNo'] = all_data['data']['applicationNo']


    url = "https://fasalrin.gov.in:443/application/updateActivityDetails"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                     "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                     "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in",
                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                     "Referer": "https://fasalrin.gov.in/loanApplicationForm", "Accept-Encoding": "gzip, deflate, br",
                     "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i"}

    d_json = {"data": {"activity": [{"activityData": {"loanSanctionedAmount": int(temp_data['activityData']['loanSanctionedAmount']),
                                                          "loanSanctionedDate": "{}".format(temp_data['activityData']['loanSanctionedDate']), "rows": row_data},
                                         "activityTypeID": "{}".format(temp_data['activityTypeID'])}],
                           "applicationNumber": "{}".format(data['applicationNo']),
                           "beneficiaryID": "{}".format(data['beneficiaryID']), "financialYear": "{}".format(data['financialYear']),
                           "loanFinancialDetailsID": "{}".format(data['loanFinancialDetailsID'])}}
    data = s.put(url, headers=headers, json=d_json)
    response = dict(data.json())
    if data.status_code == 200 and response['status'] is True:
        return response
    else:
        logging.info("====== error in fourth update")

        print(response)
        sys.exit(1)


def final(fin_year, all_data):
    url = "https://fasalrin.gov.in:443/application/submitLoanApplication"
    headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                     "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                     "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in",
                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                     "Referer": "https://fasalrin.gov.in/applicationpreview", "Accept-Encoding": "gzip, deflate, br",
                     "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i", "Connection": "close"}
    d_json = {"data": {"actionType": "SUBMIT", "financialYear": "{}".format(fin_year),
                           "loanFinancialDetailsID": "{}".format(all_data['data']['loanFinancialDetailsID'])}}
    data = s.put(url, headers=headers, json=d_json)
    response = dict(data.json())
    if data.status_code == 200 and response['status'] is True:
        return response
    else:
        logging.info("====== error in final update")
        print(response)
        sys.exit(1)


def read_data():
    ws = Worksheet()
    ws.xlsx_to_dict(path="{}".format(glob.glob('*.xlsx')[0]))
    data = ws.sheet_items
    new_data = {}
    file_content = []
    failed_file_content = []
    if os.path.isfile("./completed_list.txt"):
        with open("./completed_list.txt") as f:
            file_content = f.readlines()
        file_content = [i.split(",")[0].strip() for i in file_content]
    if os.path.isfile("./Failed_list.txt"):
        with open("./Failed_list.txt") as f:
            failed_file_content = f.readlines()
        failed_file_content = [i.split(",")[0].strip() for i in failed_file_content]
    file_content = file_content + failed_file_content
    for i in data:
        if len(i['Account No']) < 14:
            continue
        if i['Account No'] in file_content:
            continue
        elif i['Account No'] in new_data.keys():
            temp = new_data[i['Account No']]
            temp.append(i)
        else:
            temp = []
            temp.append(i)
        new_data[i['Account No']] = temp
    return new_data


def main_run():
    all_excel_data = read_data()
    try:
        if len(all_excel_data) == 0:
            logging.info("there is no data to work on")
            sys.exit(1)
        # Run the JavaScript file using Node.js
        result = subprocess.run(["node", "js_encryption.js", username, password], capture_output=True, text=True)
        encrypted_number = None
        encrypted_pass = None
        # Check if the process ran successfully
        if result.returncode == 0:
            # Print the stdout (console output) of the JavaScript file
            print("Output:")
            result = result.stdout.split("\n")
            encrypted_number = result[0]
            encrypted_pass = result[1]
        else:
            # Print any error messages if the process failed
            print("Error:", result.stderr)
        burp0_url = "https://fasalrin.gov.in:443/login"
        burp0_headers = {"Cache-Control": "max-age=0", "Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"",
                         "Sec-Ch-Ua-Mobile": "?0", "Sec-Ch-Ua-Platform": "\"macOS\"", "Upgrade-Insecure-Requests": "1",
                         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                         "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                         "Sec-Fetch-Site": "none", "Sec-Fetch-Mode": "navigate", "Sec-Fetch-User": "?1",
                         "Sec-Fetch-Dest": "document", "Accept-Encoding": "gzip, deflate, br",
                         "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=0, i", "Connection": "close"}
        s.get(burp0_url, headers=burp0_headers)

        while True:
            burp0_url = "https://fasalrin.gov.in:443/services/getServerCaptcha"
            burp0_headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                             "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                             "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                             "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
                             "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/login",
                             "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
                             "Priority": "u=1, i", "Connection": "close"}
            captcha = s.get(burp0_url, headers=burp0_headers).json()
            captcha = dict(captcha)
            token = captcha['data']['token']
            svg = captcha['data']['stream']
            try:
                result = solver.solve_captcha(svg)
                burp0_url = "https://fasalrin.gov.in:443/services/verifyServerCaptcha?token={}&text={}".format(token, result)
                burp0_headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                                 "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                                 "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                                 "Sec-Ch-Ua-Platform": "\"macOS\"", "Origin": "https://fasalrin.gov.in",
                                 "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                                 "Referer": "https://fasalrin.gov.in/login", "Accept-Encoding": "gzip, deflate, br",
                                 "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i", "Connection": "close"}
                captcha_result = s.post(burp0_url, headers=burp0_headers)
                captcha_result = dict(captcha_result.json())
                if captcha_result['data'] == 'Captcha verified successfully':
                    burp0_url = "https://fasalrin.gov.in:443/auth/login"
                    burp0_headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"",
                                     "Accept": "application/json", "Sec-Ch-Ua-Platform": "\"macOS\"", "Devicetype": "web",
                                     "Sec-Ch-Ua-Mobile": "?0",
                                     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                                     "Content-Type": "application/json", "Origin": "https://fasalrin.gov.in",
                                     "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors", "Sec-Fetch-Dest": "empty",
                                     "Referer": "https://fasalrin.gov.in/login", "Accept-Encoding": "gzip, deflate, br",
                                     "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8", "Priority": "u=1, i",
                                     "Connection": "close"}
                    burp0_json = {"otp": "",
                                  "password": encrypted_pass,
                                  "username": encrypted_number}
                    login = s.post(burp0_url, headers=burp0_headers, json=burp0_json)
                    if login.status_code == 200:
                        print("Logged in successfully")
                    else:
                        print("Logged in Failed")
                    break
            except:
                print("error in captcha solve")

        burp0_url = "https://fasalrin.gov.in:443/application/getApplicationsCount?financialYear={}".format(year)
        burp0_headers = {"Sec-Ch-Ua": "\"Not(A:Brand\";v=\"24\", \"Chromium\";v=\"122\"", "Accept": "application/json",
                         "Devicetype": "web", "Sec-Ch-Ua-Mobile": "?0",
                         "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.6261.112 Safari/537.36",
                         "Sec-Ch-Ua-Platform": "\"macOS\"", "Sec-Fetch-Site": "same-origin", "Sec-Fetch-Mode": "cors",
                         "Sec-Fetch-Dest": "empty", "Referer": "https://fasalrin.gov.in/dashboard",
                         "Accept-Encoding": "gzip, deflate, br", "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8",
                         "Priority": "u=1, i", "Connection": "close"}
        application_data = s.get(burp0_url, headers=burp0_headers)
        application_data = dict(application_data.json())
        # "data":{"draft":4,"submitted":0,"approved":430,"rejected":0,"reviewRequired":0}
        if int(application_data["data"]["draft"]) > 0:
            for excel_row_key in all_excel_data.keys():
                try:
                    print(" ****** running for {}".format(excel_row_key))
                    excel_row_data = all_excel_data[excel_row_key]
                    loan_id = None
                    draft_data = draft_section(year, excel_row_key)
                    if int(draft_data['data']['count']['applicationCount']) > 0:
                        loan_id = draft_data['data']['list'][0]['loanFinancialDetailsID']
                    else:
                        logging.exception("there is no data for account no {}".format(excel_row_key))
                        with open("./Failed_list.txt", 'a+') as f:
                            f.write('{}\n'.format(excel_row_key))
                        logging.info("Data pushed failed for {}".format(excel_row_key))
                        continue
                    loan_data = all_data_loan(loan_id, year)
                    relation_name = excel_row_data[0]['Relative Name']
                    rel_type = excel_row_data[0]['Relation Type']
                    relation_type = relation_type_dict[rel_type]
                    state_excel = excel_row_data[0]['State']
                    state_api_response = get_state()
                    state_id = None
                    for state_api in state_api_response['data']:
                        if state_api['stateName'].upper() == state_excel.upper():
                            state_id = state_api['stateID']
                            break
                    district_excel = excel_row_data[0]['District']
                    district_id = None
                    district_api_response = get_district(state_id)
                    for district_api in district_api_response['data']['hierarchy']['level3']:
                        if district_api['level3Name'].upper() == district_excel.upper():
                            district_id = district_api['level3ID']
                            break
                    if district_id is None:
                        for district_api in district_api_response['data']['hierarchy']['level4']:
                            if district_api['level4Name'].upper() == district_excel.upper():
                                district_id = district_api['level4ID']
                                break
                    sub_district_excel = excel_row_data[0]['Block / Subdistrict']
                    sub_district_id = None
                    sub_district_api_response = get_sub_district(district_id)
                    for sub_district_api in sub_district_api_response['data']['hierarchy']['level3']:
                        if sub_district_api['level3Name'].upper() == sub_district_excel.upper():
                            sub_district_id = sub_district_api['level3ID']
                            break
                    if sub_district_id is None:
                        for sub_district_api in sub_district_api_response['data']['hierarchy']['level4']:
                            if sub_district_api['level4Name'].upper() == sub_district_excel.upper():
                                sub_district_id = sub_district_api['level4ID']
                                break
                    village_id = None
                    village_excel = excel_row_data[0]['Village']
                    village_api_response = get_village(state_id, district_id, sub_district_id)
                    for village_api in village_api_response['data']['hierarchy']:
                        if village_api['labels'].upper() == village_excel.upper():
                            village_id = village_api['id']
                            break
                    res = first_update(loan_data,excel_row_data[0], state_excel, state_id, district_excel, district_id, sub_district_excel, sub_district_id, village_excel, village_id)
                    print("first run done for {}".format(excel_row_key))
                    second_update(loan_data, state_id, district_id)
                    print("second run done for {}".format(excel_row_key))
                    third_update(loan_data,excel_row_data[0], year)
                    print("third run done for {}".format(excel_row_key))
                    village_data_set = {village_excel: [district_id, village_id]}
                    rows = []
                    for i in excel_row_data:
                        if i["Village"] in village_data_set.keys():
                            district_id, village_id = village_data_set[i["Village"]]
                        else:
                            all_data_tp = get_village_district_id(i["State"], i["District"], i["Block / Subdistrict"], i["Village"])
                            village_data_set[i["Village"]] = [all_data_tp['district_id'], all_data_tp['village_id']]
                            district_id, village_id = [all_data_tp['district_id'], all_data_tp['village_id']]
                        temp_excel = {}
                        temp_excel["cropCode"] = crop_dict[i["Crop"]]
                        temp_excel["surveyNumber"] = i["Survey Number"]
                        temp_excel["khataNumber"] = i["Khata Number"]
                        temp_excel["landArea"] = round(float(i["Land Area"]), 4)
                        temp_excel["landType"] = land_type_dict[i["Land Type"]]
                        temp_excel["seasonCode"] = season_code_dict[i["Season"]]
                        temp_excel["landDistrictID"] = district_id
                        temp_excel["landVillageID"] = village_id
                        rows.append(temp_excel)
                    fourth_update(loan_data,excel_row_data[0], year, rows)
                    print("forth run done for {}".format(excel_row_key))
                    final(year, loan_data)
                    print("final run done for {}".format(excel_row_key))
                    print("time took for one entry {}".format(time.process_time() - start))
                    with open("./completed_list.txt", 'a+') as f:
                        f.write('{}\n'.format(excel_row_key))
                        logging.info("Data pushed completed for {}".format(excel_row_key))
                except:
                    with open("./Failed_list.txt", 'a+') as f:
                        f.write('{}\n'.format(excel_row_key))
                    logging.info("Data pushed failed for {}".format(excel_row_key))
        else:
            logging.info("There is no data in draft section to work on")
    except Exception as e:
        logging.info("Error in main log = /n{}".format(traceback.format_exc()))


start = time.process_time()
s = requests.Session()

main_run()

