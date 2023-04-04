import requests

cookies = {
    'CPC_Track_Searches': '7322397190105430%2CLA227939168CA',
    'OWSPRD002TRACK-REPERAGE': 'track-reperage_04028_s004ptom002',
    'at_check': 'true',
    's_vnc7': '1681177409459%26vn%3D1',
    's_ivc': 'true',
    'AMCVS_0C4E3704533345770A490D44%40AdobeOrg': '1',
    'OWSPRD003CWCCOMPONENTS': 'cwc_components_04027_s021ptom001',
    '_gcl_au': '1.1.466865068.1680572610',
    's_gpv_url': 'https%3A%2F%2Fwww.canadapost-postescanada.ca%2Ftrack-reperage%2Fen',
    'LANG': 'e',
    'LANG': 'e',
    'AMCV_0C4E3704533345770A490D44%40AdobeOrg': '-1124106680%7CMCIDTS%7C19452%7CMCMID%7C70710874911945993534367394569705780567%7CMCAAMLH-1681177409%7C3%7CMCAAMB-1681177409%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1680579810s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.2.0',
    's_lv_s': 'First%20Visit',
    's_cc': 'true',
    'ln_or': 'eyI5MTk4IjoiZCJ9',
    'QSI_HistorySession': 'https%3A%2F%2Fwww.canadapost-postescanada.ca%2Ftrack-reperage%2Fen%23%2Fhome~1680572618573',
    'cpc-gdpr-cookie-policy-accept': 'true',
    'gpv_v4': 'cpc.ca%3A%20%3E%20en%20%3E%20common%20%3E%20track%20%3E%20details%20%3E%20PIN',
    's_ppvl': 'cpc.ca%253A%2520%253E%2520en%2520%253E%2520common%2520%253E%2520track%2520%253E%2520home%2C47%2C32%2C767%2C807%2C527%2C1280%2C720%2C1.25%2CL',
    's_nr': '1680573347730-New',
    's_lv': '1680573347731',
    's_ppv': 'cpc.ca%253A%2520%253E%2520en%2520%253E%2520common%2520%253E%2520track%2520%253E%2520details%2520%253E%2520PIN%2C100%2C20%2C1847%2C1280%2C174%2C1280%2C720%2C1.25%2CL',
    'mbox': 'session#b924af8485c540fd8426de320faa2ff8#1680575238|PC#b924af8485c540fd8426de320faa2ff8.38_0#1743817411',
}

headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9,ja-JP;q=0.8,ja;q=0.7,id;q=0.6',
    'Authorization': 'Basic Og==',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
    # 'Cookie': 'CPC_Track_Searches=7322397190105430%2CLA227939168CA; OWSPRD002TRACK-REPERAGE=track-reperage_04028_s004ptom002; at_check=true; s_vnc7=1681177409459%26vn%3D1; s_ivc=true; AMCVS_0C4E3704533345770A490D44%40AdobeOrg=1; OWSPRD003CWCCOMPONENTS=cwc_components_04027_s021ptom001; _gcl_au=1.1.466865068.1680572610; s_gpv_url=https%3A%2F%2Fwww.canadapost-postescanada.ca%2Ftrack-reperage%2Fen; LANG=e; LANG=e; AMCV_0C4E3704533345770A490D44%40AdobeOrg=-1124106680%7CMCIDTS%7C19452%7CMCMID%7C70710874911945993534367394569705780567%7CMCAAMLH-1681177409%7C3%7CMCAAMB-1681177409%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1680579810s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.2.0; s_lv_s=First%20Visit; s_cc=true; ln_or=eyI5MTk4IjoiZCJ9; QSI_HistorySession=https%3A%2F%2Fwww.canadapost-postescanada.ca%2Ftrack-reperage%2Fen%23%2Fhome~1680572618573; cpc-gdpr-cookie-policy-accept=true; gpv_v4=cpc.ca%3A%20%3E%20en%20%3E%20common%20%3E%20track%20%3E%20details%20%3E%20PIN; s_ppvl=cpc.ca%253A%2520%253E%2520en%2520%253E%2520common%2520%253E%2520track%2520%253E%2520home%2C47%2C32%2C767%2C807%2C527%2C1280%2C720%2C1.25%2CL; s_nr=1680573347730-New; s_lv=1680573347731; s_ppv=cpc.ca%253A%2520%253E%2520en%2520%253E%2520common%2520%253E%2520track%2520%253E%2520details%2520%253E%2520PIN%2C100%2C20%2C1847%2C1280%2C174%2C1280%2C720%2C1.25%2CL; mbox=session#b924af8485c540fd8426de320faa2ff8#1680575238|PC#b924af8485c540fd8426de320faa2ff8.38_0#1743817411',
    'Pragma': 'no-cache',
    'Referer': 'https://www.canadapost-postescanada.ca/track-reperage/en',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
    'sec-ch-ua': '"Google Chrome";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Linux"',
}

response = requests.get(
    'https://www.canadapost-postescanada.ca/track-reperage/rs/track/json/package/7322397190105430/detail',
    cookies=cookies,
    headers=headers,
)

print(response.text)