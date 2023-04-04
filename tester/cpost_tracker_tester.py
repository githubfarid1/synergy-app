import requests
import json
from datetime import datetime
import calendar
headers = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'en-US,en;q=0.9,ja-JP;q=0.8,ja;q=0.7,id;q=0.6',
    'Authorization': 'Basic Og==',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
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

trackids = (
'LA232005509CA',
'7322397190105430',
'LX061926861CA',
'LA232005883CA',
'LA232005897CA',
'LA232005906CA',
'LA227939168CA',
'LA227939171CA',
'LA227939185CA',
'LA227939199CA',
'LA227939242CA',
'LA227939256CA',
'LA227939327CA',
'LA227939335CA',
'LA227939344CA',
'LA227939358CA',
'LA227939361CA',
'LA227939375CA',
'LA227939389CA',
'LA227939392CA',
'LA227939401CA',
'LA227939415CA',
'LA227939579CA',
'LA227939582CA',
'LA227939596CA',
'LA227939605CA',
'LA227939619CA',
'LA227939843CA',
'LA232026954CA',
'LA232026968CA',
'LA232026971CA',
'LA232026985CA',
'LA232026999CA',
'LA232027019CA',
'LA232027098CA',
'LA232027107CA',
'LA232027115CA',
'LA232027124CA',
'LA232027138CA',
)
for trackid in trackids:
    response = requests.get(
        'https://www.canadapost-postescanada.ca/track-reperage/rs/track/json/package/{}/detail'.format(trackid),
        headers=headers,
    )
    
    data = json.loads(response.text)
    newest = data['events'][0]
    regcd = newest['locationAddr']['regionCd']
    if regcd == "":
        regcd = newest['locationAddr']['countryNmEn']
    datetime_str = newest['datetime']['date'] + " " + newest['datetime']['time']
    dt = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
    
    text = f"{calendar.month_abbr[dt.month]} {dt.day} {dt.strftime('%I:%M %p')} {newest['descEn']} {newest['locationAddr']['city'].capitalize() }, {regcd}"
    print(text)