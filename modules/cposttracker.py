import requests
import json
from datetime import datetime
import calendar
import argparse
import os
import sys
import pandas as pd
import time
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




def get_information(trackid):
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
    return text

def main():
    parser = argparse.ArgumentParser(description="Canada Post Tracker")
    parser.add_argument('-i', '--csvinput', type=str,help="CSV File Input")
    args = parser.parse_args()
    isExist = os.path.exists(args.csvinput)
    if not isExist:
        input(args.csvinput + " file does not exist")
        sys.exit()
    print("#"*10, "Canada Post Tracker", "#"*10)
    print("CSV file mounting..", end=" ", flush=True)
    data = pd.read_csv(args.csvinput)
    time.sleep(1)
    print("Success")
    try:
        data.drop(columns=['Status'])
    except:
        pass

    tracklist = []
    print("")
    print("Tracking Started..")
    for idx, d in data.iterrows():
        try:
            trackstatus = get_information(d['Tracking'])
            print(d['Tracking'], trackstatus)
            tracklist.append(trackstatus)
        except:
            print(d['Tracking'], 'Failed')

            tracklist.append("Failed")
        time.sleep(0.8)
    print("Tracking Finished..", end="\n\n")
        
    data['Status'] = tracklist
    print("Trying to save to file", args.csvinput + "...", end=" ", flush=True)
    data.to_csv(args.csvinput, index=False)
    time.sleep(2)
    print("Success")
    input("Finished....")
if __name__ == '__main__':
    main()