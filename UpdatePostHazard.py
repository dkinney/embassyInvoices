# download the post hazard reports from the state department website

import requests
import pandas as pd
import numpy as np
import os
import sys
from datetime import datetime

sites = {
    'Moscow': {
        'countryCode': 1106,
        'postCode': 10196
    },
    'Kyiv': {
        'countryCode': 1113,
        'postCode': 10218
    },
    'Brussels': {
        'countryCode': 1075,
        'postCode': 10111
    },
    'Chisinau': {
        'countryCode': 1100,
        'postCode': 10178
    },
    'Hong Kong': {
        'countryCode': 1126,
        'postCode': 10261
    },
    'Guangzhou': {
        'countryCode': 1123,
        'postCode': 10255
    },
    'Shanghai': {
        'countryCode': 1123,
        'postCode': 10256
    },
    'Beijing': {
        'countryCode': 1123,
        'postCode': 10254
    },
    'Hanoi': {
        'countryCode': 1144,
        'postCode': 10314
    },
    'Ho Chi Minh City': {
        'countryCode': 1144,
        'postCode': 12631
    }
}

def getPreviousRateDates(countryCode=None, postCode=None):
    if countryCode is None:
        # print(f'Country Code is required')
        return None
    
    if postCode is None:
        # print(f'Post Code is required')
        return None
    
    previous = None

    url = r'https://aoprals.state.gov/Web920/location_action.asp?MenuHide=1'
    url += f'&CountryCode={countryCode}'
    url += f'&PostCode={postCode}'

    html = requests.get(url).content
    tables = pd.read_html(html)

    for table in tables:
        # look for the string "Previous Rates:" in the first column of the table
        searchString = 'Previous Rates:'
        data = table.loc[table[0] == searchString]

        if not data.empty:
            # the second column of the data has a list of dates for the previous rates
            # they need to be extracted and used to get the previous rates
            dateString = data.iloc[:, 1].values[0]

            if dateString is not None:
                n = 10
                previous = [dateString[i:i+n] for i in range(0, len(dateString), n)]
                break

    if previous is None:
        print(f'There are no previous rates available')
    
    return previous

def getPostHazardData(countryCode=None, postCode=None, effectiveDate=None) -> pd.DataFrame:
    if countryCode is None:
        # print(f'Country Code is required')
        return None
    
    if postCode is None:
        # print(f'Post Code is required')
        return None
    
    df = None

    url = r'https://aoprals.state.gov/Web920/location_action.asp?MenuHide=1'
    url += f'&CountryCode={countryCode}'
    url += f'&PostCode={postCode}'

    if effectiveDate is not None:
        url = url + f'&EffectiveDate={effectiveDate}'

    html = requests.get(url).content
    tables = pd.read_html(html)

    # access the table that has a header row with 'Post Name'
    for table in tables:
        if 'Post Name' in table.columns:
            df = table
            break
    
    if df is None:
        print(f'No table found with a header row containing "Post Name"')
        return None
    
    # df has the table with the relevant information
    # name the columns to make it easier to reference them
    df.columns = ['PostName', 'COLA', 'PostingRate', 'TransferZone', 'Footnote', 'HazardRate', 'EducationAllowance', 'LivingAllowance', 'ReportingSchedule']
    df = df.fillna(0)

    # the effective date is not in the table, so it needs to be added
    # if effectiveDate is not None and df['EffectiveDate'] is None:

    if effectiveDate is None:
        effectiveDate = pd.to_datetime('today').strftime('%Y-%m-%d')
    else:
        # the passed in date needs to be converted to a datetime object
        effectiveDate = pd.to_datetime(effectiveDate).strftime('%Y-%m-%d')

    df['EffectiveDate'] = effectiveDate

    df = df[['EffectiveDate', 'PostName', 'PostingRate', 'HazardRate']]
    return df

# main function
if __name__ == '__main__':
    ratesData = None
    earliestRateData = pd.to_datetime('2023-01-01')

    ratesFile = sys.argv[1] if len(sys.argv) > 1 else 'data/PostHazardRates.csv'

    # load existing data if the file exists
    if os.path.exists(ratesFile):
        # print(f'reading from {ratesFile}')
        ratesData = pd.read_csv(ratesFile)
        ratesData.sort_values(by='EffectiveDate', inplace=True)
    
    # print(f'Starting with data:')
    # print(ratesData)

    # create a timestamp for today
    now = datetime.now().strftime('%Y-%m-%d')
    nowDatetime = pd.to_datetime(now)

    for site in sites:
        print(f'Checking data for {site}')

        countryCode = sites[site]['countryCode']
        postCode = sites[site]['postCode']

        # is this the first time we are getting data for any site?
        if ratesData is None:
            print(f'Getting data for {site}')
            ratesData = getPostHazardData(countryCode, postCode)
            continue

        siteData = ratesData.loc[ratesData['PostName'] == site] if ratesData is not None else None

        if len(siteData) < 2:
            # force a second data point to be added to capture the range of dates
            print(f'Getting data for {site}')
            df = getPostHazardData(countryCode, postCode)
            ratesData = pd.concat([ratesData, df], ignore_index=True)
            siteData = ratesData.loc[ratesData['PostName'] == site]

        newest = siteData.loc[siteData['EffectiveDate'] == siteData['EffectiveDate'].max()]

        if newest.empty:
            print(f'No data found for {site}')
            df = getPostHazardData(countryCode, postCode)
            print(df)
            ratesData = pd.concat([ratesData, df], ignore_index=True)

        else:
            # we have some data for this site, did it change from the last time we checked?
            newestIndex = newest.index[0]
            newestDate = pd.to_datetime(newest['EffectiveDate'].values[0])

            if nowDatetime > newestDate:
                print(f'Getting data for {site}')
                df = getPostHazardData(countryCode, postCode)

                if df is not None:
                    comparison = np.where(df.values == newest.values, True, False)

                    if comparison.all():
                        pass
                    elif comparison[0][1] and comparison[0][2] and comparison[0][3]:
                        # update the newest ratesData with today's date
                        ratesData.loc[newestIndex, 'EffectiveDate'] = now
                    else:
                        # add the new data to the ratesData
                        ratesData = pd.concat([ratesData, df], ignore_index=True)

        previousDates = getPreviousRateDates(countryCode, postCode)

        for date in previousDates:
            datetime = pd.to_datetime(date)

            if datetime <= earliestRateData:
                break

            siteData = ratesData.loc[ratesData['PostName'] == site]

            if len(siteData) < 2:
                # force a second data point to be added to capture the range of dates
                print(f'Getting data for {site}')
                df = getPostHazardData(countryCode, postCode)
                ratesData = pd.concat([ratesData, df], ignore_index=True)
                siteData = ratesData.loc[ratesData['PostName'] == site]

            newest = siteData.loc[siteData['EffectiveDate'] == siteData['EffectiveDate'].max()]
            newestDate = pd.to_datetime(newest['EffectiveDate'].values[0]) if not newest.empty else nowDatetime

            oldest = siteData.loc[siteData['EffectiveDate'] == siteData['EffectiveDate'].min()]
            oldestDate = pd.to_datetime(oldest['EffectiveDate'].values[0]) if not oldest.empty else nowDatetime

            if oldest.empty:
                print(f'No data found for {site}')
                df = getPostHazardData(countryCode, postCode, date)
                ratesData = pd.concat([ratesData, df], ignore_index=True)
            else:
                oldestIndex = oldest.index[0]

                if datetime <= newestDate and datetime >= oldestDate:
                    pass
                else:
                    df = getPostHazardData(countryCode, postCode, date)

                    if df is not None:
                        dateFormatted = pd.to_datetime(date).strftime('%Y-%m-%d')
                        comparison = np.where(df.values == oldest.values, True, False)

                        if comparison.all():
                            pass
                        elif comparison[0][1] and comparison[0][2] and comparison[0][3]:
                            ratesData.loc[oldestIndex, 'EffectiveDate'] = dateFormatted
                        else:
                            ratesData = pd.concat([ratesData, df], ignore_index=True)

            # save the data to a file
            ratesData.sort_values(by=['PostName', 'EffectiveDate'], ascending=[True, False] ,inplace=True)
            
            with open(ratesFile, 'w') as file:
                ratesData.to_csv(file, index=False)