import datetime as dt
import pandas as pd
import tkinter as tk
from tkinter import filedialog


states = {
        'AK': 'Alaska',
        'AL': 'Alabama',
        'AR': 'Arkansas',
        'AS': 'American Samoa',
        'AZ': 'Arizona',
        'CA': 'California',
        'CO': 'Colorado',
        'CT': 'Connecticut',
        'DC': 'District of Columbia',
        'DE': 'Delaware',
        'FL': 'Florida',
        'GA': 'Georgia',
        'GU': 'Guam',
        'HI': 'Hawaii',
        'IA': 'Iowa',
        'ID': 'Idaho',
        'IL': 'Illinois',
        'IN': 'Indiana',
        'KS': 'Kansas',
        'KY': 'Kentucky',
        'LA': 'Louisiana',
        'MA': 'Massachusetts',
        'MD': 'Maryland',
        'ME': 'Maine',
        'MI': 'Michigan',
        'MN': 'Minnesota',
        'MO': 'Missouri',
        'MP': 'Northern Mariana Islands',
        'MS': 'Mississippi',
        'MT': 'Montana',
        'NA': 'National',
        'NC': 'North Carolina',
        'ND': 'North Dakota',
        'NE': 'Nebraska',
        'NH': 'New Hampshire',
        'NJ': 'New Jersey',
        'NM': 'New Mexico',
        'NV': 'Nevada',
        'NY': 'New York',
        'OH': 'Ohio',
        'OK': 'Oklahoma',
        'OR': 'Oregon',
        'PA': 'Pennsylvania',
        'PR': 'Puerto Rico',
        'RI': 'Rhode Island',
        'SC': 'South Carolina',
        'SD': 'South Dakota',
        'TN': 'Tennessee',
        'TX': 'Texas',
        'UT': 'Utah',
        'VA': 'Virginia',
        'VI': 'Virgin Islands',
        'VT': 'Vermont',
        'WA': 'Washington',
        'WI': 'Wisconsin',
        'WV': 'West Virginia',
        'WY': 'Wyoming'
}


def normalizeState(state):
    upperState = state.upper()
    return states.get(upperState,upperState).upper()

def separate(style, half):
    for x in range(len(style)-1,0,-1):
        if(style[x] == "-"):
            if(half == 0):
                return style[:x]
            else:
                return style[x+1:]
    return style

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    filePath = filedialog.askopenfilename()
    while(filePath[len(filePath)-3:] != "csv"):
        print("this is not a csv file")
        filePath = filedialog.askopenfilename()
    raw = pd.read_csv(filePath, sep = ",", skiprows = 7,infer_datetime_format = True, parse_dates = [0])

    # take out all records which aren't orders
    raw.drop(raw[raw['type'] != 'Order'].index, inplace = True)

    # expand state abbreviations to full name
    #raw['order state'] = normalizeState(raw['order state'])
    raw['order state'] = raw.apply(lambda row: normalizeState(row['order state']), axis =1 )



    raw['style with color'] = raw.apply(lambda row: separate(row['sku'],0), axis = 1)

    raw['size'] = raw.apply(lambda row: separate(row['sku'],1), axis = 1)


    raw['style'] = raw.apply(lambda row: separate(row['style with color'],0), axis = 1)
    raw['color'] = raw.apply(lambda row: separate(row['style with color'],1), axis = 1)

    raw.drop(columns=['style with color','settlement id', 'type', 'order id', 'description', 'marketplace','account type', 'fulfillment', 'order city','order postal','tax collection model'], inplace = True)
    raw.to_csv("test.csv",index=False)
