import pandas as pd
import numpy as np
import datetime
import json
from openpyxl import load_workbook

'''
    1. Change the filename below
    2. And run the file!
    (Make sure that the file is in the same folder as this script)
'''
filename = 'patientgegevens_nov3.xlsx'


'''
    Don't change anything below this comment!
'''
datetimeFormat = '%Y-%m-%d %H:%M:%S'

def load_data(filename):
    '''
        Load data and convert into JSON format
    '''
    
    wb = load_workbook(filename)['slaap periodes']
    period_data = wb.values
    period_columns = next(period_data)[0:]
    periods = pd.DataFrame(period_data, columns=period_columns)
    periods.drop(['opmerking', 'locatie nauwkeurigheid (m)'], axis=1, inplace=True)
    periods.rename(columns = {'light aan':'licht aan'}, inplace = True)
    periods.sort_values(by=['licht aan'], ascending=True)
    periods_json = json.loads(periods.to_json(orient='records'))

    records = pd.read_excel(filename, sheet_name='slaapstaat')
    records.sort_values(by=['van'], ascending=False)
    records_json = json.loads(records.to_json(orient='records'))
       
    return periods_json, records_json


def merge(data):
    '''
        Function to merge equal consecutive sleep events
    '''
    
    merged = [data[0]]
    for i, j in zip(data[1:], data[2: ]):
        if i['tot'] == merged[-1]['tot'] and i['staat'] == merged[-1]['staat'] and i['patient_id'] == merged[-1]['patient_id']:
            continue
        elif i['patient_id'] == j['patient_id'] and i['staat'] == j['staat'] and i['tot'] == j['van']:
            obj = {'id': i['id'],
                           'patient_id': i['patient_id'],
                           'staat': i['staat'],
                           'van': i['van'],
                           'tot': j['tot']}
            merged.append(dict(obj))
        else:
            merged.append(i)
    
    return merged    


def combine(periods, records):
    '''
        Function to add sleep event sequence to sleep periods
    '''
    
    for i in periods:
        states = np.array([])
        state_count = 0
        
        for j in records:
            B = False
            C = False

            belongs_to_period = i['patient_id'] == j['patient_id'] and (
                (datetime.datetime.strptime(j['van'], datetimeFormat) <= datetime.datetime.strptime(i['licht uit'], datetimeFormat) and    # Overlap with lights off
                datetime.datetime.strptime(j['tot'], datetimeFormat) >= datetime.datetime.strptime(i['licht uit'], datetimeFormat)) or 
                (datetime.datetime.strptime(j['van'], datetimeFormat) >= datetime.datetime.strptime(i['licht uit'], datetimeFormat) and    # Between lights off and lights on
                datetime.datetime.strptime(j['van'], datetimeFormat) <= datetime.datetime.strptime(i['licht aan'], datetimeFormat)) or
                (datetime.datetime.strptime(j['tot'], datetimeFormat) >= datetime.datetime.strptime(i['licht uit'], datetimeFormat) and    # Overlap with lights on
                datetime.datetime.strptime(j['tot'], datetimeFormat) <= datetime.datetime.strptime(i['licht aan'], datetimeFormat)))

            if belongs_to_period:

                if (state_count == 0 and j['staat'] == 'outOfBed'):
                    continue
                else:
                    # Start before lights off                     
                    if ( datetime.datetime.strptime(j['van'], datetimeFormat) < datetime.datetime.strptime(i['licht uit'], datetimeFormat) ):
                        # End before lights off
                        if ( datetime.datetime.strptime(j['tot'], datetimeFormat) <= datetime.datetime.strptime(i['licht uit'], datetimeFormat) ):
                            if j['staat'] == 'outOfBed':
                                continue
                            state_dict = {'state': j['staat'] + '_on', 'start': j['van'], 'end': j['tot']}
                        # End between lights off and lights on
                        elif ( datetime.datetime.strptime(j['tot'], datetimeFormat) > datetime.datetime.strptime(i['licht uit'], datetimeFormat) and 
                               datetime.datetime.strptime(j['tot'], datetimeFormat) <=  datetime.datetime.strptime(i['licht aan'], datetimeFormat) ):
                            if j['staat'] == 'outOfBed':
                                continue 
                            B = True
                            for s in range(1,3):
                                if s == 1:   
                                    state_dict = {'state': j['staat'] + '_on', 'start': j['van'], 'end': i['licht uit']}
                                else:
                                    state_dict = {'state': j['staat'] + '_off', 'start': i['licht uit'], 'end': j['tot']}
                                states= np.append(states, dict(state_dict))
                        else:
                            C = True
                            for s in range(1,4):
                                if s == 1:   
                                    state_dict = {'state': j['staat'] + '_on', 'start': j['van'], 'end': i['licht uit']}
                                elif s == 2:
                                    state_dict = {'state': j['staat'] + '_off', 'start': i['licht uit'], 'end': i['licht aan']}
                                else:
                                    state_dict = {'state': j['staat'] + '_on', 'start': i['licht aan'], 'end': j['tot']}
                                states= np.append(states, dict(state_dict))

                    # Start after lights off, before lights on
                    elif ( datetime.datetime.strptime(j['van'], datetimeFormat) >= datetime.datetime.strptime(i['licht uit'], datetimeFormat) and
                             datetime.datetime.strptime(j['van'], datetimeFormat) < datetime.datetime.strptime(i['licht aan'], datetimeFormat)): 
                        # End voor lights on
                        if ( datetime.datetime.strptime(j['tot'], datetimeFormat) <= datetime.datetime.strptime(i['licht aan'], datetimeFormat) ): 
                            state_dict = {'state': j['staat'] + '_off', 'start': j['van'], 'end': j['tot']}
                        # End na lights on
                        elif ( datetime.datetime.strptime(j['tot'], datetimeFormat) > datetime.datetime.strptime(i['licht aan'], datetimeFormat) ):
                            B = True
                            for s in range(1,3):
                                if s == 1:   
                                    state_dict = {'state': j['staat'] + '_off', 'start': j['van'], 'end': i['licht aan']}
                                else:
                                    state_dict = {'state': j['staat'] + '_on', 'start': i['licht aan'], 'end': j['tot']}
                                states= np.append(states, dict(state_dict))

                    # Start op lights on
                    else: 
                        if j['staat'] == 'outOfBed':
                            continue
                        j['new_staat'] = j['staat'] + '_on'

                    if B == True:
                        state_count += 2
                    elif C == True:
                        state_count += 3
                    else:
                        state_count += 1
                        states= np.append(states, dict(state_dict))

        # Remove last event if it is outOfBed
        if (len(states) > 0):
            while (states[-1]['state'] == 'outOfBed_off' or states[-1]['state'] == 'outOfBed_on' ):
                states = states[:-1]

        i['states'] = sorted(states, key=lambda x: datetime.datetime.strptime(x['start'], datetimeFormat), reverse=False)

    return periods


def clear_data(data):
    '''
        Function to clear the data
    '''
    
    for d in data:
        d['lights_off'] = d.pop('licht uit')
        d['lights_on'] = d.pop('licht aan')

    return data


def transform(filename):
    '''
        Final function to transform the data
    '''
    
    print('Loading the data')
    new_periods, new_records = load_data(filename)
    print('Finished loading the data')

    print('Merging the records')
    records_merged = merge(new_records)
    print('Finished merging the records')
    
    print('Combining the records and periods')
    combined = combine(new_periods, records_merged)
    print('Finished combining the records and periods')

    print('Cleaning the data')
    cleared = clear_data(combined)
    print('Finished cleaning the data')
    
    print('Saving the data')
    with open('transformed_sleepdiary.json', 'w') as fp:
        json.dump(cleared, fp)
    print('Finished saving the data! ')
        
    return

transform(filename)