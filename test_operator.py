#!/usr/bin/env python

# current-activity.py
# -------------------
# Count the number of active chats and the number of librarians that
# are staffing services.
from pprint import pprint as pprint
from datetime import datetime
import pandas as pd
import dateparser

import warnings
import dateparser

# Ignore dateparser warnings regarding pytz
warnings.filterwarnings(
    "ignore",
    message="The localize method is no longer necessary, as this time zone supports the fold attribute",
)

import lh3.api
client = lh3.api.Client()

# For each user...
client.set_options(version = 'v1')
users = client.all('users')
num_users = 0

operator_activity = list()
print('calculating user list - please wait...')
for user in users.get_list():
    query = {
        'query': {
            'operator': [user.get('name')],
            'from': '2016-01-01',
            'to': '2022-12-31'
        },
        'sort': [
            {'started': 'descending'}
        ]
    }
    operator_chats = client.api().post('v4', '/chat/_search', json = query)
    if operator_chats:
        last_chat = operator_chats[0].get('local_started', None)
        if last_chat:
            operator_activity.append({'user':user.get('name'), 'last_chat':dateparser.parse(last_chat)})
            #breakpoint()
            #import sys; sys.exit()
        else:
            operator_activity.append({'user':user.get('name'), 'last_chat':None})
    else:
        operator_activity.append({'user':user.get('name'), 'last_chat':None})


writer = pd.ExcelWriter('assignments.xlsx', engine='xlsxwriter')

df_op = pd.DataFrame(operator_activity)

df_op.to_excel(writer, index=False, sheet_name='last_chat')
###

assign = list()
print('calculating assignment - please wait...')
for user in users.get_list():
    assignments = users.one(user['id']).all('assignments').get_list()
    for assignee in assignments:
        assign.append(assignee)

df = pd.DataFrame(assign)

df['operator']=df['user']
del df['queueShow']
del df['userShow']
del df['enabled']
df.to_excel(writer, index=False, sheet_name='assignments')
print('created the Excel file assignments.xlsx')
