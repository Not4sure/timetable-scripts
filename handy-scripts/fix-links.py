import json
import requests

token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpZCI6IjYzMGNkMGNjMDg4YjE4MGM4YjdmZWE3MiIsImFjY2Vzc0dyb3VwcyI6WyJzdXBlcmFkbWluIiwiYWRtaW4iXSwiaWF0IjoxNjY0ODE1ODg4LCJleHAiOjE2NjQ4MTc2ODh9.FP6IimUq-giZWzMQiuCKImDeTWpqzJaT0AfuKVcRs-E"
headers = {'Authorization': 'lol ' + token}
origin = "https://timetable.univera.app"

divisions = requests.get(origin + "/divisions")

json = divisions.json()

ICS_divisions = []

for i in json:
    if i['name'][0] == 'А' and i['name'][2] == '2' and i['name'][3] == '1':
        print(i)
        ICS_divisions.append(i)

for division in ICS_divisions:
    div_id = division['id']
    print(division['name'])
    lessons = requests.get(origin + "/lessons/" + div_id + "/current").json()
    for lesson in lessons['lessons']:
        lesson_id = lesson['id']
        requests.delete(origin + "/lesson/" + lesson_id, headers=headers)
