import json

import requests

token = '5735681613:AAHxRfOOKeW5XxMwdG3mQmSOyBVxnLHqp9M'

file_id = "AgACAgIAAxkBAAIEp2N5jqUAAQKnrBHRmL1yyFRcgGmFtgACyMExG-DcyEsPTrvYeAOIgAEAAwIAA3kAAysE"

file_path = requests.get(f'https://api.telegram.org/bot{token}/getFile?file_id={file_id}').text
file_path = json.loads(file_path)
if file_path['result']:
    file_path = str(file_path['result']['file_path'])
    icon = requests.get(f'https://api.telegram.org/file/bot{token}/{file_path}').content
    with open(file_path.split("/")[1], 'wb') as new_file:
        new_file.write(icon)