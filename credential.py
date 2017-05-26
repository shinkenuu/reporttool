import json
import os


def get_credential(subject: str):
    with open('{}/credentials.json'.format(os.path.expanduser('~')), 'r') as json_file:
        return next((cred for cred in json.load(json_file) if cred['subject'] == subject), None)
