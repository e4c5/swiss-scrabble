from __future__ import print_function
from tsh import Tsh, Pair

import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

class TshGS(Tsh):
    def __init__(self, filename):
        super(TshGS,self).__init__(filename)
        self.credentials = get_credentials()
        http = self.credentials.authorize(httplib2.Http())
        self.discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                  discoveryServiceUrl=self.discoveryUrl)


    def save_to_xl(self, filename):
        players = self.players

        requests = [ {'addSheet': {'properties': {'title': 'Initial'}}}]
        response = self.service.spreadsheets().batchUpdate(spreadsheetId=filename,
            body={'requests': requests}).execute()

        rows = []
        rounds = [ [] for i in range(len(players[1]['opponents']))] 

        for idx, player in enumerate(players):
            if player['name'] == 'Bye':
                continue
            row = [player['name'], player['old_rating'],0,0]

            wins = 0
            spread = 0
            
            
            for i,opposite in enumerate(player['opponents']) :

                opponent = players[int(opposite)]
                row.append(opponent['name'])

                try:
                    player_score = player['scores'][i]
                    opponent_score = opponent['scores'][i]

                    if opponent['name'] == 'Bye' :
                        margin = int(player['scores'][i])
                    else :
                        margin = int(player['scores'][i]) - int(opponent['scores'][i])

                    if margin > 0 :
                        win = 1
                    elif margin == 0 :
                        win = 0.5
                    else :
                        win = 0

                    wins += win
                    row.append(win)
                    row.append(margin)

                    spread += margin
                    p = Pair(player['name'],opponent['name'], player_score, opponent_score)
                except IndexError:
                    # this round doesn't have scores yet
                    p = Pair(player['name'],opponent['name'], '','')

                if p not in rounds[i]:
                    rounds[i].append(p)

            row[2] = wins
            row[3] = spread
            rows.append(row)

        values = []
        for row in sorted(rows, key=lambda x: (x[2],x[3]), reverse = True):
#            values.append({'range': 'Initial','majorDimension': 'ROWS', 'values': row})
            self.service.spreadsheets().values().append(spreadsheetId=filename, body={'values': row })

 #       response = self.service.spreadsheets().values().batchUpdate(spreadsheetId=filename,
 #           body=values).execute()

        for i, rnd in enumerate(rounds):
            sheet = wb.create_sheet('Round{0}'.format(i))
            for p in rnd:
                sheet.append(p.to_array())

        wb.save(filename)


if __name__ == '__main__':

    tsh = TshGS(filename='/tmp/Junior/a.t')
    tsh.process_data()
    tsh.save_to_xl('16YtdFn-EmYW46ddVWjCDM54nl5wVdWDnTbSVqMBbDFc')
