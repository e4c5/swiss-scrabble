import httplib2
import os
import argh

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from tsh import Tsh, Pair
from swiss import Pairing

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


class GSheetPairing(Pairing):
    ''' Pairing from tournament results stored in a Google Sheet'''
    
    HEADER_ROWS = 1

    def __init__(self, filename, next_round=1):

        ''' 
        players - list of dicts(name, rating, spread)
        games - list of dicts(round, player(name), opponent(name||None), player_score, opponent_score, player_color, opponent_color, is_walkover)
        round_number - number of next round
        '''

        self.filename = filename
        
        self.credentials = get_credentials()
        http = self.credentials.authorize(httplib2.Http())
        self.discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                  discoveryServiceUrl=self.discoveryUrl)

        self.players = []
        self.pairs = []
        self.bye = {'name': 'Bye', 'rating': 0, 'score': 0 , 'spread': 0 }

        data = self.service.spreadsheets().values().batchGet(
                spreadsheetId=filename, ranges='Initial!A:N'.format(GSheetPairing.HEADER_ROWS)).execute()

        for pl in data['valueRanges'][0]['values']:
            opponents = []
            spread = 0
            wins = 0
            if pl[0] == None:
                break

            for game in range(1, next_round):
                spread += int(pl[game*3 + 3])
                opponents.append(pl[game*3 +1])
                wins += float(pl[game*3 + 2])
            
            try:
                if pl[0] != 'Bye':
                    self.players.append( 
                         { 'name': pl[0],
                           'spread' : spread,
                           'rating' : pl[1] or 0,
                           'score'  : wins,
                           'opponents' : opponents
                         }
                    )
            except :
                traceback.print_exc() 
                pass
        
        self.next_round = next_round
        
        # A.3
        brackets = {}
        for player in self.players:
            if player['score'] not in brackets:
                brackets[player['score']] = []
            brackets[player['score']].append(player)
        self.brackets = brackets
        print len(self.players)

    def make_it(self):
        pairs = super(GSheetPairing,self).make_it()
        requests = [ {'addSheet': {'properties': {'title': 'Round{0}'.format(self.next_round)}}} ]
        response = self.service.spreadsheets().batchUpdate(spreadsheetId=self.filename,
            body={'requests': requests}).execute()

        
        values = []
        for item in sorted(pairs, reverse=True, key=lambda x: (x[0]['score'], x[0]['spread'])):
            values.append([item[0]['name'],'', item[1]['name'],''])
        
        response = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.filename,
                body={'data': {'values': values, 'range': 'Round{0}!A1'.format(self.next_round)},
                    'valueInputOption': 'RAW'}).execute()


        return pairs
    
    
    def update_scores(self):

        players = []
        data = self.service.spreadsheets().values().batchGet(
                spreadsheetId=self.filename, ranges='Initial!A:A'.format(GSheetPairing.HEADER_ROWS)).execute()

        for pl in data['valueRanges'][0]['values']:
            players.append(pl[0])
        
        column = chr(5 + (self.next_round-1) * 3)
        
        data = self.service.spreadsheets().values().batchGet(
                spreadsheetId=self.filename, ranges='Round{0}!A:D'.format(self.next_round)).execute()

        
        print 'Updating results for round ', self.next_round, 'starting from column', column
        

        for game in data['valueRanges'][0]['values']:
            player = players.index(game[0]) + GSheetPairing.HEADER_ROWS+1

            if game[2] == 'Bye':
                opponent = None
            else :
                opponent = players.index(game[2]) + GSheetPairing.HEADER_ROWS+1

            spread = int(game[1]) - int(game[3])
            print player, column, 'opponent', opponent

            if spread == 0:
                sheet.cell(row=player,column=column+1,value=0.5)
                sheet.cell(row=opponent,column=column+1,value=0.5)
            
            elif spread > 0:
                sheet.cell(row=player,column=column+1,value=1)
                if opponent:
                    sheet.cell(row=opponent,column=column+1,value=0)
            else :
                sheet.cell(row=player,column=column+1,value=0)
                sheet.cell(row=opponent,column=column+1,value=1)
            
            sheet.cell(row=player,column=column+2,value=spread)
            sheet.cell(row=player,column=column,value=game[2].value)
            
            if opponent:
                sheet.cell(row=opponent,column=column,value=game[0].value)
                sheet.cell(row=opponent,column=column+2,value=-spread)
               


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
            values.append(row)
#            values.append({'range': 'Initial','majorDimension': 'COLUMNS', 'values': [row]})
#            self.service.spreadsheets().values().append(spreadsheetId=filename,
#                    valueInputOption='RAW',
#                    body={'values': [row] }, range="A1").execute()

        response = self.service.spreadsheets().values().batchUpdate(spreadsheetId=filename,
                body={'data': {'values': values, 'range': 'Initial!A1'}, 'valueInputOption': 'RAW'}).execute()


        for i, rnd in enumerate(rounds):
            requests = [ {'addSheet': {'properties': {'title': 'Round{0}'.format(i+1)}}}]

            response = self.service.spreadsheets().batchUpdate(spreadsheetId=filename,
                body={'requests': requests}).execute()
            values = [] 
            for p in rnd:
                values.append(p.to_array())

            response = self.service.spreadsheets().values().batchUpdate(spreadsheetId=filename,
                body={'data': {'values': values, 'range': 'Round{0}!A1'.format(i+1)},
                    'valueInputOption': 'RAW'}).execute()


def division_to_excel():
    tsh = TshGS(filename='/tmp/Junior/a.t')
    tsh.process_data()
    tsh.save_to_xl('16YtdFn-EmYW46ddVWjCDM54nl5wVdWDnTbSVqMBbDFc')

def make_pairing(round_number):
    pair = GSheetPairing('16YtdFn-EmYW46ddVWjCDM54nl5wVdWDnTbSVqMBbDFc',round_number)
    pair.make_it()


if __name__ == '__main__':
    pair = GSheetPairing('16YtdFn-EmYW46ddVWjCDM54nl5wVdWDnTbSVqMBbDFc',4)
    pair.update_scores()

