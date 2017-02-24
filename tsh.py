import re
import traceback
import json

from openpyxl import Workbook

class Tsh(object):
    '''
    Imports data from TSH
    TSH saves it's data in what are known as division files. Typically in a tournament 
    with only one division that file will be named as `a.t` hence division files are 
    sometimes known as at files as well
    '''

    def __init__(self, filename):
        self.filename = filename
        self.players = [{'name': 'bye'}]
        self.rounds = 0;


    def process_data(self):
        players = self.players

        with open(self.filename) as f:
            for line in f:
                if line and len(line) > 30 :
                    rating = re.search('[0-9]{1,4} ', line).group(0).strip()
                    name = line[0: line.index(rating)].strip()
                    newr = None
                    data = line[line.index(rating):]
                    data = data.split(';')
                    opponents = data[0].strip().split(' ')[1:]
                    scores = data[1].strip().split(' ')
                    p12 = None
                    rank = None

                    offed = False

                    for d in data:
                        obj = d.strip().split(' ')
                        obj_name = obj[0].strip()
                        itms = obj[1:]
                        
                        if obj_name == 'p12':
                            p12 = itms
                        elif obj_name == 'newr':
                            newr = d.strip().split(' ')[1:]

                        elif obj_name == 'rrank':
                            # important change on Nov 29, 2015 - it was decided to always 
                            # use rcrank instead of rrank since rrank was noticed to have 
                            # the wrong information at times example SOY3 2015

                            # Exception will be when the rcrank field has fewer data items
                            # than the rrank field
                            tmp_rank = itms 
                            if not rank:
                                rank = tmp_rank
                            else :
                                if len(tmp_rank) > len(rank):
                                    rank = tmp_rank

                        elif obj_name == 'rcrank' and rank == None:
                            rank = [0] + itms
                        elif obj_name == 'off':
                            offed = True
                           

                    if rank != None:
                        if opponents and len(rank) > len(opponents) :
                            rank = rank[1:]

                    if not p12:
                        p12 = ['3'] * (len(rank)+1)

                    players.append({'name': name, 'opponents' :opponents,'scores':scores,
                        'p12': p12, 'rank': rank,'newr': newr, 'off': offed, 
                        'old_rating': rating})
                    
                    
            if len(players) < 2:
                print 'a.t file does not contain any data'
                return

            players[0]['scores'] = [0] * len(players[1]['scores'])
            
class TshXl(Tsh):
    '''
    Saves processed TSH data into a spreadsheet
    '''

    def save_to_xl(self, filename):
        players = self.players
        wb = Workbook()
        sheet = wb.create_sheet('Standings')
        rows = []

        for idx, player in enumerate(players):
            if player['name'] == 'bye':
                continue

            row = [player['name'], player['old_rating'],0,0]

            wins = 0
            spread = 0
            
            for i,opposite in enumerate(player['opponents']) :
                opponent = players[int(opposite)]
                row.append(opponent['name'])

                player_score = player['scores'][i]
                opponent_score = opponent['scores'][i]

                if opponent['name'] == 'bye' :
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

            row[2] = wins
            row[3] = spread
            rows.append(row)

        for row in sorted(rows, key=lambda x: (x[2],x[3]), reverse = True):
            sheet.append(row)

        wb.save(filename)


tsh = TshXl('/home/raditha/SLSL/unrated/2015/Junior/a.t')
tsh.process_data()
tsh.save_to_xl('swiss.xlsx')


