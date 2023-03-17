import gspread
from oauth2client.service_account import ServiceAccountCredentials 

#Funkcja zwracająca poprzednią wartość winów/winratia
def previous(idx, is_wins, cut_point):
    result = sheet.cell(idx, 6).value
    result = result.strip()
    result = result[cut_point:-1]

    if is_wins: 
        return int(result)
    else: 
        return float(result)

#Funkcja wpisująca różnicę między dwoma wynikami do komórki
def difference(a, b, what, idx):
    diff = a - b
    if isinstance(diff, float):
        diff = round(diff, 2)

    if diff >= 0:
        sheet.update_cell(idx, 7, '(+{}{})'.format(diff, what))
        if diff == 0:
            sheet.format('G{}'.format(idx), {
            "textFormat": {
                "foregroundColor": {
                "red": 0.627,
                "green": 0.624,
                "blue": 0.6
            }}})
        else:
            sheet.format('G{}'.format(idx), {
            "textFormat": {
                "foregroundColor": {
                "red": 0.416,
                "green": 0.659,
                "blue": 0.31
            }}})
    else:
        sheet.update_cell(idx, 7, '({}{})'.format(diff, what))
        sheet.format('G{}'.format(idx), {
            "textFormat": {
                "foregroundColor": {
                "red": 1,
                "green": 0,
                "blue": 0
            }}})

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Statystyki gier w Fall Guys').sheet1

#Pobranie wartości z czwartej kolumny(wygrana/przegrana)
amount_of_rows = len(sheet.col_values(1))
rows = []
for i in range(2, amount_of_rows + 1):
    rows.append(sheet.cell(i, 4).value)

total_wins = 41
wins = 0
games = 0
finals = 0
exceptions = ['CHEATER', 'BUG', 'TECH']

for row in rows:
    if row not in exceptions: 
        games += 1
        if row == 'WIN' or row == '-': 
            finals += 1
    if row == 'WIN': 
        wins += 1

wr = round(100 * wins / games, 2)
fwr = round(100 * wins / finals, 2)
total_wins += wins

print('WINRATIO: {}%'.format(wr))
print('WINRATIO W FINAŁACH: {}%'.format(fwr))
print('ILOŚĆ WINÓW: {}'.format(total_wins))

#Sprawdzenie i zmiana ilości winów 
wins_from_sheet = sheet.cell(2, 7).value
wins_from_sheet = wins_from_sheet.strip()
wins_from_sheet = wins_from_sheet[:-1]

if total_wins != int(wins_from_sheet): 
    sheet.update_cell(2, 7, '{}👑'.format(total_wins))

#Wpisanie statystyk do arkusza 
choice = input("Dodać wyniki do arkusza? [TAK/NIE]\n")

if choice.lower() == 'tak':
    sheet.update_cell(amount_of_rows - 2, 6, 'WINS: {}👑'.format(total_wins))
    sheet.update_cell(amount_of_rows - 1, 6, 'WR: {}%'.format(wr))
    sheet.update_cell(amount_of_rows, 6, 'FWR: {}%'.format(fwr))

    sheet.format('F{}:G{}'.format(amount_of_rows - 2, amount_of_rows - 2), {
    "backgroundColor": {
        "red": 1,
        "green": 1,
        "blue": 0
    },
    "textFormat": {
        "bold": True
    }})

    sheet.format('F{}:G{}'.format(amount_of_rows - 1, amount_of_rows), {
    "backgroundColor": {
        "red": 0.67,
        "green": 0.80,
        "blue": 0.93
    },
    "textFormat": {
        "bold": True
    }})

    #Różnica względem poprzedniego dnia(+X/-X)
    previous_diff_index = len(sheet.col_values(7))

    previous_wins = previous(previous_diff_index - 2, True, 6)
    previous_wr = previous(previous_diff_index - 1, False, 4)
    previous_fwr = previous(previous_diff_index, False, 5)

    difference(total_wins, previous_wins, '👑', amount_of_rows - 2)
    difference(wr, previous_wr, '%', amount_of_rows - 1)
    difference(fwr, previous_fwr, '%', amount_of_rows)

    



