# GreenieBoard.py
# 
# Created August 1, 2022 by Niklas Elmqvist (niklas.elmqvist@gmail.com)
# Created and maintained by VF-111 (a virtual F-14A DCS squadron)

import re
import csv
import sys
import gspread

from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

# --- Configuration settings

# lsoBot Root
lsoBot_root = 'c:\\lsoBot\\'
lsoBot_logs = lsoBot_root + "Logs\\"

# The name of the DCS server
server_field = "VF-111"

# Use the below to configure where empty slots begin and end
emptySlotIndexStart = 8
emptySlotIndexEnd = 43

# Location of the main feed Google Sheet
feed_url = 'url_to_google_sheet'

# Location of squadron 
greenie_sheets = {
    'squadron name': 'url_to_google_sheet'
}

# --- Classes

class Pilot: 

    def __init__(self, name) -> None:
        self.name = name
        self.landings = 0
        self.points = 0.0
        self.months = {}

    def add_landing(self, date, grade):

        # Ignore wave offs
        if grade['score'] == -1:
            return

        # Update global stats
        self.landings = self.landings + 1
        self.points = self.points + grade['score']

        # Now update monthly stats
        if date not in self.months:
            self.months[date] = { 'landings': 0, 'points': 0 }
        self.months[date]['landings'] = self.months[date]['landings'] + 1
        self.months[date]['points'] = self.months[date]['points'] + grade['score']

    def get_stats(self):
        return (self.landings, self.points / self.landings)

    def get_stats_month(self, date):
        month_points = 0.0
        month_landings = 1.0
        if date in self.months:
            month_points = self.months[date]['points']
            month_landings = self.months[date]['landings']
        return (month_landings, month_points / month_landings)

    def get_stats_list(self, date):
        mo_ldg, mo_avg = self.get_stats_month(date)
        return [[ self.landings, mo_ldg, self.points / self.landings, mo_avg ]]


class GreenieBoard:
    
    def __init__(self):
        self.reset()

    def reset(self):

        # Initialize the data table
        self.data = []

        # Initalize new pilot events
        self.new_pilot_events = {}

        # Initialize pilot stats
        self.stats = {}

    def load_data(self, data_file):

        # Status message
        eprint('Loading data file "' + data_file + "'...")

        # Open the CSV file
        with open(data_file, 'r', encoding='utf-16-le') as csvfile:

            # Create a CSV reader
            trap_reader = csv.reader(csvfile, delimiter=',', quotechar='"')

            # Read it line by line
            for row in trap_reader:
                self.data.append(row)

        # Status message
        eprint("Loaded " + str(len(self.data)) +  " rows.")

    def process(self):

        cleaned_data = []
        eprint("Processing and cleaning data...")

        # Step through all of the rows in the field
        for row in range(len(self.data)):

            # Is this an illegal row? If so, skip it
            if (len(self.data[row]) < 3):
                continue

            # Split the grade into parts
            tokens = self.data[row][2].split(':')

            # If there are no comments, add an empty comment
            if (len(tokens) < 2):
                tokens.append('')

            # Delete the old grade and append the new one
            self.data[row].pop(2)
            self.data[row].extend(tokens)

            # Calculate the point score
            score = score_grade(tokens[0], tokens[1])
            if score['token']:
                self.data[row].append(score['token'])
            else: 

                # Is there a wire specification?
                wire = re.search('WIRE# (\d)', tokens[1])
                if wire != None:
                    self.data[row].append(wire.group(1))
                else:
                    self.data[row].append('-')

            # Add the score as a field
            self.data[row].append(str(score['score']))

            # Also add the server as a field
            self.data[row].append(server_field)

            # Trim the cells
            for cell in range(len(self.data[row])):
                self.data[row][cell] = self.data[row][cell].strip()

            # Store the stats
            callsign, _, _ = parse_pilot(self.data[row][1])
            date = parse_date(self.data[row][0])
            if callsign not in self.stats:
                self.stats[callsign] = Pilot(callsign)
            self.stats[callsign].add_landing(date, score)

            # Then copy it 
            cleaned_data.append(self.data[row])

        # Store the cleaned data
        self.data = cleaned_data
        eprint(str(len(self.data)) + " items retained after cleaning.")

    def save_event(self, callsign, event, sheet):

        grade = event[2]
        comments = event[3]

        # Parse the date
        board_name = parse_date(event[0])
        if board_name == None:
            return

        # Do we need a new worksheet for this year and month?
        greenie = get_worksheet_by_name(sheet, board_name)
        if greenie == None:
            template = get_worksheet_by_name(sheet, 'Template')
            greenie = sheet.duplicate_sheet(template.id, insert_sheet_index=0, new_sheet_name=board_name)

        # Now find the row with the pilot's name
        cell = greenie.find(callsign, case_sensitive=False)

        # If it doesn't exist, just get out
        if cell == None:
            return

        # Read the entire row
        row = greenie.row_values(cell.row)

        # Find the first free slot
        # FIXME: overwrite from the beginning
        col = len(row) + 1
        if col < emptySlotIndexStart:
            col = emptySlotIndexStart
        if col > emptySlotIndexEnd:
            col = emptySlotIndexEnd

        # Add the trap to the cell
        cell = greenie.cell(cell.row, col)
        greenie.update_acell(cell.address, event[4])

        # Figure out formatting
        score = score_grade(grade, comments)

        # Set the formatting appropriately
        greenie.format(cell.address, {
            "backgroundColor": { "red": score["red"], "green": score["green"], "blue": score["blue"] },
            "textFormat": { "underline": score["underline"] }
        })

        # Now update the stats
        pilot_stats = self.stats[callsign].get_stats_list(board_name)
        greenie.update('D' + str(cell.row) + ":G" + str(cell.row), pilot_stats)

    def save_summary(self, sheets):

        # Step through all of the pilots
        for pilot in self.new_pilot_events:

            # Parse the pilot name
            callsign, modex, squadron = parse_pilot(pilot)
            if not callsign: return

            # Find the correct greenie board
            if squadron not in sheets:
                return
            sheet_url = sheets[squadron]

            # Open the google sheet
            sheet = open_gspread(sheet_url)

            # Step through the events
            for event in self.new_pilot_events[pilot]:
                self.save_event(callsign, event, sheet)
    
    def save_feed(self, sheet_url):

        # Open the google sheet
        sheet = open_gspread(sheet_url)

        # Get the feed sheet of the spreadsheet
        feed_sheet = get_worksheet_by_name(sheet, 'Main Feed')

        # Read the entire date column
        date_col = feed_sheet.col_values(1)

        # Add the missing data to the Google Sheet
        row = 0
        new_rows = 0
        while row < len(self.data):

            # Has this trap been stored already?
            # FIXME: Small kludge; there is an infinitesmal risk that two traps happened at the same time on different servers/carriers
            curr_row = self.data[row]

            # Increment the row number
            row = row + 1

            # If it has been stored already, skip it
            if curr_row[0] in date_col:
                continue

            # Insert the new data
            feed_sheet.append_row(curr_row)
            new_rows = new_rows + 1

            # Now find the pilot worksheet
            pilot = curr_row[1]
            callsign, _, _ = parse_pilot(pilot)
            pilot_sheet = get_worksheet_by_name(sheet, callsign)

            # If no such sheet, create it
            if pilot_sheet == None: 
                template = get_worksheet_by_name(sheet, 'Template')
                pilot_sheet = sheet.duplicate_sheet(template.id, insert_sheet_index=1, new_sheet_name=callsign)

            # Now add to the pilot worksheet as well
            pilot_sheet.append_row(curr_row)

            # Finally, add it to our dictionary since it is new
            if pilot not in self.new_pilot_events:
                self.new_pilot_events[pilot] = []
            self.new_pilot_events[pilot].append(curr_row)

        # Print status message
        eprint('Added ' + str(new_rows) + ' rows to the sheet.')

# --- Functions

def open_gspread(sheet_url):

    # Define the scope
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    # Add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name('token.json', scope)

    # Authorize the clientsheet 
    client = gspread.authorize(creds)

    # Get the instance of the spreadsheet
    sheet = client.open_by_url(sheet_url)
    print('Opening Google Sheet "' + sheet_url + '"')

    return sheet


def parse_date(date_string):

    # Parse the date
    date_parse = re.search('(\d\d\d\d)-(\d\d)-(\d\d)', date_string)
    if date_parse == None:
        return None

    # Construct the date object
    date = datetime(int(date_parse.group(1)), int(date_parse.group(2)), int(date_parse.group(3)))

    # Construct the string
    return "{0:%B} {0:%Y}".format(date)


def parse_pilot(pilot):

    # Parse the pilot name
    names = pilot.split('|')
    callsign = ''
    modex = '000'
    squadron = 'VF-111'

    # Construct the callsign, modex, and squadron information
    if len(names) == 1:
        callsign = names[0].strip()
    elif len(names) == 2:
        callsign = names[0].strip()
        squadron = names[1].strip()
    elif len(names) == 3:
        callsign = names[0].strip()
        modex = names[2].strip()
        squadron = names[3].strip()

    return (callsign, modex, squadron)


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


def get_worksheet_by_name(sheet, name):
    all_sheets = sheet.worksheets()
    for curr_sheet in all_sheets:
        if curr_sheet.title == name:
            return curr_sheet
    return None


def score_grade(grade, comments):

    # Predefined grades, scores, and formatting 
    grade_list = {
        "Perfect": { "score": 5, "token": "", "red": 106 / 255.0, "green": 168 / 255.0, "blue": 79 / 255.0, "underline": True }, 
        "Acceptable": { "score": 4, "token": "", "red": 241 / 255.0, "green": 168 / 255.0, "blue": 79 / 255.0, "underline": False }, 
        "Fair": { "score": 3, "token": "", "red": 180 / 255.0, "green": 194 / 255.0, "blue": 50 / 255.0, "underline": False }, 
        "No Grade": { "score": 2, "token": "", "red": 180 / 255.0, "green": 95 / 255.0, "blue": 6 / 255.0, "underline": False }, 
        "WO\(FD\)": { "score": -1, "token": "NC", "red": 255 / 255.0, "green": 255 / 255.0, "blue": 255 / 255.0, "underline": False }, 
        "Bolter": { "score": 2, "token": "B", "red": 11 / 255.0, "green": 83 / 255.0, "blue": 148 / 255.0, "underline": False }, 
        "Wave Off": { "score": 1, "token": "-", "red": 204 / 255.0, "green": 0 / 255.0, "blue": 0 / 255.0, "underline": False }, 
        "CUT": { "score": 0, "token": "C", "red": 0 / 255.0, "green": 255 / 255.0, "blue": 255 / 255.0, "underline": False }
    }

    # Look for this grade 
    for curr in grade_list:
        if re.search(curr, grade) or re.search(curr, comments):
            return grade_list[curr]
    
    # Danger Will Robinson, couldn't find it
    return None

def main():

    # Create the board
    board = GreenieBoard()

    # Load it from CSV
    board.load_data(lsoBot_logs + 'lsoBot-data.csv')

    # Process the data
    board.process()

    # Save to the feed sheet
    board.save_feed(feed_url)

    # Save the new events the greenie summary sheet
    board.save_summary(greenie_sheets)

# Call main
if __name__ == "__main__":
    main()
