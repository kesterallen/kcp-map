
from collections import defaultdict
import copy
import json
import os
import pickle
from pprint import pprint
import random
import re
import time

import googlemaps
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

ONE_PER_QUANTITY = 'one_per_quantity'
ONE_PER_SPREADSHEET_ROW = 'one_per_spreadsheet_row'

DEBUG = False
MODE = ONE_PER_QUANTITY

class KindleDonorError(Exception):
    pass

class KindleDonorParseInputError(KindleDonorError):
    pass

class KindleDonorLocationError(KindleDonorError):
    pass

class KindleDonor(object):
    NUM_CELLS_EXPECTED = 6
    NUM_CELLS_OK_WITH_EMAIL = 5

    @classmethod
    def validate_row(cls, row):
        is_cell_count_good = len(row) == KindleDonor.NUM_CELLS_EXPECTED
        is_row_missing_donation_type = (
            len(row) == KindleDonor.NUM_CELLS_OK_WITH_EMAIL and '@' in row[1])

        if is_cell_count_good or is_row_missing_donation_type:
            pass
        else:
            raise KindleDonorParseInputError("need {} cells, got {}".format(
                KindleDonor.NUM_CELLS_EXPECTED, len(row)))

        # TODO?
        # check email format
        # check address exists
        # check donation_date validity
        # check quantity is an integer in string format
        # check donation type?

    @classmethod
    def from_row(cls, row):
        KindleDonor.validate_row(row)
        kindle_donor = KindleDonor(*row)
        return kindle_donor
        
    def __init__(self, name, email, address, donation_date, quantity, donation_type=None):
        self.name = name
        self.firstname = self.name.split(' ')[0]
        self.email = email
        self.donation_date = donation_date
        self.quantity = int(quantity)
        self.donation_type = donation_type

        self.set_city_state_address(address)
        self.geocode_location_from_address()

    def __repr__(self):
        return "{0.name} -- {0.address} -- {0.lat}, {0.lng}".format(self)

    def set_city_state_address(self, address):
        """Throw away any info but city/state"""
        address_lines = address.split('\n')
        address = ", ".join(address_lines[len(address_lines)-1:])
        address = re.sub('[ \d-]+$', '', address)
        self.address = address

    def geocode_location_from_address(self, is_fuzz_location=True):
        key = 'AIzaSyAlvkOzM7bg0MkBvBrcSX0Zn3QOEXa8Ljs'
        try:
            gmaps = googlemaps.Client(key=key)
            geocode_result = gmaps.geocode(self.address)
            loc = geocode_result[0]['geometry']['location']
            self.lat = loc['lat']
            self.lng = loc['lng']
        except IndexError as err:
            msg = "{} for {}".format(str(err), self.name)
            raise KindleDonorLocationError(msg)
            
    @property
    def features(self):
        """ Make an array of  geoJSON features, one per self.quantity"""
        features = [self.feature for _ in range(self.quantity)]
        return features

    @property
    def feature(self):
        """ Make a geoJSON feature for this location """
        geojson_feature = {
            "type": "Feature", 
            "geometry": {
                "type": "Point", 
                "coordinates": [self.lng, self.lat], # N.B. lng, then lat
            },
            "properties": {
                'name': self.firstname,
                'address': self.address,
                'donation_type': self.donation_type,
                'quantity': self.quantity,
                'date': self.donation_date,
            },
        }
        return geojson_feature

def get_creds():
    """Load cached credentials or get fresh ones if they've expired"""

    cred_file = 'credentials.json'
    scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    creds = None

    # Load existing, if present
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # Reload or prompt for login:
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_file, scopes)
            creds = flow.run_local_server(port=0)
    
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds

def rows_from_spreadsheet(spreadsheet_id, cells_range):
    """Extract the rows in cells_range from spreadsheet spreadsheet_id"""
    credentials = get_creds()
    service = build('sheets', 'v4', credentials=credentials)

    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=cells_range)
    response = request.execute()

    rows = [r for r in response['values']]
    return rows


def main():
    spreadsheet_id = '1-IxnOv-YFzh4eourc_PLUvGyl47xvKwTx9mNggtDaX8'
    range_ = 'Kindles!A2:F2000'

    rows = rows_from_spreadsheet(spreadsheet_id, range_)
    failures = defaultdict(list)

    kindle_donors = []
    for i, row in enumerate(rows):
        try:
            kindle_donor = KindleDonor.from_row(row)
            if DEBUG:
                print(i, kindle_donor)
            kindle_donors.append(kindle_donor)
        except KindleDonorError as err:
            failures[type(err)].append(row)
        except Exception as err:
            failures[type(err)].append(row)

    if DEBUG:
        # Skip JSON output, print summary:
        pprint(failures)
        for failtype, faillist in failures.items():
            print(failtype, len(faillist))
        print('successes:', len(kindle_donors))
    else:
        # Make JSON output
        output = dict(type="FeatureCollection", features=[])
        for kindle_donor in kindle_donors:
            if MODE == ONE_PER_QUANTITY:
                output['features'].extend(kindle_donor.features)
            else:
                output['features'].extend([kindle_donor.feature])
        json_str = json.dumps(output)
        print(json_str)

if __name__ == '__main__':
    main()
