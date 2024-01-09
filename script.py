import datetime
import os
from re import I
import xlrd
import requests
import json


LOCAL_PATH = './data/dane.xls'
EXCLUDED_ISSUERS_FILE_PATH = './data/excluded_issuers.txt'

TRADING_VALUE_THRESHOLD = 50.0
NOMINAL_VALUE_THRESHOLD = 1000
MARGIN_THRESHOLD_RELATIVE_TO_AVERAGE = 3
IS_FLOATING = True
MIN_MATURITY_YEARS = 1
MAX_MATURITY_YEARS = 3

OBLIGACJE_PL_URL = 'https://obligacje.pl/ajax/kalkulatorDane.php'

INSTRUMENTS_SHEET_NAME = 'lista'
INSTRUMENTS_COORPORATE_BONDS_START_ROW_INDEX = 9
INSTRUMENTS_COORPORATE_BONDS_END_ROW_INDEX = 237
INSTRUMENTS_INSTRUMENT_ID_COL_INDEX = 1
INSTRUMENTS_ISSUER_ID_COL_INDEX = 2
INSTRUMENTS_MATURITY_DATE_COL_INDEX = 3
INSTRUMENTS_TYPE_OF_INTEREST_COL_INDEX = 4
INSTRUMENTS_TRADING_CURRENCY_COL_INDEX = 5
INSTRUMENTS_NOMINAL_VALUE_COL_INDEX = 7

TRADINGS_SHEET_NAME = 'notowania'
TRADING_COORPORATE_BONDS_START_ROW_INDEX = 10
TRADING_COORPORATE_BONDS_END_ROW_INDEX = 304
TRADING_INSTRUMENT_ID_COL_INDEX = 1
TRADING_TRADING_VALUE_COL_INDEX = 8

class Bond:
  def __init__(self, id, issuer_id, maturity_date: datetime.date, type_of_interest, trading_currency, nominal_value, trading_value = 0, margin: float = 0.0, last_trading_value = 0):
    self.id = id
    self.issuer_id = issuer_id
    self.maturity_date = maturity_date
    self.type_of_interest = type_of_interest
    self.trading_currency = trading_currency
    self.nominal_value = nominal_value
    self.trading_value = trading_value
    self.margin = margin
    self.last_trading_value = last_trading_value

class ObligacjePlResponse:
    def __init__(self, kal_stopa_nominalna_marza, kal_kurs_ostatniej_transakcji):
        self.kal_stopa_nominalna_marza = kal_stopa_nominalna_marza
        self.kal_kurs_ostatniej_transakcji = kal_kurs_ostatniej_transakcji

def fetch_json(url, instrument_id):
    cache_file = './cache/{instrument_id}.json'.format(instrument_id=instrument_id)

    if os.path.isfile(cache_file):
        with open(cache_file, 'r') as f:
            return json.load(f)

    response = requests.post(url, data = 'dane=kal_kod_obligacji%3D{instrument_id}'.format(instrument_id=instrument_id), headers={'Content-Type': 'application/x-www-form-urlencoded'})
    if response.status_code != 200:
        raise Exception('There was an error. Response code: ' + str(response.status_code))

    with open(cache_file, 'w') as f:
        json.dump(response.json(), f)

    return json.loads(response.text)

def convert_obligacje_pl_json_to_response(json):
    return ObligacjePlResponse(float(json['kal_stopa_nominalna_marza']), json['kal_kurs_ostatniej_transakcji'])

def read_cells_from_instruments_sheet_and_convert_to_bond_map(instrumentsSheet, start_row_index, end_row_index, instrument_id_col_index, issuer_id_col_index, maturity_date_col_index, type_of_interest_col_index, trading_currency_col_index, nominal_value_col_index):
    bonds = {}
    for row_index in range(start_row_index, end_row_index):
        try:
          instrument_id = instrumentsSheet.cell_value(row_index, instrument_id_col_index)
          issuer_id = instrumentsSheet.cell_value(row_index, issuer_id_col_index)
          maturity_date = datetime.datetime.strptime(instrumentsSheet.cell_value(row_index, maturity_date_col_index), '%Y.%m.%d').date()
          type_of_interest = instrumentsSheet.cell_value(row_index, type_of_interest_col_index)
          trading_currency = instrumentsSheet.cell_value(row_index, trading_currency_col_index)
          nominal_value = instrumentsSheet.cell_value(row_index, nominal_value_col_index)
          bonds[instrument_id] = Bond(instrument_id, issuer_id, maturity_date, type_of_interest, trading_currency, nominal_value)
        except Exception as e:
            print(e)
    return bonds

def read_cells_from_trading_sheet_and_fill_bond_map(sheet, start_row_index, end_row_index, instrument_id_col_index, trading_value_col_index, bonds):
    for row_index in range(start_row_index, end_row_index):
        instrument_id = sheet.cell_value(row_index, instrument_id_col_index)
        trading_value = sheet.cell_value(row_index, trading_value_col_index)
        if instrument_id in bonds:
            bonds[instrument_id].trading_value = trading_value

def fetch_obligacje_pl_response_and_fill_bond_map(bonds):
    for bond in bonds.values():
        json = fetch_json(OBLIGACJE_PL_URL, bond.id)
        response = convert_obligacje_pl_json_to_response(json)
        bonds[bond.id].margin = response.kal_stopa_nominalna_marza
        bonds[bond.id].last_trading_value = response.kal_kurs_ostatniej_transakcji
    return bonds

def calculate_average_margin(bonds):
    return sum(bond.margin for bond in bonds.values()) / len(bonds)

def read_excluded_issuers_from_text_file_delimeted_by_newline(file_path):
    if not os.path.isfile(file_path):
        return []

    with open(file_path, 'r') as f:
        return [line.strip() for line in f.readlines()]

def filter_bonds(bonds, trading_value_threshold, is_floating, max_maturity_date_years, min_maturity_date_years, average_margin, excluded_issuers):
    filtered_bonds = []
    for bond in bonds:
        if bond.issuer_id in excluded_issuers:
            continue

        if bond.trading_value >= trading_value_threshold and bond.nominal_value <= NOMINAL_VALUE_THRESHOLD and bond.margin <= (average_margin + MARGIN_THRESHOLD_RELATIVE_TO_AVERAGE) and ((is_floating and bond.type_of_interest == 'zmienne/floating') or (not is_floating and bond.type_of_interest == 'stałe/fixed')) and bond.maturity_date.year <= datetime.datetime.now().year + max_maturity_date_years and bond.maturity_date.year >= datetime.datetime.now().year + min_maturity_date_years:
            filtered_bonds.append(bond)

    return filtered_bonds

def sort_bonds_by_margin(bonds):
    return sorted(bonds, key=lambda bond: bond.margin, reverse=True)

def parse_xls(file_path, instruments_sheet_name, tradings_sheet_name):
    try:
        workbook = xlrd.open_workbook(file_path)
        instrumentsSheet = workbook.sheet_by_name(instruments_sheet_name)
        bonds = read_cells_from_instruments_sheet_and_convert_to_bond_map(instrumentsSheet, INSTRUMENTS_COORPORATE_BONDS_START_ROW_INDEX, INSTRUMENTS_COORPORATE_BONDS_END_ROW_INDEX, INSTRUMENTS_INSTRUMENT_ID_COL_INDEX, INSTRUMENTS_ISSUER_ID_COL_INDEX, INSTRUMENTS_MATURITY_DATE_COL_INDEX, INSTRUMENTS_TYPE_OF_INTEREST_COL_INDEX, INSTRUMENTS_TRADING_CURRENCY_COL_INDEX, INSTRUMENTS_NOMINAL_VALUE_COL_INDEX)
        tradingsSheet = workbook.sheet_by_name(tradings_sheet_name)
        read_cells_from_trading_sheet_and_fill_bond_map(tradingsSheet, TRADING_COORPORATE_BONDS_START_ROW_INDEX, TRADING_COORPORATE_BONDS_END_ROW_INDEX, TRADING_INSTRUMENT_ID_COL_INDEX, TRADING_TRADING_VALUE_COL_INDEX, bonds)
        fetch_obligacje_pl_response_and_fill_bond_map(bonds)
        return bonds
    except Exception as e:
        print(e)
        return None

def print_bonds(bonds):
    print('ID ISSUER_ID MATURITY_DATE MARGIN LAST_TRADING_VALUE')
    for bond in bonds:
        print(bond.id, bond.issuer_id, bond.maturity_date, bond.margin, bond.last_trading_value)

    print('Bonds count: ' + str(bonds.__len__()))

excluded_issuers = read_excluded_issuers_from_text_file_delimeted_by_newline(EXCLUDED_ISSUERS_FILE_PATH)
bonds = parse_xls(LOCAL_PATH, INSTRUMENTS_SHEET_NAME, TRADINGS_SHEET_NAME)
average_margin = calculate_average_margin(bonds)
bonds = bonds.values()
bonds = filter_bonds(bonds, TRADING_VALUE_THRESHOLD, IS_FLOATING, MAX_MATURITY_YEARS, MIN_MATURITY_YEARS, average_margin, excluded_issuers)
bonds = sort_bonds_by_margin(bonds)
print_bonds(bonds)
#average_margin to string to avoid floating point errors
print('Average margin: ' + str(average_margin))