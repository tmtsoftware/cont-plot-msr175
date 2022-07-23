#!/usr/bin/env python3

import argparse
from datetime import datetime, time
from pathlib import Path

from openpyxl import load_workbook

class MSR175WorkbookLoadError(Exception):
    '''Indicates that the workbook has an error.'''
    
    def __init__(self, xlsx_path, sheetname, cell_address, message):
        xlsx_filename = Path(xlsx_path).name
        msg = f'[{xlsx_filename}]\'{sheetname}\'!{cell_address}: {message}'
        super().__init__(msg)

        self.__xlsx_path    = xlsx_path
        self.__sheetname    = sheetname
        self.__cell_address = cell_address
        self.__message      = message

    @property
    def xlsx_path(self):
        return self.__xlsx_path

    @property
    def sheetname(self):
        return self.__sheetname
    
    @property
    def cell_address(self):
        return self.__cell_address

    @property
    def message(self):
        return self.__message

class MSR175WorksheetParseError(Exception):
    '''Indicates that the worksheet is not in an expected format.'''
    
    def __init__(self, cell_address, message):
        super().__init__(message)
        self.__cell_address = cell_address
        self.__message      = message

    @property
    def cell_address(self):
        return self.__cell_address

    @property
    def message(self):
        return self.__message

class MSR175ShockEvent:
    def __init__(self, event_id, timestamp, sampling_period_ms, x_g, y_g, z_g):
        self.__xlsx_path = None
        self.__event_id  = event_id
        self.__timestamp = timestamp
        self.__sampling_period_ms = sampling_period_ms
        self.__x_g = x_g
        self.__y_g = y_g
        self.__z_g = z_g

        assert len(x_g) == len(y_g)
        assert len(y_g) == len(z_g)

    @property
    def event_id(self):
        return self.__event_id

    @property
    def timestamp(self):
        return self.__timestamp

    @property
    def sampling_period_ms(self):
        return self.__sampling_period_ms

    @property
    def sampling_frequency_Hz(self):
        return 1000.0 / self.__sampling_period_ms

    @property
    def duration_ms(self):
        return self.sampling_period_ms * self.n

    @property
    def n(self):
        return len(self.x_g)

    @property
    def x_g(self):
        return self.__x_g

    @property
    def y_g(self):
        return self.__y_g

    @property
    def z_g(self):
        return self.__z_g

    @property
    def t_ms(self):
        return [self.sampling_period_ms * i for i in range(self.n)]

    @property
    def xlsx_path(self):
        return Path(self.__xlsx_path)

    @property
    def xlsx_filename(self):
        return self.xlsx_path.name

    @staticmethod
    def validate_cell(worksheet, cell_address, expected_value):
        value = worksheet[cell_address].value
        if value != expected_value:
            message = f'Expected value was "{expected_value}", but it was "{value}".'
            raise MSR175WorksheetParseError(cell_address, message)

    @staticmethod
    def parse_date(worksheet, cell_address):
        value = worksheet[cell_address].value
        try:
            return datetime.strptime(value, '%y-%m-%d')
        except ValueError as e:
            message = f'Expected format was "%y-%m-%d" (e.g. 22-01-31), but it was "{value}". {str(e)}'
            raise MSR175WorksheetParseError(cell_address, message)
            
    @staticmethod
    def parse_time(worksheet, cell_address):
        value = worksheet[cell_address].value
        if isinstance(value, time):
            return value
        else:
            raise MSR175WorksheetParseError(cell_address, 'Data type must be time.')
        
    @classmethod
    def parse_worksheet(cls, worksheet):
        validate_cell = \
            lambda cell_address, expected_value: \
            MSR175ShockEvent.validate_cell(worksheet, cell_address, expected_value)

        # Read headers
        validate_cell('A1', 'Event ID:')
        validate_cell('D1', 'Start Date:')
        validate_cell('D2', 'Start Time:')

        event_id   = worksheet['B1'].value
        timestamp  = datetime.combine(MSR175ShockEvent.parse_date(worksheet, 'E1'),
                                      MSR175ShockEvent.parse_time(worksheet, 'E2'))
        
        validate_cell('A6', 'Time (msec)')
        validate_cell('B6', 'X (g)')
        validate_cell('C6', 'Y (g)')
        validate_cell('D6', 'Z (g)')

        sampling_period_ms = None
        prev_t_ms = None
        x_g = []
        y_g = []
        z_g = []
        
        for row in worksheet.iter_rows(min_row = 7, min_col = 1, max_col = 4):
            t_ms = row[0].value
            x_g.append(row[1].value)
            y_g.append(row[2].value)
            z_g.append(row[3].value)
            
            if sampling_period_ms is None and prev_t_ms is None:
                if t_ms != 0.0:
                    message = f'The first time must be 0, but was {t_ms:.5f} ms.'
                    raise MSR175WorksheetParseError(row[0].coordinate, message)
            elif sampling_period_ms is None:
                sampling_period_ms = round(t_ms - prev_t_ms, 6)
            else:
                dt_ms = round(t_ms - prev_t_ms, 6)
                if dt_ms != sampling_period_ms:
                    message = f'The time difference from the previous time must be {sampling_period_ms:.5f} ms, but was {dt_ms:.5f} ms.'
                    raise MSR175WorksheetParseError(row[0].coordinate, message)

            prev_t_ms = t_ms
        
        return MSR175ShockEvent(event_id  = event_id,
                                timestamp = timestamp,
                                sampling_period_ms = sampling_period_ms,
                                x_g = x_g,
                                y_g = y_g,
                                z_g = z_g)

    @staticmethod
    def load_xlsx(xlsx_path, skip_invalid_sheets = True):
        wb     = load_workbook(filename = xlsx_path)
        shock_events = []
        for sheetname in wb.sheetnames:
            worksheet = wb[sheetname]
            
            try:
                shock_event = MSR175ShockEvent.parse_worksheet(worksheet)
                shock_event.__xlsx_path = xlsx_path
                shock_events.append(shock_event)
            except MSR175WorksheetParseError as e:
                cell_address  = e.cell_address
                error_message = e.message

                if skip_invalid_sheets:
                    print(f'''
Warning: Skipped loading data in sheet "{sheetname}" in {xlsx_path}.
Cell "{cell_address}": {error_message}''')
                else:
                    raise MSR175WorkbookLoadError(xlsx_path, sheetname,cell_address, error_message)
                
        return shock_events

def parse_arguments():
    parser = argparse.ArgumentParser(
        description = 'Tool to plot MSR 175 acceleration data.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument('--output',
                        dest    = 'output_path',
                        default = 'output.html',
                        metavar = 'OUTPUT',
                        help    = 'Output HTML file path.')
    parser.add_argument('--skip-invalid-sheets',
                        dest     = 'skip_invalid_sheets',
                        action   = 'store_true',
                        help     = 'Specify this option to skip loading Excel sheets that include invalid format/value and continue processing other Excel sheets.')
                        
    parser.add_argument('xlsx_file', nargs='+')

    return parser.parse_args()

def main():
    args = parse_arguments()

    xlsx_paths = [Path(xf) for xf in args.xlsx_file]
    shock_events = []

    for xlsx_path in xlsx_paths:
        shock_events.extend(MSR175ShockEvent.load_xlsx(xlsx_path,
                                                       skip_invalid_sheets = args.skip_invalid_sheets))

    for shock_event in shock_events:
        print(f'{shock_event.xlsx_filename}: {shock_event.event_id} at {shock_event.timestamp} (Sampling frequency: {shock_event.sampling_frequency_Hz} Hz, duration: {shock_event.duration_ms} ms)')

if __name__ == "__main__":
    main()
