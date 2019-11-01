"""Script to convert receipt data from spreadsheet format to JSON."""
import argparse
import datetime
import json
import pathlib
import sys
import uuid

import pyexcel


GOOD_SMC_VALUES = [
    'full_prepayment', 'prepayment', 'advance', 'full_payment',
    'partial_payment', 'credit', 'credit_payment'
]
RECEIPT_DT_VALUES = [
    'sale', 'sale_refund'
]


def process_dt(input):
    """Return operation type (dt) from input string."""
    if input not in RECEIPT_DT_VALUES:
        raise Exception('Invalid dt value')  # TODO
    return input


def process_smc(input):
    """Return calculation form (smc) from input string."""
    if input not in GOOD_SMC_VALUES:
        raise Exception('Invalid smc value')  # TODO
    return input


def process_ts(input):
    """Return tax system (ts) from input string."""
    return 'OSN'  # TODO


def process_i_p(input):
    """Return price (i.p) from input string."""
    return float(input)  # TODO


def process_i_q(input):
    """Return quantity (i.q) from input string."""
    return float(input)  # TODO


def process_i_s(input):
    """Return sum (i.s) from input string."""
    return float(input)  # TODO


def process_i_ts(input):
    """Return tax (i.ts) from input string."""
    return input  # TODO


def process_i_tv(input):
    """Return tax amount (i.tv) from input string."""
    return float(input)  # TODO


def process_i_sco(input):
    """Return good type (i.sco) from input string."""
    return input  # TODO


def get_row_reciept(input_row):
    """Get reciept data and bill ID from row."""
    if len(input_row) < 15:
        raise Exception('Invalid row width')

    bill_id = input_row[2]
    result = {
        'id': uuid.uuid4().hex.upper(),
        'dt': process_dt(input_row[3]),
        'cr': 0,  # TODO
        'ts': process_ts(input_row[6]),
        'i': [],  # TODO
    }

    if (not isinstance(input_row[4], str)) or (not input_row[4]):
        raise Exception('No e-mail given')  # TODO
    else:
        result['em'] = input_row[4]

    if isinstance(input_row[4], str) and input_row[4]:
        result['ph'] = input_row[5]

    return result, bill_id


def get_row_good(input_row):
    """Get good data from row."""
    return {
        'n': input_row[7],  # TODO
        'p': process_i_p(input_row[8]),
        'q': process_i_q(input_row[9]),
        's': process_i_s(input_row[10]),
        'ts': process_i_ts(input_row[11]),
        'tv': process_i_tv(input_row[12]),
        'smc': process_smc(input_row[13]),
        'sco': process_i_sco(input_row[14]),
    }


def convert(input_file, input_file_type, output_file):
    """
    Convert reciept entries.

    Input file should be in spreadsheet (according to `input_file_type`).
    Output file will be in JSON format.
    """
    if input_file_type is None:
        raise Exception('No input file')  # TODO
    sheet = pyexcel.get_sheet(
        file_stream=input_file,
        file_type=input_file_type,
    )

    rows = sheet.rows()
    try:
        next(rows)
    except StopIteration:
        raise Exception('Input spreadsheet is empty')  # TODO

    receipts = dict()
    for row in rows:
        if isinstance(row[0], str):
            if not row[0]:
                break  # TODO
        row_receipt, row_bill_id = get_row_reciept(row)
        if row_bill_id not in receipts:
            receipts[row_bill_id] = row_receipt
        else:
            row_receipt = receipts[row_bill_id]
        row_good = get_row_good(row)
        row_receipt['i'].append(row_good)
        row_receipt['cr'] += row_good['s']

    json.dump(
        {
            't': datetime.datetime.now().isoformat(),
            'i': list(receipts.values()),
        },
        output_file
    )


class FileTypeAction(argparse.Action):
    """Argparse action to get input file type from file name."""

    def __call__(self, parser, namespace, values, option_string=None):
        """Do TODO."""
        # TODO
        setattr(namespace, self.dest, values)
        if getattr(namespace, 'input_file_type') is not None:
            return
        input_file_type = pathlib.PurePath(values.name).suffix[1:].lower()
        setattr(namespace, 'input_file_type', input_file_type)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='Convert receipt data from spreadsheet format to JSON.'
    )
    parser.add_argument('--input-file', '-i', metavar='input', nargs='?',
                        default=sys.stdin, action=FileTypeAction,
                        type=argparse.FileType('rb'),
                        help='input file (in XLS or CSV format')
    parser.add_argument('--input-file-type', '-t', metavar='input_type',
                        nargs='?', default=None,
                        type=str, choices=['csv', 'xls', 'xlsx'],
                        help='input file format (xls, xlsx or csv)')
    parser.add_argument('--output-file', '-o', metavar='output', nargs='?',
                        default=sys.stdout,
                        type=argparse.FileType('wt'),
                        help='output file (will be in JSON format)')
    args = parser.parse_args()

    convert(args.input_file, args.input_file_type, args.output_file)
