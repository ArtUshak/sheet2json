"""Script to convert receipt data from spreadsheet format to JSON."""
import argparse
import datetime
import ipaddress
import json
import pathlib
import re
import sys
import uuid

import dns.resolver
import pyexcel


class InvalidFileFormatException(Exception):
    """Input file format is invalid."""

    def __init__(self, message, row_id=None):
        """Create exception."""
        super().__init__(message)
        self.row_id = row_id

    def __str__(self):
        """Return exception string."""
        if self.row_id is not None:
            return 'Row {}, {}'.format(self.row_id, super().__str__())
        else:
            return super().__str__()


RECEIPT_DT_VALUES = [
    'sale', 'sale_refund'
]
GOOD_SMC_VALUES = [
    'full_prepayment', 'prepayment', 'advance', 'full_payment',
    'partial_payment', 'credit', 'credit_payment'
]
GOOD_TS_VALUES = [
    'vat10', 'vat20'
]
GOOD_TS_NUM_VALUES = {
    'vat10': 0.1,
    'vat20': 0.2,
}

# taken from Django:
# https://github.com/django/django/blob/master/django/core/validators.py
EMAIL_USER_REGEX = re.compile(
    r"(^[-!#$%&'*+/=?^_`{}|~0-9A-Z]+(\.[-!#$%&'*+/=?^_`{}|~0-9A-Z]+)*\Z"
    # dot-atom quoted-string
    r'|^"([\001-\010\013\014\016-\037!#-\[\]-\177]|'
    r'\\[\001-\011\013\014\016-\177])*"\Z)',
    re.IGNORECASE)
EMAIL_DOMAIN_REGEX = re.compile(
    # max length for domain name labels is 63 characters per RFC 1034
    r'((?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+)'
    r'(?:[A-Z0-9-]{2,63}(?<!-))\Z',
    re.IGNORECASE)
EMAIL_LITERAL_REGEX = re.compile(
    # literal form, ipv4 or ipv6 address (SMTP 4.1.3)
    r'\[([A-f0-9:\.]+)\]\Z',
    re.IGNORECASE)


def check_ipaddress(ip_address):
    """
    Check IPv4 or IPv6 address for validity.

    Return `True` if it is valid, `False` if it is invalid.
    """
    try:
        ipaddress.IPv4Address(ipaddress)
    except ValueError:
        try:
            ipaddress.IPv6Address(ipaddress)
        except ValueError:
            return False
    return True


def check_email_domain(domain_part):
    """
    Check email domain part for validity.

    Return `True` if it is valid, `False` if it is invalid.
    """
    if EMAIL_DOMAIN_REGEX.match(domain_part):
        return True

    literal_match = EMAIL_LITERAL_REGEX.match(domain_part)
    if literal_match:
        ip_address = literal_match.group(1)
        if check_ipaddress(ip_address):
            return True

    return False


def validate_email(input):
    """
    Check email for validity.

    Return tuple of user name and domain name if it is valid,
    or `None` if it is invalid.
    """
    if '@' not in input:
        return None

    user_part, domain_part = input.rsplit('@', 1)

    if not EMAIL_USER_REGEX.match(user_part):
        return None

    if not check_email_domain(domain_part):
        try:
            # domain_part = punycode(domain_part)  # TODO!
            pass
        except UnicodeError:
            return None
        if not check_email_domain(domain_part):
            return None

    return (user_part, domain_part)


def check_email_domain_mx(domain_part):
    """Check if domain name exists and have MX DNS record."""
    try:
        dns.resolver.query(domain_part, 'MX')
    except dns.resolver.NXDOMAIN:
        return False
    except dns.resolver.YXDOMAIN:
        return False
    except dns.resolver.NoAnswer:
        return False
    return True  # TODO!


def process_dt(input):
    """Return operation type (dt) from input string."""
    if input not in RECEIPT_DT_VALUES:
        raise InvalidFileFormatException('Invalid dt value')
    return input


def process_smc(input):
    """Return calculation form (smc) from input string."""
    if input not in GOOD_SMC_VALUES:
        raise InvalidFileFormatException('Invalid i.smc value')
    return input


def process_ts(input):
    """Return tax system (ts) from input string."""
    return 'OSN'  # TODO


def process_em(input):
    """Return e-mail address (em) from input string."""
    if not input:
        return ''
    validation_results = validate_email(input)
    if validation_results is None:
        raise InvalidFileFormatException(
            'Invalid i.em value: invalid e-mail format'
        )
    if not check_email_domain_mx(validation_results[1]):
        raise InvalidFileFormatException(
            'Invalid i.em value: domain name does not exist, '
            'or does not have MX record'
        )
    return input


def process_ph(input):
    """Return phone number (ph) from input string."""
    return input  # TODO


def process_i_p(input):
    """Return price (i.p) from input string."""
    try:
        return float(input)  # TODO
    except ValueError:
        raise InvalidFileFormatException('Invalid i.p value')


def process_i_q(input):
    """Return quantity (i.q) from input string."""
    try:
        return float(input)  # TODO
    except ValueError:
        raise InvalidFileFormatException('Invalid i.q value')


def process_i_s(input):
    """Return sum (i.s) from input string."""
    try:
        return float(input)  # TODO
    except ValueError:
        raise InvalidFileFormatException('Invalid i.s value')


def process_i_ts(input):
    """Return tax (i.ts) from input string."""
    if input not in GOOD_TS_VALUES:
        raise InvalidFileFormatException('Invalid i.ts value')  # TODO
    return input


def process_i_tv(input):
    """Return tax amount (i.tv) from input string."""
    try:
        return float(input)  # TODO
    except ValueError:
        raise InvalidFileFormatException('Invalid i.tv value')


def process_i_sco(input):
    """Return good type (i.sco) from input string."""
    return input  # TODO


def get_row_reciept(input_row):
    """Get reciept data and bill ID from row."""
    if len(input_row) < 15:
        raise InvalidFileFormatException('Invalid row width')

    bill_id = input_row[2]
    result = {
        'id': uuid.uuid4().hex.upper(),
        'dt': process_dt(input_row[3]),
        'cr': 0,  # TODO
        'ts': process_ts(input_row[6]),
        'i': [],  # TODO
    }

    if (not isinstance(input_row[4], str)) or (not input_row[4]):
        raise InvalidFileFormatException('No e-mail given')  # TODO
    else:
        result['em'] = process_em(input_row[4])

    if isinstance(input_row[4], str) and input_row[4]:
        result['ph'] = process_ph(input_row[5])

    return result, bill_id


def check_good(good):
    """Check input good row."""
    total_without_vat = good['p'] * good['q']
    vat = GOOD_TS_NUM_VALUES[good['ts']]
    delta = total_without_vat * vat - good['tv']
    if abs(delta) > 0.01:
        raise InvalidFileFormatException(
            'Invalid i.tv value (contradicts with i.p, i.q and i.ts)'
        )


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
        raise InvalidFileFormatException('No input file')  # TODO
    sheet = pyexcel.get_sheet(
        file_stream=input_file,
        file_type=input_file_type,
    )

    rows = sheet.rows()
    try:
        next(rows)
    except StopIteration:
        raise InvalidFileFormatException('Input spreadsheet is empty')  # TODO

    receipts = dict()
    row_id = 2
    for row in rows:
        try:
            if isinstance(row[0], str):
                if not row[0]:
                    break  # TODO
            row_receipt, row_bill_id = get_row_reciept(row)
            if row_bill_id not in receipts:
                receipts[row_bill_id] = row_receipt
            else:
                row_receipt = receipts[row_bill_id]
            row_good = get_row_good(row)
            check_good(row_good)
            row_receipt['i'].append(row_good)
            row_receipt['cr'] += row_good['s']
        except InvalidFileFormatException as e:
            e.row_id = row_id
            raise e
        row_id += 1

    json.dump(
        {
            't': datetime.datetime.now().isoformat(),
            'i': list(receipts.values()),
        },
        output_file,
        indent=2
    )


class FileTypeAction(argparse.Action):
    """Argparse action to get input file type from file name."""

    def __call__(self, parser, namespace, values, option_string=None):
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

    try:
        convert(args.input_file, args.input_file_type, args.output_file)
    except InvalidFileFormatException as e:
        print(f'Error: {e}', file=sys.stderr)  # TODO
