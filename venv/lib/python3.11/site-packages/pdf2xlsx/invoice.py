# -*- coding: utf-8 -*-
"""
Classes for different invoce types
"""

import re
from collections import namedtuple
from datetime import datetime
from .config import config
from .utility import list2row

def get_invo_type(pdf_line):
    """
    TODO add title parse to decide between invoce types
    """
    if pdf_line.startswith('HELYESB') or pdf_line.startswith('VISSZÁRU'):
        return CreditInvoice, CreditEntry
    if pdf_line.startswith('SZÁMLA'):
        return Invoice, Entry
    return None

def invo_parser(pdf_file, logger):
    """
    Factory to generate the apropriate invoce type based on the title in the PDF
    """
    invoice_type_found = False
    invo_cls = Invoice
    entry_cls = Entry
    invo = None
    entry = None
    for i in range(pdf_file.getNumPages()):
        for line in pdf_file.getPage(i).extractText().split('\n'):
            if invoice_type_found:
                if invo.parse_line(line):
                    logger.new_invo()
                if entry.parse_line(line):
                    invo.entries.append(entry)
                    entry = entry_cls(invo=invo)
                    logger.new_entr()
            else:
                tmp = get_invo_type(line)
                if get_invo_type(line):
                    invoice_type_found = True
                    invo_cls, entry_cls = tmp
                    invo = invo_cls(entries=list())
                    entry = entry_cls(invo=invo)
    return invo

EntryTuple = namedtuple('EntryTuple', ['kod', 'nev', 'ME', 'mennyiseg', 'BEgysegar',
                                       'Kedv', 'NEgysegar', 'osszesen', 'AFA'])

class Invoice():
    """
    Parse, store and write to xlsx invoce informations. Such as Invoice Number,
    Invoice Date, Payment Date, Total Sum Price. It also contains a list of Entry,
    which is also extracted form raw string.
    The parsing of the raw string is controlled by three state variables: no_parsed,
    orig_date_parsed and pay_due_parsed. These represent the structure of the pdf.

    :param int no: Invoice number, default:0
    :param str orig_date: Invoice date stored as a string YYYY.MM.DD
    :param str pay_due: Payment Date stored as string YYYY.MM.DD
    :param int total_sum: Total price of invoice
    :param list entries: List of :class:`Entry` containing each entries in invoice


    [TODO] implement state pattern for parsing ???
    [TODO] implement _to_money as a mixin class
    """
    NO_PATTERN = '([0-9]{10})'
    NO_CMP = re.compile(NO_PATTERN)
    ID_NO_PATTERN = "".join(['[ ]*Számla sorszáma:', NO_PATTERN])
    ID_NO_CMP = re.compile(ID_NO_PATTERN)

    ORIG_DATE_PATTERN = ('[ ]*Számla kelte:'
                         r'([0-9]{4}\.[0-9]{2}\.[0-9]{2}|[0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    ORIG_DATE_CMP = re.compile(ORIG_DATE_PATTERN)

    PAY_DUE_PATTERN = (r'[ ]*FIZETÉSI HATÁRIDÕ:([0-9]{4}\.[0-9]{2}\.[0-9]{2}|'
                       r'[0-9]{2}\.[0-9]{2}\.[0-9]{4})[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)')
    PAY_DUE_CMP = re.compile(PAY_DUE_PATTERN)

    AWKWARD_DATE_PATTERN = r'([0-9]{2})\.([0-9]{2})\.([0-9]{4})'
    AWKWARD_DATE_CMP = re.compile(AWKWARD_DATE_PATTERN)

    def __init__(self, no=0, orig_date="", pay_due="", total_sum=0, entries=None):
        self.id_no = no
        self.orig_date = orig_date
        self.pay_due = pay_due
        self.total_sum = total_sum
        self.entries = entries

        self.id_no_parsed = False
        self.orig_date_parsed = False
        self.pay_due_parsed = False

    def __str__(self):
        tmp_str = '\n    '.join([str(entr) for entr in self.entries])
        return ('Sorszam: {id_no}, Kelt: {orig_date}, Fizetesihatarido: {pay_due},'
                'Vegosszeg: {total_sum}\nBejegyzesek:\n    {entry_list}'
                '').format(**self.__dict__, entry_list=tmp_str)

    def __repr__(self):
        return ('{__class__.__name__}(no={no!r},orig_date={orig_date!r},'
                'pay_due={pay_due!r},total_sum={total_sum!r},'
                'entries={entries!r})').format(__class__=self.__class__, **self.__dict__)

    def _normalize_str_date(self, strdate):
        """
        The date is represented in two different format in the pdf: YYYY.MM.DD and
        DD.MM.YYYY. The second needs to be converted to the first one.

        :param str strdate: string representation of the parsed date to normalize

        :return: the normalized date
        :rtype: datetime
        """
        if self.AWKWARD_DATE_CMP.match(strdate):
            return datetime.strptime(strdate, '%d.%m.%Y')
        return datetime.strptime(strdate, '%Y.%m.%d')

    def parse_line(self, line):
        """
        Parse through a raw text which is supplied line-by-line. This is the structure
        of the pdf (the brackets() indicate what should be collected):
        <disinterested rubish>
        Számla sorszáma: (NNNNNNNN) ...
        <disinterested rubish>
        Számla kelte: (YYYY.MM.DD|DD.MM.YYYY) ...
        <disinterested rubish>
        FIZETÉSI HATÁRIDŐ:(YYYY.MM.DD|DD.MM.YYYY) (NNN[.NNN.NNN])
        <disinterested rubish>
        This is structure is paresed using the three state variable, and stored inside
        the class attributes

        :param str line: The actual line to parse

        :return: True when the parsing of the Invoice was started
        :rtype: bool
        """
        if not self.id_no_parsed:
            matchob = self.ID_NO_CMP.match(line)
            if matchob:
                self.id_no = int(matchob.group(1))
                self.id_no_parsed = True
                return True

        elif not self.orig_date_parsed:
            matchob = self.ORIG_DATE_CMP.match(line)
            if matchob:
                self.orig_date = self._normalize_str_date(matchob.group(1))
                self.orig_date_parsed = True
                return False

        elif not self.pay_due_parsed:
            matchob = self.PAY_DUE_CMP.match(line)
            if matchob:
                self.pay_due = self._normalize_str_date(matchob.group(1))
                self.total_sum = int(matchob.group(2).replace('.', ''))
                self.pay_due_parsed = True
                return False

        return False

    def xlsx_write(self, worksheet, row, col):
        """
        Write the invoice information to a template xlsx file.

        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
        """
        values = [self.id_no, datetime.strftime(self.orig_date, '%Y.%m.%d'),
                  datetime.strftime(self.pay_due, '%Y.%m.%d'), self.total_sum]
        positions = config['invo_header_ident']['value']
        row, col = list2row(worksheet, row, col, values, positions)
        return row, col


class Entry():
    """
    Parse, store and write to xlsx invoice entries. The invoice informations are
    stored in the EntryTuple namedtuple. The parsing is contolled by a state
    variable (:entry_found:) Because the invoice entries are split into two line,
    the tmp_str attribute is used to store the first part of the entire
    The ME values are configurable, so they cannot be created at class level, they
    need to be recomputed at evry instantiation

    :param EntryTuple entry_tuple: The invoice entry
    :param Invoice invo: The parent invoice containing this entry
    """

    CODE_PATTERN = '[ ]*([A-Z0-9]{2}[0-9]{4}-[0-9]{3})'
    CODE_CMP = re.compile(CODE_PATTERN)

    def __init__(self, entry_tuple=None, invo=None):
        self.entry_tuple = entry_tuple
        self.invo = invo

        self.entry_found = False
        self.tmp_str = ""

        self.me_pattern = "".join(['(', "|".join(config['ME']['value']), ')'])
        self.entry_pattern = "".join([self.CODE_PATTERN,  #termek kod
                                      "(.*)", #termek megnevezes
                                      self.me_pattern, # ME
                                      "[ ]+([0-9]+)", # mennyiseg
                                      r"[ ]+([0-9]+\.?[0-9]*)", # Brutto Egysegar
                                      "[ ]+([0-9]+)%", # Kedvezmeny
                                      r"[ ]+([0-9]+\.?[0-9]*)", # Netto Egysegar
                                      r"[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)", # Osszesen
                                      "[ ]+([0-9]+)%",]) # Afa
        self.entry_cmp = re.compile(self.entry_pattern)
        self.multiplyer = 1

    def __str__(self):
        return '{entry_tuple}'.format(**self.__dict__)

    def __repr__(self):
        return ('{__class__.__name__}(entry_tuple={entry_tuple!r}'
                ')').format(__class__=self.__class__, **self.__dict__)

    def _to_money(self, str_money):
        """
        There is no sanity check
        """
        return int(str_money.replace('.', '')) * self.multiplyer

    def line2entry(self, line):
        """
        Extracts entry information from the given line. Tries to search for nine different
        group in the line. See implementation of entry_pattern. This should match the
        following pattern:
        NNNNNN-NNN STR+WSPACE PREDEFSTR INTEGER INTEGER-. INTEGER% INTEGER-. INTEGER-.
        INTEGER%
        Where:
        N: a single digit: 0-9
        STR+WSPACE: string containing white spaces, numbers and special characters
        PREDEFSTR: string without white space ( predefined )
        INTEGER: decimal number, unknown length
        INTEGER-.: a decimal number, grouped with . by thousends e.g 1.589.674
        INTEGER%: an integer with percentage at the end

        :param str pdfline: Line to parse, this line should be begin with NNNNNNN-NNN

        :return: The actual invoice entry
        :rtype: EntryTuple
        """
        try:
            matchgp = self.entry_cmp.match(line).groups()
            return EntryTuple(matchgp[0], matchgp[1], matchgp[2],
                              int(matchgp[3]),
                              self._to_money(matchgp[4]),
                              int(matchgp[5]),
                              self._to_money(matchgp[6]),
                              self._to_money(matchgp[7]),
                              int(matchgp[8]))
        except AttributeError as exc:
            print("Entry pattern regex didn't match for line: {}".format(line))
            raise exc

    def parse_line(self, line):
        """
        Parse through raw text which is supplied line-by-line. This is the structure
        of the pdf (the brackets() indicate what should be collected):
        n times:
        <disinterested rubish>
        (NNNNNN-NNN ... \n
        ...)
        <disinterested rubish>
        When the Invoice code is found, an additional line is waited, and then it is
        sent to the line2entry converter.

        :param str line: The actual line to parse

        :return: True when an entry was found
        :rtype: bool
        """
        if self.entry_found:
            self.entry_tuple = self.line2entry(" ".join([self.tmp_str, line]))
            self.entry_found = False
            return True
        elif self.CODE_CMP.match(line):
            self.tmp_str = line
            self.entry_found = True
        return False

    def xlsx_write(self, worksheet, row, col):
        """
        Write the entry information to a template xlsx file.

        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
        """
        values = [self.invo.id_no] + list(self.entry_tuple._asdict().values())
        row, col = list2row(worksheet, row, col, values)
        return row, col


class CreditInvoice(Invoice):
    """
    Creadit invoice class
    """

    ID_NO_PATTERN = '[ ]*Helyesbítõ számla sorszáma([0-9]{10})'
    ID_NO_CMP = re.compile(ID_NO_PATTERN)

    ORIG_DATE_PATTERN = ('[ ]*Helyesbítõ számla kelte'
                         r'([0-9]{4}\.[0-9]{2}\.[0-9]{2}|[0-9]{2}\.[0-9]{2}\.[0-9]{4})')
    ORIG_DATE_CMP = re.compile(ORIG_DATE_PATTERN)

    ORIG_INVO_START = 'Eredeti számla sorszáma'

    AMOUNT_START = 'Adó részletezés'
    AMOUNT_PATTERN = r"[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)-"
    AMOUNT_CMP = re.compile(AMOUNT_PATTERN)

    def __init__(self, no=0, orig_date="", pay_due="", total_sum=0, entries=None,
                 orig_invo_no=0):
        super().__init__(no=no,
                         orig_date=orig_date,
                         pay_due=pay_due,
                         total_sum=total_sum,
                         entries=entries)
        self.pay_due = self._normalize_str_date("2017.01.01")
        self.orig_invo_no = orig_invo_no

        self.orig_invo_no_found = False
        self.orig_invo_no_parsed = False

        self.total_sum_found = False
        self.total_sum_parsed = False

    def parse_line(self, line):
        """

        :param str line: The actual line to parse

        :return: True when the parsing of the Invoice was started
        :rtype: bool
        """
        if not self.orig_date_parsed:
            return super().parse_line(line)

        elif not self.orig_invo_no_parsed:
            if not self.orig_invo_no_found:
                if line.startswith(self.ORIG_INVO_START):
                    self.orig_invo_no_found = True
                    return False
            else:
                matchob = self.NO_CMP.match(line)
                self.orig_invo_no = int(matchob.group(1))
                self.orig_invo_no_parsed = True
                return False

        elif not self.total_sum_parsed:
            if not self.total_sum_found:
                if line.startswith(self.AMOUNT_START):
                    self.total_sum_found = True
                    return False
            else:
                matchob = self.AMOUNT_CMP.search(line)
                self.total_sum = -int(matchob.group(1).replace('.', ''))
                self.total_sum_parsed = True
            return False
        return False

    def xlsx_write(self, worksheet, row, col):
        """
        Write the invoice information to a template xlsx file.

        :param Worksheet worksheet: Worksheet class to write info
        :param int row: Row number to start writing
        :param int col: Column number to start writing

        :return: the next position of cursor row,col
        :rtype: tuple of (int,int)
        """
        values = [self.id_no,"Credit note",
                  "EREDETI: " + str(self.orig_invo_no), self.total_sum]
        positions = config['invo_header_ident']['value']
        row, col = list2row(worksheet, row, col, values, positions)
        return row, col

class CreditEntry(Entry):
    """
    These entries contain negative prices as these are creadit invoices Dummy!
    """
    def __init__(self, entry_tuple=None, invo=None):
        super().__init__(entry_tuple, invo)

        self.entry_pattern = "".join([self.CODE_PATTERN,  #termek kod
                                      "(.*)", #termek megnevezes
                                      self.me_pattern, # ME
                                      "[ ]+([0-9]+)", # mennyiseg
                                      r"[ ]+([0-9]+\.?[0-9]*)-", # Brutto Egysegar
                                      "[ ]+([0-9]+)%", # Kedvezmeny
                                      r"[ ]+([0-9]+\.?[0-9]*)-", # Netto Egysegar
                                      r"[ ]+([0-9]+\.?[0-9]*\.?[0-9]*)-", # Osszesen
                                      "[ ]+([0-9]+)%",]) # Afa
        self.entry_cmp = re.compile(self.entry_pattern)
        self.multiplyer = -1
