# -*- coding: utf-8 -*-
"""
Statistics collector helper class for pdf2xlsx
"""
class StatLogger():
    """
    Collect statistic about the zip to xlsx process. Assembles a list containin invoice
    number of items. Every item is the number of entries found during the invoice parsing.
    It implements a simple API: new_invo(), new_entr() and __str__()
    A new instance contains an empty list: invo_list
    """
    def __init__(self):
        self.invo_list = []

    def __str__(self):
        return '{invo_list}'.format(**self.__dict__)

    def new_invo(self):
        """
        When a new invoice was found create a new invoice log instance
        The current implementation is a simple list of numbers
        """
        self.invo_list.append(0)

    def new_entr(self):
        """
        When a new entry was found increase the entry counter for the current
        invoice.
        """
        self.invo_list[-1] += 1
