# -*- coding: utf-8 -*-
"""
Configuration structure, loading and storing
"""
from json import dumps, loads
from collections import OrderedDict, Mapping
import os


"""
Calculate path to store config at
"""
try:
    HOME = os.environ['HOME']
except KeyError:
    print("Home environment variable not set, use current working directory for config")
    HOME = os.getcwd()

CONF_DEFAULT_PATH = os.path.join(HOME, '.pdf2xlsx', 'config.txt')


class JsonDict(OrderedDict):
    """
    OrderedDict class extended with serialization functions, store and load.
    The configuration will be stored in an orderedDictionary, each value in it will be
    a regular dictionary containing 'value' and 'text'. Text could be used during
    GUI implementation, to show what is stored in the value.
    """

    @classmethod
    def _update2(cls, dictionary, update):
        for keys, values in update.items():
            if isinstance(values, Mapping):
                tmp_dict = cls._update2(dictionary.get(keys, {}), values)
                dictionary[keys] = tmp_dict
            else:
                dictionary[keys] = update[keys]
        return dictionary

    def store(self, path=CONF_DEFAULT_PATH):
        """
        Store the actual configuration to config file (path)

        :param str path: Path and filename of the config file
        """
        with open(path, 'w', encoding="utf-8") as conf_out:
            conf_out.write(dumps(self, indent=4, ensure_ascii=False))

    def load(self, path=CONF_DEFAULT_PATH):
        """
        Update the config from the config file (path)

        :param str path: Path and filename of the config file
        """
        with open(path, 'r', encoding="utf-8") as conf_in:
            self._update2(self, loads(conf_in.read()))

_KEYS = ('value', 'text', 'conf_method', 'Display')


def _create_dict(values, _keys=_KEYS):
    return dict(zip(_keys, values))

config = JsonDict([
    ('tmp_dir', _create_dict(['tmp', 'tmp dir', 'Entry', False])),
    ('file_extension', _create_dict(['pdf', 'file ext', 'Entry', False])),
    ('xlsx_name', _create_dict(['invoices.xlsx', 'xlsx name', 'Entry', False])),
    ('invo_header_ident', _create_dict([[1, 2, 3, 4], 'invo header pos', 'Entry', False])),
    ('ME', _create_dict([['Pár', 'Darab'], 'Me category', 'Entry', True])),
    ('excel_path', _create_dict(
        [r'C:\Program Files (x86)\Microsoft Office\Office14\excel.exe', 'Excel:', 'filedialog',
         True])),
    ('last_path', _create_dict(
        [r'C:\\', 'Last Dir:', 'filedialog', True]))
])


def init_conf(conf=config, cfg_path=CONF_DEFAULT_PATH):
    """
    Load the config file from $HOME/pdf2xlsx/cfg_name. If it doesn't exist
    try to create it. First create the pdf2xlsx directory, and then write out
    the default config
    """
    try:
        conf.load(cfg_path)
    except FileNotFoundError:
        cfg_dir, cfg_name = os.path.split(cfg_path)
        try:
            os.mkdir(cfg_dir)
        except FileExistsError:
            pass
        finally:
            conf.store(cfg_path)
