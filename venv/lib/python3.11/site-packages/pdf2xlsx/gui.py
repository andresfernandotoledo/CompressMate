# -*- coding: utf-8 -*-
"""
Not so simple tkinter based gui around the pdf2xlsx.do_it function.
"""
from tkinter import Tk, ttk, filedialog, messagebox, StringVar, Toplevel, END
import os
import shutil
from .managment import do_it, do_it2
from .config import config

__version__ = '0.2.0'


class ConfOption:
    """
    This widget is used to place the configuration options to the ConfigWindow. It contains
    a label to show what is the configuration and an entry with StringVar to provide override
    possibility. The value of the config :class:`JsonDict` is converted to a string for the entry.
    If the value of a configuration is a list, it is converted to a comma separated string.

    :param Frame root: Tk parent frame
    :param str key: Key to the "config" :class:`JsonDict`
    :param int row: Parameter for grid window manager
    """
    def __init__(self, root, key, row):
        self.key = key
        dict_value = config[key]
        ttk.Label(root, text=dict_value['text']).grid(row=row, column=0, sticky='w')
        self.sv = StringVar()
        if isinstance(dict_value['value'], list):
            self.sv.set(", ".join(map(str, dict_value['value'])))
        else:
            self.sv.set(str(dict_value['value']))
        self.entry = ttk.Entry(root, textvariable=self.sv, width=54)
        self.entry.grid(row=row, column=1, sticky='e')
        if dict_value['conf_method'] == 'filedialog':
            ttk.Button(root, text='Sel',
                       command=self.browse_callback,
                       width=4).grid(row=row, column=2, sticky='w')

    def update_config(self):
        """
        Write the current entry value to the configuration. The original type of the
        config value is checked, and the string is converted to this value (int, list of
        int, list of string...)
        """
        if isinstance(config[self.key]['value'], list):
            if isinstance(config[self.key]['value'][0], int):
                config[self.key]['value'] = list(map(int, self.sv.get().split(', ')))
            else:
                config[self.key]['value'] = self.sv.get().split(', ')
        elif isinstance(config[self.key]['value'], int):
            config[self.key]['value'] = int(self.sv.get())
        else:
            config[self.key]['value'] = self.sv.get()

    def browse_callback(self):
        """
        Asks for the source zip file, the opened dialog filters for zip files by default
        The src_entry attribute is updated based on selection
        """
        path = filedialog.askopenfilename(initialdir='.\\',
                                          title="Choose file...",)
        self.entry.delete(0, END)
        self.entry.insert(0, path)


class ConfigWindow:
    """
    Sub window for settings. The window is hidden by default, when the user clicks  to the settings
    button it is activated. It contains the configuration options.
    There are two buttons the Save ( which hides the window ), and the Accept, both of them updates
    the configuration file. The configuration items are stored in a list.

    :param master: Tk parent class
    """
    def __init__(self, master):
        self.master = master
        self.window = Toplevel(self.master)
        self.window.resizable(False, False)
        self.window.withdraw()
        self.window.protocol("WM_DELETE_WINDOW", self._on_closing)
        self.window.title('Settings...')
        self.conf_list = []

        self.main_frame = ttk.Frame(self.window)
        self.main_frame.pack(padx=5, pady=5)
        self.main_frame.grid_columnconfigure(1, minsize=20, weight=1)

        ttk.Label(self.main_frame, text='Configuration:').grid(row=0, column=0,
                                                               columnspan=2, sticky='w')

        row = 1
        for conf_element in config:
            if config[conf_element]['Display'] is True:
                self.conf_list.append(
                    ConfOption(root=self.main_frame, key=conf_element, row=row))
                row += 1

        ttk.Button(self.main_frame, text='Save',
                   command=self.save_callback).grid(row=row, column=0, sticky='e')

        ttk.Button(self.main_frame, text='Accept',
                   command=self.accept_callback).grid(row=row, column=1, sticky='w')

    def save_callback(self):
        """
        Hides the ConfigWindow and updates and stores the configuration
        """
        self.window.withdraw()
        self.accept_callback()

    def accept_callback(self):
        """
        Goes through on every configuration item and updates them one by one. Stores the updated
        configuration.
        """
        for conf in self.conf_list:
            conf.update_config()
        config.store()

    def _on_closing(self):
        self.window.withdraw()


class PdfXlsxGui:
    """
    Simple GUI which lets the user select the source file zip and the destination directory
    for the xlsx file. Contains a file dialog for selecting the zip file to work with.
    There is a button to start the conversion, and also a Settings button to open the
    settings window

    :param master: Tk parent class
    """

    def __init__(self, master):
        self.master = master
        self.master.title('Convert Zip -> Pdf -> Xlsx')
        self.master.resizable(False, False)

        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(padx=5, pady=5)

        self.option_list = ["zip pdf to xlsx", "order details"]
        self.selected_task = StringVar(self.main_frame)
        self.selected_task.set("zip Pdf to xlsx")  # default value

        self.box = ttk.Combobox(self.main_frame, textvariable=self.selected_task, values=self.option_list)
        self.box.bind("<<ComboboxSelected>>", self.update_task)
        self.box.grid(row=0, column=0, columnspan=2)

        self.task_do = self.process_pdf

        ttk.Label(self.main_frame, text='Source File:').grid(row=1, column=0, sticky='w')
        self.src_entry = ttk.Entry(self.main_frame, width=54)
        self.src_entry.grid(row=1, column=0, sticky='e')
        self.src_entry.insert(0, '.\\src.zip')
        ttk.Button(self.main_frame, text='Browse...',
                   command=self.browse_src_callback).grid(row=1, column=1, sticky='w')

        ttk.Button(self.main_frame, text='Start conversion',
                   command=self.execute_task).grid(row=5, column=0, sticky='w')

        ttk.Button(self.main_frame, text='Settings',
                   command=self.config_callback).grid(row=5, column=1, columnspan=1, sticky='e')

        self.config_window = ConfigWindow(self.master)
        self.filetypes = (("zip files", "*.zip"), ("all files", "*.*"))

    def update_task(self, event):
        print(event.widget.get())
        if event.widget.get() == self.option_list[0]:
            self.task_do = self.process_pdf
            self.filetypes = (("zip files", "*.zip"), ("all files", "*.*"))
        elif event.widget.get() == self.option_list[1]:
            self.task_do = self.convert_xlsx
            self.filetypes = (("xlsx files", "*.xlsx"), ("all files", "*.*"))
        else:
            self.task_do = self.unknown_task

    def config_callback(self):
        """
        Bring the configuration window up
        """
        self.config_window.window.state('normal')
        self.config_window.window.lift(self.master)

    def browse_src_callback(self):
        """
        Asks for the source zip file, the opened dialog filters for zip files by default
        The src_entry attribute is updated based on selection
        """
        path = filedialog.askopenfilename(initialdir=config['last_path']['value'],
                                          title="Choose the Zip file...",
                                          filetypes=self.filetypes)
        config['last_path']['value'] = os.path.dirname(path)
        config.store()
        self.src_entry.delete(0, END)
        self.src_entry.insert(0, path)

    def execute_task(self):
        self.task_do()

    def process_pdf(self):
        """
        Facade for the do_it function. Only the src file and destination dir is updated
        the other parameters are left for defaults.
        """
        try:
            logger = do_it(src_name=self.src_entry.get(),
                           dst_dir=config['tmp_dir']['value'],
                           xlsx_name=config['xlsx_name']['value'],
                           tmp_dir=config['tmp_dir']['value'],
                           file_extension=config['file_extension']['value'])
            # tmp_str = '{1} Invoices were found with the following number of Entries:\n{0!s}'
            # messagebox.showinfo(title='Conversion Completed',
            #                    message=tmp_str.format(logger, len(logger.invo_list)))
        except PermissionError as exc:
            messagebox.showerror('Exception', exc)

    def convert_xlsx(self):
        print("Convert those xlsx: {}".format(self.box.get()))
        try:
            logger = do_it2(src_name=self.src_entry.get(),
                            dst_dir=config['tmp_dir']['value'],
                            xlsx_name=config['xlsx_name']['value'],
                            tmp_dir=config['tmp_dir']['value'])
            # tmp_str = '{1} Invoices were found with the following number of Entries:\n{0!s}'
            # messagebox.showinfo(title='Conversion Completed',
            #                    message=tmp_str.format(logger, len(logger.invo_list)))
        except PermissionError as exc:
            messagebox.showerror('Exception', exc)

    def unknown_task(self):
        print("Unknown task selected: {}".format(self.box.get()))


def main():
    root = Tk()

    def _post_clean_up():
        try:
            shutil.rmtree(config['tmp_dir']['value'])
        except FileNotFoundError:
            print("You did nothing, you dummy, why did you start me up???")
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", _post_clean_up)
    gui = PdfXlsxGui(root)
    root.mainloop()

if __name__ == '__main__':
    main()
