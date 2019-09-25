#!/usr/bin/env python
# (c) 2017 - 2019, Chris Perkins
# Takes CUCM Route Plan Report exported as CSV or uses AXL, parses the regexs for the dial plan to find
# unused numbers in a given direct dial range. Number range to match against is defined in JSON format in dialplan.json
# Won't parse dial plan entries with * or # as they're invalid for a direct dial range

# v1.5 - code tidying
# v1.4 - GUI adjustments & fixes some edge cases
# v1.3 – added AXL support
# v1.2 – bug fixes
# v1.1 – added GUI
# v1.0 – initial release with only CSV file support and CLI usage

# Original AXL SQL query code courtesy of Jonathan Els - https://afterthenumber.com/2018/04/27/serializing-thin-axl-sql-query-responses-with-python-zeep/

# To Do:
# Improve the GUI
# Add number classification, e.g. bronze, silver, gold & platinum

import itertools, csv, sys, json
import tkinter as tk
import requests
from tkinter import ttk
from tkinter import filedialog, simpledialog, messagebox
from collections import OrderedDict
from zeep import Client
from zeep.cache import SqliteCache
from zeep.transports import Transport
from zeep.plugins import HistoryPlugin
from zeep.exceptions import Fault
from zeep.helpers import serialize_object
from requests import Session
from requests.auth import HTTPBasicAuth
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
from lxml import etree

# Stores information about numbers in a range
class DirectoryNumbers:
    def __init__(self, start_num, end_num):
        """Constructor initialises attributes"""
        self.number = []
        self.is_used = []
        self.classification = []

        for num in range(int(start_num), int(end_num) + 1):
            num_str = str(num)
            # For numbers with preceeding 0, conversion to int will strip, so prepend with 0 to match
            #  length of source string
            if len(num_str) < len(end_num):
                pad_str = ''
                for x in range(0, len(end_num) - len(num_str)):
                    pad_str += '0'
                num_str = pad_str + num_str
            self.number.append(num_str)
            self.is_used.append(False)
            self.classification.append(0)

# GUI and main code
class GUIFrame(tk.Frame):
    def __init__(self, parent):
        """Constructor checks parameters and initialise variables"""
        self.range_descriptions = []
        self.numbers = []
        self.input_filename = None
        self.use_axl = False
        self.axl_password = ''

        try:
            with open("dialplan.json") as f:
                self.json_data = json.load(f)
                for range_data in self.json_data:
                    try:
                        if len(range_data['range_start']) != len(range_data['range_end']):
                            tk.messagebox.showerror(title="Error", message="The first and last numbers"
                                " in range must be of equal length.")
                            sys.exit()
                        elif int(range_data['range_start']) >= int(range_data['range_end']):
                            tk.messagebox.showerror(title="Error", message="The last number in range"
                                " must be greater than the first.")
                            sys.exit()
                    except (TypeError, ValueError, KeyError):
                        tk.messagebox.showerror(title="Error", message="Number range parameters incorrectly"
                            " formatted.")
                        sys.exit()
                    try:
                        if not range_data['description']:
                            tk.messagebox.showerror(title="Error", message="Description must be specified.")
                            sys.exit()
                        # Uncomment to disallow DNs not in a partition
                        #elif not range_data['partition']:
                        #    tk.messagebox.showerror(title="Error", message="Partition must be specified.")
                        #    sys.exit()
                        self.range_descriptions.append(range_data['description'])
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="Description must be specified.")
                        sys.exit()
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            sys.exit()
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            sys.exit()

        self.range_descriptions = sorted(self.range_descriptions)
        for item in self.json_data:
            if item['description'].upper() == self.range_descriptions[0].upper():
                self.range_description = item['description']
                self.range_start = int(item['range_start'])
                self.range_end = int(item['range_end'])
                self.range_partition = item['partition']
                self.directory_numbers = DirectoryNumbers(item['range_start'], item['range_end'])
                break

        tk.Frame.__init__(self, parent)
        parent.geometry("320x480")
        self.pack(fill=tk.BOTH, expand=True)
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load AXL", command=self.open_json_file_dialog)
        file_menu.add_command(label="Load CSV", command=self.open_csv_file_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        parent.config(menu=menu_bar)
        tk.Label(self, text="DN Range:").place(relx=0.4, rely=0.0, height=22, width=62)
        self.range_combobox = ttk.Combobox(self, values=self.range_descriptions, state="readonly")
        self.range_combobox.current(0)
        self.range_combobox.bind("<<ComboboxSelected>>", self.combobox_update)
        self.range_combobox.place(relx=0.02, rely=0.042, relheight=0.06, relwidth=0.96)
        tk.Button(self, text="Find Unused DNs", command=self.find_unused_dns).place(relx=0.35, rely=0.12,
            height=22, width=100)
        self.unused_label_text = tk.StringVar()
        self.unused_label_text.set("Unused DNs: ")
        tk.Label(self, textvariable=self.unused_label_text).place(relx=0.35, rely=0.18, height=22, width=110)
        list_box_frame = tk.Frame(self, bd=2, relief=tk.SUNKEN)
        list_box_scrollbar_y = tk.Scrollbar(list_box_frame)
        list_box_scrollbar_x = tk.Scrollbar(list_box_frame, orient=tk.HORIZONTAL)
        self.list_box = tk.Listbox(list_box_frame, xscrollcommand=list_box_scrollbar_x.set,
            yscrollcommand=list_box_scrollbar_y.set)
        list_box_frame.place(relx=0.02, rely=0.22, relheight=0.73, relwidth=0.96)
        list_box_scrollbar_y.place(relx=0.94, rely=0.0, relheight=1.0, relwidth=0.06)
        list_box_scrollbar_x.place(relx=0.0, rely=0.94, relheight=0.06, relwidth=0.94)
        self.list_box.place(relx=0.0, rely=0.0, relheight=0.94, relwidth=0.94)
        list_box_scrollbar_y.config(command=self.list_box.yview)
        list_box_scrollbar_x.config(command=self.list_box.xview)
        self.entries_label_text = tk.StringVar()
        self.entries_label_text.set("Dial Plan Entries Parsed: ")
        tk.Label(self, textvariable=self.entries_label_text).place(relx=0.21, rely=0.95, height=22, width=220)

    def parse_regex(self, pattern, range_start, range_end):
        """Parse CUCM regex pattern and return list of the digit strings the regex matches within the
         number range specified"""
        is_slice = False
        is_range = False
        is_negate = False
        num_digits = 0
        digits = []
        numbers_in_use = []

        # Parse regex and store digits in jagged list
        for column in range(16):
            digits.append([])
        for char in pattern:
            if char == '[':
                is_slice = True
            elif char == '^' and is_slice == True:
                is_negate = True
            elif char == ']':
                is_slice = False
                if is_negate == True:
                    negate_slice = []
                    for range_char in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                        if range_char not in digits[num_digits]:
                            negate_slice.append(range_char)
                    digits[num_digits] = negate_slice[:]
                    is_negate = False
                num_digits += 1
            elif char in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']:
                if is_range == False:
                    digits[num_digits].append(char)
                    if is_slice == False:
                        num_digits += 1
                else:
                    for range_char in range(int(digits[num_digits][-1]) + 1, int(char) + 1):
                        digits[num_digits].append(str(range_char))
                    is_range = False
            elif char == '-' and is_slice == True:
                is_range = True
            elif char == 'X':
                digits[num_digits] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
                num_digits += 1
            elif char == '*' or char == '#':
                # Strings containing * or # can't be parsed as an integer so return empty list as also
                # not a valid PSTN number
                return []

        # Strip empty lists
        digits2 = [x for x in digits if x != []]

        # Use itertools.product() to convert jagged list of digits to list of permutation strings
        #  >= range_start & <= range_end
        for list in itertools.product(*digits2):
            char_string = ''
            for char in list:
                char_string += str(char)
            if char_string:
                number = int(char_string)
                if number >= range_start and number <= range_end:
                    numbers_in_use.append(char_string)

        return numbers_in_use

    def element_list_to_ordered_dict(self, elements):
        """Convert list to OrderedDict"""
        return [OrderedDict((element.tag, element.text) for element in row) for row in elements]


    def sql_query(self, service, sql_statement):
        """Execute SQL query via AXL and return results"""
        try:
            axl_resp = service.executeSQLQuery(sql=sql_statement)
            try:
                return self.element_list_to_ordered_dict(serialize_object(axl_resp)["return"]["rows"])
            except KeyError:
                # Single tuple response
                return self.element_list_to_ordered_dict(serialize_object(axl_resp)["return"]["row"])
            except TypeError:
                # No SQL tuples
                return serialize_object(axl_resp)["return"]
        except requests.exceptions.ConnectionError as e:
            tk.messagebox.showerror(title="Error", message=str(e))
            return None

    def read_axl(self):
        """Read and parse Route Plan via AXL"""
        try:
            self.list_box.delete(0, tk.END)
            with open(self.input_filename) as f:
                axl_json_data = json.load(f)
                for axl_json in axl_json_data:
                    try:
                        if not axl_json['fqdn']:
                            tk.messagebox.showerror(title="Error", message="FQDN must be specified.")
                            return
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="FQDN must be specified.")
                        return
                    try:
                        if not axl_json['username']:
                            tk.messagebox.showerror(title="Error", message="Username must be specified.")
                            return
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="Username must be specified.")
                        return
                    try:
                        if not axl_json['wsdl_file']:
                            tk.messagebox.showerror(title="Error", message="WSDL file must be specified.")
                            return
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="WSDL file must be specified.")
                        return
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            return
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            return

        sql_statement = "SELECT n.dnorpattern, p.name FROM numplan n LEFT JOIN routepartition p ON " \
            "n.fkroutepartition=p.pkid"
        axl_binding_name = "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding"
        axl_address = f"https://{axl_json['fqdn']}:8443/axl/"
        session = Session()
        session.verify = False
        session.auth = HTTPBasicAuth(axl_json['username'], self.axl_password)
        transport = Transport(cache=SqliteCache(), session=session, timeout=60)
        history = HistoryPlugin()
        try:
            client = Client(wsdl=axl_json['wsdl_file'], transport=transport, plugins=[history])
        except FileNotFoundError as e:
            tk.messagebox.showerror(title="Error", message=str(e))
            return
        axl = client.create_service(axl_binding_name, axl_address)

        try:
            raw_route_plan = []
            for row in self.sql_query(service=axl, sql_statement=sql_statement):
                # Ignore entries not in the correct partition and update directory_numbers with numbers
                # found to be in use
                if row['name'] is None:
                    pname = ''
                else:
                    pname = row['name']
                if pname.upper() == self.range_partition.upper():
                    for char_string in self.parse_regex(row['dnorpattern'], self.range_start, self.range_end):
                        raw_route_plan.append(char_string)
                        try:
                            dn_index = self.directory_numbers.number.index(char_string)
                            self.directory_numbers.is_used[dn_index] = True
                        except (IndexError, ValueError):
                            continue
        except TypeError:
            return
        except Fault as thin_axl_error:
            tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
            return

        # Update TKinter display objects with results
        self.entries_label_text.set(f"Dial Plan Entries Parsed: {str(len(raw_route_plan))}")
        cntr = 0
        for num in range(0, len(self.directory_numbers.number)):
            if self.directory_numbers.is_used[num] == False:
                cntr += 1
                self.list_box.insert(tk.END, f"{self.directory_numbers.number[num]} / {self.range_partition}")
        self.unused_label_text.set(f"Unused DNs: {str(cntr)}")

    def read_csv_file(self):
        """Read and parse Route Plan Report CSV file"""
        column_index = []

        try:
            self.list_box.delete(0, tk.END)
            # encoding="utf-8-sig" is necessary for correct parsing fo UTF-8 encoding of CUCM Route
            # Plan Report CSV file
            with open(self.input_filename, encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header_row = next(reader)
                for index, column_header in enumerate(header_row):
                    if column_header == "Pattern or URI":
                        column_index.append(index)
                    elif column_header == "Pattern/Directory Number":
                        column_index.append(index)
                    elif column_header == "Partition":
                        column_index.append(index)
                if len(column_index) != 2:
                    tk.messagebox.showerror(title="Error", message="Unable to parse CSV file.")
                    return
                raw_route_plan = []
                for row in reader:
                    # Ignore entries not in the correct partition and update directory_numbers with
                    # numbers found to be in use
                    if row[column_index[1]].upper() == self.range_partition.upper():
                        for char_string in self.parse_regex(row[column_index[0]], self.range_start, self.range_end):
                            raw_route_plan.append(char_string)
                            try:
                                dn_index = self.directory_numbers.number.index(char_string)
                                self.directory_numbers.is_used[dn_index] = True
                            except (IndexError, ValueError):
                                pass
        except FileNotFoundError:
            tk.messagebox.showerror(title="Error", message="Unable to open CSV file.")
            return

        # Update TKinter display objects
        self.entries_label_text.set(f"Dial Plan Entries Parsed: {str(len(raw_route_plan))}")
        cntr = 0
        for num in range(0, len(self.directory_numbers.number)):
            if self.directory_numbers.is_used[num] == False:
                cntr += 1
                self.list_box.insert(tk.END, f"{self.directory_numbers.number[num]} / {self.range_partition}")
        self.unused_label_text.set(f"Unused DNs: {str(cntr)}")

    def find_unused_dns(self):
        """Check AXL or CSV selected and hand over to correct method to handle"""
        if self.use_axl:
            if not self.input_filename:
                tk.messagebox.showerror(title="Error", message="No AXL file selected.")
                return
            else:
                self.read_axl()
        else:
            if not self.input_filename:
                tk.messagebox.showerror(title="Error", message="No CSV file selected.")
                return
            else:
                self.read_csv_file()

    def open_csv_file_dialog(self):
        """Dialogue to prompt for CSV file to open"""
        self.input_filename = tk.filedialog.askopenfilename(initialdir='/', filetypes=(("CSV files",
            "*.csv"),("All files", "*.*")))
        self.use_axl = False
        self.axl_password = ""

    def open_json_file_dialog(self):
        """Dialogue to prompt for JSON file to open and AXL password"""
        self.input_filename = tk.filedialog.askopenfilename(initialdir='/', filetypes=(("JSON files",
            "*.json"),("All files", "*.*")))
        self.use_axl = True
        self.axl_password = tk.simpledialog.askstring("Input", "AXL Password?", show='*')

    def combobox_update(self, event):
        """Populate range variables when Combobox item selected"""
        self.list_box.delete(0, tk.END)
        self.unused_label_text.set("Unused DNs: ")
        self.entries_label_text.set("Dial Plan Entries Parsed: ")
        value = self.range_combobox.get()
        for item in self.json_data:
            if item['description'].upper() == value.upper():
                self.range_description = item['description']
                self.range_start = int(item['range_start'])
                self.range_end = int(item['range_end'])
                self.range_partition = item['partition']
                self.directory_numbers = DirectoryNumbers(item['range_start'], item['range_end'])
                break

if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    # Initialise TKinter GUI objects
    root = tk.Tk()
    root.title("Dial Plan Analyser v1.5")
    GUIFrame(root)
    root.mainloop()