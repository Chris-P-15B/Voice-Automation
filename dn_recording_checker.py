#!/usr/bin/env python
# (c) 2018 - 2019, Chris Perkins
# For a list of DNs in a CSV file, find phones (tkclass=1) & device profiles (tkclass=254) where built-in
# bridge isn’t on or privacy isn’t off, automatic call recording isn't enabled, recording profile doesn't
# match & recording media source isn't phone preferred. Optionally output to another CSV file

# v1.2 - code tidying
# v1.1 - fixes some edge cases
# v1.0 - original release

# Original AXL SQL query code courtesy of Jonathan Els - https://afterthenumber.com/2018/04/27/serializing-thin-axl-sql-query-responses-with-python-zeep/

# To Do:
# Improve the GUI

import sys, json, csv
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

# GUI and main code
class GUIFrame(tk.Frame):

    def __init__(self, parent):
        """Constructor checks parameters and initialise variables"""
        self.axl_input_filename = None
        self.axl_password = ''
        self.csv_input_filename = None
        tk.Frame.__init__(self, parent)
        parent.geometry("320x480")
        self.pack(fill=tk.BOTH, expand=True)
        menu_bar = tk.Menu(self)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Load AXL", command=self.open_json_file_dialog)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menu_bar.add_cascade(label="File", menu=file_menu)
        parent.config(menu=menu_bar)
        tk.Label(self, text="Output Filename:").place(relx=0.2, rely=0.0, height=22, width=200)
        self.output_csv_text = tk.StringVar()
        tk.Entry(self, textvariable=self.output_csv_text).place(relx=0.2, rely=0.05, height=22, width=200)
        tk.Button(self, text="Check Recording Config", command=self.check_recording).place(relx=0.265,
            rely=0.12, height=22, width=160)
        self.results_count_text = tk.StringVar()
        self.results_count_text.set("Results Found: ")
        tk.Label(self, textvariable=self.results_count_text).place(relx=0.35, rely=0.18, height=22, width=110)
        list_box_frame = tk.Frame(self, bd=2, relief=tk.SUNKEN)
        list_box_scrollbar_y = tk.Scrollbar(list_box_frame)
        list_box_scrollbar_x = tk.Scrollbar(list_box_frame, orient=tk.HORIZONTAL)
        self.list_box = tk.Listbox(list_box_frame, xscrollcommand=list_box_scrollbar_x.set,
            yscrollcommand=list_box_scrollbar_y.set)
        list_box_frame.place(relx=0.02, rely=0.22, relheight=0.75, relwidth=0.96)
        list_box_scrollbar_y.place(relx=0.94, rely=0.0, relheight=1.0, relwidth=0.06)
        list_box_scrollbar_x.place(relx=0.0, rely=0.94, relheight=0.06, relwidth=0.94)
        self.list_box.place(relx=0.0, rely=0.0, relheight=0.94, relwidth=0.94)
        list_box_scrollbar_y.config(command=self.list_box.yview)
        list_box_scrollbar_x.config(command=self.list_box.xview)

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

    def read_axl(self, dn_list, output_filename):
        """Check configuration via AXL SQL query"""
        try:
            self.list_box.delete(0, tk.END)
            self.results_count_text.set("Results Found: ")
            with open(self.axl_input_filename) as f:
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
                    try:
                        if not axl_json['subquery']:
                            tk.messagebox.showerror(title="Error", message="Subquery must be specified.")
                            return
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="Subquery must be specified.")
                        return
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            return
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            return

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

        # For each DN read from CSV file
        cntr = 0
        result_list = [["Device Name", "Device Description", "DN", "DN Description"]]
        for dn in dn_list:
            # Check for phones (tkclass=1)
            sql_statement = f"SELECT d.name, d.description, n.dnorpattern, n.description AS ndescription " \
                f"FROM device d INNER JOIN devicenumplanmap dnmap ON dnmap.fkdevice=d.pkid INNER JOIN numplan n " \
                f"ON dnmap.fknumplan=n.pkid INNER JOIN deviceprivacydynamic dpd ON dpd.fkdevice=d.pkid " \
                f"INNER JOIN recordingdynamic rd ON rd.fkdevicenumplanmap=dnmap.pkid WHERE (d.tkclass=1 " \
                f"AND n.dnorpattern='{dn}') AND (d.tkstatus_builtinbridge!=1 OR dpd.tkstatus_callinfoprivate!=0 " \
                f"OR {axl_json['subquery']} OR dnmap.fkrecordingprofile IS NULL OR dnmap.tkpreferredmediasource!=2 " \
                f"OR rd.tkrecordingflag!=1) ORDER BY d.name"
            try:
                for row in self.sql_query(service=axl, sql_statement=sql_statement):
                    try:
                        # Handle None results
                        if row['name'] is None:
                            d_name = ''
                        else:
                            d_name = row['name']
                        if row['description'] is None:
                            d_description = ''
                        else:
                            d_description = row['description']
                        if row['dnorpattern'] is None:
                            n_dnorpattern = ''
                        else:
                            n_dnorpattern = row['dnorpattern']
                        if row['ndescription'] is None:
                            n_description = ''
                        else:
                            n_description = row['ndescription']

                        self.list_box.insert(tk.END, f'{d_name} "{d_description}", {n_dnorpattern} "{n_description}"')
                        result_list.append(list(row.values()))
                        cntr += 1
                    except TypeError:
                        continue
            except TypeError:
                pass
            except Fault as thin_axl_error:
                tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
                return

            # Check for device profiles (tkclass=254)
            sql_statement = f"SELECT d.name, d.description, n.dnorpattern, n.description AS ndescription " \
                f"FROM device d INNER JOIN devicenumplanmap dnmap ON dnmap.fkdevice=d.pkid INNER JOIN numplan n " \
                f"ON dnmap.fknumplan=n.pkid INNER JOIN deviceprivacydynamic dpd ON dpd.fkdevice=d.pkid " \
                f"INNER JOIN recordingdynamic rd ON rd.fkdevicenumplanmap=dnmap.pkid WHERE (d.tkclass=254 " \
                f"AND n.dnorpattern='{dn}') AND (dpd.tkstatus_callinfoprivate!=0 OR {axl_json['subquery']} " \
                f"OR dnmap.fkrecordingprofile IS NULL OR dnmap.tkpreferredmediasource!=2 OR rd.tkrecordingflag!=1) " \
                f"ORDER BY d.name"
            try:
                for row in self.sql_query(service=axl, sql_statement=sql_statement):
                    try:
                        # Handle None results
                        if row['name'] is None:
                            d_name = ''
                        else:
                            d_name = row['name']
                        if row['description'] is None:
                            d_description = ''
                        else:
                            d_description = row['description']
                        if row['dnorpattern'] is None:
                            n_dnorpattern = ''
                        else:
                            n_dnorpattern = row['dnorpattern']
                        if row['ndescription'] is None:
                            n_description = ''
                        else:
                            n_description = row['ndescription']

                        self.list_box.insert(tk.END, f'{d_name} "{d_description}", {n_dnorpattern} "{n_description}"')
                        result_list.append(list(row.values()))
                        cntr += 1
                    except TypeError:
                        continue
            except TypeError:
                pass
            except Fault as thin_axl_error:
                tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
                return
        self.results_count_text.set(f"Results Found: {str(cntr)}")

        # Output to CSV file if required
        try:
            if len(output_filename) != 0:
                with open(output_filename, 'w', newline='') as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerows(result_list)
        except OSError:
            tk.messagebox.showerror(title="Error", message="Unable to write CSV file.")

    def check_recording(self):
        """Validate parameters, read CSV file of DNs and then call AXL query"""
        if not self.axl_input_filename:
            tk.messagebox.showerror(title="Error", message="No AXL file selected.")
            return
        if not self.csv_input_filename:
            tk.messagebox.showerror(title="Error", message="No CSV file selected.")
            return
        # Parse input CSV file
        dn_list = []
        try:
            with open(self.csv_input_filename, encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                for row in reader:
                    dn_list.append(row[0])
        except FileNotFoundError:
            tk.messagebox.showerror(title="Error", message="Unable to open CSV file.")
            return

        output_string = self.output_csv_text.get()
        if len(output_string) == 0:
            self.read_axl(dn_list, '')
        else:
            self.read_axl(dn_list, output_string)

    def open_json_file_dialog(self):
        """Dialogue to prompt for JSON file to open and AXL password"""
        self.axl_input_filename = tk.filedialog.askopenfilename(initialdir='/', filetypes=(("JSON files",
            "*.json"),("All files", "*.*")))
        self.axl_password = tk.simpledialog.askstring("Input", "AXL Password?", show='*')
        self.csv_input_filename = tk.filedialog.askopenfilename(initialdir='/', filetypes=(("CSV files",
            "*.csv"),("All files", "*.*")))

if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    # Initialise TKinter GUI objects
    root = tk.Tk()
    root.title("DN Recording Checker v1.2")
    GUIFrame(root)
    root.mainloop()