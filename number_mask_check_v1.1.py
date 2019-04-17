#!/usr/bin/env python
# v1.0 - written by Chris Perkins in 2019
# Finds & fixes primary DN's in specified range(s) with an External Phone Number Masks that doesn't match the approved list

# v1.1 - fixed CSV output to UTF-8
# v1.0 – initial release

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
        self.axl_password = ""
        self.csv_input_filename = None

        try:
            with open("dialplan.json") as f:
                self.json_data = json.load(f)
                for range_data in self.json_data:
                    try:
                        if len(range_data['range_start']) != len(range_data['range_end']):
                            tk.messagebox.showerror(title="Error", message="The first and last numbers in range must be of equal length.")
                            sys.exit()
                        elif int(range_data['range_start']) >= int(range_data['range_end']):
                            tk.messagebox.showerror(title="Error", message="The last number in range must be greater than the first.")
                            sys.exit()
                    except (TypeError, ValueError, KeyError):
                        tk.messagebox.showerror(title="Error", message="Number range parameters incorrectly formatted.")
                        sys.exit()
                    try:
                        if not range_data['mask']:
                            tk.messagebox.showerror(title="Error", message="Number mask must be specified.")
                            sys.exit()
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="Number mask must be specified.")
                        sys.exit()
                    try:
                        if not range_data['partition']:
                            tk.messagebox.showerror(title="Error", message="Partition must be specified.")
                            sys.exit()
                    except KeyError:
                        tk.messagebox.showerror(title="Error", message="Partition must be specified.")
                        sys.exit()
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            sys.exit()
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            sys.exit()

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
        tk.Button(self, text="Check Number Masks", command=self.check_masks).place(relx=0.08, rely=0.12, height=22, width=135)
        tk.Button(self, text="Update Number Masks", command=self.update_masks).place(relx=0.52, rely=0.12, height=22, width=135)
        self.results_count_text = tk.StringVar()
        self.results_count_text.set("Results Found: ")
        tk.Label(self, textvariable=self.results_count_text).place(relx=0.20, rely=0.18, height=22, width=210)
        list_box_frame = tk.Frame(self, bd=2, relief=tk.SUNKEN)
        list_box_scrollbar_y = tk.Scrollbar(list_box_frame)
        list_box_scrollbar_x = tk.Scrollbar(list_box_frame, orient=tk.HORIZONTAL)
        self.list_box = tk.Listbox(list_box_frame, xscrollcommand=list_box_scrollbar_x.set, yscrollcommand=list_box_scrollbar_y.set)
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

    def sql_update(self, service, sql_statement):
        """Execute SQL update via AXL and return rows updated"""
        try:
            axl_resp = service.executeSQLUpdate(sql=sql_statement)
            return serialize_object(axl_resp)["return"]["rowsUpdated"]
        except requests.exceptions.ConnectionError as e:
            tk.messagebox.showerror(title="Error", message=str(e))
            return None

    def read_axl(self, output_filename):
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
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            return
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            return

        axl_binding_name = "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding"
        axl_address = "https://{fqdn}:8443/axl/".format(fqdn=axl_json['fqdn'])
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

        # List each primary DN in specified range(s) with an External Phone Number Mask that doesn't match the approved list
        cntr = 0
        result_list = [["DN", "Partition", "Device Name", "Device Description", "Number Mask", "New Number Mask", "pkid"]]
        self.list_box.insert(tk.END, "DN, Partition, Device Name, Device Description, Number Mask, New Number Mask, pkid\n")
        sql_statement = "SELECT n.dnorpattern, p.name AS pname, d.name, d.description, dnmap.e164mask, dnmap.pkid FROM device d INNER JOIN devicenumplanmap dnmap ON dnmap.fkdevice=d.pkid INNER JOIN numplan n ON dnmap.fknumplan=n.pkid LEFT JOIN routepartition p ON n.fkroutepartition=p.pkid WHERE (d.tkclass=1 OR d.tkclass=254) AND dnmap.numplanindex=1 ORDER BY n.dnorpattern"
        try:
            for row in self.sql_query(service=axl, sql_statement=sql_statement):
                try:
                    # Handle None results
                    if row['pkid'] is None:
                        dnmap_pkid = ""
                    else:
                        dnmap_pkid = row['pkid']
                    if row['e164mask'] is None:
                        dnmap_e164mask = ""
                    else:
                        dnmap_e164mask = row['e164mask']
                    if row['description'] is None:
                        d_description = ""
                    else:
                        d_description = row['description']
                    if row['name'] is None:
                        d_name = ""
                    else:
                        d_name = row['name']
                    if row['pname'] is None:
                        p_name = ""
                    else:
                        p_name = row['pname']
                    if row['dnorpattern'] is None:
                        n_dnorpattern = ""
                    else:
                        n_dnorpattern = row['dnorpattern']

                    is_valid_mask = False
                    is_in_range = False
                    correct_mask = ""
                    for range_data in self.json_data:
                        try:
                            range_start = int(range_data['range_start'])
                            range_end = int(range_data['range_end'])
                            dn = int(n_dnorpattern)
                            if p_name.upper() == range_data['partition'].upper() and dn >= range_start and dn <= range_end:
                                if dnmap_e164mask.upper() == range_data['mask'].upper():
                                    is_valid_mask = True
                                    is_in_range = True
                                    break
                                else:
                                    is_in_range = True
                                    correct_mask = range_data['mask']
                                    break
                        except TypeError:
                            continue

                    if is_in_range == True and is_valid_mask == False:
                        self.list_box.insert(tk.END, n_dnorpattern + ', ' + p_name + ', ' + d_name + ', ' + d_description + ', ' + dnmap_e164mask + ', ' + correct_mask + ', ' + dnmap_pkid)
                        result_list.append([n_dnorpattern, p_name, d_name, d_description, dnmap_e164mask, correct_mask, dnmap_pkid])
                        cntr += 1
                except TypeError:
                    continue
        except TypeError:
            pass
        except Fault as thin_axl_error:
            tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
            return

        self.results_count_text.set("Results Found: " + str(cntr))
        # Output to CSV file if required
        try:
            if len(output_filename) != 0:
                with open(output_filename, 'w', newline='', encoding='utf-8-sig') as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerows(result_list)
        except OSError:
            tk.messagebox.showerror(title="Error", message="Unable to write CSV file.")

    def write_axl(self, output_filename):
        """Update configuration via AXL SQL query"""
        try:
            self.list_box.delete(0, tk.END)
            self.results_count_text.set("Updates Made: ")
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
        except FileNotFoundError:
            messagebox.showerror(title="Error", message="Unable to open JSON file.")
            return
        except json.decoder.JSONDecodeError:
            messagebox.showerror(title="Error", message="Unable to parse JSON file.")
            return

        axl_binding_name = "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding"
        axl_address = "https://{fqdn}:8443/axl/".format(fqdn=axl_json['fqdn'])
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

        # Update External Phone Number Masks contained in CSV file
        cntr = 0
        result_list = [["DN", "Partition", "Device Name", "Device Description", "Number Mask", "New Number Mask", "pkid"]]
        self.list_box.insert(tk.END, "DN, Partition, Device Name, Device Description, Number Mask, New Number Mask, pkid\n")

        # Parse input CSV file & make updates based on the content
        try:
            with open(self.csv_input_filename, encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                header_row = next(reader)
                if header_row[5] != "New Number Mask" or header_row[6] != "pkid":
                    tk.messagebox.showerror(title="Error", message="Unable to parse CSV file.")
                    return
                for row in reader:
                    try:
                        # Check replacement mask has only valid characters
                        is_valid = True
                        for mask_char in row[5]:
                            if mask_char not in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'X']:
                                self.list_box.insert(tk.END, row[0] + ', ' + row[1] + ', ' + row[2] + ', ' + row[3] + ', ' + row[4] + ', ' + row[5] + ', ' + row[6])
                                result_list.append(row)
                                is_valid = False
                                break
                        if is_valid == False:
                            continue

                        sql_statement = "UPDATE devicenumplanmap SET e164mask='" + row[5] + "' WHERE pkid='" + row[6] + "'"
                        num_results = self.sql_update(service=axl, sql_statement=sql_statement)
                        # List updates that failed
                        if num_results < 1:
                            self.list_box.insert(tk.END, row[0] + ', ' + row[1] + ', ' + row[2] + ', ' + row[3] + ', ' + row[4] + ', ' + row[5] + ', ' + row[6])
                            result_list.append(row)
                        else:
                            cntr += 1
                    except TypeError:
                        continue
                    except Fault as thin_axl_error:
                        tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
                        break
        except KeyError:
            tk.messagebox.showerror(title="Error", message="Unable to parse CSV file.")
            pass
        except FileNotFoundError:
            tk.messagebox.showerror(title="Error", message="Unable to open CSV file.")
            return

        self.results_count_text.set("Updates Made: " + str(cntr) + " (failures below)")
        # Output to CSV file if required
        try:
            if len(output_filename) != 0:
                with open(output_filename, 'w', newline='', encoding='utf-8-sig') as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerows(result_list)
        except OSError:
            tk.messagebox.showerror(title="Error", message="Unable to write CSV file.")

    def check_masks(self):
        """Validate parameters and then call AXL query"""
        if not self.axl_input_filename:
            tk.messagebox.showerror(title="Error", message="No AXL file selected.")
            return

        output_string = self.output_csv_text.get()
        if len(output_string) == 0:
            self.read_axl('')
        else:
            self.read_axl(output_string)

    def update_masks(self):
        """Validate parameters and then call AXL update"""
        if not self.axl_input_filename:
            tk.messagebox.showerror(title="Error", message="No AXL file selected.")
            return

        self.open_csv_file_dialog()
        if not self.csv_input_filename:
            tk.messagebox.showerror(title="Error", message="No CSV file selected.")
            return

        output_string = self.output_csv_text.get()
        if len(output_string) == 0:
            self.write_axl('')
        else:
            self.write_axl(output_string)

    def open_json_file_dialog(self):
        """Dialogue to prompt for JSON file to open and AXL password"""
        self.axl_input_filename = tk.filedialog.askopenfilename(initialdir="/", filetypes=(("JSON files", "*.json"),("All files", "*.*")))
        self.axl_password = tk.simpledialog.askstring("Input", "AXL Password?", show='*')

    def open_csv_file_dialog(self):
        """Dialogue to prompt for CSV file to open"""
        self.csv_input_filename = tk.filedialog.askopenfilename(initialdir="/", filetypes=(("CSV files", "*.csv"),("All files", "*.*")))

if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    # Initialise TKinter GUI objects
    root = tk.Tk()
    root.title("External Number Mask Checker v1.0")
    GUIFrame(root)
    root.mainloop()