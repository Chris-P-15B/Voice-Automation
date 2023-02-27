#!/usr/bin/env python3

"""
Copyright (c) 2019, Chris Perkins
Licence: BSD 3-Clause

Finds & fixes Line Text Labels not in the standard of Initial Last Name-Extension

v1.3 - code tidying
v1.2 - fixed CSV output to UTF-8
v1.1 - fixed single word alerting/display name handling
v1.0 – initial release

Original AXL SQL query code courtesy of Jonathan Els - https://afterthenumber.com/2018/04/27/serializing-thin-axl-sql-query-responses-with-python-zeep/

To Do:
Improve the GUI
"""

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
        tk.Label(self, text="Output Filename:").place(
            relx=0.2, rely=0.0, height=22, width=200
        )
        self.output_csv_text = tk.StringVar()
        tk.Entry(self, textvariable=self.output_csv_text).place(
            relx=0.2, rely=0.05, height=22, width=200
        )
        tk.Button(self, text="Check Line Labels", command=self.check_labels).place(
            relx=0.1, rely=0.12, height=22, width=120
        )
        tk.Button(self, text="Update Line Labels", command=self.update_labels).place(
            relx=0.5, rely=0.12, height=22, width=120
        )
        self.results_count_text = tk.StringVar()
        self.results_count_text.set("Results Found: ")
        tk.Label(self, textvariable=self.results_count_text).place(
            relx=0.20, rely=0.18, height=22, width=210
        )
        list_box_frame = tk.Frame(self, bd=2, relief=tk.SUNKEN)
        list_box_scrollbar_y = tk.Scrollbar(list_box_frame)
        list_box_scrollbar_x = tk.Scrollbar(list_box_frame, orient=tk.HORIZONTAL)
        self.list_box = tk.Listbox(
            list_box_frame,
            xscrollcommand=list_box_scrollbar_x.set,
            yscrollcommand=list_box_scrollbar_y.set,
        )
        list_box_frame.place(relx=0.02, rely=0.22, relheight=0.75, relwidth=0.96)
        list_box_scrollbar_y.place(relx=0.94, rely=0.0, relheight=1.0, relwidth=0.06)
        list_box_scrollbar_x.place(relx=0.0, rely=0.94, relheight=0.06, relwidth=0.94)
        self.list_box.place(relx=0.0, rely=0.0, relheight=0.94, relwidth=0.94)
        list_box_scrollbar_y.config(command=self.list_box.yview)
        list_box_scrollbar_x.config(command=self.list_box.xview)

    def element_list_to_ordered_dict(self, elements):
        """Convert list to OrderedDict"""
        return [
            OrderedDict((element.tag, element.text) for element in row)
            for row in elements
        ]

    def sql_query(self, service, sql_statement):
        """Execute SQL query via AXL and return results"""
        try:
            axl_resp = service.executeSQLQuery(sql=sql_statement)
            try:
                return self.element_list_to_ordered_dict(
                    serialize_object(axl_resp)["return"]["rows"]
                )
            except KeyError:
                # Single tuple response
                return self.element_list_to_ordered_dict(
                    serialize_object(axl_resp)["return"]["row"]
                )
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
                        if not axl_json["fqdn"]:
                            tk.messagebox.showerror(
                                title="Error", message="FQDN must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="FQDN must be specified."
                        )
                        return
                    try:
                        if not axl_json["username"]:
                            tk.messagebox.showerror(
                                title="Error", message="Username must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="Username must be specified."
                        )
                        return
                    try:
                        if not axl_json["wsdl_file"]:
                            tk.messagebox.showerror(
                                title="Error", message="WSDL file must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="WSDL file must be specified."
                        )
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
        session.auth = HTTPBasicAuth(axl_json["username"], self.axl_password)
        transport = Transport(cache=SqliteCache(), session=session, timeout=60)
        history = HistoryPlugin()
        try:
            client = Client(
                wsdl=axl_json["wsdl_file"], transport=transport, plugins=[history]
            )
        except FileNotFoundError as e:
            tk.messagebox.showerror(title="Error", message=str(e))
            return
        axl = client.create_service(axl_binding_name, axl_address)

        # List each Line Text Label for Phones or Device Profiles that doesn't include the DN
        cntr = 0
        result_list = [
            [
                "Device Name",
                "DN",
                "Alerting Name",
                "Display Name",
                "Line Text Label",
                "New Line Label",
                "pkid",
            ]
        ]
        self.list_box.insert(
            tk.END,
            "Device Name, DN, Alerting Name, Display Name, Line Text Label, "
            "New Line Label, pkid\n",
        )
        sql_statement = (
            "SELECT d.name, n.dnorpattern, n.alertingname, dnmap.display, dnmap.label, dnmap.pkid "
            "FROM device d INNER JOIN devicenumplanmap dnmap ON dnmap.fkdevice=d.pkid INNER JOIN numplan n "
            "ON dnmap.fknumplan=n.pkid WHERE (d.tkclass=1 OR d.tkclass=254) ORDER BY d.name"
        )
        try:
            for row in self.sql_query(service=axl, sql_statement=sql_statement):
                try:
                    # Handle None results
                    dnmap_pkid = row["pkid"] if row["pkid"] else ""
                    dnmap_label = row["label"] if row["label"] else ""
                    dnmap_display = row["display"] if row["display"] else ""
                    n_alertingname = row["alertingname"] if row["alertingname"] else ""
                    n_dnorpattern = row["dnorpattern"] if row["dnorpattern"] else ""
                    d_name = row["name"] if row["name"] else ""
                except TypeError:
                    continue
                # First choice to generate Initial & Last Name is display name, then alerting name
                if n_dnorpattern not in dnmap_label:
                    new_label = ""
                    name_words = ""
                    if dnmap_display:
                        name_words = dnmap_display.split()
                    elif n_alertingname:
                        name_words = n_alertingname.split()
                    if len(name_words) > 1:
                        new_label = (
                            f"{name_words[0][0]} {name_words[-1]}-{n_dnorpattern}"
                        )
                    elif len(name_words) == 1:
                        new_label = f"{name_words[0]}-{n_dnorpattern}"
                    new_label = new_label.replace("'", "")
                    self.list_box.insert(
                        tk.END,
                        f"{d_name}, {n_dnorpattern}, {n_alertingname}, "
                        f"{dnmap_display}, {dnmap_label}, {new_label}, {dnmap_pkid}",
                    )
                    result_list.append(
                        [
                            d_name,
                            n_dnorpattern,
                            n_alertingname,
                            dnmap_display,
                            dnmap_label,
                            new_label,
                            dnmap_pkid,
                        ]
                    )
                    cntr += 1
        except TypeError:
            pass
        except Fault as thin_axl_error:
            tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
            return

        self.results_count_text.set(f"Results Found: {str(cntr)}")
        # Output to CSV file if required
        try:
            if len(output_filename) != 0:
                with open(
                    output_filename, "w", newline="", encoding="utf-8-sig"
                ) as csv_file:
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
                        if not axl_json["fqdn"]:
                            tk.messagebox.showerror(
                                title="Error", message="FQDN must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="FQDN must be specified."
                        )
                        return
                    try:
                        if not axl_json["username"]:
                            tk.messagebox.showerror(
                                title="Error", message="Username must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="Username must be specified."
                        )
                        return
                    try:
                        if not axl_json["wsdl_file"]:
                            tk.messagebox.showerror(
                                title="Error", message="WSDL file must be specified."
                            )
                            return
                    except KeyError:
                        tk.messagebox.showerror(
                            title="Error", message="WSDL file must be specified."
                        )
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
        session.auth = HTTPBasicAuth(axl_json["username"], self.axl_password)
        transport = Transport(cache=SqliteCache(), session=session, timeout=60)
        history = HistoryPlugin()
        try:
            client = Client(
                wsdl=axl_json["wsdl_file"], transport=transport, plugins=[history]
            )
        except FileNotFoundError as e:
            tk.messagebox.showerror(title="Error", message=str(e))
            return
        axl = client.create_service(axl_binding_name, axl_address)

        # Update Line Text Labels contained in CSV file
        cntr = 0
        result_list = [
            [
                "Device Name",
                "DN",
                "Alerting Name",
                "Display Name",
                "Line Text Label",
                "New Line Label",
                "pkid",
            ]
        ]
        self.list_box.insert(
            tk.END,
            "Device Name, DN, Alerting Name, Display Name, Line Text Label, "
            "New Line Label, pkid\n",
        )

        # Parse input CSV file & make updates based on the content
        try:
            with open(self.csv_input_filename, encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                header_row = next(reader)
                if header_row[5] != "New Line Label" or header_row[6] != "pkid":
                    tk.messagebox.showerror(
                        title="Error", message="Unable to parse CSV file."
                    )
                    return
                for row in reader:
                    try:
                        row[5] = row[5].replace("'", "")
                        sql_statement = f"UPDATE devicenumplanmap SET label='{row[5]}' WHERE pkid='{row[6]}'"
                        num_results = self.sql_update(
                            service=axl, sql_statement=sql_statement
                        )
                        # List updates that failed
                        if num_results < 1:
                            self.list_box.insert(
                                tk.END,
                                f"{row[0]}, {row[1]}, {row[2]}, {row[3]}, {row[4]}, "
                                f"{row[5]}, {row[6]}",
                            )
                            result_list.append(row)
                        else:
                            cntr += 1
                    except TypeError:
                        continue
                    except Fault as thin_axl_error:
                        tk.messagebox.showerror(
                            title="Error", message=thin_axl_error.message
                        )
                        break
        except KeyError:
            tk.messagebox.showerror(title="Error", message="Unable to parse CSV file.")
            pass
        except FileNotFoundError:
            tk.messagebox.showerror(title="Error", message="Unable to open CSV file.")
            return

        self.results_count_text.set(f"Updates Made: {str(cntr)} (failures below)")
        # Output to CSV file if required
        try:
            if len(output_filename) != 0:
                with open(
                    output_filename, "w", newline="", encoding="utf-8-sig"
                ) as csv_file:
                    writer = csv.writer(csv_file)
                    writer.writerows(result_list)
        except OSError:
            tk.messagebox.showerror(title="Error", message="Unable to write CSV file.")

    def check_labels(self):
        """Validate parameters and then call AXL query"""
        if not self.axl_input_filename:
            tk.messagebox.showerror(title="Error", message="No AXL file selected.")
            return

        output_string = self.output_csv_text.get()
        if len(output_string) == 0:
            self.read_axl("")
        else:
            self.read_axl(output_string)

    def update_labels(self):
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
            self.write_axl("")
        else:
            self.write_axl(output_string)

    def open_json_file_dialog(self):
        """Dialogue to prompt for JSON file to open and AXL password"""
        self.axl_input_filename = tk.filedialog.askopenfilename(
            initialdir="/", filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
        )
        self.axl_password = tk.simpledialog.askstring(
            "Input", "AXL Password?", show="*"
        )

    def open_csv_file_dialog(self):
        """Dialogue to prompt for CSV file to open"""
        self.csv_input_filename = tk.filedialog.askopenfilename(
            initialdir="/", filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )


if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    # Initialise TKinter GUI objects
    root = tk.Tk()
    root.title("Line Text Label Checker v1.3")
    GUIFrame(root)
    root.mainloop()
