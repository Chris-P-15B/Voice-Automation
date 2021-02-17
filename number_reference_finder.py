#!/usr/bin/env python3

"""
(c) 2018 - 2019, Chris Perkins
Licence: BSD 3-Clause

Checks NumPlan for CFA, CFB, CFNA, CFNC, CFUR, AAR Destination Mask or Called Party Transformation
that reference a given number, SQL wildcard % can be used

v1.2 - code tidying
v1.1 - fixes some edge cases
v1.0 - original release

Original AXL SQL query code courtesy of Jonathan Els - https://afterthenumber.com/2018/04/27/serializing-thin-axl-sql-query-responses-with-python-zeep/

To Do:
Improve the GUI
"""

import sys, json
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

    # tkPatternUsage Mappings
    pattern_usage = {
        "0": "Call Park",
        "1": "Conference",
        "2": "Directory Number",
        "3": "Translation Pattern",
        "4": "Call Pick Up Group",
        "5": "Route Pattern",
        "6": "Message Waiting",
        "7": "Hunt Pilot",
        "8": "Voice Mail Port",
        "9": "Domain Routing",
        "10": "IP Address Routing",
        "11": "Device Template",
        "12": "Directed Call Park",
        "13": "Device Intercom",
        "14": "Translation Intercom",
        "15": "Translation Calling Party Number",
        "16": "Mobility Handoff",
        "17": "Mobility Enterprise Feature Access",
        "18": "Mobility IVR",
        "19": "Device Intercom Template",
        "20": "Called Party Number Transformation",
        "21": "Call Control Discovery Learned Pattern",
        "22": "URI Routing",
        "23": "ILS Learned Enterprise Number",
        "24": "ILS Learned E164 Number",
        "25": "ILS Learned Enterprise Numeric Pattern",
        "26": "ILS Learned E164 Numeric Pattern",
        "27": "Alternate Number",
        "28": "ILS Learned URI",
        "29": "ILS Learned PSTN Failover Rule",
        "30": "ILS Imported E164 Number",
        "104": "Centralized Conference Number",
        "105": "Emergency Location ID Number",
    }

    def __init__(self, parent):
        """Constructor checks parameters and initialise variables"""
        self.input_filename = None
        self.axl_password = ""
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
        tk.Label(self, text="Number Pattern to Find:").place(
            relx=0.2, rely=0.0, height=22, width=200
        )
        self.search_pattern_text = tk.StringVar()
        tk.Entry(self, textvariable=self.search_pattern_text).place(
            relx=0.2, rely=0.05, height=22, width=200
        )
        tk.Button(self, text="Find References", command=self.find_references).place(
            relx=0.35, rely=0.12, height=22, width=100
        )
        self.records_label_text = tk.StringVar()
        self.records_label_text.set("Dial Plan Records: ")
        tk.Label(self, textvariable=self.records_label_text).place(
            relx=0.35, rely=0.18, height=22, width=110
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

    def read_axl(self, search_string):
        """Read and parse NumPlan via AXL"""
        try:
            self.list_box.delete(0, tk.END)
            self.records_label_text.set("Dial Plan Records: ")
            with open(self.input_filename) as f:
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

        sql_statement = (
            f"SELECT n.DNOrPattern, n.Description, n.tkPatternUsage FROM NumPlan n LEFT JOIN "
            f"CallForwardDynamic cfd ON cfd.fkNumPlan=n.pkid WHERE n.CFAptDestination LIKE '"
            f"{search_string}' OR n.CFBDestination LIKE '"
            f"{search_string}' OR n.CFBIntDestination LIKE '"
            f"{search_string}' OR n.CFNADestination LIKE '"
            f"{search_string}' OR n.CFNAIntDestination LIKE '"
            f"{search_string}' OR n.PFFDestination LIKE '"
            f"{search_string}' OR n.PFFIntDestination LIKE '"
            f"{search_string}' OR n.CFURDestination LIKE '"
            f"{search_string}' OR n.CFURIntDestination LIKE '"
            f"{search_string}' OR n.AARDestinationMask LIKE '"
            f"{search_string}' OR n.CalledPartyTransformationMask LIKE '"
            f"{search_string}' OR cfd.CFADestination LIKE '"
            f"{search_string}' ORDER BY n.DNOrPattern"
        )
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

        # Update TKinter display objects with results
        cntr = 0
        try:
            for row in self.sql_query(service=axl, sql_statement=sql_statement):
                try:
                    # Handle None results
                    n_dnorpattern = row["dnorpattern"] if row["dnorpattern"] else ""
                    n_description = row["description"] if row["description"] else ""
                    n_tkpatternusage = (
                        row["tkpatternusage"] if row["tkpatternusage"] else "2"
                    )  # Assume DN if unknown
                    self.list_box.insert(
                        tk.END,
                        f'{n_dnorpattern} "{n_description}", '
                        f"{self.pattern_usage[n_tkpatternusage]}",
                    )
                    cntr += 1
                except TypeError:
                    continue
        except TypeError:
            pass
        except Fault as thin_axl_error:
            tk.messagebox.showerror(title="Error", message=thin_axl_error.message)
            return
        self.records_label_text.set(f"Dial Plan Records: {str(cntr)}")

    def find_references(self):
        """Validate parameters then call AXL query"""
        if not self.input_filename:
            tk.messagebox.showerror(title="Error", message="No AXL file selected.")
            return
        # Check for invalid characters
        search_string = self.search_pattern_text.get()
        if len(search_string) == 0:
            tk.messagebox.showerror(title="Error", message="Search pattern is blank.")
            return
        for range_char in search_string:
            if range_char not in [
                "0",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "*",
                "#",
                "X",
                "%",
            ]:
                tk.messagebox.showerror(
                    title="Error", message="Invalid characters in search pattern."
                )
                return

        self.read_axl(search_string)

    def open_json_file_dialog(self):
        """Dialogue to prompt for JSON file to open and AXL password"""
        self.input_filename = tk.filedialog.askopenfilename(
            initialdir="/", filetypes=(("JSON files", "*.json"), ("All files", "*.*"))
        )
        self.axl_password = tk.simpledialog.askstring(
            "Input", "AXL Password?", show="*"
        )


if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    # Initialise TKinter GUI objects
    root = tk.Tk()
    root.title("Number Reference Finder v1.2")
    GUIFrame(root)
    root.mainloop()