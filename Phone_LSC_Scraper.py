#!/usr/bin/env python3

"""
Copyright (c) 2022 - 2023, Chris Perkins
Licence: BSD 3-Clause

Dynamic auditing of certificates installed on phones. Running against the publisher finds all the phones in a cluster.
First pulls list of SEP devices from AXL API, then uses this list to retrieve IP addresses of registered phones via the RIS API.
Then connects via HTTPS to each IP address & outputs the certificate's issuer, subject & the expiry date.
Application user requires Standard AXL API Access, Standard RealtimeAndTraceCollection & Standard Serviceability roles.

v1.3 - implemented proper rate limited of API requests
v1.2 - switched to displaying the full certificate issuer & subject to provide more information
v1.1 - added fallback from TLS v1.2 to v1.0 for older phones
v1.0 - original release

Portions of this code from cucm-compare-reg-status, (c) Paul Tursan 2018, https://github.com/ptursan/cucm-compare-reg-status/ & used under the MIT license
Portions of this code from https://stackoverflow.com/questions/16903528/how-to-get-response-ssl-certificate-from-requests-in-python

I have no idea how the OpenSSL stuff works, it's magic ;)
"""

import sys
import requests
import urllib3
import socket
import json
import OpenSSL
import time
from zeep import Client
from zeep.cache import SqliteCache
from zeep.transports import Transport
from zeep.exceptions import Fault
from zeep.plugins import HistoryPlugin
from requests import Session
from requests.auth import HTTPBasicAuth
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
from lxml import etree
from OpenSSL.SSL import Connection, Context, SSLv23_METHOD, TLSv1_METHOD, TLSv1_2_METHOD
from datetime import datetime, time
from time import sleep
from OpenSSL.crypto import X509
from getpass import getpass

TLS_METHODS = (TLSv1_2_METHOD, TLSv1_METHOD, SSLv23_METHOD)
MAX_API_CALLS_A_MINUTE = 15


def show_history(history):
    """Output error messages from Zeep"""
    for hist in [history.last_sent, history.last_received]:
        print(etree.tostring(hist["envelope"], encoding="unicode", pretty_print=True))


def main():
    """Program entry point, reads config"""
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <AXL config JSON>")
        sys.exit(1)

    # Load JSON configuration parameters
    try:
        with open(sys.argv[1]) as f:
            axl_json_data = json.load(f)
            for axl_json in axl_json_data:
                try:
                    if not axl_json["fqdn"]:
                        print("Config Error: CUCM FQDN must be specified.")
                        sys.exit(1)
                except KeyError:
                    print("Config Error: CUCM FQDN must be specified.")
                    sys.exit(1)
                try:
                    if not axl_json["username"]:
                        print("Config Error: AXL username must be specified.")
                        sys.exit(1)
                except KeyError:
                    print("Config Error: AXL username must be specified.")
                    sys.exit(1)
                try:
                    if not axl_json["wsdl_file"]:
                        print("Config Error: WSDL file must be specified.")
                        sys.exit(1)
                except KeyError:
                    print("Config Error: WSDL file must be specified.")
                    sys.exit(1)
    except FileNotFoundError:
        print(f"Error: Unable to open JSON config file {sys.argv[1]}.")
        sys.exit(1)
    except json.decoder.JSONDecodeError:
        print(f"Error: Unable to parse JSON config file {sys.argv[1]}.")
        sys.exit(1)

    username = axl_json["username"]
    password = getpass("Password: ")
    server = axl_json["fqdn"]
    axl_wsdl = axl_json["wsdl_file"]

    # Common Plugins
    history = HistoryPlugin()

    # Build Client object for AXL Service
    axl_location = f"https://{server}:8443/axl/"
    axl_binding = "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding"

    axl_session = Session()
    axl_session.verify = False
    axl_session.auth = HTTPBasicAuth(username, password)

    axl_transport = Transport(cache=SqliteCache(), session=axl_session, timeout=20)
    axl_client = Client(wsdl=axl_wsdl, transport=axl_transport, plugins=[history])
    axl_service = axl_client.create_service(axl_binding, axl_location)

    # Build Client object for RisPort70 Service
    wsdl = f"https://{server}:8443/realtimeservice2/services/RISService70?wsdl"

    session = Session()
    session.verify = False
    session.auth = HTTPBasicAuth(username, password)

    transport = Transport(cache=SqliteCache(), session=session, timeout=20)
    client = Client(wsdl=wsdl, transport=transport, plugins=[history])
    service = client.create_service(
        "{http://schemas.cisco.com/ast/soap}RisBinding",
        f"https://{server}:8443/realtimeservice2/services/RISService70",
    )

    # Get list of Phones to query via AXL, required when using SelectCmDeviceExt
    try:
        resp = axl_service.listPhone(
            searchCriteria={"name": "SEP%"}, returnedTags={"name": ""}
        )
    except Fault:
        show_history(history)
        raise

    # Build item list for RisPort70 SelectCmDeviceExt
    items = []
    for phone in resp["return"].phone:
        items.append(phone.name)
    print(f"{len(items)} SEP devices found in configuration.\n")
    # Run SelectCmDeviceExt on each Phone
    cntr_success = 0
    cntr_fail = 0
    cntr_iterations = 0
    timer = 0.0
    for phone in items:
        last_time = time.perf_counter()
        cntr_iterations += 1
        CmSelectionCriteria = {
            "MaxReturnedDevices": "1",
            "DeviceClass": "Phone",
            "Model": "255",
            "Status": "Registered",
            "NodeName": "",
            "SelectBy": "Name",
            "SelectItems": {"item": phone},
            "Protocol": "Any",
            "DownloadStatus": "Any",
        }

        StateInfo = ""
        try:
            resp = service.selectCmDeviceExt(
                CmSelectionCriteria=CmSelectionCriteria, StateInfo=StateInfo
            )
        except Fault:
            show_history(history)
            raise

        CmNodes = resp.SelectCmDeviceResult.CmNodes.item
        for CmNode in CmNodes:
            if len(CmNode.CmDevices.item) > 0:
                # If the node has returned CmDevices
                for item in CmNode.CmDevices.item:
                    # Older phones don't support TLS 1.2
                    for method in TLS_METHODS:
                        try:
                            try:
                                ssl_connection_setting = Context(method)
                            except ValueError:
                                continue
                            ssl_connection_setting.set_timeout(1)
                            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                                s.connect((item["IPAddress"]["item"][0]["IP"], 443))
                                c = Connection(ssl_connection_setting, s)
                                c.set_tlsext_host_name(
                                    str.encode(item["IPAddress"]["item"][0]["IP"])
                                )
                                c.set_connect_state()
                                c.do_handshake()
                                cert = c.get_peer_certificate()
                                # Convert issuer & subject into dictionaries, parse expiry date + time
                                issuer_list = cert.get_issuer().get_components()
                                cert_issuer = {}
                                for thing in issuer_list:
                                    cert_issuer.update(
                                        {
                                            thing[0]
                                            .decode("utf-8"): thing[1]
                                            .decode("utf-8")
                                        }
                                    )
                                subject_list = cert.get_subject().get_components()
                                cert_subject = {}
                                for thing in subject_list:
                                    cert_subject.update(
                                        {
                                            thing[0]
                                            .decode("utf-8"): thing[1]
                                            .decode("utf-8")
                                        }
                                    )
                                end_date = datetime.strptime(
                                    str(cert.get_notAfter().decode("utf-8")),
                                    "%Y%m%d%H%M%SZ",
                                )
                                diff = end_date - datetime.now()
                                # if cert.has_expired() or diff.days <= 7:
                                #    print(f"FIX ME! {item['Name']}, {item['IPAddress']['item'][0]['IP']}, issuer {cert_issuer}, subject {cert_subject}, expires {str(end_date)}.")
                                print(
                                    f"{item['Name']}, {item['IPAddress']['item'][0]['IP']}, issuer {cert_issuer}, subject {cert_subject}, expires {str(end_date)}."
                                )
                                c.shutdown()
                                s.close()
                                cntr_success += 1
                                break
                        except (
                            TimeoutError,
                            ConnectionRefusedError,
                            socket.timeout,
                            urllib3.exceptions.ConnectTimeoutError,
                            urllib3.exceptions.MaxRetryError,
                            requests.exceptions.ConnectTimeout,
                        ):
                            print(
                                f"{item['Name']}, {item['IPAddress']['item'][0]['IP']}, unable to connect."
                            )
                            cntr_fail += 1
                            break
                        except OpenSSL.SSL.Error:
                            continue
                    else:
                        print(
                            f"{item['Name']}, {item['IPAddress']['item'][0]['IP']}, unable to connect."
                        )
                        cntr_fail += 1
        timer += time.perf_counter() - last_time
        if cntr_iterations >= MAX_API_CALLS_A_MINUTE and timer < 60:
            wait_time = 60.0 - timer
            # print(f"{cntr_iterations} iterations in {timer}s, waiting {wait_time}s")
            time.sleep(wait_time)
            cntr_iterations = 0
            timer = 0.0

    # Summarise
    print(
        f"\nOut of {cntr_success + cntr_fail} registered devices - {cntr_success} certificate confirmed, {cntr_fail} unable to connect via HTTPS."
    )


if __name__ == "__main__":
    disable_warnings(InsecureRequestWarning)
    main()
