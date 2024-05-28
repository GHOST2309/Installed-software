# -*- coding: utf8 -*-
import re
import subprocess
import json
import sys

SENDER_PATH = r"C:\Program Files\Zabbix Agent\zabbix_sender.exe"
AGENT_CONF = r"C:\Program Files\Zabbix Agent\zabbix_agentd.conf"

def send_data(SENDER_PATH_, AGENT_CONF_, senderData):
    p = subprocess.Popen([SENDER_PATH_, '-c', AGENT_CONF_, '-i', '-'],
                                  stdin=subprocess.PIPE, universal_newlines=True)

    senderDataNStr = '\n'.join(senderData)



    if sys.argv[1] == 'get':
        senderProc = subprocess.Popen([SENDER_PATH_, '-c', AGENT_CONF_, '-i', '-'],
                                      stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, universal_newlines=True)

    elif sys.argv[1] == 'getverb':
        print(senderDataNStr)
        senderProc = subprocess.Popen([SENDER_PATH_, '-vv', '-c', AGENT_CONF_, '-i', '-'],
                                      stdin=subprocess.PIPE, universal_newlines=True)

    else:
        print(sys.argv[0] + " : Not supported. Use 'get' or 'getverb'.")
        sys.exit(1)


    p.communicate(input=senderDataNStr)

    return p




def get_output():
    p = subprocess.check_output(["powershell.exe", r"Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallDate, InstallLocation | Format-List"], universal_newlines=True)

    return p

def find_values(p_out_):
    l = p_out_.split("\n\n")

    all_subjects = []
    for i in l:
        dictionary = {}

        name = re.search("DisplayName\s+:([^\n]+)", i)
        if name:
            name_s = name.group(1).strip()
            if name_s:
                dictionary["DisplayName"] = name_s

        version = re.search("DisplayVersion\s+:([^\n]+)", i)
        if version:
            version_s = version.group(1).strip()
            if version_s:
                dictionary["DisplayVersion"] = version_s

        install_date = re.search("InstallDate\s+:([^\n]+)", i)
        if install_date:
            install_s = install_date.group(1).strip()
            if install_s:
                dictionary["InstallDate"] = install_s

        install_location = re.search("InstallLocation\s+:([^\n]+)", i)
        if install_location:
            location_s = install_location.group(1).strip()
            if location_s:
                dictionary["InstallLocation"] = location_s

        if len(dictionary.keys()) == 1:
            display_name_val = dictionary["DisplayName"]
            if display_name_val.startswith("Update for Microsoft") or \
                display_name_val.startswith("Security Update for Microsoft") or \
                display_name_val.startswith("Service Pack 2 for Microsoft"):
                continue

        if dictionary:
            all_subjects.append(dictionary)

    return all_subjects

def form_data(all_lines):
    host = sys.argv[1]

    stop_chars = [('!', '_'), (',', '_'), ('[', '_'), ('~', '_'), ('  ', '_'),
        (']', '_'), ('+', '_'), ('/', '_'), ('\\', '_'), ('\'', '_'),
        ('`', '_'), ('@', '_'), ('#', '_'), ('$', '_'), ('%', '_'),
        ('^', '_'), ('&', '_'), ('*', '_'), ('(', '_'), (')', '_'),
        ('{', '_'), ('}', '_'), ('=', '_'), (':', '_'), (';', '_'),
        ('"', '_'), ('?', '_'), ('<', '_'), ('>', '_'), (' ', '_'),]

    macros_json = []
    sender_data = []

    for dct in all_lines:
        if "DisplayName" in dct:
            software_id_key = dct["DisplayName"].replace(" ", "_")
            software_id = dct["DisplayName"]
            macros_json.append({'{#SOFTWAREID}}': software_id_key})
            sender_data.append(f"'{host}' 'software.name[{software_id_key}]' '{software_id}' ")

            if "DisplayVersion" in dct:
                version = dct["DisplayVersion"]
                sender_data.append(f"'{host}' 'software.version[{software_id_key}]' '{version}' ")

            if "InstallDate" in dct:
                date = dct["InstallDate"]
                sender_data.append(f"'{host}' 'software.date[{software_id_key}]' '{date}' ")

            if "InstallLocation" in dct:
                location = dct["InstallLocation"]
                sender_data.append(f"'{host}' 'software.location[{software_id_key}]' '{location}' ")


    # for i in sender_data:
    #    print(i)

    json_dumps = json.dumps(macros_json, indent=4)

    # d = "\u0431\u0432\u00a0\u00ad\u00a4\u00a0\u0430\u0432\u00ad\u043b\u00a9"
    # d.decode('cp1251').encode('utf8')
    # print(d)

    return json_dumps, sender_data

if __name__ == '__main__':
    p_out = get_output()
    all_lines = find_values(p_out)
    output = form_data(all_lines)
    json = output[0]
    sender_items = output[1]
    send_data(SENDER_PATH, AGENT_CONF, sender_items)

    #print(json)