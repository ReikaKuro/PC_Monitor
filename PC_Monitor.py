import json
import re
import time
from winrm.protocol import Protocol
from ping3 import ping
import threading
import os
import logging


class SetupInfo:
    def __init__(self, setup_ip, login, password, pdu_list, computer_to_username):
        """
        Parameters below stores...
        :param self.setup_ip: IP of PC
        :param self.login: login of PC
        :param self.password: password of PC
        :param self.pdu_list: list of PDUs with their IPs and ports that
         corresponds to the RM that are connected [[PDU_IP,PDU_Port,self.pdu_controller,[],[]]
        :param self.computertousername: list of mapped computer names to person
        :param self.pdu_amount: how many RM are connected to unique PDUs and ports
        :param self.pdu_ip: IP of PDU (used together with self.pdu_port)
        :param self.pdu_port: Port of PDU (used together with self.pdu_ip)
        :param self.pdu_state: information if specific port on PDU is 'on' or 'off'
        :param self.online: True = PC is pingable, False = PC is not pingable
        :param self.pdu_controller: object for each pdu which let us control it
        :param self.is_logged_in: Active = someone is logged on PC, Disc = no one is logged on PC
        :param self.user: PC name of user that is logged on PC
        :param self.idle: time since last action on PC
        """

        self.setup_ip = setup_ip
        self.login = login
        self.password = password
        self.pdu_list = pdu_list
        self.computer_to_username_dict = computer_to_username
        self.pdu_amount = len(pdu_list)
        self.pdu_state = None
        self.online = None
        self.pdu_controller = pdu.interface.pdu()
        for index, pdu in enumerate(pdu_list):
            pdu_ip = pdu[0]
            self.pdu_list[index].append(self.pdu_controller.setup_pdu(model='GUDE', version='8220-1', host=pdu_ip, timeout=15))
        self.is_logged_in = None
        self.user = None
        self.idle = None
        self.shell_id_evaluation = True  # Make description above

    def ping_pc(self):
        """
        Simple ping to PC to check if there is connection

        Uses variables from __init__:
        self.setup_ip

        Fills variables from __init__:
        self.online
        :return:
        """
        self.online = True if type(ping(self.setup_ip)) is float else False

    def get_setup_user_idle_info(self):
        """
        Gets info about setup.
        Uses variables from __init__:
        self.setup_ip
        self.login
        self.password

        Fill variables from __init__:
        self.is_logged_in
        self.idle
        self.user

        :return:
        """
        p = Protocol(
            endpoint='https://' + self.setup_ip + ':5986/wsman',
            transport='ntlm',
            username=self.login,
            password=self.password,
            read_timeout_sec=4,
            operation_timeout_sec=2,
            server_cert_validation='ignore')

        self.shell_id_evaluation = True
        try:
            print(f'Opening shell {self.setup_ip}')
            shell_id = p.open_shell()

        except Exception:
            logging.warning(f'shell_id evaluation failed')
            self.shell_id_evaluation = False
            self.online = False
            self.idle = 'Offline'
            self.user = 'Offline'

        if self.shell_id_evaluation:
            command_id = p.run_command(shell_id, 'query user', [])
            std_out, std_err, status_code = p.get_command_output(shell_id, command_id)
            std_out = std_out.decode('UTF-8')
            logging.debug(f'Output from query user command: {std_out}')
            x = re.search(r'(\d+)\s+(Active|Disc)\s+([\d.:+]+)', std_out)
            try:
                id_user = x.group(1)
                self.is_logged_in = x.group(2)
                self.idle = x.group(3)

                command_id = p.run_command(shell_id, 'REG QUERY "HKCU\\Volatile Environment\\' + id_user + '" /v "CLIENTNAME"', [])
                std_out, std_err, status_code = p.get_command_output(shell_id, command_id)
                std_out = std_out.decode('UTF-8')
                x = re.search(r'CLIENTNAME\s+[a-zA-Z_]+\s+([a-zA-Z-\d]+)', std_out)

                self.user = x.group(1)
                if self.user in self.computer_to_username_dict and self.is_logged_in != 'Disc':
                    self.computer_to_username()
                else:
                    self.user = x.group(1) if self.is_logged_in == 'Active' else 'Free'

                if self.idle == '.':
                    self.idle = '0' if self.user == 'Free' else 'Active'
                elif not self.idle.__contains__('+' and ':'):
                    pass
                else:
                    self.idle = '60+'

                p.cleanup_command(shell_id, command_id)
                p.close_shell(shell_id)
            except Exception:
                logging.warning(f'PROBLEM with get_setup_user_idle_info')
                self.online = False
                self.idle = 'Offline'
                self.user = 'Offline'
        else:
            pass

    def computer_to_username(self):
        try:
            self.user = self.computer_to_username_dict[self.user]
        finally:
            pass

    def pdu_check(self):
        """
        Check if port on PDU is ON or OFF
        Uses variables from __init__:
        self.pdu_list

        Fills variables from __init__:
        self.pdu_state

        :param state[]: list of results 'on' or 'off' that contains all PDU mapped to PC
        :return: state
        """
        state = []
        for index, pdu in enumerate(self.pdu_list):
            pdu_port = pdu[1]
            state.append(self.pdu_list[index][2].get_port_status(port=int(pdu_port)))
        self.pdu_state = state
        logging.debug(f'PDU: {self.pdu_state}')
        return state

    def pdu_switch_off(self):
        """
        Turn off port on PDU
        Uses variables from __init__:
        self.pdu_list

        :return: nothing
        """
        for index, pdu in enumerate(self.pdu_list):
            pdu_port = pdu[1]
            self.pdu_list[index][2].power_off(int(pdu_port))
        logging.debug(f'RM turned off because of idle time: {self.idle}')

    def evaluate_usage(self):
        if 'on' not in self.pdu_check():
            pass
        else:
            # if setup blocked then do nothing
            if self.is_logged_in == 'Disc':
                if int(self.idle.replace('+', '')) <= 10:
                    pass
                else:
                    self.pdu_switch_off()

            elif self.is_logged_in == 'Active':
                if self.idle == "Active" or int(self.idle.replace('+', '')) <= 50:
                    pass
                else:
                    self.pdu_switch_off()
            pass

    def monitor_setup(self):
        self.ping_pc()
        if not self.online:
            logging.debug(f'No connection with {self.setup_ip}')
        else:
            self.get_setup_user_idle_info()
            if self.shell_id_evaluation:
                logging.debug('\t\t\t\tSETUP INFO')
                logging.debug(f'IP: {self.setup_ip}\tUser: {self.user}\tAFK: {self.idle}')
                logging.debug(f'PDU: {self.pdu_list}')
                self.evaluate_usage()
            else:
                logging.debug(f'\t\t\t\tSETUP INFO')
                logging.debug(f'IP: {self.setup_ip}\tUser: {self.user}\tAFK: {self.idle}')
                logging.debug(f'PDU: {self.pdu_list}')


class Html:
    def __init__(self):
        self.htmlText = '''<html>
<head>
<style>
body{
background-color:black;
text-align:center;
}
table{
width:100%;
}
h1{
color:orange;
}
tr{
height:75px;
}
th{
border:1px solid gray;
color:white;
font-size:24;
}
</style>
<title>Setup_Monitor_WRO7</title>
<meta http-equiv="refresh" content="5">
</head>
<body>
<h1>Setup_Monitor_WRO7</h1>
<table>
<tr style="height:75px; background-color:darkblue;">
<th style="font-size:38px;">Setup</th>
<th style="font-size:38px;">User</th>
<th style="width:15%; font-size:38px;">Idle Time</th>
</tr>'''
        self.endOfHtml = '''</table>
</body>
</html>'''

    def append_setup(self, setup_ip, user, idle):
        setup = f'''<tr>
<th>{setup_ip}</th>
<th>{user}</th>
<th>{idle}</th>
</tr>'''
        self.htmlText += f'\n{setup}'

    def create_html(self):
        self.htmlText += f'\n{self.endOfHtml}'
        with open('setupmonitor.html', 'w', encoding='utf-8') as f:
            f.write(self.htmlText)
        self.htmlText = '''<html>
        <head>
        <style>
        body{
        background-color:black;
        text-align:center;
        }
        table{
        width:100%;
        }
        h1{
        color:orange;
        }
        tr{
        height:75px;
        }
        th{
        border:1px solid gray;
        color:white;
        font-size:24;
        }
        </style>
        <title>Setup_Monitor_WRO7</title>
        </head>
        <body>
        <h1>Setup_Monitor_WRO7</h1>
        <meta http-equiv="refresh" content="5">
        <table>
        <tr style="height:75px; background-color:darkblue;">
        <th style="font-size:38px;">Setup</th>
        <th style="font-size:38px;">User</th>
        <th style="width:15%; font-size:38px;">Idle Time</th>
        </tr>'''


def load_json_file():
    """
    Opens json file with mapped setups
    :return: mapped setups in form of dictionary
    """
    with open('Setups.json', 'r') as f:
        setups_json = json.load(f)

    with open('ComputerToUsername.json', 'r', encoding='utf-8') as f:
        computer_to_username_json = json.load(f)

    return setups_json, computer_to_username_json


def parse_json_file(setups_file, computer_to_username):
    """
    Takes argument in form json file parsed into dictionary
    :param setups_file: contains mapped setups in form of dictionary
    :param computer_to_username: contains mapped pc name to person in form of dictionary

    :return: Mapped setups in form of list which contains elements needed to create object SetupInfo()
    """
    setups_json = setups_file
    parsed_json = []
    for setup_keyword in setups_json['setups']:
        setup_ip = setups_json['setups'][setup_keyword]['setup_ip']
        login = setups_json['setups'][setup_keyword]['login']
        password = setups_json['setups'][setup_keyword]['password']
        pdu_list = []
        pdu_count = 1
        for pdu in setups_json['setups'][setup_keyword]:
            if re.match(r'pdu\d+', pdu):
                pdu = setups_json['setups'][setup_keyword][f'pdu{pdu_count}'].split(',')
                pdu[1] = pdu[1].replace(' ', '')
                pdu_list.append(pdu)
                pdu_count += 1
            else:
                pass
        parsed_json.append(SetupInfo(setup_ip, login, password, pdu_list, computer_to_username))

    return parsed_json


if __name__ == "__main__":
    print('Setup Monitor is running... better catch it before it run away :)')
    logging.basicConfig(format='%(levelname)s %(asctime)s: %(message)s',
                        datefmt='%Y-%m-%d-%H:%M:%S',
                        level=logging.DEBUG,
                        handlers=[
                            logging.FileHandler('Setup_Monitor_LOGS.log', encoding='utf8'),
                            logging.StreamHandler()
                        ])

    htmlFile = Html()
    os.startfile('run_webpage.bat')
    while True:
        start = time.perf_counter()
        threads = []
        setups, computerToUsername = load_json_file()
        setups_list = parse_json_file(setups, computerToUsername)

        for i in range(len(setups_list)):
            t = threading.Thread(target=setups_list[i].monitor_setup())
            t.start()
            threads.append(t)

        for thread in threads:
            thread.join()
        finish = time.perf_counter()

        for i in range(len(setups_list)):
            htmlFile.append_setup(setups_list[i].setup_ip, setups_list[i].user, setups_list[i].idle)

        htmlFile.create_html()
        print(f'Finished in {round(finish - start, 2)} seconds(s)')
        time.sleep(5)
