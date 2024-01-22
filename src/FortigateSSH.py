#!/usr/bin/env python3

from netmiko import ConnectHandler, redispatch
import time
import re
# from netmiko.fortinet import FortinetSSH


class FortigateSSH():
    def __init__(self, ip: str = "", port: int = 22, user: str = "admin", password: str = "", bastion_ip: str = "", bastion_user: str = "admin", bastion_password: str = "", bastion_port: int = 22) -> None:
        self.vdom = []
        # ENDPOINT_FIREWALL_VDOM = ["/system/vdom", "vdom"]
        # ENDPOINT_FIREWALL_INTERFACE = ["/system/interface", "interface"]
        # ENDPOINT_FIREWALL_SERVICE = ["/firewall.service/custom", "service"]
        # ENDPOINT_FIREWALL_SERVICE_GROUP = ["/firewall.service/group", "service_group"]
        # ENDPOINT_FIREWALL_ADDRESS = ["/firewall/address", "address"]
        # ENDPOINT_FIREWALL_ADDRESS_GROUP = ["/firewall/addrgrp", "address_group"]
        # ENDPOINT_FIREWALL_VIP = ["/firewall/vip", "vip"]
        # ENDPOINT_FIREWALL_VIP_GROUP = ["/firewall/vipgrp", "vip_group"]
        # ENDPOINT_FIREWALL_IPPOOL = ["/firewall/ippool", "ippool"]
        # ENDPOINT_FIREWALL_POLICY = ["/firewall/policy", "policy"]
        bastion = {
            'device_type': 'terminal_server',
            'ip': bastion_ip,
            'username': bastion_user,
            'password': bastion_password,
            'port': bastion_port,
            'session_log': './output.log'
        }

        fortigate = {
            'device_type': 'fortinet',
            'host': ip,
            'username': user,
            'password': password,
            'port': port,
        }
        if bastion_ip:
            self.net_connect = ConnectHandler(**bastion)
            time.sleep(2)
            self.net_connect.write_channel(f"ssh {ip} -p {port}\r\n")
            time.sleep(2)
            self.net_connect.write_channel(password)
            redispatch(self.net_connect, device_type="fortinet")
        else:
            self.net_connect = ConnectHandler(**fortigate)

        if self.vdoms_enabled():
            # vdoms = self.net_connect.send_command('config global', cmd_verify=False)
            # vdoms = self.net_connect.send_command('show system vdom-property | grep edit', cmd_verify=False)
            # vdoms = vdoms.split('edit "')
            # print(vdoms)
            self.vdom = self.get_vdoms()
        self.net_connect.cleanup()
        self.get_policies()
        self.net_connect.disconnect()

    def vdoms_enabled(self) -> bool:
        self.net_connect.send_command('config global', expect_string=r"[#$]")
        output = self.net_connect.send_command("get system status | grep Virtual", expect_string=r"[#$]")
        return bool(re.search(r"Virtual domain configuration: (multiple|enable)", output))

    def get_vdoms(self):
        output = self.net_connect.send_command("get system vdom-property | grep name", expect_string=r"[#$]")
        return re.findall(r"name: (?P<vdom_name>\S+)", output)

    def get_policies(self):
        print(self.vdom)
        if self.vdom:
            self.net_connect.send_command("config vdom", expect_string=r"[#$]")
            for vdom in self.vdom:
                print(vdom)
                self.net_connect.send_command(f"edit {vdom}", expect_string=r"[#$]")
                output = self.net_connect.send_command("show firewall policy", expect_string=r"[#$]")
                match_params = re.finditer(r"edit\D(?P<id>\d*)[\s\S]*?next\n", output, flags=re.M)
                for params in match_params:
                    # print(params.group('id'))
                    matches = re.findall(r"set\D(?P<param>\S*)\s(?P<desc>.*)\n", params.group(), flags=re.M)
                    # il faut ajouter l'ID dans le tuple matches
                    # id = ('id', params.group('id'))
                    print(type(matches))  # .append(tuple(('id', params.group('id'))))
                    # print()
                self.net_connect.send_command("next", expect_string=r"[#$]")
            self.net_connect.send_command("end", expect_string=r"[#$]")
        else:
            output = self.net_connect.send_command("show firewall policy", expect_string=r"[#$]")
            match_params = re.finditer(r"edit\D(?P<id>\d*)[\s\S]*?next\n", output, flags=re.M)
            for params in match_params:
                matches = re.findall(r"set\D(?P<param>\S*)\s(?P<desc>.*)\n", params.group(), flags=re.M)
