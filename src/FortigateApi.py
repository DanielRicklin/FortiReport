#!/usr/bin/env python3

import requests
from .FortigateToXlsxFromAPI import FortigateToXlsxFromAPI
requests.packages.urllib3.disable_warnings()


class FortiGateApi():
    """Get fortigate configuration by API"""

    ENDPOINT_FIREWALL_VDOM = ["/system/vdom", "vdom"]
    ENDPOINT_FIREWALL_INTERFACE = ["/system/interface", "interface"]
    ENDPOINT_FIREWALL_SERVICE = ["/firewall.service/custom", "service"]
    ENDPOINT_FIREWALL_SERVICE_GROUP = ["/firewall.service/group", "service_group"]
    ENDPOINT_FIREWALL_ADDRESS = ["/firewall/address", "address"]
    ENDPOINT_FIREWALL_ADDRESS_GROUP = ["/firewall/addrgrp", "address_group"]
    ENDPOINT_FIREWALL_VIP = ["/firewall/vip", "vip"]
    ENDPOINT_FIREWALL_VIP_GROUP = ["/firewall/vipgrp", "vip_group"]
    ENDPOINT_FIREWALL_IPPOOL = ["/firewall/ippool", "ippool"]
    ENDPOINT_FIREWALL_POLICY = ["/firewall/policy", "policy"]

    def __init__(self, ip: str = "", port: int = 40403, api_key: str = "", proxy_local_port: int = 0) -> None:
        self.ip = ip
        self.port = port
        self.base_url = f"https://{ip}:{port}/api/v2/cmdb"
        self.api_key = api_key
        self.proxy = dict(https=f'socks5://@127.0.0.1:{proxy_local_port}') if proxy_local_port != 0 else dict()
        self.headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
        self.vdom = self.get_vdoms()
        self.configuration = []

    def get_vdoms(self) -> list[dict]:
        vdoms = requests.get(f"{self.base_url}{FortiGateApi.ENDPOINT_FIREWALL_VDOM[0]}?access_token={self.api_key}", headers=self.headers, verify=False, proxies=self.proxy)
        return vdoms.json()['results']

    def get_api_request(self, endpoint) -> dict:
        self.configuration = self.vdom
        for index, vdom in enumerate(self.configuration):
            req = requests.get(f"{self.base_url}{endpoint[0]}?vdom={vdom['name']}&access_token={self.api_key}", headers=self.headers, verify=False, proxies=self.proxy)
            self.configuration[index][endpoint[1]] = req.json()['results']
        return self.configuration

    def get_xlsx_file(self):
        file = FortigateToXlsxFromAPI()
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_INTERFACE)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_SERVICE)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_SERVICE_GROUP)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_ADDRESS)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_ADDRESS_GROUP)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_VIP)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_VIP_GROUP)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_IPPOOL)
        self.get_api_request(FortiGateApi.ENDPOINT_FIREWALL_POLICY)
        file.export_xlsx_file(self.configuration)


# proxy_local_port = input("Un port proxy (si non presse entrer): ")
# ip = input("IP : ")
# api_key = input("clé d'API : ")
# if proxy_local_port:
#     forti = FortiGateApi(ip=ip, api_key=api_key, proxy_local_port=proxy_local_port)
# else:
#     forti = FortiGateApi(ip=ip, api_key=api_key)
# forti = FortiGateApi(ip="10.48.124.36", api_key="8frtpzxmp91rjmxGdccwtrysxkyc9j", proxy_local_port=7080)
# forti.get_xlsx_file()
