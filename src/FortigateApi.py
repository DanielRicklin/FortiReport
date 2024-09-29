#!/usr/bin/env python3

import requests
import logging
from .FortigateToXlsxFromAPI import FortigateToXlsxFromAPI

requests.packages.urllib3.disable_warnings()


class FortiGateApi:
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

    def __init__(
        self,
        ip: str = "",
        port: int = 443,
        api_key: str = "",
        proxy_local_port: int = 0,
        filename: str = "",
    ) -> None:
        self.ip = ip
        self.port = port
        self.base_url = f"https://{ip}:{port}/api/v2/cmdb"
        self.api_key = api_key
        self.filename = filename
        if proxy_local_port:
            self.proxy = (
                dict(https=f"socks5://@127.0.0.1:{proxy_local_port}")
                if proxy_local_port != 0
                else dict()
            )
        else:
            self.proxy = 0
        self.headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {api_key}",
        }
        self.vdom = self.get_vdoms()
        self.configuration = []

    def get_vdoms(self) -> list[dict]:
        logging.debug("Requesting vdoms list")
        if self.proxy:
            vdoms = requests.get(
                f"{self.base_url}{FortiGateApi.ENDPOINT_FIREWALL_VDOM[0]}?access_token={self.api_key}",
                headers=self.headers,
                verify=False,
                proxies=self.proxy,
            )
        else:
            vdoms = requests.get(
                f"{self.base_url}{FortiGateApi.ENDPOINT_FIREWALL_VDOM[0]}?access_token={self.api_key}",
                headers=self.headers,
                verify=False,
            )
        logging.debug(vdoms.json()["results"])
        return vdoms.json()["results"]

    def get_api_request(self, endpoint) -> dict:
        self.configuration = self.vdom
        for index, vdom in enumerate(self.configuration):
            logging.debug(
                f"For VDOM {vdom['name']}, Requesting {endpoint[1]}, Endpoint : {endpoint[0]}"
            )
            if self.proxy:
                req = requests.get(
                    f"{self.base_url}{endpoint[0]}?vdom={vdom['name']}&access_token={self.api_key}",
                    headers=self.headers,
                    verify=False,
                    proxies=self.proxy,
                )
            else:
                req = requests.get(
                    f"{self.base_url}{endpoint[0]}?vdom={vdom['name']}&access_token={self.api_key}",
                    headers=self.headers,
                    verify=False,
                )
            self.configuration[index][endpoint[1]] = req.json()["results"]
        return self.configuration

    def get_xlsx_file(self):
        file = FortigateToXlsxFromAPI(self.filename)
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
