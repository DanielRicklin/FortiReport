import ipaddress
from .FortigateToXlsx import FortigateToXlsx
import xlsxwriter


class FortigateToXlsxFromAPI(FortigateToXlsx):
    """Here is created and exported the xlsx file with your configuration"""

    def __init__(self, filename) -> None:
        super().__init__()
        if filename:
            self.workbook = xlsxwriter.Workbook(f"Report_{filename}.xlsx")

    def get_security_profile_format(self, sp: str):
        colors = {
            "WEB": "#189fba",
            "AV": "#ff4d00",
            "SSL": "#be9e6f",
            "APP": "#009848",
            "IPS": "#aeb948",
        }
        return self.workbook.add_format(
            {
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "font_color": "white",
                "fg_color": colors[sp],
            }
        )

    def style_interface_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "J"),
            "ZONES",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )

        self.lign += 1

        self.worksheet.write(
            self.lign,
            self.column,
            "Nomes des zones",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "IP",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 2,
            "Nom interface physique",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 3,
            "Type d'interface",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 4,
            "Vlan ID",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("H", "J"),
            "Description",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )

        self.lign += 1

        for interface in self.configuration["interface"]:
            if (
                "vdom" in interface and self.vdom_name == interface["vdom"]
            ) or "intrazone" in interface:
                temp_lign_interface = 1
                if "intrazone" in interface:
                    temp_lign_interface = len(interface["interface"])
                if "alias" in interface:
                    if temp_lign_interface == 1:
                        self.worksheet.write(
                            self.lign,
                            self.column,
                            (
                                interface.get("alias")
                                if interface.get("alias")
                                else interface.get("name")
                            ),
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                    else:
                        self.worksheet.merge_range(
                            f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+temp_lign_interface}",
                            (
                                interface.get("alias")
                                if interface.get("alias")
                                else interface.get("name")
                            ),
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                else:
                    if temp_lign_interface == 1:
                        self.worksheet.write(
                            self.lign,
                            self.column,
                            interface["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                    else:
                        self.worksheet.merge_range(
                            f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+temp_lign_interface}",
                            interface["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                if "ip" in interface:
                    self.worksheet.write(
                        self.lign,
                        self.column + 1,
                        str(
                            ipaddress.ip_network(
                                interface["ip"].replace(" ", "/"), strict=False
                            )
                        ),
                        self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 2,
                        interface["interface"],
                        self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 3,
                        interface["type"],
                        self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 4,
                        interface["vlanid"],
                        self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                    )
                if "intrazone" in interface:
                    for index, itf in enumerate(interface["interface"]):
                        interface_zone = list(
                            filter(
                                lambda vdom_interface: vdom_interface["name"]
                                == itf["interface-name"],
                                self.configuration["interface"],
                            )
                        )
                        self.worksheet.write(
                            self.lign + index,
                            self.column + 1,
                            str(
                                ipaddress.ip_network(
                                    interface_zone[0]["ip"].replace(" ", "/"),
                                    strict=False,
                                )
                            ),
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign + index,
                            self.column + 2,
                            interface_zone[0]["alias"],
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign + index,
                            self.column + 3,
                            interface_zone[0]["type"],
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign + index,
                            self.column + 4,
                            interface_zone[0]["vlanid"],
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                if temp_lign_interface == 1:
                    columns_to_merge = self.get_merge_string("H", "J")
                else:
                    columns_to_merge = f"{chr(self.column+70)}{self.lign+1}:{chr(self.column+72)}{self.lign+temp_lign_interface}"
                self.worksheet.merge_range(
                    columns_to_merge,
                    interface["description"],
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += temp_lign_interface
        self.lign += 2

    def style_service_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "SERVICES",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "Services utilisés",
            self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
        )
        self.lign += 1
        self.worksheet.write(
            self.lign,
            self.column,
            "Nom du service",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "Protocole",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 2,
            "Port",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("F", "H"),
            "Commentaire",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.lign += 1

        for service in self.configuration["service"]:
            is_service_in_policies = False
            is_service_in_group = False
            for policy in self.configuration["policy"]:
                if list(
                    filter(
                        lambda policy_service: policy_service["name"]
                        == service["name"],
                        policy["service"],
                    )
                ):
                    is_service_in_policies = True
            for grp_service in self.configuration["service_group"]:
                if list(
                    filter(
                        lambda service_in_grp: grp_service["name"] == service["name"],
                        grp_service["member"],
                    )
                ):
                    is_service_in_group = True
                #  il faut rajouter si le service qui est dans un groupe est aussi dans une policy
            if (
                "tcp-portrange" in service
                and "udp-portrange" in service
                and "sctp-portrange" in service
                and service["visibility"] == "enable"
                and (is_service_in_policies or is_service_in_group)
            ):
                self.worksheet.write(
                    self.lign,
                    self.column,
                    service["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                protocols = ""
                ports = ""
                udp = False
                sctp = False
                if service["tcp-portrange"]:
                    protocols += "TCP"
                    ports += service["tcp-portrange"].replace(" ", "/")
                if service["udp-portrange"] and protocols == "":
                    protocols += "UDP"
                    udp = True
                    ports += service["udp-portrange"].replace(" ", "/")
                if service["udp-portrange"] and protocols != "" and not udp:
                    protocols += "/UDP"
                    if ports != service["udp-portrange"].replace(" ", "/"):
                        ports += f"/{service['udp-portrange'].replace(' ','/')}"
                if service["sctp-portrange"] and protocols == "":
                    protocols += "SCTP"
                    ports += service["sctp-portrange"].replace(" ", "/")
                if service["sctp-portrange"] and protocols != "" and not sctp:
                    protocols += "/SCTP"
                    ports += f"/{service['sctp-portrange'].replace(' ', '/')}"
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    protocols,
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 2,
                    ports,
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("F", "H"),
                    service["comment"],
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1
            elif is_service_in_policies or is_service_in_group:
                self.worksheet.write(
                    self.lign,
                    self.column,
                    service["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    service["protocol"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 2,
                    "",
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )  # pas sûr de mon coup
                self.worksheet.merge_range(
                    self.get_merge_string("F", "H"),
                    service["comment"],
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1
        self.lign += 2

    def style_service_group_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "Groupe de services",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1
        self.worksheet.write(
            self.lign,
            self.column,
            "Nom du groupe",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "Membres",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("E", "H"),
            "Commentaires",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.lign += 1

        for grp_service in self.configuration["service_group"]:
            if (
                (grp_service["name"] != "Email Access")
                and (grp_service["name"] != "Exchange Server")
                and (grp_service["name"] != "Web Access")
            ):
                self.worksheet.merge_range(
                    f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+len(grp_service['member'])}",
                    grp_service["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                for index, service in enumerate(grp_service["member"]):
                    self.worksheet.write(
                        self.lign + index,
                        self.column + 1,
                        service["name"],
                        self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                    )
                self.worksheet.merge_range(
                    f"{chr(self.column+67)}{self.lign+1}:{chr(self.column+70)}{self.lign+len(grp_service['member'])}",
                    grp_service["comment"],
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.lign += len(grp_service["member"])
        self.lign += 2

    def style_address_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "OBJETS RESEAU",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1

        for interface in self.configuration["interface"]:
            addresses_list = list(
                filter(
                    lambda address: address["associated-interface"]
                    == interface["name"],
                    self.configuration["address"],
                )
            )
            if addresses_list:
                self.worksheet.merge_range(
                    self.get_merge_string("C", "H"),
                    f"Zone {interface.get('alias') if interface.get('alias') else interface.get('name')} - Objects",
                    self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
                )
                self.lign += 1
                self.worksheet.write(
                    self.lign,
                    self.column,
                    "Nom",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    "subnet/FQDN/range",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("E", "H"),
                    "Commentaire",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1

                for address in self.configuration["address"]:
                    if interface["name"] == address["associated-interface"]:
                        self.worksheet.write(
                            self.lign,
                            self.column,
                            address["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        if address["type"] == "ipmask":
                            self.worksheet.write(
                                self.lign,
                                self.column + 1,
                                str(
                                    ipaddress.ip_network(
                                        address.get("subnet").replace(" ", "/"),
                                        strict=False,
                                    )
                                ),
                                self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                            )
                        if address["type"] == "fqdn":
                            self.worksheet.write(
                                self.lign,
                                self.column + 1,
                                address.get("fqdn"),
                                self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                            )
                        if address["type"] == "iprange":
                            self.worksheet.write(
                                self.lign,
                                self.column + 1,
                                f"{address['start-ip']}-{address['end-ip']}",
                                self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                            )
                        self.worksheet.merge_range(
                            self.get_merge_string("E", "H"),
                            address["comment"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                        self.lign += 1

        #  les sans interfaces
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "Zone sans interface - Objects",
            self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
        )
        self.lign += 1
        self.worksheet.write(
            self.lign,
            self.column,
            "Nom",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "subnet/FQDN/range",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("E", "H"),
            "Commentaire",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.lign += 1
        for address in list(
            filter(
                lambda address: address["associated-interface"] == "",
                self.configuration["address"],
            )
        ):
            self.worksheet.write(
                self.lign,
                self.column,
                address["name"],
                self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
            )
            if address["type"] == "ipmask":
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    str(
                        ipaddress.ip_network(
                            address.get("subnet").replace(" ", "/"), strict=False
                        )
                    ),
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            if address["type"] == "fqdn":
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    address.get("fqdn"),
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            if address["type"] == "iprange":
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    f"{address['start-ip']}-{address['end-ip']}",
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            self.worksheet.merge_range(
                self.get_merge_string("E", "H"),
                address["comment"],
                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
            )
            self.lign += 1

        self.lign += 2

    def style_address_group_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "GROUPE D'OBJETS RESEAU",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1
        self.worksheet.write(
            self.lign,
            self.column,
            "Nom",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "Membres",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("E", "H"),
            "Commentaire",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.lign += 1

        for grp_address in self.configuration["address_group"]:
            if len(grp_address["member"]) > 1:
                self.worksheet.merge_range(
                    f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+len(grp_address['member'])}",
                    grp_address["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            else:
                self.worksheet.write(
                    self.lign,
                    self.column,
                    grp_address["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            for index, address in enumerate(grp_address["member"]):
                self.worksheet.write(
                    self.lign + index,
                    self.column + 1,
                    address["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
            self.worksheet.merge_range(
                f"{chr(self.column+67)}{self.lign+1}:{chr(self.column+70)}{self.lign+len(grp_address['member'])}",
                grp_address["comment"],
                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
            )
            self.lign += len(grp_address["member"])

        self.lign += 2

    def style_vip_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "J"),
            "VIP",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1

        for interface in self.configuration["interface"]:
            vip_list = list(
                filter(
                    lambda vip: vip["extintf"] == interface["name"],
                    self.configuration["vip"],
                )
            )
            if vip_list:
                self.worksheet.merge_range(
                    self.get_merge_string("C", "J"),
                    f"Zone {interface.get('alias') if interface.get('alias') else interface.get('name')} - VIP",
                    self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
                )
                self.lign += 1
                self.worksheet.write(
                    self.lign,
                    self.column,
                    "Nom de la VIP",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    "IP src",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 2,
                    "Port src",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 3,
                    "IP dest",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 4,
                    "Port dest",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("H", "J"),
                    "Commentaire",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1
                for vip in vip_list:
                    self.worksheet.write(
                        self.lign,
                        self.column,
                        vip["name"],
                        self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 1,
                        vip["extip"],
                        self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 2,
                        vip["extport"],
                        self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 3,
                        vip["mappedip"][0]["range"],
                        self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                    )
                    self.worksheet.write(
                        self.lign,
                        self.column + 4,
                        vip["mappedport"],
                        self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                    )
                    self.worksheet.merge_range(
                        self.get_merge_string("H", "J"),
                        vip["comment"],
                        self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                    )
                    self.lign += 1

        vip_any_list = list(
            filter(lambda vip: vip["extintf"] == "any", self.configuration["vip"])
        )
        if vip_any_list:
            self.worksheet.merge_range(
                self.get_merge_string("C", "J"),
                "Zone Sans Interface - VIP",
                self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
            )
            self.lign += 1
            self.worksheet.write(
                self.lign,
                self.column,
                "Nom de la VIP",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 1,
                "IP src",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 2,
                "Port src",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 3,
                "IP dest",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 4,
                "Port dest",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.worksheet.merge_range(
                self.get_merge_string("H", "J"),
                "Commentaire",
                self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
            )
            self.lign += 1

            for vip in vip_any_list:
                self.worksheet.write(
                    self.lign,
                    self.column,
                    vip["name"],
                    self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    vip["extip"],
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 2,
                    vip["extport"],
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 3,
                    vip["mappedip"][0]["range"],
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 4,
                    vip["mappedport"],
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("H", "J"),
                    vip["comment"],
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1

        self.lign += 2

    def style_vip_group_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "G"),
            "Groupe de VIP",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1

        for interface in self.configuration["interface"]:
            grp_vip_list = list(
                filter(
                    lambda grp_vip: grp_vip["interface"] == interface["name"],
                    self.configuration["vip_group"],
                )
            )
            if grp_vip_list:
                self.worksheet.merge_range(
                    self.get_merge_string("C", "G"),
                    f"Zone {interface.get('alias') if interface.get('alias') else interface.get('name')} - VIP",
                    self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
                )
                self.lign += 1
                self.worksheet.write(
                    self.lign,
                    self.column,
                    "Nom du groupe",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    "Membres",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("E", "G"),
                    "Commentaire",
                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                )
                self.lign += 1
                for grp_vip in grp_vip_list:
                    nbr_lign = len(grp_vip["member"])
                    if nbr_lign > 1:
                        self.worksheet.merge_range(
                            f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+nbr_lign}",
                            grp_vip["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        for index, vip in enumerate(grp_vip["member"]):
                            self.worksheet.write(
                                self.lign + index,
                                self.column + 1,
                                vip["name"],
                                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                            )
                        self.worksheet.merge_range(
                            f"{chr(self.column+67)}{self.lign+1}:{chr(self.column+69)}{self.lign+nbr_lign}",
                            grp_vip["comments"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                    else:
                        self.worksheet.write(
                            self.lign,
                            self.column,
                            grp_vip["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign,
                            self.column + 1,
                            grp_vip["member"][0]["name"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            self.get_merge_string("E", "G"),
                            grp_vip["comments"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                    self.lign += nbr_lign

        self.lign += 2

    def style_ippool_zone(self):
        self.worksheet.merge_range(
            self.get_merge_string("C", "H"),
            "IP POOL",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1
        self.worksheet.write(
            self.lign,
            self.column,
            "Nom du pool",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 1,
            "Type",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.write(
            self.lign,
            self.column + 2,
            "Range",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.worksheet.merge_range(
            self.get_merge_string("F", "H"),
            "Commentaire",
            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
        )
        self.lign += 1

        for ippool in self.configuration["ippool"]:
            self.worksheet.write(
                self.lign,
                self.column,
                ippool["name"],
                self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 1,
                ippool["type"],
                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
            )
            self.worksheet.write(
                self.lign,
                self.column + 2,
                f"{ippool['startip']}/{ippool['endip']}",
                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
            )
            self.worksheet.merge_range(
                self.get_merge_string("F", "H"),
                ippool["comments"],
                self.get_format(FortigateToXlsx.ALIGN_FORMAT),
            )
            self.lign += 1

        self.lign += 2

    def style_policy_zone(self):
        array_interfaces_policy = []

        # On forme les paquets de policy car sinon c'est le bordel
        for interface_src in self.configuration["interface"]:
            for interface_dest in self.configuration["interface"]:
                if interface_src["name"] != interface_dest["name"]:
                    interface_src_name = list(
                        filter(
                            lambda interface: interface["name"]
                            == interface_src["name"],
                            self.configuration["interface"],
                        )
                    )
                    interface_dest_name = list(
                        filter(
                            lambda interface: interface["name"]
                            == interface_dest["name"],
                            self.configuration["interface"],
                        )
                    )
                    array_interfaces_policy.append(
                        {
                            "sens": f"{interface_src_name[0].get('alias') if interface_src_name[0].get('alias') else interface_src['name']} vers {interface_dest_name[0].get('alias') if interface_dest_name[0].get('alias') else interface_dest['name']}",
                            "policy": [],
                        }
                    )

        # On range les policies dans nos magnifiques paquets
        for policy in self.configuration["policy"]:
            interface_src_name = list(
                filter(
                    lambda interface: interface["name"] == policy["srcintf"][0]["name"],
                    self.configuration["interface"],
                )
            )
            interface_dest_name = list(
                filter(
                    lambda interface: interface["name"] == policy["dstintf"][0]["name"],
                    self.configuration["interface"],
                )
            )
            for interface_policy in array_interfaces_policy:
                if (
                    f"{interface_src_name[0].get('alias') if interface_src_name[0].get('alias') else policy['srcintf'][0]['name']} vers {interface_dest_name[0].get('alias') if interface_dest_name[0].get('alias') else policy['dstintf'][0]['name']}"
                    == interface_policy["sens"]
                ):
                    interface_policy["policy"].append(policy)

        self.worksheet.merge_range(
            self.get_merge_string("H", "I"),
            "A=Activé, D=Désactivé",
            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
        )
        self.lign += 1
        self.worksheet.merge_range(
            self.get_merge_string("B", "K"),
            "POLICIES",
            self.get_format(FortigateToXlsx.MERGE_FORMAT),
        )
        self.lign += 1

        for sens_flux in array_interfaces_policy:
            if sens_flux["policy"] != []:
                self.worksheet.merge_range(
                    self.get_merge_string("B", "K"),
                    sens_flux["sens"],
                    self.get_format(FortigateToXlsx.SUBTITLE_FORMAT),
                )
                self.lign += 1
                self.worksheet.write(
                    self.lign,
                    self.column - 1,
                    "ID",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column,
                    "Nom",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 1,
                    "Source",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 2,
                    "Destination",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 3,
                    "Protocole",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 4,
                    "NAT",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 5,
                    "Profils de sécurités",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.write(
                    self.lign,
                    self.column + 6,
                    "A/D",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.worksheet.merge_range(
                    self.get_merge_string("J", "K"),
                    "Commentaire",
                    self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                )
                self.lign += 1

                for policy in sens_flux["policy"]:
                    length_src = 0
                    length_dest = 0
                    length_service = 0
                    nat = ""

                    if policy["internet-service-src"] == "disable":
                        length_src = len(policy["srcaddr"])
                    else:
                        length_src = len(policy["internet-service-src-name"])
                    if policy["internet-service"] == "disable":
                        length_dest = len(policy["dstaddr"])
                        length_service = len(policy["service"])
                    else:
                        length_dest = len(policy["internet-service-name"])

                    # Secutiry profiles
                    security_profiles = []
                    if policy["ssl-ssh-profile"]:
                        security_profiles.append(f"SSL: {policy['ssl-ssh-profile']}")
                    if policy["av-profile"]:
                        security_profiles.append(f"AV: {policy['av-profile']}")
                    if policy["webfilter-profile"]:
                        security_profiles.append(f"WEB: {policy['webfilter-profile']}")
                    if policy["dnsfilter-profile"]:
                        security_profiles.append(f"DNS: {policy['dnsfilter-profile']}")
                    if policy["application-list"]:
                        security_profiles.append(f"APP: {policy['application-list']}")
                    if policy["ips-sensor"]:
                        security_profiles.append(f"IPS: {policy['ips-sensor']}")

                    nbr_lign = max(
                        [
                            length_src,
                            length_dest,
                            length_service,
                            len(security_profiles),
                        ]
                    )
                    id_lign_source_temp = self.lign
                    id_lign_destination_temp = self.lign
                    id_lign_service_temp = self.lign
                    id_lign_secu_profile_temp = self.lign

                    if policy["internet-service-src"] == "disable":
                        for source in policy["srcaddr"]:
                            if length_src == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+66)}{id_lign_source_temp+1}:{chr(self.column+66)}{id_lign_source_temp+nbr_lign}",
                                    source["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            else:
                                self.worksheet.write(
                                    id_lign_source_temp,
                                    self.column + 1,
                                    source["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_source_temp += 1
                    else:
                        for source in policy["internet-service-src-name"]:
                            if length_src == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+66)}{id_lign_source_temp+1}:{chr(self.column+66)}{id_lign_source_temp+nbr_lign}",
                                    source["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            else:
                                self.worksheet.write(
                                    id_lign_source_temp,
                                    self.column + 1,
                                    source["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_source_temp += 1
                    if policy["internet-service"] == "disable":
                        for destination in policy["dstaddr"]:
                            if length_dest == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+67)}{id_lign_destination_temp+1}:{chr(self.column+67)}{id_lign_destination_temp+nbr_lign}",
                                    destination["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            else:
                                self.worksheet.write(
                                    id_lign_destination_temp,
                                    self.column + 2,
                                    destination["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_destination_temp += 1
                        for service in policy["service"]:
                            if length_service == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+68)}{id_lign_service_temp+1}:{chr(self.column+68)}{id_lign_service_temp+nbr_lign}",
                                    service["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            else:
                                self.worksheet.write(
                                    id_lign_service_temp,
                                    self.column + 3,
                                    service["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_service_temp += 1
                    else:
                        for destination in policy["internet-service-name"]:
                            if length_dest == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+67)}{id_lign_destination_temp+1}:{chr(self.column+67)}{id_lign_destination_temp+nbr_lign}",
                                    destination["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            else:
                                self.worksheet.write(
                                    id_lign_destination_temp,
                                    self.column + 2,
                                    destination["name"],
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_destination_temp += 1
                        for service in policy["internet-service-name"]:
                            if length_service == 1 and nbr_lign > 1:
                                self.worksheet.merge_range(
                                    f"{chr(self.column+68)}{id_lign_service_temp+1}:{chr(self.column+68)}{id_lign_service_temp+nbr_lign}",
                                    "Internet Service",
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                                break
                            else:
                                self.worksheet.write(
                                    id_lign_service_temp,
                                    self.column + 3,
                                    "Internet Service",
                                    self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                                )
                            id_lign_service_temp += 1

                    for sp in security_profiles:
                        format = self.get_format(FortigateToXlsx.ALIGN_FORMAT)
                        if "WEB:" in sp:
                            format = self.workbook.add_format(
                                {
                                    "border": 1,
                                    "align": "center",
                                    "valign": "vcenter",
                                    "font_color": "white",
                                    "fg_color": "#189fba",
                                }
                            )
                        if "AV:" in sp:
                            format = self.workbook.add_format(
                                {
                                    "border": 1,
                                    "align": "center",
                                    "valign": "vcenter",
                                    "font_color": "white",
                                    "fg_color": "#ff4d00",
                                }
                            )
                        if "APP:" in sp:
                            format = self.workbook.add_format(
                                {
                                    "border": 1,
                                    "align": "center",
                                    "valign": "vcenter",
                                    "font_color": "white",
                                    "fg_color": "#009848",
                                }
                            )
                        if "SSL:" in sp:
                            format = self.workbook.add_format(
                                {
                                    "border": 1,
                                    "align": "center",
                                    "valign": "vcenter",
                                    "font_color": "white",
                                    "fg_color": "#be9e6f",
                                }
                            )
                        if "IPS:" in sp:
                            format = self.workbook.add_format(
                                {
                                    "border": 1,
                                    "align": "center",
                                    "valign": "vcenter",
                                    "font_color": "white",
                                    "fg_color": "#aeb948",
                                }
                            )
                        if len(security_profiles) == 1 and nbr_lign > 1:
                            self.worksheet.merge_range(
                                f"{chr(self.column+70)}{id_lign_secu_profile_temp+1}:{chr(self.column+70)}{id_lign_secu_profile_temp+nbr_lign}",
                                sp,
                                format,
                            )
                        else:
                            self.worksheet.write(
                                id_lign_secu_profile_temp, self.column + 5, sp, format
                            )
                        id_lign_secu_profile_temp += 1

                    # NAT
                    if policy["nat"] == "enable" and policy["ippool"] == "disable":
                        nat = "A"
                    elif policy["nat"] == "enable" and policy["ippool"] == "enable":
                        nat = policy["poolname"][0]["name"]
                    elif policy["nat"] == "disable":
                        nat = "D"

                    if nbr_lign > 1:
                        self.worksheet.merge_range(
                            f"{chr(self.column+64)}{self.lign+1}:{chr(self.column+64)}{self.lign+nbr_lign}",
                            policy["policyid"],
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            f"{chr(self.column+65)}{self.lign+1}:{chr(self.column+65)}{self.lign+nbr_lign}",
                            policy["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            f"{chr(self.column+69)}{self.lign+1}:{chr(self.column+69)}{self.lign+nbr_lign}",
                            nat,
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            f"{chr(self.column+71)}{self.lign+1}:{chr(self.column+71)}{self.lign+nbr_lign}",
                            f"{'A' if policy['status'] == 'enable' else 'D'}",
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            f"{chr(self.column+72)}{self.lign+1}:{chr(self.column+73)}{self.lign+nbr_lign}",
                            policy["comments"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                    else:
                        self.worksheet.write(
                            self.lign,
                            self.column - 1,
                            policy["policyid"],
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign,
                            self.column,
                            policy["name"],
                            self.get_format(FortigateToXlsx.BLUE_BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign,
                            self.column + 4,
                            nat,
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.write(
                            self.lign,
                            self.column + 6,
                            f"{'A' if policy['status'] == 'enable' else 'D'}",
                            self.get_format(FortigateToXlsx.BOLD_ALIGN_FORMAT),
                        )
                        self.worksheet.merge_range(
                            self.get_merge_string("J", "K"),
                            policy["comments"],
                            self.get_format(FortigateToXlsx.ALIGN_FORMAT),
                        )
                    self.lign += nbr_lign

    def export_xlsx_file(self, vdoms):
        for vdom in vdoms:
            self.lign = 3
            self.column = 2
            self.vdom_name = vdom["name"]
            self.configuration = vdom

            self.worksheet = self.workbook.add_worksheet(vdom["name"])
            self.worksheet.set_column("C:C", 35)
            self.worksheet.set_column("D:D", 30)
            self.worksheet.set_column("E:E", 30)
            self.worksheet.set_column("F:F", 20)
            self.worksheet.set_column("G:G", 15)
            self.worksheet.set_column("H:H", 25)
            self.worksheet.set_column("I:I", 5)
            self.worksheet.set_column("J:J", 50)

            if self.configuration.get("interface"):
                self.style_interface_zone()
            if self.configuration.get("service"):
                self.style_service_zone()
            if self.configuration.get("service_group"):
                self.style_service_group_zone()
            if self.configuration.get("address"):
                self.style_address_zone()
            if self.configuration.get("address_group"):
                self.style_address_group_zone()
            if self.configuration.get("vip"):
                self.style_vip_zone()
            if self.configuration.get("vip_group"):
                self.style_vip_group_zone()
            if self.configuration.get("ippool"):
                self.style_ippool_zone()
            if self.configuration.get("policy"):
                self.style_policy_zone()

        self.workbook.close()
