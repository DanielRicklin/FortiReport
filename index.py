#!/usr/bin/env python3
import argparse
from getpass import getpass
from src.FortigateApi import FortiGateApi
from src.FortigateSSH import FortigateSSH

parser = argparse.ArgumentParser()
parser.add_argument('-s', '--ssh', action="store_true")
parser.add_argument('-b', '--bastion', action="store_true")
parser.add_argument('-a', '--api', action="store_true")
args = parser.parse_args()

ip = input("IP : ")

if args.ssh:
    port = input("Port SSH : ")
    user = input("Nom d'utilisateur : ")
    password = getpass("Mot de passe : ")
    if args.bastion:
        bastion_ip = input("IP/FQDN du bastion : ")
        bastion_port = input("Port SSH du bastion : ")
        q = ""
        while q not in ["Y", "N"]:
            q = input("L'utilisateur est le même ? (Y/N) : ")
            if q == "Y":
                forti = FortigateSSH(ip=ip, port=port, user=user, password=password, bastion_port=bastion_port, bastion_ip=bastion_ip, bastion_user=user, bastion_password=password)
            elif q == "N":
                bastion_user = input("Nom d'utilisateur Bastion: ")
                bastion_password = getpass("Mot de passe Bastion: ")
                forti = FortigateSSH(ip=ip, port=port, user=user, password=password, bastion_port=bastion_port, bastion_ip=bastion_ip, bastion_user=bastion_user, bastion_password=bastion_password)
    else:
        forti = FortigateSSH(ip=ip, port=port, user=user, password=password)
        # forti.get_xlsx_file()
elif args.api:
    api_key = input("Clé d'API : ")
    proxy_port = input("Port proxy : ")
    forti = FortiGateApi(ip=ip, api_key=api_key, proxy_local_port=proxy_port)
    forti.get_xlsx_file()
else:
    print("Il faut ajouter l'option --ssh ou --api pour avancer")
