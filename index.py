#!/usr/bin/env python3
import argparse
from getpass import getpass
import logging
from src.FortigateApi import FortiGateApi
from src.FortigateSSH import FortigateSSH

logging.basicConfig(format='%(asctime)s %(levelname)s: %(message)s')
logging.getLogger().setLevel(logging.INFO)

# def list_of_config(arg):
#     return arg.split(',')

parser = argparse.ArgumentParser(description="Make nice Excel Fortigate Report")
parser.add_argument('--ip', required=True, help="set the IP you want to access")
parser.add_argument('--args-mode', action="store_true", help="to run the program only with arguments")
parser.add_argument('--ssh', action="store_true", help="if you want to request your Fortigate through SSH")
parser.add_argument('--bastion', action="store_true", help="Only for SSH if you need to access you Fortigate through a bastion")
parser.add_argument('--api', action="store_true", help="if you want to request your Fortigate through the API")
parser.add_argument('--api-key', help="set the API KEY")
parser.add_argument('--proxy-port', help="set the proxy port is you need to access you Fortigate through a proxy")
parser.add_argument('--override-filename', help="set the name you want")
# parser.add_argument('--config-list', help="set the list of configs you want in you report", type=list_of_config)
parser.add_argument('--verbose', action="store_true", help="Verbose logging mode, will log all actions")
args = parser.parse_args()

if args.verbose:
    logging.getLogger().setLevel(logging.DEBUG)
    logging.debug(args)

if args.api and args.ssh:
    logging.error(f"--api and --ssh are both set")
    exit(1)
    # parser.error("You need to choose between SSH and API...")
elif args.api is None and args.ssh is None:
    logging.error(f"--api and --ssh are not set")
    exit(1)
    # parser.error("You need to choose between SSH and API...")

if args.args_mode:
    if args.api:
        logging.info("Running program through API")
        forti = FortiGateApi(ip=args.ip, api_key=args.api_key, proxy_local_port=args.proxy_port, filename=args.override_filename)
        forti.get_xlsx_file()
else:
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
        filename = input("Override filename : ")
        forti = FortiGateApi(ip=ip, api_key=api_key, proxy_local_port=proxy_port, filename=filename)
        forti.get_xlsx_file()
    else:
        print("Il faut ajouter l'option --ssh ou --api pour avancer")
