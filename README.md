# Script python qui automatise la création d'une matrice Excel pour les Firewall Fortigate via l'API
Tu peux l'utiliser via un proxy, genre PuTTY

C'est la V2 de ça : https://gitlab.adista.intra/dricklin/matrice_hds

# Ça extrait quoi ?
- interfaces
- services
- groupes de services
- objets
- groupe d'objets
- VIP
- groupe de VIP
- IP Pool
- policies

# Comment l'utiliser ?
1. Créer un utilisateur API sur le Fortigate
2. ssh berthe2.exploit.intra
3. Lance la commande : 
```git clone git@gitlab.adista.intra:dricklin/fortireport.git```.
S'il te demande un mot de passe et que ça passe pas, c'est que t'as pas de clé publique d'enregistrée dans gitlab pour berthe2.
Du coup sur Berthe2 tu vas lancer : ```ssh-keygen -b 2048 -t rsa```.
Puis tu affiches le contenu du fichier : ```cat /home/dricklin/.ssh/id_rsa.pub```.
Pour finir, va ici : https://gitlab.adista.intra/-/profile/keys.
Ajoute le contenu du fichier, valide et recommence le git clone
4. Lance les commandes : 
```bash
cd fortireport
chmod +x main.py
pip3 install -r requirements.txt
./main.py
```
5. Récupère le fichier "Matrice_FLUX.xlsx" via WinSCP ou direct sur ton pc si t'utilise un proxy
