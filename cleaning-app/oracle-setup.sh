#!/usr/bin/env bash
# Instalarea aplicatiei de facturi pe un server Ubuntu (Oracle Cloud Always Free).
# Rulare (pe server, o singura data):
#   curl -fsSL https://raw.githubusercontent.com/AlexVladX/cleaning-app/main/cleaning-app/oracle-setup.sh | sudo bash
# Se poate rula din nou oricand (ex. ca sa schimbi parola): pastreaza ce exista deja daca apesi Enter.
#
# Ce face:
#   - instaleaza Python, git si Caddy (HTTPS automat, gratuit)
#   - descarca aplicatia din GitHub in /opt/facturi si o porneste ca serviciu (porneste singura la restart)
#   - cheile (Google, parola, Groq) le cere aici si le pune in /etc/facturi, citibile doar de aplicatie
#   - deschide porturile 80 / 443 in firewall-ul serverului
#   - la fiecare 5 minute verifica GitHub: daca ai urcat fisiere noi, actualizeaza si reporneste singura
set -euo pipefail

REPO="${REPO:-https://github.com/AlexVladX/cleaning-app.git}"
BRANCH="${BRANCH:-main}"
APP_DIR=/opt/facturi
CONF_DIR=/etc/facturi
ENV_FILE=$CONF_DIR/facturi.env
KEY_FILE=$CONF_DIR/google.json
SVC_USER=facturi

say()  { printf '\n\033[1;32m==> %s\033[0m\n' "$*"; }
warn() { printf '\033[1;33m!! %s\033[0m\n' "$*"; }
ask()  { # ask VAR "intrebare" [secret]
  local __v="$1" __q="$2" __s="${3:-}" __a=""
  if [ -n "$__s" ]; then read -r -s -p "$__q: " __a </dev/tty; echo; else read -r -p "$__q: " __a </dev/tty; fi
  printf -v "$__v" '%s' "$__a"
}

[ "$(id -u)" -eq 0 ] || { echo "Ruleaza cu sudo"; exit 1; }
. /etc/os-release
[ "${ID:-}" = "ubuntu" ] || warn "Scriptul e facut pentru Ubuntu (gasit: ${PRETTY_NAME:-?}). Continui oricum."

say "1/7 Pachete (Python, git, Caddy)"
export DEBIAN_FRONTEND=noninteractive
apt-get update -qq
apt-get install -y -qq python3 python3-venv python3-pip git curl sudo iptables-persistent >/dev/null
if ! apt-get install -y -qq caddy >/dev/null 2>&1; then
  # Ubuntu mai vechi (22.04) nu are Caddy in depozitele proprii: il luam din depozitul oficial Caddy
  apt-get install -y -qq debian-keyring debian-archive-keyring apt-transport-https gnupg >/dev/null
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor --yes -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' > /etc/apt/sources.list.d/caddy-stable.list
  apt-get update -qq && apt-get install -y -qq caddy >/dev/null
fi

# Pe serverul mic (1 GB memorie) adaugam 2 GB swap, ca citirea PDF-urilor sa nu ramana fara memorie
MEM_MB=$(awk '/MemTotal/ {print int($2/1024)}' /proc/meminfo)
if [ "$MEM_MB" -lt 2000 ] && ! swapon --show | grep -q /swapfile; then
  say "Memorie mica (${MEM_MB} MB): adaug 2 GB swap"
  fallocate -l 2G /swapfile && chmod 600 /swapfile && mkswap /swapfile >/dev/null && swapon /swapfile
  grep -q '^/swapfile' /etc/fstab || echo '/swapfile none swap sw 0 0' >> /etc/fstab
fi

say "2/7 Utilizator separat pentru aplicatie"
id "$SVC_USER" >/dev/null 2>&1 || useradd --system --home "$APP_DIR" --shell /usr/sbin/nologin "$SVC_USER"

say "3/7 Aplicatia din GitHub ($REPO)"
if [ -d "$APP_DIR/.git" ]; then
  git -C "$APP_DIR" fetch -q origin "$BRANCH" && git -C "$APP_DIR" reset -q --hard "origin/$BRANCH"
else
  rm -rf "$APP_DIR"; git clone -q --branch "$BRANCH" "$REPO" "$APP_DIR"
fi
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install -q --upgrade pip
"$APP_DIR/venv/bin/pip" install -q -r "$APP_DIR/cleaning-app/requirements.txt"
chown -R "$SVC_USER:$SVC_USER" "$APP_DIR"
git config --system --add safe.directory "$APP_DIR" 2>/dev/null || true

say "4/7 Cheile si setarile (raman doar pe server, in $CONF_DIR)"
mkdir -p "$CONF_DIR"; chmod 750 "$CONF_DIR"
# valorile existente (la o rulare noua, Enter = le pastrezi)
old() { [ -f "$ENV_FILE" ] && sed -n "s/^$1=\"\(.*\)\"$/\1/p" "$ENV_FILE" | sed 's/\\"/"/g; s/\\\\/\\/g' || true; }
O_PASS=$(old APP_PASSWORD); O_FOLDER=$(old DRIVE_FOLDER_ID); O_GROQ=$(old GROQ_API_KEY); O_DOMAIN=$(old APP_DOMAIN)

echo "Le gasesti in Render → serviciul tau → Environment (sau unde le-ai notat)."
ask PASS   "APP_PASSWORD (parola paginii)${O_PASS:+ [Enter = pastrez]}" secret;   PASS=${PASS:-$O_PASS}
ask FOLDER "DRIVE_FOLDER_ID (folderul cu facturi din Drive)${O_FOLDER:+ [Enter = $O_FOLDER]}"; FOLDER=${FOLDER:-$O_FOLDER}
ask GROQ   "GROQ_API_KEY (poate ramane gol)${O_GROQ:+ [Enter = pastrez]}" secret;  GROQ=${GROQ:-$O_GROQ}
[ -n "$PASS" ] || { echo "Parola e obligatorie (altfel oricine ar vedea facturile)"; exit 1; }
[ -n "$FOLDER" ] || { echo "DRIVE_FOLDER_ID e obligatoriu"; exit 1; }

if [ -s "$KEY_FILE" ]; then
  ask REKEY "Cheia Google exista deja. O inlocuiesti? (d/N)"
else REKEY=d; fi
if [ "${REKEY,,}" = "d" ]; then
  echo "Lipeste acum TOT continutul cheii Google (GOOGLE_CREDENTIALS_JSON, incepe cu { si se termina cu }),"
  echo "apoi apasa Enter. Citirea se opreste singura cand JSON-ul e complet."
  python3 - "$KEY_FILE" <<'PY'
import json, sys
buf = ""
with open("/dev/tty") as tty:
    for line in tty:
        buf += line
        try:
            info = json.loads(buf)
        except ValueError:
            continue
        if info.get("type") != "service_account" or "private_key" not in info:
            sys.exit("Nu pare cheia unui service account (lipseste type / private_key)")
        with open(sys.argv[1], "w", encoding="utf-8") as f:
            json.dump(info, f)
        print("Cheie salvata pentru:", info.get("client_email"))
        break
    else:
        sys.exit("JSON incomplet")
PY
fi

IP=$(curl -fsS -4 --max-time 10 https://ifconfig.me || curl -fsS -4 --max-time 10 https://api.ipify.org || true)
DEF_DOMAIN=${O_DOMAIN:-${IP:+${IP//./-}.sslip.io}}
echo "Adresa paginii. Fara domeniu propriu, folosim una gratuita facuta din IP: $DEF_DOMAIN"
ask DOMAIN "Domeniu [Enter = $DEF_DOMAIN]"; DOMAIN=${DOMAIN:-$DEF_DOMAIN}
[ -n "$DOMAIN" ] || { echo "Nu am putut afla IP-ul public; ruleaza din nou si scrie domeniul"; exit 1; }

q() { printf '"%s"' "$(printf '%s' "$1" | sed 's/\\/\\\\/g; s/"/\\"/g')"; }
umask 077
{
  echo "APP_PASSWORD=$(q "$PASS")"
  echo "DRIVE_FOLDER_ID=$(q "$FOLDER")"
  [ -n "$GROQ" ] && echo "GROQ_API_KEY=$(q "$GROQ")"
  echo "GOOGLE_CREDENTIALS_FILE=\"$KEY_FILE\""
  echo "APP_DOMAIN=$(q "$DOMAIN")"
  echo "APP_SERVER_NAME=\"oracle\""
  echo "PORT=\"8000\""
  echo "POLL_INTERVAL=\"120\""
  # variabile in plus (ex. MONITOR=0): le adaugi in /etc/facturi/extra.env, nu sunt atinse de script
} > "$ENV_FILE"
touch "$CONF_DIR/extra.env"
chown -R root:"$SVC_USER" "$CONF_DIR"; chmod 640 "$ENV_FILE" "$KEY_FILE" "$CONF_DIR/extra.env"
umask 022

say "5/7 Serviciul aplicatiei (porneste singur, si dupa restart)"
cat > /etc/systemd/system/facturi.service <<EOF
[Unit]
Description=Analizator facturi SimNettoyage
After=network-online.target
Wants=network-online.target

[Service]
User=$SVC_USER
WorkingDirectory=$APP_DIR/cleaning-app
EnvironmentFile=$ENV_FILE
EnvironmentFile=-$CONF_DIR/extra.env
ExecStart=$APP_DIR/venv/bin/gunicorn app:app --config gunicorn.conf.py --bind 127.0.0.1:8000
Restart=always
RestartSec=5
PrivateTmp=true
NoNewPrivileges=true

[Install]
WantedBy=multi-user.target
EOF

# actualizare automata din GitHub (la 5 minute): urci fisierele in GitHub, serverul le ia singur
cat > /usr/local/bin/facturi-update <<EOF
#!/usr/bin/env bash
set -e
cd $APP_DIR
OLD=\$(git rev-parse HEAD)
sudo -u $SVC_USER git fetch -q origin $BRANCH
NEW=\$(git rev-parse origin/$BRANCH)
[ "\$OLD" = "\$NEW" ] && [ "\${1:-}" != "--force" ] && exit 0
sudo -u $SVC_USER git reset -q --hard origin/$BRANCH
if ! git diff --quiet "\$OLD" "\$NEW" -- cleaning-app/requirements.txt || [ "\${1:-}" = "--force" ]; then
  sudo -u $SVC_USER $APP_DIR/venv/bin/pip install -q -r cleaning-app/requirements.txt
fi
systemctl restart facturi
echo "Actualizat: \${OLD:0:7} -> \${NEW:0:7}"
EOF
chmod 755 /usr/local/bin/facturi-update
cat > /etc/systemd/system/facturi-update.service <<'EOF'
[Unit]
Description=Actualizare aplicatie facturi din GitHub
[Service]
Type=oneshot
ExecStart=/usr/local/bin/facturi-update
EOF
cat > /etc/systemd/system/facturi-update.timer <<'EOF'
[Unit]
Description=Verifica GitHub la 5 minute
[Timer]
OnBootSec=2min
OnUnitActiveSec=5min
[Install]
WantedBy=timers.target
EOF

say "6/7 HTTPS (Caddy) si firewall"
cat > /etc/caddy/Caddyfile <<EOF
$DOMAIN {
	encode gzip
	reverse_proxy 127.0.0.1:8000
}
EOF
# imaginile Ubuntu de la Oracle blocheaza tot in afara de SSH: deschidem 80 si 443
for p in 80 443; do
  if ! iptables -C INPUT -p tcp --dport $p -m state --state NEW -j ACCEPT 2>/dev/null; then
    n=$(iptables -L INPUT --line-numbers -n | awk '/REJECT/ {print $1; exit}')
    if [ -n "$n" ]; then iptables -I INPUT "$n" -p tcp --dport $p -m state --state NEW -j ACCEPT
    else iptables -A INPUT -p tcp --dport $p -m state --state NEW -j ACCEPT; fi
  fi
done
netfilter-persistent save >/dev/null 2>&1 || true

say "7/7 Pornire"
systemctl daemon-reload
systemctl enable -q --now facturi.service facturi-update.timer
systemctl restart facturi.service
systemctl enable -q caddy && systemctl restart caddy
sleep 6
if curl -fsS --max-time 10 http://127.0.0.1:8000/healthz >/dev/null; then
  say "Gata! Aplicatia merge."
else
  warn "Aplicatia nu raspunde inca. Vezi ce scrie: sudo journalctl -u facturi -n 50"
fi
cat <<EOF

  Adresa:   https://$DOMAIN
  (prima deschidere poate dura ~1 minut: Caddy cere certificatul HTTPS)

  Daca pagina nu se deschide: in Oracle Cloud → reteaua serverului (Security List)
  trebuie sa fie deschise porturile 80 si 443 (vezi pasii primiti).

  Comenzi utile:
    sudo journalctl -u facturi -f          jurnalul aplicatiei (Ctrl+C iese)
    sudo systemctl restart facturi         repornire
    sudo facturi-update --force            actualizare imediata din GitHub
    curl -fsSL https://raw.githubusercontent.com/AlexVladX/cleaning-app/main/cleaning-app/oracle-setup.sh | sudo bash
                                           reconfigurare (parola, chei, domeniu)
EOF
