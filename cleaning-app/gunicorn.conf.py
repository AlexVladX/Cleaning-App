# Configurarea serverului (gunicorn o citeste singur din folderul aplicatiei).
# UN SINGUR proces: setarile (reguli, formule, coloane) sunt tinute in memoria procesului si
# salvate in Drive; cu 2+ procese fiecare ar avea propria copie si s-ar suprascrie una pe alta.
# (Render poate seta WEB_CONCURRENCY > 1; valoarea de aici o inlocuieste.)
workers = 1
threads = 4
timeout = 120
