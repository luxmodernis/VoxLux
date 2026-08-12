#!/bin/bash
# ═══════════════════════════════════════════════════════════════
#  VoxLux — Installateur
#  Installe le moteur dans ~/Library/Application Support/VoxLux
#  et crée une vraie application « VoxLux.app » dans Applications.
# ═══════════════════════════════════════════════════════════════

SRC_DIR="$(cd "$(dirname "$0")" && pwd)"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
BLUE='\033[0;34m'; BOLD='\033[1m'; NC='\033[0m'

clear
echo ""
echo -e "${BOLD}  VoxLux — Installation${NC}"
echo "═══════════════════════════════════════════"
echo ""

# ── Emplacements ─────────────────────────────────────────────
APP_SUPPORT="$HOME/Library/Application Support/VoxLux"
mkdir -p "$APP_SUPPORT"
cp "$SRC_DIR/VoxLux.py" "$APP_SUPPORT/VoxLux.py"
xattr -cr "$APP_SUPPORT" 2>/dev/null

# ── 1. Vérifier Python ───────────────────────────────────────
echo -e "${BLUE}[1/5]${NC} Vérification de Python…"

PYTHON=""
for cmd in python3.12 python3.11 python3.10 python3.9 python3; do
    if command -v "$cmd" &>/dev/null; then
        VERSION=$("$cmd" --version 2>&1 | awk '{print $2}')
        MAJOR=$(echo "$VERSION" | cut -d. -f1)
        MINOR=$(echo "$VERSION" | cut -d. -f2)
        if [ "$MAJOR" -ge 3 ] && [ "$MINOR" -ge 9 ]; then
            PYTHON="$cmd"
            echo -e "    ${GREEN}✅ Python $VERSION détecté${NC}"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo -e "    ${RED}❌ Python 3.9+ introuvable.${NC}"
    echo ""
    echo "    Installez Python depuis : https://www.python.org/downloads/macos/"
    echo "    Puis relancez cet installateur."
    echo ""
    read -p "    Appuyez sur Entrée pour fermer…"
    exit 1
fi

# ── 2. Environnement virtuel ─────────────────────────────────
echo ""
echo -e "${BLUE}[2/5]${NC} Création de l'environnement VoxLux…"

VENV_DIR="$APP_SUPPORT/.voxlux_env"

if [ -d "$VENV_DIR" ]; then
    echo -e "    ${GREEN}✅ Environnement existant réutilisé${NC}"
else
    "$PYTHON" -m venv "$VENV_DIR"
    echo -e "    ${GREEN}✅ Environnement créé${NC}"
fi

PYTHON_ENV="$VENV_DIR/bin/python3"
PIP_ENV="$VENV_DIR/bin/pip"

# ── 3. Librairies Python (ffmpeg fourni par imageio-ffmpeg) ──
echo ""
echo -e "${BLUE}[3/5]${NC} Installation des librairies Python…"
echo "    (5 à 10 minutes selon la connexion)"
echo ""

"$PIP_ENV" install --quiet --upgrade pip
"$PIP_ENV" install --quiet \
    "openai-whisper" \
    "flask>=2.1" \
    "deep-translator" \
    "openpyxl" \
    "torch" \
    "tqdm" \
    "imageio-ffmpeg"

if [ $? -ne 0 ]; then
    echo -e "    ${RED}❌ Erreur lors de l'installation.${NC}"
    echo "    Vérifiez votre connexion internet et relancez l'installateur."
    read -p "    Appuyez sur Entrée pour fermer…"
    exit 1
fi
echo -e "    ${GREEN}✅ Librairies installées${NC}"

# ── 4. Modèle Whisper ────────────────────────────────────────
echo ""
echo -e "${BLUE}[4/5]${NC} Téléchargement du modèle de transcription…"
echo "    (~500 Mo — une seule fois)"
echo ""

"$PYTHON_ENV" -c "
import whisper, sys
print('    Téléchargement en cours…', flush=True)
try:
    whisper.load_model('small')
    print('    Modèle prêt.', flush=True)
except Exception as e:
    print(f'    Erreur : {e}', flush=True)
    sys.exit(1)
"

if [ $? -ne 0 ]; then
    echo -e "    ${RED}❌ Erreur lors du téléchargement du modèle.${NC}"
    read -p "    Appuyez sur Entrée pour fermer…"
    exit 1
fi
echo -e "    ${GREEN}✅ Modèle Whisper prêt${NC}"

# ── 5. Création de l'application VoxLux.app ──────────────────
echo ""
echo -e "${BLUE}[5/5]${NC} Création de l'application VoxLux…"

# Cible : /Applications si accessible en écriture, sinon ~/Applications
if [ -w "/Applications" ]; then
    APP_DEST="/Applications/VoxLux.app"
    APP_LOCATION="Applications"
else
    mkdir -p "$HOME/Applications"
    APP_DEST="$HOME/Applications/VoxLux.app"
    APP_LOCATION="$HOME/Applications"
fi

rm -rf "$APP_DEST"
mkdir -p "$APP_DEST/Contents/MacOS"
mkdir -p "$APP_DEST/Contents/Resources"

# Info.plist
cat > "$APP_DEST/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key><string>VoxLux</string>
    <key>CFBundleDisplayName</key><string>VoxLux</string>
    <key>CFBundleIdentifier</key><string>com.voxlux.app</string>
    <key>CFBundleVersion</key><string>1.0</string>
    <key>CFBundleShortVersionString</key><string>1.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>CFBundleExecutable</key><string>VoxLux</string>
    <key>CFBundleIconFile</key><string>VoxLux</string>
    <key>LSMinimumSystemVersion</key><string>10.13</string>
    <key>NSHighResolutionCapable</key><true/>
</dict>
</plist>
PLIST

# Exécutable de l'app. Points clés :
#  • Architecture native forcée : lancé par le Finder, le python universel
#    peut démarrer en x86_64 (Rosetta) et ne pas charger les libs arm64.
#  • Serveur démarré en arrière-plan DÉTACHÉ (l'app se ferme, le serveur vit) ;
#    s'il tourne déjà, on rouvre juste le navigateur (reclic = pas de conflit
#    de port). Arrêt propre via le bouton « Quitter » de l'interface.
cat > "$APP_DEST/Contents/MacOS/VoxLux" << LAUNCH
#!/bin/bash
export PATH="/opt/homebrew/bin:/usr/local/bin:\$PATH"
export VOXLUX_NO_AUTOOPEN=1
PY="$APP_SUPPORT/.voxlux_env/bin/python3"
VX="$APP_SUPPORT/VoxLux.py"
LOG="$APP_SUPPORT/voxlux.log"
URL="http://127.0.0.1:7860"

# Déjà lancé ? → rouvrir simplement le navigateur
if curl -s -o /dev/null -m 1 "\$URL"; then
    open "\$URL"
    exit 0
fi

# Démarrer le moteur en arrière-plan détaché (architecture native)
if [ "\$(sysctl -n hw.optional.arm64 2>/dev/null)" = "1" ]; then
    nohup arch -arm64 "\$PY" "\$VX" </dev/null >"\$LOG" 2>&1 &
else
    nohup "\$PY" "\$VX" </dev/null >"\$LOG" 2>&1 &
fi

# Attendre que le serveur réponde puis ouvrir le navigateur
for i in \$(seq 1 60); do
    sleep 1
    if curl -s -o /dev/null -m 1 "\$URL"; then open "\$URL"; break; fi
done
exit 0
LAUNCH
chmod +x "$APP_DEST/Contents/MacOS/VoxLux"

# Icône
if [ -f "$SRC_DIR/VoxLux.icns" ]; then
    cp "$SRC_DIR/VoxLux.icns" "$APP_DEST/Contents/Resources/VoxLux.icns"
fi

# Nettoyer la quarantaine sur l'app créée + rafraîchir l'icône
xattr -cr "$APP_DEST" 2>/dev/null
touch "$APP_DEST"

echo -e "    ${GREEN}✅ VoxLux installé dans $APP_LOCATION${NC}"

# ── Fin ──────────────────────────────────────────────────────
echo ""
echo "═══════════════════════════════════════════"
echo -e "${GREEN}${BOLD}  VoxLux est installé !${NC}"
echo "═══════════════════════════════════════════"
echo ""
echo -e "  VoxLux est dans vos ${BOLD}Applications${NC} (et dans le Launchpad)."
echo "  Lancez-le comme n'importe quelle app, d'un double-clic sur son icône."
echo ""

read -p "  Ouvrir VoxLux maintenant ? (o/n) : " REP
if [[ "$REP" =~ ^[oOyY] ]]; then
    open "$APP_DEST"
fi
