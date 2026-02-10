#!/bin/bash

# Script by @sglbl & yt-formatendo and zip by yt-formateando
set -e
set -o pipefail

echo "==============================================="
echo "  Office 365 Auto Installer - PlayOnLinux/WineCX"
echo "==============================================="

# ---------------------------------------------------------
# 1) Setup Paths and Variables
# ---------------------------------------------------------
# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
DOWNLOADS_DIR="$SCRIPT_DIR"

POL_ROOT="$HOME/.PlayOnLinux"
WINE_VERSIONS_DIR="$POL_ROOT/wine/linux-x86"
WINECX_DIR="$WINE_VERSIONS_DIR/winecx"
WINE_BIN="$WINECX_DIR/bin"
PREFIX_DIR="$POL_ROOT/wineprefix/office365"

echo "Script location: $SCRIPT_DIR"
echo "Target Prefix: $PREFIX_DIR"
echo "Target Wine: $WINECX_DIR"

# ---------------------------------------------------------
# 2) Enable 32-bit architecture (requires sudo)
# ---------------------------------------------------------
echo "Enabling 32-bit architecture..."
sudo dpkg --add-architecture i386
sudo apt update

# ---------------------------------------------------------
# 3) Install Dependencies
# ---------------------------------------------------------
echo "Installing dependencies..."
sudo apt install -y unzip || true
sudo apt install -y build-essential gcc-multilib g++-multilib flex bison || true
sudo apt install -y git wget curl pkg-config gettext || true
sudo apt install -y cups-daemon cups-client printer-driver-all system-config-printer cups-pdf printer-driver-cups-pdf || true
sudo apt install -y msitools || true
sudo apt install -y clang lld || true
sudo apt install -y libc6:i386 libgcc1:i386 libstdc++6:i386 || true
sudo apt install -y libfreetype6:i386 libx11-6:i386 libxext6:i386 libxrender1:i386 libxrandr2:i386 || true
sudo apt install -y winbind samba-common samba-libs gnutls-bin || true
sudo apt install -y ttf-mscorefonts-installer || true
sudo apt install -y wine32:i386 winetricks playonlinux  || true

# ---------------------------------------------------------
# 4) Check for installation files
# ---------------------------------------------------------
cd "$DOWNLOADS_DIR"

if [ -d "MSO365" ]; then
    echo "Found MSO365 folder."
elif [ -f "MSO365.zip" ]; then
    echo "Unzipping MSO365.zip..."
    unzip -o MSO365.zip
else
    echo "Error: MSO365.zip or MSO365 folder not found in $DOWNLOADS_DIR"
    exit 1
fi

# ---------------------------------------------------------
# 3) Enter decompressed content
# ---------------------------------------------------------
cd "$DOWNLOADS_DIR/MSO365"

# ---------------------------------------------------------
# 6) Install WineCX to PlayOnLinux folder
# ---------------------------------------------------------
if [ ! -d "$WINECX_DIR" ]; then
    echo "Installing WineCX to $WINECX_DIR..."
    DEB_PATH=""
    if [ -f "$DOWNLOADS_DIR/winecx.deb" ]; then
        DEB_PATH="$DOWNLOADS_DIR/winecx.deb"
    elif [ -f "$DOWNLOADS_DIR/MSO365/winecx.deb" ]; then
        DEB_PATH="$DOWNLOADS_DIR/MSO365/winecx.deb"
    elif [ -f "../winecx.deb" ]; then
        DEB_PATH="../winecx.deb"
    fi

    if [ -n "$DEB_PATH" ]; then
        mkdir -p temp_winecx
        dpkg -x "$DEB_PATH" temp_winecx
        mkdir -p "$WINE_VERSIONS_DIR"
        # The deb extracts to opt/winecx. We move that folder to become our target.
        if [ -d "temp_winecx/opt/winecx" ]; then
             mv temp_winecx/opt/winecx "$WINECX_DIR"
             echo "WineCX installed to $WINECX_DIR"
        else
             echo "Error: unexpected deb structure in winecx.deb"
        fi
        rm -rf temp_winecx
    else
        echo "Warning: winecx.deb not found. Assuming WineCX is installed."
    fi
else
    echo "WineCX already exists at $WINECX_DIR"
fi

# ---------------------------------------------------------
# 7) Setup Prefix
# ---------------------------------------------------------
echo "Setting up prefix at $PREFIX_DIR..."
mkdir -p "$PREFIX_DIR"

if [ -d ".Microsoft_Office_365" ]; then
    # Copy contents enabling overwrite
    cp -r .Microsoft_Office_365/. "$PREFIX_DIR/"
else
    echo "Error: .Microsoft_Office_365 folder not found inside MSO365."
    # exit 1 ? We might want to continue if it was already there?
fi

# ---------------------------------------------------------
# 8) Install Icons
# ---------------------------------------------------------
echo "Installing application icons..."
ICON_DIR="$HOME/.local/share/icons/hicolor/256x256/apps"
mkdir -p "$ICON_DIR"
cp "$DOWNLOADS_DIR/MSO365/Office2016Icons/"*365.svg "$ICON_DIR/"
gtk-update-icon-cache "$HOME/.local/share/icons/hicolor/" || true

# ---------------------------------------------------------
# 9) Create Launchers Folder
# ---------------------------------------------------------
LAUNCHERS_DIR="$PREFIX_DIR/launchers"
mkdir -p "$LAUNCHERS_DIR"
chmod 755 "$LAUNCHERS_DIR"

# ---------------------------------------------------------
# 10) Function to create prefix launchers & desktop files
# ---------------------------------------------------------
create_launcher() {
local name="$1"
local exe="$2"
local fancy_name="$3"
local icon_name="$4"
local categories="$5"
local mime_type="$6"

echo "Creating launcher for $name..."

# Shell script wrapper
cat > "$LAUNCHERS_DIR/${name}.sh" <<EOF
#!/bin/bash
set -e
export PATH="$WINE_BIN:\$PATH"
export WINEPREFIX="$PREFIX_DIR"
export LANG=C.UTF-8
export WINEDEBUG=-all

app="C:\\\\Program Files\\\\Microsoft Office\\\\root\\\\Office16\\\\${exe}"
"$WINE_BIN/wineserver" -p >/dev/null 2>&1 || true

if [ \$# -eq 0 ]; then
    exec "$WINE_BIN/wine" "\$app"
else
    for file in "\$@"; do
        fullpath=\$(realpath "\$file")
        winpath="Z:\${fullpath//\//\\\\}"
        "$WINE_BIN/wine" "\$app" "\$winpath"
    done
fi
EOF

chmod +x "$LAUNCHERS_DIR/${name}.sh"

# Desktop Entry
DESKTOP_FILE="$HOME/.local/share/applications/${name}.desktop"
cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Name=$fancy_name
Comment=$fancy_name (Office 365)
Exec="$LAUNCHERS_DIR/${name}.sh" %F
Type=Application
StartupNotify=true
Terminal=false
Icon=$icon_name
Categories=$categories
MimeType=$mime_type
EOF

}

# ---------------------------------------------------------
# 11) Create Launchers
# ---------------------------------------------------------
create_launcher "word365" "WINWORD.EXE" "Microsoft Word 365" "Word365" "Office;WordProcessor;" "application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;application/vnd.ms-word.document.macroEnabled.12;application/rtf;text/plain;"

create_launcher "excel365" "EXCEL.EXE" "Microsoft Excel 365" "Excel365" "Office;Spreadsheet;" "application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.ms-excel.sheet.macroEnabled.12;text/csv;"

create_launcher "powerpoint365" "POWERPNT.EXE" "Microsoft PowerPoint 365" "Powerpoint365" "Office;Presentation;" "application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;application/vnd.ms-powerpoint.presentation.macroEnabled.12;"

create_launcher "outlook365" "OUTLOOK.EXE" "Microsoft Outlook 365" "Outlook365" "Office;Email;" "application/vnd.ms-outlook;application/mbox;message/rfc822;"

create_launcher "access365" "MSACCESS.EXE" "Microsoft Access 365" "Access365" "Office;Database;" "application/vnd.ms-access;application/x-msaccess;"

create_launcher "publisher365" "MSPUB.EXE" "Microsoft Publisher 365" "Publisher365" "Office;Publishing;" "application/x-mspublisher;"


# ---------------------------------------------------------
# 12) Install Main App Launcher (Menu)
# ---------------------------------------------------------
echo "Installing main launcher..."
mkdir -p "$HOME/bin"
# Copy from SCRIPT_DIR/office365-launcher.sh
if [ -f "$SCRIPT_DIR/office365-launcher.sh" ]; then
    cp "$SCRIPT_DIR/office365-launcher.sh" "$HOME/bin/"
    chmod +x "$HOME/bin/office365-launcher.sh"
else
    echo "Warning: office365-launcher.sh not found in script directory."
fi

mkdir -p "$HOME/.local/share/applications"
if [ -f "$SCRIPT_DIR/office365-launcher.desktop" ]; then
    cp "$SCRIPT_DIR/office365-launcher.desktop" "$HOME/.local/share/applications/"
    # Update username in the desktop file
    sed -i "s|/home/sglbl|$HOME|g" "$HOME/.local/share/applications/office365-launcher.desktop"
else
    echo "Warning: office365-launcher.desktop not found in script directory."
fi

# Icon for main launcher
if [ -f "$SCRIPT_DIR/office.svg" ]; then
    mkdir -p "$HOME/.local/share/applications/icons"
    cp "$SCRIPT_DIR/office.svg" "$HOME/.local/share/applications/icons/office.svg"
fi

update-desktop-database "$HOME/.local/share/applications" || true

# ---------------------------------------------------------
# 13) Fix Perms
# ---------------------------------------------------------
chmod -R u+rwX "$PREFIX_DIR"

# ---------------------------------------------------------
# 14) Rebuild DOSDEVICES (PlayOnLinux Style)
# ---------------------------------------------------------
echo "Configuring dosdevices..."
rm -rf "$PREFIX_DIR/dosdevices"
mkdir -p "$PREFIX_DIR/dosdevices"

ln -s ../drive_c "$PREFIX_DIR/dosdevices/c:"
ln -s / "$PREFIX_DIR/dosdevices/z:"
ln -s /dev/null "$PREFIX_DIR/dosdevices/c::"
ln -s /dev/null "$PREFIX_DIR/dosdevices/z::"
ln -s /media "$PREFIX_DIR/dosdevices/d:"
ln -s "$HOME" "$PREFIX_DIR/dosdevices/e:"

# ---------------------------------------------------------
# 15) Crossover User Folders
# ---------------------------------------------------------
mkdir -p "$PREFIX_DIR/drive_c/users/crossover/AppData/Local"
mkdir -p "$PREFIX_DIR/drive_c/users/crossover/AppData/Roaming"

# ---------------------------------------------------------
# 16) Update Prefix
# ---------------------------------------------------------
echo "Updating prefix..."
WINEPREFIX="$PREFIX_DIR" "$WINE_BIN/wine" wineboot -u

# ---------------------------------------------------------
# 17) Restart Wine Services
# ---------------------------------------------------------
WINEPREFIX="$PREFIX_DIR" "$WINE_BIN/wineserver" -k
# WINEPREFIX="$PREFIX_DIR" "$WINE_BIN/wineserver" -w # -w means persistent, maybe just stick to -k to kill

# ---------------------------------------------------------
# 18) Copy Fonts
# ---------------------------------------------------------
echo "Copying fonts..."
mkdir -p "$PREFIX_DIR/drive_c/windows/Fonts"
if [ -d "$DOWNLOADS_DIR/MSO365/Fuentes Office365" ]; then
    cp "$DOWNLOADS_DIR/MSO365/Fuentes Office365/"* "$PREFIX_DIR/drive_c/windows/Fonts/" || true
fi

# ---------------------------------------------------------
# 19) Create PlayOnLinux Config
# ---------------------------------------------------------
echo "Creating playonlinux.cfg..."
cat > "$PREFIX_DIR/playonlinux.cfg" <<EOF
ARCH=x86
VERSION=winecx
OPEN_IN=xdg-open
WINEDEBUG=
EOF

echo "Installation Complete!"
