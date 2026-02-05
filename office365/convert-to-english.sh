#!/bin/bash
# -----------------------------------------------------------------------------
# Microsoft Office 365/2016 English Language Switcher for PlayOnLinux/Wine
# -----------------------------------------------------------------------------
# This script forces the Office UI to use English (ID 1033) by modifying the 
# Windows Registry. It only works if the English language files are already 
# present in the installation (e.g. from a multi-language installer).
# If not, download from here, install with this wine and wineprefix, even it gives error, then run this script.
# https://support.microsoft.com/en-gb/office/install-the-english-language-pack-for-32-bit-office-94ba2e0b-638e-4a92-8857-2cb5ac1d8e17
# -----------------------------------------------------------------------------

# --- Configuration ---
# Update these paths if sharing with others who use different paths
export PREFIX="${1:-$HOME/.PlayOnLinux/wineprefix/office365}"
export WINE_BIN="${2:-$HOME/.PlayOnLinux/wine/linux-x86/winecx/bin/wine}"

# --- Checks ---
if [ ! -d "$PREFIX" ]; then
    echo "Error: Wine prefix not found at: $PREFIX"
    echo "Usage: ./convert-to-english.sh [PREFIX_PATH] [WINE_PATH]"
    exit 1
fi

if [ ! -f "$WINE_BIN" ]; then
    echo "Error: Wine binary not found at: $WINE_BIN"
    exit 1
fi

echo "Checking for English language files (ID 1033)..."
ENGLISH_FOUND=$(find "$PREFIX/drive_c/Program Files/Microsoft Office" -type d -name "1033" | head -n 1)

if [ -z "$ENGLISH_FOUND" ]; then
    echo "Error: English language files (folder '1033') not found in this installation."
    echo "You must install the English Language Pack first."
    exit 1
else
    echo "Found English files at: $ENGLISH_FOUND"
fi

# --- Create Temporary Registry File ---
echo "Generating registry configuration..."
cat <<EOF > /tmp/office_english_fix.reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\LanguageResources]
"UILanguage"=dword:00000409
"HelpLanguage"=dword:00000409
"FollowSystemUI"=dword:00000000

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\LanguageResources\EnabledLanguages]
"1033"="On"

[HKEY_LOCAL_MACHINE\Software\Microsoft\Office\16.0\Common\LanguageResources]
"UILanguage"=dword:00000409
"HelpLanguage"=dword:00000409
EOF

# --- Apply Registry Fix ---
echo "Applying registry settings..."
export WINEPREFIX="$PREFIX"

# Kill running processes to ensure registry unlocks (optional but safer)
# wineserver -k 

"$WINE_BIN" regedit /tmp/office_english_fix.reg

if [ $? -eq 0 ]; then
    echo "--------------------------------------------------------"
    echo "SUCCESS: Office language set to English (US)."
    echo "Registry file imported successfully."
    echo "--------------------------------------------------------"
else
    echo "Error: Failed to apply registry settings."
    exit 1
fi

rm /tmp/office_english_fix.reg
