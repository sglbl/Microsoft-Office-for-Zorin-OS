#!/bin/bash

# Office 365 Launcher Script
# This script provides a menu to launch Office 365 applications

PREFIX_DIR="$HOME/.PlayOnLinux/wineprefix/office365"
LAUNCHERS_DIR="$PREFIX_DIR/launchers"

choice=$(GTK_THEME=ZorinGreen-Light zenity --list \
    --title="Office 365 Launcher" \
    --text="Select an Office 365 application:" \
    --column="Application" \
    --height=310 \
    --width=350 \
    "Word" \
    "Excel" \
    "PowerPoint")

case $choice in
    "Word")
        "$LAUNCHERS_DIR/word365.sh" &
        ;;
    "Excel")
        "$LAUNCHERS_DIR/excel365.sh" &
        ;;
    "PowerPoint")
        "$LAUNCHERS_DIR/powerpoint365.sh" &
        ;;
    *)
        # If cancelled or invalid choice, exit
        exit 0
        ;;
esac
