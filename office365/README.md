# Office 365 Auto Installer

This directory contains resources to install Spanish version (by formateando) of Office 365  
If you prefer English language, please refer to `office16` folder or run the english conversion script after installation.

## How to Install

1. Download the zip from [Mediafire](https://www.mediafire.com/file/qo3lpo0wr3fnu28/MSO365.zip/file) or [Drive](https://drive.google.com/file/d/1M-xee0XswaPINOPwSxYA-FlK-_qdEc9N/view) and move the zip (without extracting) to this directory that contains the scripts.
2. Open a terminal.
3.  Navigate to this directory and and give execution access:
    ```bash
    cd office365
    chmod +x install-office365-playonlinux.sh
    ```
4.  Run the installation script:
    ```bash
    ./install-office365-playonlinux.sh
    ```
5. After installation completes, you can search for Office 365 (or word, excel, powerpoint, access..) on menu. The sign-in required apps have problems due to browser compatibility issues.

---

For other Windows features that Linux doesn't have by default, you can visit this repository:
https://github.com/sglbl/gnulinux-config/
- Clipboard (using Windows + v)
- Emoji (using Windows + .)
- Drag and drop items from Files onto apps on dash
- Auto connect to bluetooth device (headphones) on startup
- Feature to be able to view recent folders
- Capslock delay fix
- Open .url files
