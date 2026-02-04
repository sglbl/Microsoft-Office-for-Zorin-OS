# Microsoft Office 2016 Installation Script for Linux (Wine/PlayOnLinux)

A complete automated installation script for Microsoft Office 2016 on Linux systems using Wine and PlayOnLinux.

## Step 1: Download Office 2016 

Download the offline version of Office 2016 (Word, Excel Powerpoint and OneNote) HomeBusinessRetail from the following link:  
[https://massgrave.dev/office_c2r_links#english-en-us](https://massgrave.dev/office_c2r_links#english-en-us)

## Step 2: Install Dependencies

```bash
(sudo apt update &&
sudo apt-get install wine:i386 wine-stable:i386 winetricks -y &&
sudo apt install smbclient samba samba-common winbind -y &&
sudo apt install python-wxtools python3-pyasyncore -y &&
sudo apt install libncurses6:i386 playonlinux -y)

# Fix .NET installation issue by adjusting ptrace_scope temporarily
echo 0 | sudo tee /proc/sys/kernel/yama/ptrace_scope

# Create a symlink for libncurses
sudo ln -s /usr/lib/i386-linux-gnu/libncurses.so.6 /usr/lib/i386-linux-gnu/libncurses.so.5
sudo ln -s /usr/lib/i386-linux-gnu/libtinfo.so.6 /usr/lib/i386-linux-gnu/libtinfo.so.5
```

The commands are grouped with parentheses to ensure that the script will exit if any command fails.

> Mount the installation iso/img and continue with the next step.

## Step 3: Import the PlayOnLinux Script

This step involves importing the `office16.pol` script into PlayOnLinux to automate the installation.

### How to Import the Script

1.  Open PlayOnLinux from terminal using (make sure you're not in a virtual env.): 
    ```code
    GTK_THEME=ZorinBlue-Light playonlinux
    ```
2.  Go to `Tools` > `Run a local script`.
3.  Select the `office16.pol` file from the `office16` directory.
4.  Follow the on-screen instructions provided by the script. (Accept mono depedendency installations and if you get .NET errors during installation, you can skip those and continue by clicking on next button.)

### What the `office16.pol` Script Does

The script automates the installation and configuration of Microsoft Office 2016. Here's a summary of its actions:

*   **Creates a Windows 7 environment using 32-bit Wine 5.8 :** It sets up an isolated environment for Office.
*   **Installs dependencies:** It installs necessary components like `msxml6`, `riched20`, and `dotnet45` and add overrides libraries as native (windows) and then builtin (wine).
*   **Runs the installer:** It launches the Office 2016 installer.
*   **Applies fixes:** It copies necessary DLLs (`AppvIsvSubsystems32`, `C2R32`, `sppc` and `sppcs`) to the required (office or system32) folders to ensure compatibility.
*   **Creates shortcuts:** It creates desktop shortcuts for Word, Excel, PowerPoint and OneNote.
<!-- 
### Post Installation

In the Main Play on Linux (POL) screen, click Configure -> Office16 -> Wine Tab -> Registry Editor:
- Under HKEY_CURRENT_USER\Software\Wine, create the key `Direct2D` by right click.
- Under Direct2D, create DWORD `max_version_factory` and keep it as 0 (default) -->


### Add OneNote Support

https://dn721300.ca.archive.org/0/items/windows_xp_files/windows_xp_tablet_edition_sdk_17.exe

For OneNote, you have to download this and install it from:
In the Main Play on Linux (POL) screen, click Configure -> Office16 -> Miscellaneous Tab -> Run a exe.
After the installation, OneNote will work. When you open the app, right click on File and unselect Collapse the Ribbon.

<details>
<summary><b>Extras</b></summary>

### Fix "Multiple Open With Options" Issue
If you have installed Office multiple times, you might see duplicate entries in the "Open With" menu (e.g., 2 "Microsoft PowerPoint" options). This is because Wine creates file associations that persist even after deleting a prefix.

To remove the old "zombie" shortcuts, run the following command in your terminal. 

**Note:** Replace `office16` with the name of your **OLD** prefix that you want to clean up.

```bash
grep -l "office16" ~/.local/share/applications/wine-extension-*.desktop | xargs rm
```

## Known Issues

- Outlook is disabled due to browser compatibility issues
- Hardware acceleration is disabled to prevent graphics crashes
- Some .NET errors during installation are normal and can be skipped
- Multi-monitor setup: <del>The applications might only work on the number of displays present during installation</del> (Fixed with disabling hardware acc. regedit)


## License

This project is licensed under the MIT License with Attribution Requirement - see the [LICENSE](LICENSE) file for details.

</details>

