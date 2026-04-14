@echo off
REM Access DevCon Vienna 2026 — Install GBLauncher Demo
REM Downloads, installs, and creates a desktop shortcut for the GBLauncher demo app
REM
REM Prerequisites: Microsoft Access (any version) must be installed

GBLauncher.exe --install --app gblauncher-demo --dropbox-share-link "https://www.dropbox.com/scl/fi/leykfo24dvbs7o2z7hmrb/gblauncher-demo.version?rlkey=d6kfonfhdruizh0xc1tb8x01z&st=nzlpsmmn&dl=0" --name "GBLauncher Demo"
pause
