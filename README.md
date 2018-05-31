# KMS-VL-ALL-CMD
Online/Offline KMS Activator for Microsoft Windows/Office/VisualStudio VL Products

How to use ?
For KMS Activation of Volume Licensed Products:
1. Run KMS-VL-ALL.cmd as Administrator.
2. Done.
For VL conversion of Office 2016 C2R:
1. Run Convert-Office2016C2R.cmd as Administrator.
2. Done.

Defaults:
Manual mode/No Auto-renewal Task (_Task=0); For auto-renewal task to be installed change it to _Task=1
Offline mode (_OfflineMode=1); For Online mode change it to _OfflineMode=0
Non-Debug mode (_Debug=0); For debug mode change it to _Debug=1
Custom ePIDs (_RandomLevel=0); For random ePIDs change it to _RandomLevel=1 or 2

Install Auto-Renewal Task:
- Open KMS-VL-ALL.cmd in Notepad, change _Task=0 to _Task=1, save and run.
- Keep KMS-VL-ALL folder in the same location
- Exclude the folder in Anti-Virus software

Did not Work ?
- Open KMS-VL-ALL.cmd in Notepad
- Change _Debug=0 to _Debug=1
- Run as Admin
- Copy Paste the log file contents in CODE tags or upload the log file

Supported Microsoft Products:
(32-bit and 64-bit)
Windows Vista/7/8/8.1/10 (v1709 RS3) All VL Editions
Windows Server 2008/2008R2/2012/2012R2/2016 (v1709 RS3) All Editions
Office 2010 Family on Windows XP SP3 or later
Office 2013 Family on Windows 7 or later
Office 2016 Family on Windows 7 SP1 or later
Visual Studio 2013 Ultimate
Visual Studio 2015 Enterprise
Visual Studio 2017 Enterprise

Retail/OEM/MAK Genuine Activations are UNAFFECTED and Converts Notice Period/OOBE-Grace period windows to VL IF they are supported and are then activated.

Supported Retail/MAK Unactivated Editions:
Win Vista(Business/Enterprise) 7/8/8.1/10 Pro Retail/MAK and their Enterprise editions
Server 2008/2008 R2/2012/2012R2/2016 Retail, MAK editions
Office 2010/2013/2016 MAK editions
Office 2016 C2R edition

Credits (In chronological order) :
- ZWT, nosferati87, CODYQX4, letsgoawayhell, Phazor, mikmik38, deagles, FreeStyler, ColdZero, Hotbird64 and everyone who contributed to KMS Server emulator development
- MasterDisaster for original script
- abbodi1406 for Huge Contribution and Co-Authoring
- s1ave77 for batch help
- os51 for Retail/MAK checks examples
- cynecx and qad for DLL Injection method and KMS Integration
- xinso for interesting ideas
- ratzlefatz for testing and helping with various aspects of the script
- Tito for general support, VS and SQL support idea (partially done)
- LostED For Mirrors
- csihcs for code improvements
- l33tisw00t for user support
- moo807 for constant updates to missing GVLKs
and MDL Community for feedback and bug reports.

Changelog:
7.0RC [2018-03-28]
- Modified firewall rules (Thanks to abbodi1406)
- Office product check bug fixes (Thanks to abbodi1406)
- Office 2019 Professional Plus C2R keys added (Thanks to abbodi1406)
- OneDrive for Business 2013 key added (Thanks to abbodi1406)
- SharePoint Designer 2010/2013 Retail keys removed (Thanks to abbodi1406)
- Windows Server 2016 ARM64 key added (Thanks to abbodi1406)
- Changed default to Manual mode due to popular vote
- Fixed VS correct OS detection bug (Thanks to asfomp)
- replaced xcopy with robocopy (Thanks to rpo), but kept DEL lines for now, as Manual Users do not require them, need to think about merits and demerits of this.
- Minor Cosmetics
