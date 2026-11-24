IMPORTANT INFORMATION – PLEASE READ BEFORE INSTALLING

Please Note: 

This Setup Program uses Jordan Russell's Inno Setup.

What are the system requirements for the SeleniumVBA DLL?

64-bit MS Windows
32- or 64-bit MS Office (Excel and/or Access)

Where should the SeleniumVBA DLL be installed on my User Account?

This installer will install SeleniumVBA in any location that you select. It is recommended that it be installed in C:\Users\[user name]\AppData\Local, which is the default location, in order to prevent the DLL file from being inadvertently moved or deleted after it has been registered. Moving or deleting the DLL after installation without first properly uninstalling will break any code that references the DLL, and render it unusable.

The installer will place a SeleniumVBA shortcut on your Desktop to provide easy access to the installation folder, uninstaller file, readme instructions, and example resources.  

What happens if I upgrade to a new version, or reinstall to a new location?

If upgrading to a new version, or re-installing with the same version but choosing a different install location, this installer will first unregister and then remove the older version.

How do I uninstall the SeleniumVBA DLL from my system?

To uninstall SeleniumVBA from your system, you can run the unins000.exe file in the installation directory. This unregisters the SeleniumVBA DLL and removes all the original files from your system.

Once I have installed the SeleniumVBA DLL, what is the quickest way to get started?

Use the SeleniumVBA shortcut of your Desktop to navigate to the examples folder in the installation directory. There is an MS Excel macro file and an MS Access database file that are loaded with examples. You can move or copy these example files (or the entire example directory) to another location on your system. But be sure NOT to move any other files, such as the DLL, ini settings file, and/or the uninstaller file.

Open one of the example documents and go to the Visual Basic Editor. Select Tools>References and scroll down the list of references until you see SeleniumVBA. Click to set the checkmark and then hit OK from the References dialog. Save the file, and then run any of the test procedures. Once the file is saved, this step will not have to be repeated.
