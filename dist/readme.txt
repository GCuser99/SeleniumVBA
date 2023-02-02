SUMMARY:

This repository folder contains 3 different SeleniumVBA solutions, depending on user requirements:

- An MS Excel Add-in called SeleniumVBA.xlam. This file contains all of the source and test modules and can optionally be called from another Excel workbook.

- An MS Access database called SeleniumVBA.accdb Like the Excel version, this file contains all of the source and test modules and can optionally be called from another MS Access database.

- An EXPERIMENTAL ActiveX DLL called SeleniumVBA.dll. This DLL must be installed and registered using the SeleniumVBADLLSetup.exe Inno setup program. Once installed, the SeleniumVBA DLL can be referenced by your VBA projects in either MS Excel or MS Access to expose the SeleniumVBA object model without having to manage the SeleniumVBA source code. The ActiveX DLL was compiled using the (currently in Beta) twinBasic compiler. 

All three solutions above allow for storing and referencing SeleniumVBA functionality from a centralized file location. This might be of use if the user does not intend to integrate the SeleniumVBA classes directly into their own code and wants a convenient means of updating SeleniumVBA with newer versions of the code library.

NOTES:

Instructions for setting up the add-in versions of SeleniumVBA:

1) Unzip/copy the SeleniumVBA.xlam and/or SeleniumVBA.accdb into a folder that is accessible to all VBA projects that will reference the source library.
2) Open your Excel (or Access) macro project that will reference the add-in (for testing, just copy-paste some of the macro examples provided in the test_* modules of SeleniumVBA.xlam)
3) In the Visual Basic for Applications window, select a code module, then click on Tools tab, References.
4) On the References Dialog, click on Browse, select Microsoft Excel Files (or Microsoft Access Files) as File Type, then browse to the add-in folder location and select the add-in.
5) Save the Excel (or Access) macro project.

See the provided test_* modules for many code usage cases.

For the ActiveX DLL, more detailed instructions on how to install and use the DLL will be presented during installation. The setup program was compiled using Inno Setup. After installation, be aware that when it is first called during a VBA session, SeleniumVBA will display a twinBasic banner for 5 seconds. . Subsequent calls during the session will not show the banner.  

There is an optional settings file in this repo folder called SeleniumVBA.ini. It allows for advanced customization of SeleniumVBA, if the user has the need. This file only takes effect if it is in the same folder as any of the three solutions above. The user can edit this file to alter the way SeleniumVBA works without having to build runtime code for customization.
