### Solution Summary:

This repository folder makes available 3 different SeleniumVBA solutions, depending on user requirements:

- **An MS Excel Workbook called SeleniumVBA.xlsm.** This file contains all of the source and test modules and can optionally be changed to an Addin (.xlam) so that it can be called from another Excel workbook.
- **An MS Access database called SeleniumVBA.accdb.** Like the Excel version, this file contains all of the source and test modules and can optionally be called from another MS Access database.
- **An ActiveX DLL called SeleniumVBA_win64.dll/SeleniumVBA_win32.dll.** This DLL can be installed and registered using the SeleniumVBADLLSetup.exe setup program. Once installed, the SeleniumVBA code library can be referenced by your VBA projects in either MS Excel, MS Access, or MS VBScript to expose the SeleniumVBA object model without having to manage the SeleniumVBA source code. The ActiveX DLL was compiled using the (in Beta) [twinBASIC](https://twinbasic.com) compiler.

All three solutions above allow for storing and referencing SeleniumVBA functionality from a centralized file location. This can be of use if the user does not intend to integrate the SeleniumVBA source directly into their own code and wants a convenient means of updating SeleniumVBA with newer versions of the code library.

Below is a table showing the compatibility for each solution with various versions of Office:

|Solution|<= Office 2007|Office 2010|Office 2013|Office 2016|Office 2019|Office 365|
| ---------------- | ------------- | ------------- |------------- |------------- |------------- |------------- |
|Excel Workbook|Not|32/64-bit|32/64-bit|32/64-bit|32/64-bit|32/64-bit|
|Access DB|Not|32/64-bit|32/64-bit|32/64-bit|32/64-bit|32/64-bit|
|ActiveX DLL*|32-bit**|32/64-bit|32/64-bit|32/64-bit|32/64-bit|32/64-bit|

*the [twinBASIC](https://twinbasic.com) ActiveX DLL can be called from MS VBScript, as well as MS Excel and MS Access

**only limited testing

### Excel Workbook and Access DB Installation:

The Excel and Access solutions are self-contained - they include both source code and test routines. However, it is also possible to reference these solutions externally from another Excel/Access document.

In cases where the intent is to run SeleniumVBA from multiple Workbooks and the user cannot or wishes not to install the DLL solution, it may make sense to convert the Workbook solution to an Addin. To change the Excel Workbook (.xlsm) to an Addin (.xlam), open the workbook, go to the Microsoft VBA IDE. In the Project Viewer, click on ThisWorkbook, and then scroll down to the IsAddin property in the Properties Window. Change the property value to True, and then save the Workbook to the Addin type of ".xlam". The resulting Addin code library can now be referenced from other Excel Workbooks. 

Instructions for referencing add-in versions of SeleniumVBA from another MS Excel/Access document:

1) Open your Excel (or Access) macro project that will reference the add-in (for testing, just copy-paste some of the macro examples provided in the test_* modules of SeleniumVBA.xlsm)
2) In the Visual Basic for Applications window, select a code module, then click on Tools tab, References.
3) On the References Dialog, click on Browse, select Microsoft Excel Files (or Microsoft Access Files) as File Type, then browse to the add-in folder location and select the add-in.
4) Save the Excel (or Access) macro project.

### ActiveX DLL Installation:

For the [twinBASIC](https://twinbasic.com) ActiveX DLL, more detailed instructions on how to install and use the DLL will be presented during installation. The setup program, which was compiled using Inno Setup, will install and register the DLL, and copy test Excel, Access, and VBScript documents to the installation folder.

The [twinBASIC](https://twinbasic.com) ActiveX DLL solution requires no dependencies (such as .Net Framework).

### Advanced Customization - SeleniumVBA.ini File:

There is an optional settings file in this repo folder called SeleniumVBA.ini. It allows for [advanced customization](https://github.com/GCuser99/SeleniumVBA/wiki#advanced-customization) of SeleniumVBA, if the user has the need. This file only takes effect if it is located in the same folder as any of the three solutions above. The user can edit this file to alter the way SeleniumVBA works without having to build runtime code for customization.

