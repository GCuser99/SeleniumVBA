The add-in version of SeleniumVBA allows for storing and referencing SeleniumVBA functionality from a centralized file location. This might be of use if the user does not intend to integrate the SeleniumVBA classes directly into their own code and, wants a convenient means of updating SeleniumVBA with newer versions.

Instructions for setting up the Add-in version of SeleniumVBA:

1) Unzip/copy the seleniumVBA.xlam add-in into a folder that is accessible to all VBA projects that will reference the add-in.
2) Open your Excel macro project that will reference the add-in (for testing, just copy-paste in it some of the macro examples provided in the test_* modules of seleniumVBA.xlam)
3) In the Visual Basic for Applications window, select a code module, then click on Tools tab, References.
4) On the References Dialog, click on Browse, select Microsoft Excel Files as File Type, then browse to the add-in folder location and select the add-in.
5) Save the Excel macro project.

See the provided test_addin.xlsm for code usage cases.