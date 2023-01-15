This folder contains an EXPERIMENTAL SeleniumVBA ActiveX Dll created in twinBasic. Once registered, the Dll can be referenced from VBA in either MS Excel or MS Access and will expose the SeleniumVBA Object Model for use in your VBA code, without having to manage the SeleniumVBA source.

There are two ways that you can register the Dll:

Method 1 - Use Regserver32 on supplied Dll

To register for the first time:

1) Copy the Dll file (and optional SeleniumVBA.ini file) to a location of your choice
2) Open up a command terminal and CD to the location of the Dll
3) Register by entering into terminal: %systemroot%\System32\regsvr32.exe SeleniumVBA_win64.dll
4) Close the command terminal

To reference the Dll for first time after registration:

1) Open a document holding your VBA project 
2) Go to Tools>References; scroll down to and check SeleniumVBA
3) Save the document

To upgrade an already existing registered Dll:

1) Open up a command terminal and CD to the location of the Dll
2) Unregister by entering into terminal: %systemroot%\System32\regsvr32.exe /u SeleniumVBA_win64.dll
3) After unregistering (not before!), replace the old file with the new one
4) Register by entering into terminal: %systemroot%\System32\regsvr32.exe SeleniumVBA_win64.dll

Method 2 - Using the twinBasic compiler with the supplied SeleniumVBA project folder

1) Copy the entire SeleniumVBA folder found in this twinBasic repo folder to a location of your choice
2) Download the latest version of twinBasic compiler from https://github.com/twinbasic/twinbasic/releases
3) Open the compiler twinBASIC.exe
4) Click Cancel on the New Project pop-up
5) Go to Tools->IDE Options; then click on "Compiler: Start in 64-bit Mode"
6) On the IDE toolbar, make sure to select "win64" on the build configuration toggle (right side)
7) Go to File->New Project; choose "Import from folder..." and click Open
8) Navigate to and select the SeleniumVBA folder from step 1)
9) Go to File->Save project; locate and name the project to be saved
10) Go to File->Make/Build - you should see success messages in the immediate window
11) Exit compiler

The newly compiled and registered Dll should be found in a folder "Build" under folder where the project was saved.

You can optionally copy the SeleniumVBA.ini file in this repo folder to the same folder as the Dll.

If this is the first time you have compiled the Dll, then you will have to follow steps under "To reference the Dll for first time after registration" to use the Dll in your VBA project.

See the provided Excel and Access test files under test_documents folder for many code usage cases. Remember that you will need to add the Dll reference before using these.
