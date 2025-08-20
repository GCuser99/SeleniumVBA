This version is a full functionality version of SeleniumVBA for AV false-positive testing:

* Workbook protected with password "123"
* Both an .xlam add-in and .xlsm version



Note: To reference the add-in from another workbook:

1. Developer Tab -> Tools -> References
2. Hit Browse button
3. Change the file type in Browse Dialog to "Microsoft Excel Files"
4. Browse to your add-in file and select it
5. Save the Workbook



If referencing the add-in from another workbook, you must instantiate SeleniumVBA objects using the Class Factory syntax as show in the [Object Instantiation Section](https://github.com/GCuser99/SeleniumVBA/wiki#object-instantiation) of the Wiki.

