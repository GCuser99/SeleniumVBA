Notes on Binary Compatibility

Note the Project Id in Project Settings: "XXXXXXXX-XXXX-XXXX-XXXX-YYYYYYYYYYYY"
Set "Use Project Id for Typelib Id" to Yes in project settings
For all Public classes, define [InterfaceId("XXXXXXXX-XXXX-XXXX-XXXX-YYYYYYYYYYYY")]
For all Public classes that are COM creatable, define [ClassId("XXXXXXXX-XXXX-XXXX-XXXX-YYYYYYYYYYYY")]
Where XXXXXXXX-XXXX-XXXX-XXXX is the first 23 characters of the Project Id/Type-Library Id
and YYYYYYYYYYYY is a random sequence unique* to each GUID assignment - insure no repeats!

This way, relevant GUID keys can be easily searched for in the Registry via XXXXXXXX-XXXX-XXXX-XXXX

Then keep these GUID assignments for the life of the project... Can add new classes but should not change existing GUID's 

*unique except for Project Id = Type-Library Id