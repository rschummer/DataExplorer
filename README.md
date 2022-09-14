
# Data Explorer 3
**Managing all kinds of data in the Visual FoxPro IDE using an extensible tool!** 

**Project Manager**: [Rick Schummer](https://github.com/rschummer)

The Data Explorer lets you examine data and components in Visual FoxPro databases, SQL Server databases, VFP free tables, or any other ODBC or OLE DB compliant database via an ADO connection. It can run as a task pane or as a standalone tool. Those familiar with SQL Server's Management Studio or Visual Studio's Server Explorer immediately notice some similarities, but this tool works with all kinds of data, is completely integrated in the Visual FoxPro IDE, and is extensible in true VFP tradition. 

Version 3 extends on the functionality provided by the Data Explorer released with Sedna. Specifically:

* Export/Import of Data Explorer features to share with others
* BROWSE for SQL Server with data update capability
* Toggling the Dock State of the Data Explorer form when standalone
* New LIST STRU off the shortcut menu
* New connection type for Directory shows list of all DBFs in a folder
* Corrected bug in Display Show Plan option added to Sedna version of Data Explorer

New file called ChangeLog.TXT included in the project and source with latest changes in each release.

The repository also contains a file called WLCDxFeatureManifest.XML, which contains extensions written by the project manager, Rick Schummer. These were previously hosted on http://vfpxrepository.com, which no longer exists. 

The Import Features is available on the shortcut menu of the Data Explorer node on the tree view. You are first prompted to pick a manifest file to import. If you do not pick one, the process naturally does nothing. If the file exists and is the proper format, the Import Features dialog is displayed. You can select all the features or just the ones you are interested in using.