# Excel-to-Access
Connecting Excel spreed sheets with VBA-scripts to an Access DB

![Insert from Access 2](https://github.com/Byggvir/Excel-to-Access/raw/master/Images/TAE2A002.png)

## Demo for data transfer between MS Excel and MS Access

You can import or link MS Access tables and queries into MS Excel spreed sheets with simple on board tools. But these tools do not allow to set parameters in the query. To pass parameters to queries you have to use VBA scripts.

This is a collection of more or less useful examples.

## Files

In DB/Datenbanken.7z you will find a test database _Spielwiese.accdb_ and the spreasheet _Spielwiese.xslm_.

A VBA class for accessing Access databases is in _TAAccessDB.cls_ (old name AccessDB)  ready for import.

The files 

* _InsertQueryForm.frm_,
* _InsertQueryForm.frx_ 
* _UpdateQueryForm.frm_,
* _UpdateQueryForm.frx_ 

contain two userforms. The file _RibbonButtons.bas_ contains macros to call the sub routines from a RibbonButton.

The files _TAExel2Access.xlsm_ and _TAExel2Access.xlam_ contain the full application modules, the userforms and a ribbon "Access".


![Insert from Access](https://github.com/Byggvir/Excel-to-Access/raw/master/Images/TAE2A001.png)
![Insert from Access 2](https://github.com/Byggvir/Excel-to-Access/raw/master/Images/TAE2A002.png)
![Insert from Access 3](https://github.com/Byggvir/Excel-to-Access/raw/master/Images/TAE2A003.png)
![Insert from Access 4](https://github.com/Byggvir/Excel-to-Access/raw/master/Images/TAE2A004.png)

_under construction_
