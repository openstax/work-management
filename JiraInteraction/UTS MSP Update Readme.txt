ReadMe UTS MSP Update
----------------------------

Prerequisites
1) In Program.cs ensure that switch statement for parms set by username has a case for your username with the correct parm values.

To run
2) Start JiraInteraction.exe
3) Select project

Run parameters
Prjct: Trello project for update
UpdtMsp 
	True to update MSP.  
	False to create xls without updating MSP.
PstAllChckLstItms: 
	True to post all Trello checkitems to MSP as tasks.  
	False to only post those checkitems that contain "ar:".
PstChckItmNm: 
	True to post the first 100 characters of the Trello checkitem name as the task name in MSP.  
	False to post only the task name portion of the Trello checkitem name.  
XlsTmpltFlNm: Input Excel template file
XlsOutptDrctry: Directory for output Excel file
LstsInclddStr: Trello lists to include in MSP update.  Separate list names with commas.  If this parm is blank then all lists are included.

Other parameters (Set in code)
DtToUpdt: Update Date.  Change in actual hrs is posted to the day before this date.  After MSP Update is run then the PM reschedules uncompleted work to start on this date. 
IncldCrdsChngdAftr: Cards changed after this date are included in the scan. Typically set to DtToUpdt-2d.
PrjctMsp: Project to update in MSP data store.
XlsFlNm: Output xls file.
 