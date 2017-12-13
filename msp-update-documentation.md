# MSP Update Documentation

## Parameters

Parameters which control program execution.  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of them for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Update MSP Actuals | If TRUE then run the Actuals Update. | FALSE | y |
| Update MSP Projected | If TRUE then run the Projection Update. | FALSE | y |
| Update MSP Measures | If TRUE then run the Measures Update. | FALSE | y |

Parameters for Update Actual.  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of them for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Boards | Trello boards to be scanned. | blank | n |
| Include Cards Changed After | Trello scan is limited to cards changed after the specified date. | 3 days before today | n |
| Post All Checklist Items | If TRUE then all checklist items on each card will be posted to Project Online.  If FALSE then only task checklist items (those containing "ar:") will be posted. | FALSE | n |
| Post Checkitem Name | If TRUE then the entire Trello checkitem name will be posted as the task name in Project, truncated to 255 chars. | FALSE | n |
| Trello Lists Included | Trello lists to be included in the scan.  List names are separated by semi-colons. | blank | y |
| Trello Lists Excluded | Trello lists to be excluded from the scan.  List names are separated by semi-colons. | blank | y |
| Trello Lists Rejected | Trello lists containing rejected work items.  Cards on these lists will be deleted from the schedule by Update Actuals. List names are separated by semi-colons. | blank | y |
| Update Date | Date which specifies when hrs are posted in Project.  Actual work is posted the day before this date.  Remaining work is posted on this date. | today | n |

Parameters for Update Measures.  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of them for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Measure Labels | Labels for which measures will be calculated. A set of measures is calculated and posted for each measure label plus all tasks. | blank | n |
| Trello Lists Not Open | Trello lists containing work items that are not open.  List names are separated by semi-colons.  Cards on these lists will not be included in the counts on the Bug & Change Open report.  All non-open lists should be included, even if they are included in the Excluded or Rejected lists.  | blank | y |


Parameters for Update Projection.  They are prefixed with the project name, like this: BIT:Boards.  There should be a set of them for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Baseline Date Complete | Baseline completion date used to calculate the daily change in completion date. | blank | y |
| Projected Actual Work Date Window Start - <measure label> | Start of date window used to calculate the least-squares projection for Actual Work. | Date of earliest data point | y |
| Projected Actual Work Date Window End - <measure label> | End of date window used to calculate the least-squares projection for Actual Work. | Date of latest data point | y |
| Projected Actual Work Change Per Day - <measure label> | Projection line slope for Actual Work.  Entering a value will override the least-squares projection. | blank | y |
| Projected Total Work Date Window Start - <measure label> | Start of date window used to calculate the least-squares projection for Total Work. | Date of earliest data point | y |
| Projected Total Work Date Window End - <measure label> | End of date window used to calculate the least-squares projection for Total Work. | Date of latest data point | y |
| Projected Total Work Change Per Day - <measure label> | Projection line slope for Total Work.  Entering a value will override the least-squares projection. | blank | y |

These parameters are used for all projects.  The parameter name is entered without the project name prefix:

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| MS Project Exe | Path to the MS Project client executable. | None | y |
| Trello AppKey | Application Key required for Trello access. | None | y |
| Trello User Token | User Token required for Trello access. | None | y |
| Xls Output Directory | Directory where output xls file will be created. | None | y |
