# MSP Update Documentation

## Parameters

**Global parameters used for all projects.**  The parameter name is entered without the project name prefix:

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| MS Project Exe | Path to the MS Project client executable. | None | y |
| Trello AppKey | Application Key required for Trello access. | None | y |
| Trello User Token | User Token required for Trello access. | None | y |
| Xls Output Directory | Directory where output xls file will be created. | None | y |

**Parameters which control program execution.**  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of these parameters for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| MSP Project Name | Name of the project to be updated in Project Online. | FALSE | y |
| Update MSP Actuals | If TRUE then run the Actuals Update. | FALSE | y |
| Update MSP Measures | If TRUE then run the Measures Update. | FALSE | y |
| Update MSP Projected | If TRUE then run the Projection Update. | FALSE | y |
| Update MSP KDs | If TRUE then run the KD Update. | FALSE | y |
| Points/Hours | If Points then run points-based update. If Hours then run hours-based update (obsolete). | none | y |
| Debug | If TRUE then print run progress notes to the console. | FALSE | y |

**Parameters for Update Actual.**  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of these parameters for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Boards | Trello boards to be scanned. | blank | y |
| Include Cards Changed After | Trello scan is limited to cards changed after the specified date. | 1/1/1900 | n |
| Trello Lists Included | Trello lists to be included in the scan.  See note below. | blank | y |
| Trello Lists Excluded | Trello lists to be excluded from the scan.  List names are separated by semi-colons. | blank | y |
| Xls File Name | File name for output xls file. | blank | y |

**Parameters for Update Measures.**  In the config file the parameter name is prefixed with the project name, like this: BIT:Boards.  There should be a set of these parameters for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Measure Condition | Condition for which measures will be calculated. Each condition is in this format: (LabelIncluded;LabelIncluded) AND NOT (LabelExcluded;LabelExcluded).  A set of measures is calculated and posted for each measure condition plus for all tasks. For a card/task to be included in a measure, it must have all the specified LabelIncluded and none of the specified LabelExcluded.  LabelIncluded and LabelExcluded entries are case-sensitive. | blank | n |
| Trello Lists Not Open | Trello lists containing work items that are not open.  List names are separated by semi-colons.  Cards on these lists will not be included in the counts on the Bug & Change Open report.  Lists with names starting with "Release" or "Hotfix" are automatically included.  | blank | y |

**Parameters for Update Projection.**  They are prefixed with the project name, like this: BIT:Boards.  There should be a set of them for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Baseline Date Complete | Baseline completion date used to calculate the daily change in completion date. | blank | y |
| Projected Actual Work Date Window Start - <measure label> | Start of date window used to calculate the least-squares projection for Actual Work. | Date of earliest data point | y |
| Projected Actual Work Date Window End - <measure label> | End of date window used to calculate the least-squares projection for Actual Work. | Date of latest data point | y |
| Projected Actual Work Change Per Day - <measure label> | Projection line slope for Actual Work.  Entering a value will override the least-squares projection. | blank | y |
| Projected Total Work Date Window Start - <measure label> | Start of date window used to calculate the least-squares projection for Total Work. | Date of earliest data point | y |
| Projected Total Work Date Window End - <measure label> | End of date window used to calculate the least-squares projection for Total Work. | Date of latest data point | y |
| Projected Total Work Change Per Day - <measure label> | Projection line slope for Total Work.  Entering a value will override the least-squares projection. | blank | y |

**Parameters for Update KDs.**  In the config file the parameter name is prefixed with the project name, like this: BIT:KD Boards.  There should be a set of these parameters for each project.

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| KD Boards | Trello boards to be scanned for the KD update. | blank | y |
| KD Lists Included | Trello lists to be included in the KD scan.  See note below. | blank | y |
| KD Lists Excluded | Trello lists to be excluded from the KD scan.  List names are separated by semi-colons. | blank | y |

## Notes
* Setting all Update parms to FALSE will generate the xls but not update Project.
* When the Project update starts, it's a good idea to bring up the Project window.  If the schedule does not appear it's a good idea to start over.  The job can be interrupted without hurting either the xls file or the schedule in the Project client.
* The MSP project should contain these header tasks in order with OutlineLevel = 1
  * WorkItem Tasks
  * Measure Tasks
  * Projected Tasks
  * KD Tasks
  
* Processing of Lists Included (LI) and Lists Excluded (LE)

  * List names are separated by semi-colons.
  * Conditions
    * LI nonblank LE blank: Lists on LI are included; others are excluded. 
    * LI blank LE nonblank: Lists on LE are excluded; others are included.
    * LI and LE nonblank: Lists on LI are included, then lists on LE are excluded.

**Obsolete Parameters**

| Parameter | Description | Default | Required |
| ---------- | --------------------------- | :--------: | :--------: |
| Points/Hours | If Points then run points-based update. If Hours then run hours-based update (obsolete). | none | y |
| Post All Checklist Items | If TRUE then all checklist items on each card will be posted to Project Online.  If FALSE then only task checklist items (those containing "ar:") will be posted. | FALSE | n |
| Post Checkitem Name | If TRUE then the entire Trello checkitem name will be posted as the task name in Project, truncated to 255 chars. | FALSE | n |
