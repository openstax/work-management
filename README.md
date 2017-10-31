# Work Management
## Metric Dashboard
[OpenStax Development Metrics](https://openstax.github.io/work-management-reports/openstax_development_metrics)
## To Update Burnup Chart
### Prerequisites
1. Trello cards must contain the following elements
   * Label for project
   * Tasks and hours on cards (format in Work Management for Software Development). 
1. MSP Update installed.
   * Get deploy.zip from https://github.com/openstax/work-management
   * Unzip in install directory (recommended location is MyDocuments\bin\Work Management). This will create a folder “Deploy”.  Executable is JiraInteraction.exe.
   * Configuration file UTS MSP Update Config.txt (available here) installed in install directory 
1. Excel client installed.
1. MS Project client 2016 installed.  
   * Log on to Office 365
   * Click the gear at the top right corner to open Settings
   * Under Your app settings click Office 365 to bring up the Settings page
   * Click Software
   * Under Install Project 2016, select Language and Version (32 bit recommended) and click Install.  If Project 2016 is not visible see Bruce Pike.
1. Project created in Project Online.  
1. Edit permission for project in Project Online
### Run MSP Update - Actuals
1. In config file set parameters for the project to be updated.
   * Update MSP Actuals = TRUE
   * Update MSP Projected = FALSE
   * Update MSP Measures = FALSE
1. Scan Trello cards
   * Run JiraInteraction
   * Possible errors
      * Started with MS Project client or xls file already open.
      * Schedule previously saved with filter on.
1. Fix task issues
   * Open schedule in MSProject client
      * Drag pane divider all the way to the right to expose data fields.
      * Task name is colored yellow for new rows.
1. Fix task errors (task name colored red; error message in field text30).  Click [Hyperlink Address] to bring up Trello card for task.
1. Progress unfinished work
   * Resource Tab > Clear Leveling > Entire Project
   * Filter for [Progress Unfinished Work] = Yes
   * Select all 
   * Project Tab > Change Status Date to yesterday.
   * Project Tab > Update Project > Reschedule uncompleted work to start after > Selected tasks
1. Fix tasks listed as errors
   * Open xls.  
   * Fix errors listed on tabs [Errors] and [MSP Task Errors].  Tasks are listed by MS Project ID. 
   * Save project and close MS Project client
### Run MSP Update - Measures
1. In config file set parameters
   * Update MSP Actuals = FALSE
   * Update MSP Projected = FALSE
   * Update MSP Measures = TRUE
1. Run update
1. When finished open in check measures tasks added.  
   * Ensure that one and only one set of measure tasks is added for today.
   * Ensure that measure tasks are created for all measure labels specified.
1. If ok then publish and close MS Project client. Click Yes to check in.
### Update and publish report
1. In Power BI log in as Portfolio Viewer
1. Open report file.
1. Refresh data
1. Publish report
1. Verify that report is updated on OpenStax Development Metrics.
