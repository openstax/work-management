#  Update MSP from Trello
param([string] $Prjct, [bool] $UpdtMsp, [bool] $PstAllChckLstItms, [bool] $PstChckItmNm, $XlsTmpltFlNm, [string] $XlsFlNm, [string] $MspExe, [string] $LstsInclddStr)

# Start time
write-host ("Started at " + (get-date).tostring('hh:mm:ss') )
write-host ("Prjct=$Prjct UpdtMsp=$UpdtMsp PstAllChckLstItms=$PstAllChckLstItms PstChckItmNm=$PstChckItmNm" )
write-host ("XlsTmpltFlNm=$XlsTmpltFlNm" )
write-host ("XlsFlNm=$XlsFlNm" )
write-host ("MspExe=$MspExe" )
write-host ("LstsInclddStr=$LstsInclddStr" )

#$Prjct = "UTS Test"
#$Prjct = "UTS Test CTE"
#$Prjct = "UTS Test Tutor"
#$UpdtMsp = $true

switch ($Prjct) {
    "UTS Test" {
        $DtToUpdt = "2017-01-04"
     }
    "UTS Test CTE" {
        $DtToUpdt = (get-date).ToString('yyyy-MM-dd')
     }
    "UTS Test OSWeb" {
        $DtToUpdt = (get-date).ToString('yyyy-MM-dd')
    }
    "UTS Test Tutor" {
        $DtToUpdt = (get-date).ToString('yyyy-MM-dd')
    }
}

# Update MSP
& "C:\Users\Bruce Pike Rice\Documents\Visual Studio 2015\Projects\JiraInteraction\JiraInteraction\bin\Debug\JiraInteraction.exe" `
	$Prjct, $UpdtMsp, $PstAllChckLstItms, $PstChckItmNm, $DtToUpdt, $XlsTmpltFlNm, $XlsFlNm, $MspExe, $LstsInclddStr

Write-Host ("`n`rAll done at " + (get-date).tostring('hh:mm:ss') )