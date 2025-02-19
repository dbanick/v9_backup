######################################
# Assessment Powershell Script       #
############################################
# Created 07/14/2020                       #
# By: Nick Patti                           #
########################################################################################
# This Powershell script attempts to rerun failed assessment scripts         	       #
# Reviews failed SQL scripts against each server and attempts to rerun them            #
########################################################################################
############################## Change Control ##########################################
#Version 5.0                                                                           #
#Date 09/2020                                                                          #
#Modified By: Nick Patti                                                               #
#  1. Code cleanup
#  2. Fixing problems
#  
########################################################################################
#determine the assessment folder based of where this is executed from; you can also explicitly define it here
$directoryPath = Split-Path $MyInvocation.MyCommand.Path
#$directoryPath = "C:\Navisite\Powershell\Assessment"

cd $directoryPath

cls


#Load Functions and Assemblies
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null


. .\functions\GetSQLVersion.ps1 #determine version of SQL to run and set variable
. .\functions\OutputStartup.ps1 #write output to cmd prompt
. .\functions\OutputSummary.ps1 #write output to cmd prompt
. .\functions\LogScriptStatus.ps1 #log script failures
. .\functions\LogAssessmentStatus.ps1 #log overall assessment failures
. .\functions\SetRErunOptions.ps1 #determines CSV location and timeout setting

#Begin TRY/CATCH to add snap-ins if needed
TRY{
    .\functions\Initialize-SqlpsEnvironment.ps1
    Write-Host "Initializing Powershell"
} CATCH {Write-Host "Cannot add Windows PowerShell snap-in SqlServerCmdletSnapin100 because it is already added."}



cd $directoryPath

########DEFAULT VARIALBES#################
$outputCSV = Resolve-Path "CSV"
$logPath = "$directoryPath\logs"
$extendedQueryTimeout = 600
#########################################

#Get input from user
$RunOptions = SetRERunOptions $outputCSV $extendedQueryTimeout
$extendedQueryTimeout = $RunOptions[0]
if ($RunOptions[1] -eq ""){$outputCSV = $RunOptions[1].Path} else {$outputCSV = $RunOptions[1]}

$folders = gci $outputCSV -recurse | Where-Object { !$PsIsContainer -and [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -match "META" -and [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -notmatch "retry"}
$i = 0

ForEach ($meta in $folders){
    
    #update progress variable
    if($i = 0){$percentOuter = 1}
    else {$percentOuter = $i/$folders.count*100}

    #reset inner progress loop each time we start new server
    $percentInner = 0
    $c = 1

    #provide progress in %  
    Write-Progress -Activity Updating -Status 'Progress->' -PercentComplete $percentOuter -CurrentOperation SeverLoop
    
    $metaName = $meta.FullName
    $contents = import-csv $metaName

    #split meta data
    $instance = $contents | ? datatype -like *version* | select-object servername #get instance name
    $instance = $instance.servername
    $version = $contents | ? datatype -like *version* | select-object value #get version
    $version = $version.value
    $assessmentType = $contents | ? datatype -like *AssessmentType* | select-object value #get assessment type
    $failedScripts = $contents | ? datatype -like *FailedScript* | select-object value #get list of failed scripts
    
    #Create a friendly name for named instance
	$instance_friendly = $instance.replace("\", "_") 
    
    #Designate an output folder for each instance
	$outputFolder = [io.path]::Combine($outputCSV, $instance_friendly)+'\'

    #create new meta file
    $metaRetry = $outputFolder + [io.path]::GetFileNameWithoutExtension($metaName) + "_retry.csv"

    #populate the new meta file with version & instance
     New-Object -TypeName PSCustomObject -Property @{
		Servername = $instance
		DataType = "Version"
		Value = $version
	} | export-csv $metaRetry -NoTypeInformation

    #pipe the non failure rows back into the new meta file
    $contents | ? datatype -notlike *FailedScript* | ? datatype -notlike *version* | export-csv -path $metaRetry -NoTypeInformation -append
        

    foreach ($query in $failedScripts){  
        #provide progress in %  
        $percentInner = ($c/($failedScripts.Count))*100
        Write-Progress -Id 1 -Activity Updating -Status 'Progress' -PercentComplete $percentInner -CurrentOperation SQLScripts
        
        $qname = $query.Value      
        try{
            invoke-sqlcmd2 -inputfile $qname -serverinstance $instance -QueryTimeout $extendedQueryTimeout | export-csv $outputFile -notype

            #log successful message to new meta file
            LogScriptStatus $instance $LogPath $qname "SuccessfulScript" $metaRetry
        }
        catch{
            #log failure message to new meta file
            LogScriptStatus $instance $logPath $qname "FailedScript" $metaRetry

        }
        $c++
    }
    #replace old meta file with new meta file
    remove-item $metaName
    Move-Item $metaRetry $metaName

    $i++
}

