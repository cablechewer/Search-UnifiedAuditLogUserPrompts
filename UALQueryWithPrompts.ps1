#newLine shortcut
$script:nl = "`r`n"
$nl

#==================================Common questions asked=====================
#prompt users to specify a path to store the output files

$outputpathExists = $false

[int]$invalidDiagPath = 0
do {
  if ($invalidDiagPath -gt 0){ 
    write-host "Please provide a valid file path!" -foregroundcolor Red
  } #EndIf
  " "
  
  write-host "This script retrieves audit logs. Please specify a local folder to use for the output.  Please use a separate folder for each execution of the script. Example c:\temp" -nonewline;
  $outputPath = Read-Host " "
  $invalidDiagPath++
  
  #$outputPath -like "*.txt" -or $outputPath -like "*.csv" -or $outputPath -like "*.log"
  
  if(($outputPath.length -gt 2) -and (":\" -eq $outputPath.substring(1,2))) { #Shortest possible path is c:\ - 3 characters.  If less than 3 chars don't try test-path 
                                                                              #Make sure we have a potentially valid absolute local path.  
                                                                              #A valid absolute path should always be in the format <drive letter>:\
    $outputpathExists = Test-Path -Path "$outputPath"
	
    if($outputpathExists) {
	" "
      write-host "Path that will be used for output files is " $outputPath
    }
    else {
	" "
      Write-host "The path does not exist, would you like to create it now?" -Nonewline; Write-host " (Y/N)" -foregroundcolor Yellow -nonewline; $pathCreate = read-host " "
      if (($pathCreate -eq "Y") -or ($pathCreate -eq  "y") -or ($pathCreate -eq "yes") -or ($pathCreate -eq "YES")) {
        $error.clear()
        New-Item $outputPath -type directory | out-null
        if ($error.count -eq 0 ) { #A little imprecise, but if there is an error after only one line it is likely a failure to create the path.
          $outputpathExists = $true
        }
        else {
          write-host "Attempt to create the path failed."
        } #EndIf
      } #Yes Create a path
      else { #Path not present and user doesn't want to create it.
        Write-host " "
        Write-host "The path does not exist and you've chosen not to have it created. Create the directory and run the script again." -foregroundcolor Yellow
        do {
          Write-Host
          $continue = Read-Host "Please use the <Enter> key to exit..."
        } While ($continue -notmatch $null)
        exit  #Terminate script.
      } #EndIf No Path
    } #EndIf $pathexists -eq false
  } #EndIf 
} while ($outputpathExists -eq $false) #Don't exit the loop until we have a valid path for output files.

#prompt users to specify the base search name
$BadSearchName = $true
do {
  " " 
  write-host "Specify a name for the new Audit Log search.  ie: AuditLogJohnDoe"-nonewline;$search = Read-host " "
  write-host "Note: Script does not check whether you have used the name before." -ForegroundColor Cyan

  $search = $search.trim()
  if ($search.length -ge 1) { #Prevent empty string in $search
    $charDisappointmentPending = $true
	$name_pos = $search.length
    while (($name_pos -gt 0) -and ($charDisappointmentPending) ) {
      $name_pos--
      if(($search[$name_pos] -eq "*" ) -or ($search[$name_pos] -eq "?")) { #If either of the known bad characters appears in the name flag it.
                                                                           #Notlike and notcontains don't work with these chars.  Thus the manual loop.
        $charDisappointmentPending = $false	 #name contains unacceptable characters.
        write-host "Search name entered contains invalid characters such as ? or *." -foregroundColor red 
      } #EndIf 
    } #Loop 
	if($charDisappointmentPending -and ($name_pos -eq 0)) { #if we did not trip loop exit condition prematurely name is considered good.
      $BadSearchName = $false 
    } #EndIf
  } #$search has at least 1 char that is not a space 
  else { #$search was empty or only had spaces 
    write-host "Name was empty." -foregroundColor red
  } #EndIf 
} while ($BadSearchName) 
#write-host "name acceptable"



#Modify the values for the following variables to configure the audit log search.
$logFile = $outputPath+"\"+$search+".txt"
$outputFile = $outputPath+"\"+$search+".csv"
write-host $nl
write-host "Log of this scripts excution will be saved to " $logFile
write-host "Audit data will be saved to " $outputFile


#prompt to ask if the search will use a START date
$InvalidDate=$true
$DisplayDate = (get-date -format "MM-dd-yyyy HH:mm")
$DisplayDateString = $DisplayDate.tostring() 
$CurrentDate = (get-date ).ToUniversalTime()
$OldestDate = ((get-date).ToUniversalTime()).adddays(-90)
$currentDateString = $CurrentDate.tostring("MM-dd-yyyy HH:mm") 
$OldestDateString = $OldestDate.tostring("MM-dd-yyyy HH:mm")

do {
  Write-host "What is the Start Date? Enter a date between $CurrentDateString and $OldestDateString" 
  write-host "Time is UTC.  " -foregroundcolor cyan
  write-host "format: " $DisplayDateString -foregroundcolor green -nonewline;$StartDate = read-host " "
  " "

  if ($StartDate -as [DateTime]) {
    $StartDate = [DateTime]::Parse($StartDate)
    if ($StartDate -le $CurrentDate ) {
      $InvalidDate = $false
    }
    else {
      write-host "Date is in the future.  Date must be today or earlier." -foregroundColor red 
    }
  } #Date valid
  else {
    write-host "Start Date is in the wrong Format.  The expected format is:" $DisplayDateString -foregroundColor red 
  } #EndIf StartDate
} while ($InvalidDate)

#prompt to ask if the search will use an END date
$InvalidDate=$true
do {
  Write-host "What is the End Date? Enter a date between $CurrentDate and $StartDate" 
  write-host "Time is UTC.  PRESS ENTER TO ACCEPT THE CURRENT DATE." -foregroundcolor cyan
  write-host "format: " $DisplayDateString -foregroundcolor green -nonewline;$EndDateInput = read-host " "
  " "

  if($EndDateInput -eq [char]13 -or $EndDateInput -eq $NULL -or $EndDateInput.length -le 1) {
	$Enddate = $CurrentDate
    $InvalidDate = $false
  }
  else {
	if (($EndDateInput -as [DateTime]) -AND $InvalidDate) {
      $EndDate = [DateTime]::Parse($EndDateInput)
      if ($EndDate -le $CurrentDate ) {
        $InvalidDate = $false
      }
      else {
        write-host "Date is in the future.  Date must be today or earlier." -foregroundColor red 
      }
    } #Date valid
    else {
      write-host "End Date is in the wrong Format.  The expected format is:" $DisplayDate -foregroundColor red 
    } #EndIf EndDate
  }
  
  If ($StartDate -ge $EndDate ) {
	$InvalidDate = $true 
    write-host "Start Date specified is greater than or equal to the End Date." -foregroundColor red
    write-host "Supplied starting date is " $StartDate
    write-host "Supplied ending date is " $EndDate
    " "
  } #EndIf 
} while ($InvalidDate)
write-host "Supplied starting date is " $StartDate
write-host "Supplied ending date is " $EndDate

[int]$invalidDiscoverypath = 0
$invalidFile=$true
do{
  do {
    if($invalidDiscoverypath -gt 0) {
      write-host "Please provide a valid file path!" -foregroundcolor Red
    } #EndIf 
    " "
    write-host $nl
    write-host "Please supply the name of a txt file that lists the users you are querying for.  File should list"
    write-host "one user per line and have no header."
    write-host "  Example path c:\temp\UserList.txt" -nonewline; $UserFile = read-host " "
#      Write-host "Specify the discovery search mailbox list location. example "C:\temp\discoverylist.txt" " -nonewline;$path1 = Read-host " "
    if( $UserFile.LastIndexOf(".txt") -eq ($UserFile.length -4)) {
      $pathExists1 = Test-Path -Path "$UserFile"
    } #EndIf 
    $invalidDiscoverypath++
  } while($pathexists1 -eq $false)
  #"Current Path: $UserFile"
  if ($pathExists1 -eq $true) {
    Write-host "The file exists continuing..." -foregroundcolor cyan
    " "
    [array]$UserList=(get-content -path $UserFile ).trim() | select-object -unique  #NOTE We do not verify the User Mailboxes. Trusting script user to supply valid userprincipalnamnes.
    $invalidFile = $false 
  }  
} while ($invalidFile)

#[DateTime]$start = [DateTime]::UtcNow.AddDays(-1)
#[DateTime]$end = [DateTime]::UtcNow
#$record = "AzureActiveDirectory"
$resultSize = 1000
$intervalMinutes = 15

#Start script
[DateTime]$currentStart = [DateTime]::Parse($StartDate)
[DateTime]$currentEnd = [DateTime]::Parse($StartDate)
write-host "Dates"
$currentStart
$currentEnd


Function Write-LogFile ([String]$Message)
{
    $final = [DateTime]::Now.ToUniversalTime().ToString("s") + ":" + $Message
    $final | Out-File $logFile -Append
}

write-host $nl
Write-LogFile "BEGIN: Retrieving audit records between $($StartDate) and $($EndDate), RecordType=$record, PageSize=$resultSize."
Write-Host "Retrieving audit records for the date range between $($StartDate) and $($End)Date, RecordType=$record, ResultsSize=$resultSize"

$totalCount = 0
while ($true)
{
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
	
	$currentend
	
    if ($currentEnd -gt $EndDate)
    {
        $currentEnd = $EndDdate
    }

    if ($currentStart -eq $currentEnd)
    {
        break
    }
#sleep -seconds 60

    $sessionID = [Guid]::NewGuid().ToString() + "_" +  "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    Write-LogFile "INFO: Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    $currentCount = 0

    $sw = [Diagnostics.StopWatch]::StartNew()
    do
    {
		$results = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -userids $UserList -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize

        if (($results | Measure-Object).Count -ne 0)
        {
            $results | export-csv -Path $outputFile -Append -NoTypeInformation

            $currentTotal = $results[0].ResultCount
            $totalCount += $results.Count
            $currentCount += $results.Count
            Write-LogFile "INFO: Retrieved $($currentCount) audit records out of the total $($currentTotal)"

            if ($currentTotal -eq $results[$results.Count - 1].ResultIndex)
            {
                $message = "INFO: Successfully retrieved $($currentTotal) audit records for the current time range. Moving on!"
                Write-LogFile $message
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor cyan
                ""
                break
            }
        } else {write-logFile "No data returned."}
    }
    while (($results | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}

Write-LogFile "END: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$resultSize, total count: $totalCount."
Write-Host "Script complete! Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green