
#***********************************************************************************************#
#                                                                                               #
#    Title       : Phishing Email Deletion using Azure Compliance Search                        #
#    Author      : Udit Mahajan                                                                 #
#    Date        : 03/06/2023                                                                   #
#    Code version: V3                                                                           #
#                                                                                               #
#***********************************************************************************************#

#***********************************************************************************************#
#                                                                                               #
#    Pre-requisites:                                                                            #
#       1. Ensure Execution policy is set to remotely signed                                    #     
#       2. Azure Compliance Manager or Global Administrator roles                               #
#       3. A connected Exchange Online account                                                  #
#                                                                                               #
#    READ ME                                                                                    #
#    The code is able to identify emails based on Date,Subject,Body and Sender email address    #
#    and delete all emails based on the citeria.                                                #
#    Please check if Content search results are as desired as it is possible to delete all      #
#    organisation mailboxes accidentally                                                        #
#    I have intentially commented delete option to prevent accidental deletes                   #
#                                                                                               #  
#***********************************************************************************************#

#Global variables users will enter later
$dateRange = $null
$startDateInput = $null
$phishidentifier = $null

function GetPhishInfo{
    #Please provide a unique name, ideally with the existing naing conventions.
    $global:casename= Read-Host -Prompt "Enter a new case name"
    $global:phishidentifier= Read-Host -Prompt "Enter a phishing email body"
}

function GetDateRange {
    #This function can accept dates in following formats: D/M, DD/MM, DD/MM/YY, DD/MM/YYYY
    # Not entering the date will skip the date filter when searching for emails
    # Prompts the user to enter the start and end dates
    $startDateInput = Read-Host "Enter the start date (DD/MM or DD/MM/YY or DD/MM/YYYY):"
    $endDateInput = Read-Host "Enter the end date (DD/MM or DD/MM/YY or DD/MM/YYYY):"

    # Remove any leading or trailing spaces
    $startDateInput = $startDateInput.Trim()
    $endDateInput = $endDateInput.Trim()

    # Check if input strings are empty
    if ([string]::IsNullOrWhiteSpace($startDateInput) -or [string]::IsNullOrWhiteSpace($endDateInput)) {
        Write-Host "No date range entered. Skipping date processing."
        return
    }

    # Check the date format and modify if necessary
    $startDateObj = $null
    $endDateObj = $null

    if ($startDateInput -match '^\d{2}/\d{2}$') {
        # Convert start date to DD/MM/YYYY format
        $startDateInput = "$startDateInput/$((Get-Date).Year)"
    }
    elseif ($startDateInput -match '^\d{2}/\d{2}/\d{2}$') {
        # Convert start date to DD/MM/YYYY format
        $startDateInput = [DateTime]::ParseExact($startDateInput, 'dd/MM/yy', $null).ToString('dd/MM/yyyy')
    }

    if ($endDateInput -match '^\d{2}/\d{2}$') {
        # Convert end date to DD/MM/YYYY format
        $endDateInput = "$endDateInput/$((Get-Date).Year)"
    }
    elseif ($endDateInput -match '^\d{2}/\d{2}/\d{2}$') {
        # Convert end date to DD/MM/YYYY format
        $endDateInput = [DateTime]::ParseExact($endDateInput, 'dd/MM/yy', $null).ToString('dd/MM/yyyy')
    }

    # Convert the input strings to DateTime objects
    $startDateObj = $startDateInput -as [DateTime]
    $endDateObj = $endDateInput -as [DateTime]

    # Validate the date range
    if ($startDateObj -eq $null -or $endDateObj -eq $null) {
        Write-Host "Invalid date format! Please provide dates in the specified format."
        return
    }
    elseif ($startDateObj -gt $endDateObj) {
        Write-Host "Invalid date range! The start date cannot be greater than the end date."
        return
    }

    # Create the date range in string to use in Content Search
    $global:dateRange = "Received:$($startDateObj.ToString('dd/MM/yyyy'))..$($endDateObj.ToString('dd/MM/yyyy'))"

    Write-Host "Selected date range:"
    Write-Host $global:dateRange
}

function RunComplianceSeach{
    Write-Host "Creating and executing Content Search"
    #If date is not provided, the search will skip it
    if ($global:dateRange -ne $null) {
       #$query = "(Sender:$sender OR Subject:$subject OR Received:$dateRange OR Body:$keywords)"
       $query = "($global:dateRange AND Body:$global:phishidentifier)"
       #Write-Host $query
       #$Search = New-ComplianceSearch -Name $global:casename -ExchangeLocation All -ContentMatchQuery $query
    } else {
       #$Search = New-ComplianceSearch -Name $global:casename -ExchangeLocation All -ContentMatchQuery $global:phishidentifier
    }
    #Start-ComplianceSearch -Identity $Search.Identity
}

function SoftDelete{
    #HardDelete is also possible however, there is no recovery after it
    # Wait for 120 seconds
    Write-Host "Sleeping 120 seconds for Content Search to finish"
    #Sleeping Animation
    for($i = 0; $i -le 100; $i++)
        {
          Write-Progress -Activity "Activity" -PercentComplete $i -Status "Processing";
          Sleep -Milliseconds 1200;
        }
    Start-Sleep -Seconds 120
    Write-Host "Deleting Emails Now"
    #New-ComplianceSearchAction -SearchName $global:casename -Purge -PurgeType SoftDelete
}


# Function calls and command calls
Connect-IPPSSession
GetPhishInfo
GetDateRange
RunComplianceSeach
SoftDelete
Read-Host -Prompt "Press Enter to exit"
