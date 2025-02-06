<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 20250206.1406

#Instructions
#$OutputPath is a directory where the results will be saved.
#$AddInCsvPath is the location of a csv file that contains known add-ins using legacy tokens. 
#Ideally use "add-ins-using-exchange-tokens.csv" from https://github.com/OfficeDev/office-js/blob/release/add-in-ids/add-ins-using-exchange-tokens.csv


param (
    [Parameter(Mandatory=$true,HelpMessage="The OutputPath parameter specifies the directory where the results are written")]
    [ValidateScript( {Test-Path $_})]
    [string]$OutputPath,

    [Parameter(Mandatory=$true,HelpMessage="The AddInCsvPath parameter specifies the directory where you have saved the CSV of known legacy auth token add-ins from Microsoft")]
    [ValidateScript( {Test-Path $_}, ErrorMessage = "{0} CSV of known legacy token add-ins was not found.")]
    [string]$AddInCsvPath
)

#Create a variable for this minute in local time
$date = (Get-Date).ToString('yyyyMMddHHmmss')

#Creating output file to log mailboxes as they are processed
$LogFile = New-Item -Path "$($OutputPath)\Get-LegacyExchangeTokenAddins-$($date).log" -ItemType File

$Addins = New-Object System.Collections.ArrayList
#Set the progress preference to continue for PowerShell 7.0 and later
$CurrentProgress = $ProgressPreference
$ProgressPreference = "Continue"
#Loop through char 97 through 122 which is A-Z
foreach ($char in [char[]](97..122)) {
    #Store the mailboxes that start with each letter in to Mailboxes
    Write-Host "Generating list of mailboxes starting with $($char)..." -ForegroundColor Green -NoNewline
    $Mailboxes = Get-Mailbox -Filter "alias -like '$($char)*'" -ResultSize unlimited
    $Mailboxes = $Mailboxes | Sort-Object -Property Alias
    Write-Host "$($Mailboxes.Count) mailboxes found"
    #Loop through each mailbox in Mailboxes
    $x = 1
    $TotalMailboxes = $Mailboxes.Count
    if($TotalMailboxes -ge 1){
        foreach ($mailbox in $Mailboxes){
            Write-Progress -Activity "Getting addins for mailboxes" -CurrentOperation "Processing $($mailbox.alias)" -PercentComplete (($x / $TotalMailboxes) * 100)
            #Get the Addins for each mailbox
            #Note: This only returns user installed and default add-ins. It doesn't return add-ins installed by admins from Integrated Apps.
            $Addins += Get-App -Mailbox $mailbox.identity | Select-Object -Property DisplayName,Enabled,AppVersion,AppId,MarketplaceAssetId
            $x++
            #Filter out the duplicates
            $Addins = $Addins | Sort-Object -Property AppId,AppVersion | Get-Unique -AsString
            #Find only unique addins and export them to CSV
            $Addins | Export-Csv "$($OutputPath)\Add-in-List-$($date).csv" -NoTypeInformation
            $LogFile | Add-Content -Value "Mailbox $($mailbox.alias) has been processed"
        }
    }
};
   

#Import csvs of legacy token addins and addins to check
$LegacyTokenAddins = Import-Csv $AddInCsvPath

#Create a hashtable of the legacy token addins
$AddinsToCheck = @{}
foreach($addin in $LegacyTokenAddins) {$AddinsToCheck[$addin.AppId] = $addin.Name}
#Check the addins against the legacy token addins
foreach($addin in $Addins) {
    if($AddinsToCheck.ContainsKey($addin.appid)) { 
        write-host "Found an add-in that matches the CSV of known add-ins using legacy tokens: $($addin.AppId) - $($addin.DisplayName)" -ForegroundColor Yellow
    }
}
$ProgressPreference = $CurrentProgress
