function Set-SO {
    [CmdletBinding()]

    param (
        [parameter(Mandatory=$true)]
        [string[]]$SONumber,
                    
        [parameter(Mandatory=$false)]
        [string[]]$SOBriefDescription,
  
        [parameter(Mandatory=$false)]
        [string[]]$ClientFirstName,
 
        [parameter(Mandatory=$false)]
        [string[]]$ClientLastName,
   
        [parameter(Mandatory=$false)]
        [string[]]$ClientEmail,
   
        [parameter(Mandatory=$false)]
        [string[]]$SOZone,
  
        [parameter(Mandatory=$false)]
        [string[]]$SOStatus,
 
        [parameter(Mandatory=$false)]
        [string[]]$SOType,
 
        [parameter(Mandatory=$false)]
        [string[]]$SOPriority,

        [parameter(Mandatory=$false)]
        [string[]]$SOBoard,
 
        [parameter(Mandatory=$false)]
        [string[]]$SOWorkRequested,

        [parameter(Mandatory=$false)]
        [string[]]$SOInternalComments,

        [parameter(Mandatory=$false)]
        [string[]]$SOAssignedTech,

        [parameter(Mandatory=$false)]
        [Switch]$SOTask,

        [parameter(Mandatory=$false)]
        [DateTime]$SOTaskStartTime,
  
        [parameter(Mandatory=$false)]
        [DateTime]$SOTaskEndTime,

        [parameter(Mandatory=$false)]
        [string[]]$SOTaskWorkPerformed,

        [parameter(Mandatory=$false)]
        [string[]]$SOBillingType,

        [parameter(Mandatory=$false)]
        [Switch]$SOFollowUp,

        [parameter(Mandatory=$false)]
        [Switch]$FollowUpType
    )
 


    BEGIN {}

PROCESS {
set-Location $PSScriptRoot
. .\Start-TextBoxInputForm.ps1
Import-Module AutoITX

#Set TP Window Title Here
$TPWindowTitle = "Tigerpaw [ServerNameorIP;Tigerpaw11]"

#Open SO
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
$TPHandle = Get-AU3WinHandle -Title $TPWindowTitle
Show-AU3WinActivate -WinHandle $TPHandle
[System.Windows.Forms.SendKeys]::SendWait("{F9}") 
start-sleep -s 1
$SavePrompt = Get-au3winhandle -title "Tigerpaw"

If ($SavePrompt -ne $null){
Show-AU3WinActivate -WinHandle $SavePrompt
[System.Windows.Forms.SendKeys]::SendWait("%{y}")
}
start-sleep -s 1
Show-AU3WinActivate -Title "Goto Service Order"
start-sleep -Milliseconds 800
[System.Windows.Forms.SendKeys]::SendWait("$SONumber") 
Start-Sleep -Milliseconds 800
[System.Windows.Forms.SendKeys]::SendWait("{Enter}") 
Start-sleep -Seconds 5
$TPSIWindowTitle = "Special Instructions"
$TPSIHandle = Get-AU3WinHandle -Title $TPSIWindowTitle
Show-AU3WinActivate -WinHandle $TPSIHandle
Invoke-AU3ControlClick -Title $TPSIWindowTitle -Control WindowsForms10.BUTTON.app.0.83dd9_r17_ad11

#Get Client Name
$ControlConst = "WindowsForms10.Window.8.app.0.83dd9_r17_ad1"
$ClientOrgMatch = "\(\d+\)"
[int]$IDCVar = '1'

While ($IDCVar -lt 200){
$ClientOrgTestControl = $ControlConst + $IDCvar
$ClientOrgTestName = get-au3controltext -Title $TPWindowTitle -Control $ClientOrgTestControl
     If ($ClientOrgTestName -Match $ClientOrgMatch){
        write-output "Match Found:`nControl: $ClientOrgTestControl`n$ClientOrgTestName`n"
        $ClientOrgName = $ClientOrgTestName -Replace '\(([^\)]+)\)', ''
        write-output $ClientOrgName`n
         }
             else{
             }

$IDCVar++

}

#Get Service order Number
$ControlConst = "WindowsForms10.Window.8.app.0.83dd9_r17_ad1"
$ServiceOrderMatch = "Service Order (([\d]{6}))"
[int]$IDSOVar = '1'

While ($IDSOVar -lt 200){
$ServiceOrderTestControl = $ControlConst + $IDSOvar
$ServiceOrderTestName = get-au3controltext -Title $TPWindowTitle -Control $ServiceOrderTestControl
     If ($ServiceOrderTestName -Match $ServiceOrderMatch){
        write-output "Match Found:`nControl: $ServiceOrderTestControl`n$ServiceOrderTestName`n"
        $ServiceOrderNumber = $ServiceOrderTestName 
        $ServiceOrderNumberNumber = $ServiceOrderTestName -replace "Service Order ", ""
        write-output $ServiceOrderNumber`n
        write-output $ServiceOrderNumberNumber`n
         }
             else{
             }

$IDSOVar++

}



#Open Work Performed Input Box if task value is true but work performed is not
If (($SOTask -eq $True) -and ($SOTaskWorkPerformed -eq $Null)){
$TextArg1 = "Client: $ClientOrgName `| $ServiceOrderNumber"
Start-TextBoxInputForm "$TextArg1" "Enter Work Performed"
$SOTaskWorkPerformed = $UserInput
If ($UserInput -eq "CANCELED"){
return
}
}


#Set Brief Description If Present
If ($SOBriefDescription -ne $Null){
Write-Output "There's a value here!: $SOBriefDescription"
Start-Sleep -Milliseconds 900
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit24
[System.Windows.Forms.SendKeys]::SendWait("$SOBriefDescription")
}
    else{ write-output "There's no SOBriefDescription value, doing nothing"
    }

#Set Client Name If Present
IF ($ClientFirstName -ne $Null){
    If ($ClientLastName -eq $null){
    Write-Output "Please input Client's Last Name"
    $ClientLastName = Read-Host
    Show-AU3WinActivate -WinHandle $TPSIHandle  
    }
Show-AU3WinActivate -WinHandle $TPSIHandle  
start-sleep -s 2
Write-Output "There's a value here!: $ClientFirstName"
Write-Output "There's a value here!: $ClientLastName"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit19 
[System.Windows.Forms.SendKeys]::SendWait("$ClientFirstName" + " " + "$ClientLastName")
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
}
    else { Write-Output "There's no ClientFirstName value, doing nothing"

    }

#Set Client Email If Present
If ($ClientEmail -ne $Null){
Write-Output "There's a value here!: $ClientEmail"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit22
[System.Windows.Forms.SendKeys]::SendWait("$ClientEmail")
}
else{Write-Output "There's no ClientEmail value, doing nothing"
}

#Set Zone If Present
If ($SOZone -ne $Null){
Write-Output "There's a value here!: $SOZone"
Start-Sleep -s 4
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit21
[System.Windows.Forms.SendKeys]::SendWait("$SOZone")
}
else {Write-Output "There's no SOZone value, doing nothing"
}

#Set Status If Present
If ($SOStatus -ne $Null){
Write-Output "There's a value here!: $SOStatus"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit1
[System.Windows.Forms.SendKeys]::SendWait("$SOStatus")
}
else {Write-Output "There's no SOStatus value, doing nothing"

}

If ($SOType -ne $Null){
Write-Output "There's a value here!: $SOType"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit20 
[System.Windows.Forms.SendKeys]::SendWait("$SOType")
}
else{Write-Output "There's no SOType value, doing nothing"
}

If ($SOPriority -ne $Null){
Write-Output "There's a value here!: $SOPriority"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit2
[System.Windows.Forms.SendKeys]::SendWait("$SOPriority")
}
else{Write-Output "There's no SOPriority value, doing nothing"
}

If ($SOBoard -ne $Null){
Write-Output "There's a value here!: $SOBoard"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit23
[System.Windows.Forms.SendKeys]::SendWait("$SOBoard")
}
else{Write-Output "There's no SOBoard value, doing nothing"
}

If ($SOWorkRequested -ne $Null){
Write-Output "There's a value here!: $SOWorkRequested"
Start-Sleep -Milliseconds 200
Set-AU3ControlText -Title $TPWindowTitle -Control RichTextWndClass1 -NewText "$SOWorkRequested"
#[System.Windows.Forms.SendKeys]::SendWait("^{a}")
#[System.Windows.Forms.SendKeys]::SendWait("$SOWorkRequested")

}
else{Write-Output "There's no SOWorkRequested value, doing nothing"
}

If ($SOInternalComments -ne $Null){
Write-Output "There's a value here!: $SOInternalComments"
Start-Sleep -Milliseconds 200
$SOInternalExistingComments = Get-AU3ControlText -Title $TPWindowTitle -Control RichTextWndClass3
$SOInternalCommentsToAdd = "$("$(Get-Date)," + " " + "Logan Harris:" + " " + "$SOInternalComments")"
$SOInternalExistingComments += "`n"
$SOInternalExistingComments += $SOInternalCommentsToAdd
Set-AU3ControlText -Title $TPWindowTitle -Control RichTextWndClass3 -NewText "$SOInternalExistingComments"

#[System.Windows.Forms.SendKeys]::SendWait("$(Get-Date)," + " " + "Logan Harris:" + " " + "$SOInternalComments")
}
else {Write-Output "There's no SOInternalComments value, doing nothing"
}

#Assign Tech if value present
If ($SOAssignedTech -ne $Null){
Write-Output "There's a value here!: $SOAssignedTech"
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit16
[System.Windows.Forms.SendKeys]::SendWait("$SOAssignedTech")
}
else {
Write-Output "There's no SOAssignedTech value, doing nothing."
}

Switch ($SOTask){
$True {
$SOAssignedTech = "Logan Harris"
Write-Output "Creating task..."
Invoke-AU3ControlClick -Title $TPWindowTitle -Control Edit16
[System.Windows.Forms.SendKeys]::SendWait("$SOAssignedTech")
Start-Sleep -s 1

Show-AU3WinActivate -WinHandle $TPHandle
[System.Windows.Forms.SendKeys]::SendWait("{F5}")
Start-sleep 5
$TPTaskWindowTitle = "Schedule Task"
$TPTaskHandle = Get-AU3WinHandle -Title $TPTaskWindowTitle
Invoke-AU3ControlClick -Title $TPTaskWindowTitle -Control ThunderRT6TextBox7
[System.Windows.Forms.SendKeys]::SendWait("$(($SOTaskStartTime).ToString("hh:mm tt"))")

If ($SOTaskEndTime -ne $Null){
Write-Host "There's a value here!: $SOTaskEndTime"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPTaskWindowTitle -Control ThunderRT6TextBox11
[System.Windows.Forms.SendKeys]::SendWait("$(($SOTaskEndTime).ToString("hh:mm tt"))")

}
else{ 
Write-Host "There's no value, writing end time based on current time"
Start-Sleep -Milliseconds 200
Invoke-AU3ControlClick -Title $TPTaskWindowTitle -Control ThunderRT6TextBox11
[System.Windows.Forms.SendKeys]::SendWait($(Get-Date).ToString("hh:mm tt"))

}
$SOTaskComplete = "1"
If (($SOTaskComplete -eq "1") -and ($SOTask -eq $True)){
Write-Host "Setting task to complete..."
Invoke-AU3ControlClick -Title $TPTaskWindowTitle -Control ThunderRT6CheckBox3

}

#Enter Time Log from Schedule Task
$SOCreateTimeLog = "1"
If (($SOCreateTimeLog -eq "1") -and ($SOTask -eq $True)){
Write-Host "Creating time log entry"
Start-sleep -Milliseconds 200
[System.Windows.Forms.SendKeys]::SendWait("`%{f}")
[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 5
$TPTimeLogWindowTitle = "New Time Entry"
$TPTimeLogHandle = Get-AU3WinHandle -Title $TPTimeLogWindowTitle
Invoke-AU3ControlClick -Title $TPTimeLogWindowTitle -Control WindowsForms10.Window.8.app.0.83dd9_r17_ad147


Switch ($SOBillingType){
Billable{
Start-sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("CSCare Block Hourly Labor")
Start-sleep -Milliseconds 500
}

MSP{
Start-sleep -Milliseconds 500
[System.Windows.Forms.SendKeys]::SendWait("CSCare MSP Included Labor")
Start-sleep -Milliseconds 500
}

}
Start-sleep -Milliseconds 500
Set-Au3ControlText -Title $TPTimeLogWindowTitle -Control WindowsForms10.EDIT.app.0.83dd9_r17_ad12 -NewText ""
Invoke-AU3ControlClick -Title $TPTimeLogWindowTitle -Control WindowsForms10.EDIT.app.0.83dd9_r17_ad12
Start-sleep -Milliseconds 500
Write-Output "Work Performed: $SOTaskWorkPerformed"
Send-AU3ControlKey -Title $TPTimeLogWindowTitle -Control WindowsForms10.EDIT.app.0.83dd9_r17_ad12 -Key $SOTaskWorkPerformed.Replace("`r`n", "`n")


start-sleep -Milliseconds 500
Invoke-AU3ControlClick -Title $TPTimeLogWindowTitle -Control WindowsForms10.EDIT.app.0.83dd9_r17_ad12
[System.Windows.Forms.SendKeys]::SendWait("^{s}")
start-sleep -Seconds 3
[System.Windows.Forms.SendKeys]::SendWait("%{n}")
start-sleep -Seconds 3
[System.Windows.Forms.SendKeys]::SendWait("%{n}")


start-sleep -Seconds 2
}

Invoke-AU3ControlClick -Title $TPTaskWindowTitle -Control ATL:13B256283
[System.Windows.Forms.SendKeys]::SendWait("%{s}")

}
default {
Write-Output "No task entries for this update"
}


}

Switch ($SOFollowUp){
    $True{

    $GetWorkRequested = Get-AU3ControlText -Title $TPWindowTitle -Control RichTextWndClass1
    $GetBriefWorkDescription = Get-AU3ControlText -Title $TPWindowTitle -Control Edit24
    $GetFullClientName = Get-AU3ControlText -Title $TPWindowTitle -Control Edit19
    $GetClientEmail = Get-AU3ControlText -Title $TPWindowTitle -Control Edit22
    $GetServiceOrderNumber = $ServiceOrderNumber

    $SMTPServer = "mail.cscloud.com"
    $SMTPPort = "587"

    #Service Order / Client Specific
    $EmailServiceOrderNumber = "$GetServiceOrderNumber"
    $EmailClientName = "$ClientOrgName"
    $EmailServiceOrderDescription = "$GetWorkRequested"
    $EmailBriefWorkDescription = $GetBriefWorkDescription
    $EmailClientFirstName = $GetFullClientName.split(" ")[0]
    $EmailClientLastName = $GetFullClientName.split(" ")[1]
    $EmailClientEmail = "$GetClientEmail"

    #Mail Items
    $MailFrom = "Logan Harris <logan@emailhere.com"
    $MailSubject = "$EmailBriefWorkDescription" + " " + "$EMailServiceOrderNumber"
    $MailCC = @("")
    
    $Date = Get-Date
    If ($Date.TimeOfDay -lt "12:00:00.0000000") {$TimeOfDay = "Morning"} else {$TimeOfDay = "Afternoon"}

   
$MailBody = 
"Hello $EmailClientFirstName,`n
Good $TimeOfDay to you! I just wanted to follow up on this Service Order to make sure that I've taken care of everything we need to on this ticket.`n
Let me know if there is any remaining work on this ticket, otherwise please let me know if everything is good to go and I will close this out!

Thank you and have a great $TimeOfDay!`n
Sincerely,
Logan Harris

$EmailSErviceOrderNumber Summary:
Brief Description: $EmailBriefWorkDescription
Work Requested: $EmailServiceOrderDescription
"
    Write-host "$FollowUpType"
    Send-MailMessage -To $EmailClientEmail  -Subject $MailSubject -Body $MailBody -SmtpServer $SMTPServer -From $MailFrom -Cc $MailCC -Port $SMTPPort -Credential $(Get-Credential) -Bcc "logan@emailhere.com"

            }
 

}

Show-AU3WinActivate -WinHandle $TPHandle
    

[System.Windows.Forms.SendKeys]::SendWait("%(f)")
[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Write-Output "SO Saved"

}


}

