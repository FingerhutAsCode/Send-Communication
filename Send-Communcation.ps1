
<#PSScriptInfo

.VERSION 0.1.0

.GUID 399d5314-fa07-45e4-a7c1-10a80469520e

.AUTHOR FingerhutAsCode

.COMPANYNAME

.COPYRIGHT

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/FingerhutAsCode/Send-Communication

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#
.SYNOPSIS
The Send-Communcation cmdlet is designed to send targeted communication using a communication templates sent from central notification account such as a helpdesk or communcations group.

.DESCRIPTION

The Send-Communcation cmdlet is designed to send targeted communication using a communication templates sent from central notification account such as a helpdesk or communcations group.

The script works similar to a Mail-Merge and can replace the following values with matching values from the RecepientList.CSV

--- Template Value ---- CSV Value ---
    #Username           Username
    #FirstName          FirstName
    #LastName           LastName
    #Email              Email
    #UPN                UPN
    #Password           Password
    #Extension          Extension
    #PhoneNumber        PhoneNumber

.NOTES

How to create a new communication template:

1. Copy an existing template folder and the contents
2. Delete all items in the Attachements and Output folders
3. Update CommunicationTemplate.html
4. Update Recepientlist.csv
5. Update Subject.txt

.INPUTS

The Send-Communication cmdlet is reliant on four items to exist in the Communication Template Path provided at the time of execution.

1. CommunicationTemplate.html which should contain the body of the email to be sent
2. RecepientList.csv which should contain at minimum a column labeled [Email] which contains a list of email addresses to send the communication to
3. Subject.txt which should contain the subject line of the email
4. An Output folder, generally empty, which a copy of each sent email will be stored

Optional:

If attachments are required such as images embedded in the HTML or documents to be attached to the email an Attachments folder should exist which contains all items to be attached to the communication.

.OUTPUTS

The Send-Communication cmdlet will produce a copy of each email sent stored in HTML format into the Output folder within the Communication Template Path.


.PARAMETER CommunicationTemplatePath

The path to the folder the communication key files are stored within

.PARAMETER SendToRecepients

This parameter sets the communication to be sent to the targeted recepients.  If this parameter is not used the communication will be sent to the email address indicated in the TestEmailAddress parameter.

.PARAMETER Credentials

Credentials of the user account authorized to send as the desired mailbox

.PARAMETER ContainsAttachments

This parameter indicates if there are attachments to be included in the communication

.PARAMETER TestEmailAddress

This parameter sets the email address that test emails should be sent to and should be used whenever the SendToRecepients parameter is not set.

.PARAMETER FromEmailAddress

This parameter sets the email address the communication will send from.

.EXAMPLE

PS >.\Send-Communcation.ps1 -CommunicationTemplatePath .\Communications\Example_Communication -FromEmailAddress "ITSystems@fingerhut.us" -TestEmailAddress "recepient@test.com" -Credentials $Credentials

This example will send the Example_Communication to the TestEmailAddress recepient@test.com using the Credentials stored in $Credentials variable.

.EXAMPLE

PS >.\Send-Communcation.ps1 -CommunicationTemplatePath .\Communications\Example_Communication -FromEmailAddress "ITSystems@fingerhut.us" -TestEmailAddress "recepient@test.com" -ContainsAttachments -Credentials $Credentials

This example will send the Example_Communication to the TestEmailAddress recepient@test.com using the Credentials stored in $Credentials variable and include any attachments stored in the Communcation Template Path.

#>

[CmdletBinding()]
param(
    [Alias("Communication")]
    [Parameter(Mandatory=$true)]
    [string] $CommunicationTemplatePath,

    [Parameter(Mandatory=$false)]
    [switch] $SendToRecepients = $false,

    [Parameter(Mandatory=$true)]
    [string] $FromEmailAddress,

    [Parameter(Mandatory=$true)]
    [PSCredential]$Credentials = $null,

    [Parameter(Mandatory=$false)]
    [switch] $ContainsAttachments = $false,

    [Parameter(Mandatory=$false)]
    [string] $TestEmailAddress
)

$RecepientsList = "$CommunicationTemplatePath\RecepientList.csv"
$SubjectLine = Get-Content -Path "$CommunicationTemplatePath\Subject.txt"
$OutputPath = "$CommunicationTemplatePath\Output"
$TemplateItemsList = @("CommunicationTemplate.html", "RecepientList.csv", "Subject.txt")
$TestCountMax = 5
[xml]$XML = Get-Content .\Config.xml -Raw

# Ensure Output Path Exists
If (-not(Test-Path($OutputPath)))
{
    New-Item -Type Directory -Path $OutputPath
}

# Load Template
$Template = (Get-Content -Path "$CommunicationTemplatePath\CommunicationTemplate.html" -Raw)

# Load Receipients List
$CSV = Import-Csv $RecepientsList

if (-not($SendToRecepients)) {
    Write-Verbose "Executing as a test communcation"
    $TestRecepientsArray = @()
    if ($CSV.count -lt $TestCountMax) {
        Write-Verbose "Number of records imported from RecepientsList.csv [$($CSV.count)]"
        $TestCountMax = $CSV.count
        Write-Warning "The number of addresses in the RecepientsList.csv is less than the TestCountMax and therefore only [$TestCountMax] test communications will be sent"
    }
    $TestRecepientsArray = for ($i=0; $i -lt $TestCountMax; $i++) {
        $Record = $CSV[$i]
        $Record.mail = "$TestEmailAddress"
        Write-Host "Record is $Record"
        $Record
    }
    $CSV = $TestRecepientsArray
}

foreach ($Record in $CSV)
{
    # Process Template
    $Output = $Template
    
    Write-Verbose "Template [$Template]"

    foreach ($Value in $XML.Config.ReplacementValues.Value) {
        $Output = $Output -replace "#$($Value.Name)", "$($Record.$($Value.Replacement))"
    }

    Write-Verbose "Output [$Output]"

    $EmailHashTable = @{
        To = "$($Record.mail)"
        From = $FromEmailAddress
        Subject = $SubjectLine
        Body = $Output
        BodyAsHtml = $True
        usessl = $True
        Port = 587
        SmtpServer = "smtp.office365.com"
        Credential = $Credentials
    }

    if ($ContainsAttachments) {
        $AttachmentsList = Get-ChildItem -Path $CommunicationTemplatePath -File | Where-Object {$_.Name -notin $TemplateItemsList}
        Write-Verbose "Attachments List [$AttachmentsList]"
        $EmailHashTable = $EmailHashTable + @{
            Attachments = $AttachmentsList 
        }
    }

    Write-Host "Sending Communication to $($Record.mail)"
    Send-MailMessage  @EmailHashTable
    Start-Sleep -s  2
}
