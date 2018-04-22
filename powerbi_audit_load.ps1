Param(
   [Parameter(Mandatory=$true)]
   [string]$StartDate,
   [string]$EndDate
) #end param
if(-not($StartDate)) { Throw "The StartDate cannot be blank. Date is the format of MM/DD/YYYY HH:MM:SS AM" }
if(-not($EndDate)) { Throw "The EndDate cannot be blank. Date is the format of MM/DD/YYYY HH:MM:SS AM" }

Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

#Import-PSSession $Session
$result = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType PowerBI  -ResultSize 5000

$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = "Data Source=77ST-GPSQL01;Initial Catalog=EADB;Integrated Security=SSPI;"
$conn.open()

$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.connection = $conn

ForEach($audit in $result)
{
    # DEBUG
    # All Columns 
    #   PSComputerName	RunspaceId	PSShowComputerName	RecordType	CreationDate	UserIds	Operations	AuditData	ResultIndex	ResultCount	Identity	IsValid	ObjectState
    #Write-Host $audit.CreationDate, $audit.UserIds, $audit.Operations, $audit.AuditData
    $cmd.commandtext = "INSERT INTO powerbi.AuditData (AuditIdentity, CreationDate, UserIds, Operations, ResultIndex, AuditData) VALUES('{0}','{1}','{2}','{3}','{4}','{5}')" -f
        $audit.Identity, [datetime]$audit.CreationDate, $audit.UserIds, $audit.Operations, $audit.ResultIndex, $audit.AuditData
    $cmd.executenonquery()
}

$From = "from@email.com"
$To = "to@email.com"
$Cc = ""
$Attachment = ""
$Subject = "Power BI Audit Report"
$Body = "Power BI Audit Data from <b> " + $StartDate + "</b> to <b>" + $EndDate + "</b> with total number of records of <b>" + $result.count + "</b>." 
$SMTPServer = "smtp.server.com"
Send-MailMessage -From $From -to $To -Subject $Subject -BodyAsHtml $Body -SmtpServer $SMTPServer -Credential $UserCredential

$conn.close()
