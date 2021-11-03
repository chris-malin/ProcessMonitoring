$PingResultPath = ".\PingResults.xls"
$processname = "processname"
$listofmachines = "C:\Monitoring\MiningStatus\listofmachines.txt"

$adminpass = ConvertTo-SecureString “PaintextPassword” -AsPlainText -Force
$adminCred = New-Object System.Management.Automation.PSCredential (“Administrator”, $adminpass)
  
$ResultExcel = new-object -comobject excel.application 
  
if (Test-Path $PingResultPath) 
{ 
$ResultWorkbook = $ResultExcel.WorkBooks.Open($PingResultPath) 
$ExcelWorksheet = $ResultWorkbook.Worksheets.Item(1) 
}
  
else { 
$ResultWorkbook = $ResultExcel.Workbooks.Add() 
$ExcelWorksheet = $ResultWorkbook.Worksheets.Item(1)
}
  
$ResultExcel.Visible = $True
  
$downServerArray = [System.Collections.ArrayList]@()
  
$ExcelWorksheet.Cells.Item(1, 1) = "MachineName"
$ExcelWorksheet.Cells.Item(1, 2) = "Ping Result"
  
$Servers = Get-Content $listofmachines
$count = $Servers.count
 
$ExcelWorksheet.Cells.Item(1,1).Interior.ColorIndex = 6
$ExcelWorksheet.Cells.Item(1,2).Interior.ColorIndex = 6
 
$row=2
  
$Servers | foreach-object{
$pingResult=$null
$Server = $_
$pingResult = Invoke-Command -ComputerName $Server -ScriptBlock {
        @(Get-Process -Name $Using:processname -ErrorAction SilentlyContinue -ErrorVariable ProcessError).count}
                          

if($pingResult -gt 0) {
  
$ExcelWorksheet.Cells.Item($row,1) = $Server
$ExcelWorksheet.Cells.Item($row,2) = "UP"
$ExcelWorksheet.Cells.Item($row,1).Interior.ColorIndex = 17
$ExcelWorksheet.Cells.Item($row,2).Interior.ColorIndex = 4
  
$row++}
else {
  
$ExcelWorksheet.Cells.Item($row,1) = $Server
$ExcelWorksheet.Cells.Item($row,2) = "DOWN"
$ExcelWorksheet.Cells.Item($row,1).Interior.ColorIndex = 17
$ExcelWorksheet.Cells.Item($row,2).Interior.ColorIndex = 3
 
 $arrayID  = $downServerArray.Add($Server)
  
$row++}
}

  
# Format the excel
   
# To wrap the text           
$d = $ExcelWorksheet.UsedRange 
$null = $d.EntireColumn.AutoFit()
  
##to set column width and cell alignment
$ExcelWorksheet.Columns.Item(1).columnWidth = 24
$ExcelWorksheet.Columns.Item(2).columnWidth = 24
$ExcelWorksheet.Columns.HorizontalAlignment = -4131
$ExcelWorksheet.rows.item(1).Font.Bold = $True 
  
##to apply filter
$headerRange = $ExcelWorksheet.Range("a1","n1")
$headerRange.AutoFilter() | Out-Null
 
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$enddate = (get-date -f yyyy-MM-dd).tostring()
$timeStamp= get-date -f MM-dd-yyyy_HH_mm_ss
$ResultExcel.ActiveWorkbook.SaveAs( (get-location).tostring() +"\" +$timeStamp +".xlsx", $xlFixedFormat)
 
$ResultExcel.Workbooks.Close()
$ResultExcel.Quit()
 
 
Write-Host ' Number of processes that are down at the moment : ' $downServerArray.Count
 
if($downServerArray.Count -gt 0)
{
   $Body ='<style>body{background-color:lightgrey;}</style>' 
 $Body = $Body + " <body> Hi World, <br /><br /> The below machines' <b> $($processname) </b> process is not running at the moment. Please check on these machines:  <br /><br /> "
  $Body = $Body + " <table style='border-width: 1px;padding: 0px;border-style: solid;border-color: black;'><th style='background-color:black;color : white; font-weight:bolder;'> Machine </th> "
foreach($item in $downServerArray)
{
     Write-Host $item -ForegroundColor red -BackgroundColor white   
     $Body = $Body + " <tr style='background-color:red;color:white'><td>  $item </td></tr>"
}
$Body = $Body + "</table><br />Regards,<br />IT Management</body>"
$PSEmailServer = "smtp.gmail.com"
$SmtpUser = "SmtpUser"   
$smtpPassword = "eMailpassword"  
$MailTo = "emailaddress"
 
$MailFrom = 'emailaddress@gmail.com'   
$MailSubject = "Machine Process Status - $processname"  
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $SmtpUser, $($smtpPassword | ConvertTo-SecureString -AsPlainText -Force)   

Send-MailMessage -SmtpServer $PSEmailServer -Port 587 -UseSsl -From $MailFrom -To $MailTo -Subject $MailSubject -Body $Body -BodyAsHtml -Credential $Credentials

$SMSText = "Hi World, the following machines' $processname process is down at the moment. $downServerArray . The Restart command will be sent momentarily."

C:\Monitoring\ServerDown-Twilio_v1\ServerDown_v2_Twilio-SMS-Helper.ps1 -AccountSid "AccountSid" -authToken "authToken" -fromNumber "+15555555555" -toNumber  "+15555555555" -message $SMSText

write-Output " `r`n Custom Message : Machine Status Email Sent" 


foreach($item in $downServerArray)
{
	Restart-Computer -ComputerName $item -Credential $adminCred -Force
}

foreach($item in $downServerArray)
{
     Write-Host $item -ForegroundColor red -BackgroundColor white   
     "Restart request for Machine " + " <tr style='background-color:red;color:white'><td>  $item </td></tr>" + " sent."
}

}

else{
Write-Host 'All Machines' processes are up' -ForegroundColor red -BackgroundColor white   
}
