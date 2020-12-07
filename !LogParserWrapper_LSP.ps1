# Readme..
	# This script will generate SAMLogon's Status/User/Machine activties summary Excel sheet from netlogon.log(s) using LogParser and Excel via COM objects.
	#		1. To enable Netlogon debug logging: run NLtest. Logging will start right after NLtest for later OS, only restart Netlogon service if debug info is not present. 
	#					Nltest /DBFlag:2080FFFF
	#		2. Output in: %windir%\debug\netlogon.log & netlogon.bak
	#		3. To stop netlogon debug logging: 
	#					Nltest /DBFlag:0
	#		4. No need to delete Netlogon.* since OS continues log essential netlogon info.
	#		5. More info https://docs.microsoft.com/en-us/troubleshoot/windows-client/windows-security/enable-debug-logging-netlogon-service
	#
	# LogParserWrapper_Netlogon.ps1 v0.9 12/4 (skipped rename, keeping netlogon untouch)
	# 	Steps:
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#    			Info on LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)
	#   	2. Copy Netlogon.log & Netlogon.bak from traget's %windir%\debug directory to same directory as this script.
	#					Note: Script will process all *.log & *.bak in script directory when run.
	#   	3. Run script (right click script, 'run with powershell')
	# 
function Invoke-WorkbookTasks { [CmdletBinding()] param ( 
		$WorkBook = 1, $TotalColumn = 4, $SheetTitle = $null, $LogsInfoText = $null
	)
	$iSheet = $Excel.Workbooks[$WorkBook].Worksheets[1]
		$iSheet.Columns.Item($TotalColumn).numberformat = "###,###,###,###,###"
		$iSheet.Name = ($SheetTitle)
			$iSheet.Application.ActiveWindow.SplitRow=1  
			$null = $iSheet.Range("A1").AutoFilter() 
			$iSheet.Application.ActiveWindow.FreezePanes = $true
			$null = $iSheet.Columns.AutoFit() 
		$iSheet.Cells.Item(1,$TotalColumn+1)= "[NOTE]" #--Add log info
			$null = $iSheet.Cells.Item(1,$TotalColumn+1).addcomment()      
			$null = $iSheet.Cells.Item(1,$TotalColumn+1).comment.text($LogsInfoText)
			$iSheet.Cells.Item(1,$TotalColumn+1).comment.shape.textframe.Autosize = $true
		$iChart = $iSheet.shapes.addChart().chart #--Add Chart
			$iChart.chartType = 5
			$iChart.SizeWithWindow = $iChart.HasTitle = $true  
			$iChart.ChartTitle.text = $SheetTitle
			$iChart.ChartArea.Top = $iSheet.Cells.Item(2,$TotalColumn+1).Top
			$iChart.ChartArea.Left = $iSheet.Cells.Item(2,$TotalColumn+1).Left
			$iChart.ChartArea.Width = $iChart.ChartArea.Height = 300
	$Excel.Workbooks[$WorkBook].SaveAs($ScriptPath+'\'+$TimeStamp+'-'+$SheetTitle,51)
		$iCSV = $ScriptPath+'\'+$TimeStamp+'-'+$SheetTitle+'.csv'
		Remove-Item $iCSV
}
#------Main---------------------------------
	$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
	# $null = Get-ChildItem -Path $ScriptDirectory -Filter '*.bak' | Rename-Item -NewName {$_.name -replace '\.bak$', "_bak.log"} -ErrorAction Stop 
		$InFiles = $ScriptPath+'\*.log, '+$ScriptPath+'\*.bak'
		$InputFormat = New-Object -ComObject MSUtil.LogQuery.TextLineInputFormat
		$TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
		$LPQuery = New-Object -ComObject MSUtil.LogQuery
		$OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
#--SamLogon-Machine_
	$OutTitle1 = 'SAM-Logon-Machine'
	$OutFile1 = "$ScriptPath\$TimeStamp-$OutTitle1.csv"
	$Query = @"
		SELECT 
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 				WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'				WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'	WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'				WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, 'from ')) as User, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'from '), 0, 'Returns')) as MachineName, 
			COUNT(*) as Total
		INTO $OutFile1
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Status, User, MachineName ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle1 CSV using Log Parser.." -PercentComplete (30)
	$null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)
#--SamLogon-Domain_
	$OutTitle2 = 'SAM-Logon-Domain'
	$OutFile2 = "$ScriptPath\$TimeStamp-$OutTitle2.csv"
	$Query = @"
		SELECT
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 			WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'			WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'		WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'			WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, '\\')) as Domain, 
			COUNT(*) AS Total 
		INTO $OutFile2
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Domain,Status ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle2 CSV using Log Parser.." -PercentComplete (60)
	$null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)
#--SamLogon-User_
	$OutTitle3 = 'SAM-Logon-User'
	$OutFile3 = "$ScriptPath\$TimeStamp-$OutTitle3.csv"
	$Query = @"
		SELECT 
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 			WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'			WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'		WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'			WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, 'from ')) as User, 
			COUNT(*) AS Total
		INTO $OutFile3
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Status, User ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle3 CSV using Log Parser.." -PercentComplete (90)
	$null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery) 
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) 
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat) 
#---------Find logs's time range Info----------
	$OldestTimeStamp = $NewestTimeStamp = $LogsInfo = $null
	(Get-ChildItem -Path $ScriptPath\* -include ('*.log', '*.bak') ).foreach({
		$FirstLine = (Get-Content $_ -Head 1) -split ' '
		$LastLine  = (Get-Content $_ -Tail 1) -split ' '
			$FirstTimeStamp = [datetime]::ParseExact($FirstLine[0]+' '+$FirstLine[1],"MM/dd HH:mm:ss",$Null)
			$LastTimeStamp = [datetime]::ParseExact($LastLine[0]+' '+$LastLine[1],"MM/dd HH:mm:ss",$Null)
			if ($OldestTimeStamp -eq $null) { $OldestTimeStamp = $NewestTimeStamp = $FirstTimeStamp }
			If ($OldestTimeStamp -gt $FirstTimeStamp) {$OldestTimeStamp = $FirstTimeStamp }
			If ($NewestTimeStamp -lt $LastTimeStamp) {$NewestTimeStamp = $LastTimeStamp }
		$LogsInfo = $LogsInfo + ($_.name+"`n   "+$FirstTimeStamp+' ~ '+$LastTimeStamp+"`t   Log range = "+($LastTimeStamp-$FirstTimeStamp).Totalseconds+" Seconds`n`n")
	})
		$LogTimeRange = ($NewestTimeStamp-$OldestTimeStamp)
		$LogRangeText = ("Netlogon info:`n`n")
		$LogRangeText += ("5E_NO_LOGON_SERVERS Ref: MaxConcurrentApi`n   https://support.microsoft.com/en-us/topic/how-to-do-performance-tuning-for-ntlm-authentication-by-using-the-maxconcurrentapi-setting-92228a96-6874-b52e-1e9f-4a9503ca4fda`n`n") 
		$LogRangeText += ("(NULL) Ref: LsaLookupRestrictIsolatedNameLevel`n   https://support.microsoft.com/en-us/help/818024/how-to-restrict-the-lookup-of-isolated-names-to-external-trusted-domai`n`n") 
		$LogRangeText += ("(NULL) Ref: NeverPing`n   https://support.microsoft.com/en-us/help/923241/the-lsass-exe-process-may-stop-responding-if-you-have-many-external-tr`n`n") 
		$LogRangeText += ("#-------------------------------`n  [Overall EventRange]: "+$OldestTimeStamp+' ~ '+$NewestTimeStamp+"`n  [Overall TimeRange]: "+$LogTimeRange.Days+' Days '+$LogTimeRange.Hours+' Hours '+$LogTimeRange.Minutes+' Minutes '+$LogTimeRange.Seconds+" Seconds `n`n") + $LogsInfo 
#---------Excel--------------------------------
	If (Test-Path $OutFile1) { # Check if LogParser generated CSV.
		$Excel = New-Object -ComObject excel.application  # https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model
		Write-Progress -Activity "Generating Excel worksheets" -PercentComplete (95)
		# $Excel.visible = $true
				$Excel.Workbooks.OpenText($OutFile1)
					Invoke-WorkbookTasks -WorkBook 1 -TotalColumn 4 -SheetTitle $OutTitle1 -LogsInfoText $LogRangeText
				$null = $Excel.Workbooks.Open($OutFile2) 	
					Invoke-WorkbookTasks -WorkBook 2 -TotalColumn 3 -SheetTitle $OutTitle2 -LogsInfoText $LogRangeText
				$null = $Excel.Workbooks.Open($OutFile3) 	
					Invoke-WorkbookTasks -WorkBook 3 -TotalColumn 3 -SheetTitle $OutTitle3 -LogsInfoText $LogRangeText
		$Excel.visible = $true
			$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) 
			# Stop-process -Name Excel 
		} else {
			Write-Host 'No LogParser result found. Please verify log type is Netlogon.log.' -ForegroundColor Red
		}
