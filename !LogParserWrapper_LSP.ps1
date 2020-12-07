# Readme..
	# This script will generate LsarLookup*'s IP/SiD/User/Process summary Excel sheet from LSP log(s) using LogParser and Excel via COM objects.
	#		1. To enable LSP logging, save below to .REG and run, logging starts as soon as REG set. 
		#			Windows Registry Editor Version 5.00 
		#			[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa] 
		#			"LspDbgInfoLevel"=dword:40000800 
		#			"LspDbgTraceOptions"=dword:00000001 
	#		2. Resulting logs in %windir%\debug\lsp.log & lsp.bak
	#		3. To stop, delete REG in (1) and files in (2)
	#		4. More info http://technet.microsoft.com/en-us/library/ff428139(v=WS.10).aspx#BKMK_LsaLookupNames 
	#
	# LogParserWrapper_LSP.ps1 v0.8 12/4 (skip lsp logs modifications by using *.bak and LogParser's iCodePage)
	#		Steps: 
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#     	Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
	#   	2. Copy LSP.log & LSP.bak from target's %windir%\debug directory to same directory as this script.
	#					Note: Script will process all *.log & *.bak in script directory when run.
	#   	3. Run script
	# 
#------Main---------------------------------
$ErrorActionPreference = "SilentlyContinue"
	$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
	$TotalSteps = 4
	$Step=1
	$InFiles = $ScriptPath+'\*.log, '+$ScriptPath+'\*.bak'
	$InputFormat = New-Object -ComObject MSUtil.LogQuery.TextLineInputFormat
		$InputFormat.iCodePage=-1
	$TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
	$OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
	$OutTitle = 'LSP-IP_Sid_Name_App'
	$OutFile = "$ScriptPath\$TimeStamp-$OutTitle.csv"
		$Query = @"
			SELECT Top 1000
				EXTRACT_SUFFIX(SUBSTR(TEXT, INDEX_OF (TEXT, 'Network Address = '), STRLEN(TEXT)), 0, '= ') as Remote_IP,
				EXTRACT_SUFFIX(SUBSTR(TEXT, INDEX_OF (TEXT, 'Sids[ 0 ] = '), STRLEN(TEXT)), 0, '= ') as LookupSID,
				EXTRACT_SUFFIX(SUBSTR(TEXT, INDEX_OF (TEXT, 'Names[ 0 ] = '), STRLEN(TEXT)), 0, '= ') as LookupName,
				EXTRACT_SUFFIX(SUBSTR(TEXT, INDEX_OF (TEXT, 'Process Name = '), STRLEN(TEXT)), 0, '= ') as Process,
				Count (*) as Total
			INTO $OutFile
			FROM $InFiles
			Where 
			Index_of(text, 'Network Address')>0 or Index_of(text,'Sids[ 0 ]')>0 or Index_of(text,'Names[ 0 ]')>0 or Index_of(text,'Process Name = ')>0 
			Group By 
				Remote_IP, LookupSID, LookupName, Process
			Order By 
				Total, Remote_IP, LookupSID, LookupName, Process DESC
"@
		Write-Progress -Activity "Generating $OutTitle CSV using LogParser" -PercentComplete (($Step++/$TotalSteps)*100)
		$LPQuery = New-Object -ComObject MSUtil.LogQuery
		$null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery) 
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) 
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat) 
#---------Find logs's time range Info----------
	$OldestTimeStamp = $NewestTimeStamp = $LogsInfo = $null
	(Get-ChildItem -Path $ScriptPath\* -include ('*.log', '*.bak') ).foreach({
    $FirstLine = (Get-Content $_ -Head 1 -Encoding Unicode) -split ' '
		$LastLine  = (Get-Content $_ -Tail 1 -Encoding Unicode) -split ' '
		$FirstTimeStamp = [datetime]::ParseExact($FirstLine[0]+' '+$FirstLine[1],"[MM/dd HH:mm:ss]",$Null)
		$LastTimeStamp = [datetime]::ParseExact($LastLine[0]+' '+$LastLine[1],"[MM/dd HH:mm:ss]",$Null)
			If ($OldestTimeStamp -eq $null) { $OldestTimeStamp = $NewestTimeStamp = $FirstTimeStamp }
			If ($OldestTimeStamp -gt $FirstTimeStamp) {$OldestTimeStamp = $FirstTimeStamp }
			If ($NewestTimeStamp -lt $LastTimeStamp) {$NewestTimeStamp = $LastTimeStamp }
		$LogsInfo = $LogsInfo + ($_.name+"`n   "+$FirstTimeStamp+' ~ '+$LastTimeStamp+"`t   Log range = "+($LastTimeStamp-$FirstTimeStamp).Totalseconds+" Seconds`n`n")
	})
		$LogTimeRange = ($NewestTimeStamp-$OldestTimeStamp)
		$LogRangeText = ("LSP info:`n`n")
		$LogRangeText += ("Ref: LsaLookupCacheMaxSize`n  https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/ff428139(v=ws.10)`n`n") 
		$LogRangeText += ("Ref: Well-Known SID`n  https://docs.microsoft.com/en-us/troubleshoot/windows-server/identity/security-identifiers-in-windows`n`n") 
		$LogRangeText += ("#-------------------------------`n  [Overall EventRange]: "+$OldestTimeStamp+' ~ '+$NewestTimeStamp+"`n  [Overall TimeRange]: "+$LogTimeRange.Days+' Days '+$LogTimeRange.Hours+' Hours '+$LogTimeRange.Minutes+' Minutes '+$LogTimeRange.Seconds+" Seconds `n`n") + $LogsInfo 
#---------Excel--------------------------------
If (Test-Path $OutFile) { # Check if LogParser generated CSV.
	$Excel = New-Object -ComObject excel.application  # https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model
	Write-Progress -Activity "Generating Excel worksheets" -PercentComplete (($Step++/$TotalSteps)*100)
		# $Excel.visible = $true
		$Excel.Workbooks.OpenText("$OutFile")
		$Sheet = $Excel.Workbooks[1].Worksheets[1]
			$null = $Sheet.Range("A1").AutoFilter()
			$Sheet.Application.ActiveWindow.SplitRow=1  
			$Sheet.Application.ActiveWindow.FreezePanes = $true
			$Sheet.Columns.Item(1).columnwidth = $Sheet.Columns.Item(2).columnwidth = $Sheet.Columns.Item(3).columnwidth = $Sheet.Columns.Item(4).columnwidth = $Sheet.Columns.Item(5).columnwidth = 25
			$Sheet.Columns.Item(5).numberformat = "###,###,###,###,###"
			$Sheet.Name = $OutTitle
		$Sheet.Cells.Item(1,6)= "[NOTE]" #--Add log info
			$null = $Sheet.Cells.Item(1,6).addcomment()
			$null = $Sheet.Cells.Item(1,6).comment.text($LogRangeText)
			$Sheet.Cells.Item(1,6).comment.shape.textframe.Autosize = $true
		$Chart = $Sheet.shapes.addChart().chart # https://codewala.net/2016/09/20/how-to-create-excel-chart-using-powershell-part-1/, https://codewala.net/2016/09/23/how-to-create-excel-chart-using-powershell-part-2/, https://codewala.net/2016/09/27/how-to-create-excel-chart-using-powershell-part-3/
			$Chart.chartType = -4120 
			$Chart.SizeWithWindow = $Chart.HasTitle=$true  
			$Chart.ChartTitle.text = $OutTitle
			$Chart.ChartArea.Top = $Sheet.Cells.Item(2,6).Top
			$Chart.ChartArea.Left = $Sheet.Cells.Item(2,6).Left
			$Chart.ChartArea.Width = $Chart.ChartArea.Height = 300
		$Excel.Workbooks[1].SaveAs($ScriptPath+'\'+$TimeStamp+'-'+$OutTitle,51)
		Remove-Item ($ScriptPath+'\'+$TimeStamp+'-'+$OutTitle+'.csv')
	$Excel.visible = $true
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
	# Stop-process -Name Excel 
} else {
	Get-ChildItem $ScriptPath 
	Write-Host "`nNo LogParser CSV generated. Did we copied LSP logs to the same directory as this script?" -ForegroundColor Red
}
