# Readme..
	# This script will generate LsarLookup*'s IP/SiD/User/Process summary Excel sheet from LSP log(s) using LogParser and Excel via COM objects.
	#		1. To enable LSP logging, save below to .REG and run, logging starts as soon as REG set. 
		#			Windows Registry Editor Version 5.00 
		#			[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa] 
		#			"LspDbgInfoLevel"=dword:40000800 
		#			"LspDbgTraceOptions"=dword:00000001 
	#		2. Output in %windir%\debug\lsp.log & lsp.bak
	#		3. To stop, delete REG in (1) and files in (2)
	#		4. More info http://technet.microsoft.com/en-us/library/ff428139(v=WS.10).aspx#BKMK_LsaLookupNames 
	#
	# LogParserWrapper_LSP.ps1 v0.7 11/14 (auto convert UniCode > Ascii)
	#		Steps: 
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#     	Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
	#   	2. Copy LSP.log & LSP.bak from target's %windir%\debug directory to same directory as this script.
	#     		Note1: Script will rename LSP.bak to LSP_bak.log.
	#					Note2: Script will process all *.log(s) in script directory when run.
	#   	3. Run script
	# 
#------Main---------------------------------
$ErrorActionPreference = "SilentlyContinue"
	$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
	$null = Get-ChildItem -Path $ScriptPath -Filter '*.bak' | Rename-Item -NewName {$_.name -replace '\.bak$', "_bak.log"} -ErrorAction Stop
	$TotalSteps = ((Get-ChildItem -Path $ScriptPath -Filter '*.log').count)+4
  $Step=1
	(Get-ChildItem -Path $ScriptPath -Filter '*.log').foreach({
		Write-Progress -Activity "Convert $_ from UniCode to Ascii" -PercentComplete (($Step++/$TotalSteps)*100)
		If ((Get-Content -Encoding byte -ReadCount 2 -TotalCount 2 -Path $_)[1] -eq 0){ 
			Get-Content $_ -Encoding Unicode |  Set-Content "$ScriptPath\Tmp-$_" -Encoding ascii 
			Remove-Item $_ 
			Rename-Item "$ScriptPath\Tmp-$_" "$ScriptPath\$_" 
		}
	})
	$InFiles = $ScriptPath+'\*.log'
	$InputFormat = New-Object -ComObject MSUtil.LogQuery.TextLineInputFormat
	$TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
	$OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
	$OutTitle = 'LSP-IP_Sid_Name_App'
	$OutFile = "$ScriptPath\$TimeStamp-$OutTitle.csv"
		Write-Progress -Activity "Generating CSV"  -PercentComplete (($Step++/$TotalSteps)*100)
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
		Write-Progress -Activity "Generating $OutTitle report" -PercentComplete (($Step++/$TotalSteps)*100)
		$LPQuery = New-Object -ComObject MSUtil.LogQuery
		$null = $LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery) 
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) 
		$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat) 
#---------Find logs's time range Info----------
	$OldestTimeStamp = $NewestTimeStamp = $LogsInfo = $null
	(Get-ChildItem -Path $ScriptPath -Filter '*.log').foreach({
    $FirstLine = (Get-Content $_ -Head 1) -split ' '
      if ($FirstLine[1] -eq $null) { $FirstLine = (Get-Content $_ -Head 2)[1] -split ' ' } # incase first line is blank
		$LastLine  = (Get-Content $_ -Tail 1) -split ' '
		$FirstTimeStamp = [datetime]::ParseExact($FirstLine[0]+' '+$FirstLine[1],"[MM/dd HH:mm:ss]",$Null)
		$LastTimeStamp = [datetime]::ParseExact($LastLine[0]+' '+$LastLine[1],"[MM/dd HH:mm:ss]",$Null)
			if ($OldestTimeStamp -eq $null) { $OldestTimeStamp = $NewestTimeStamp = $FirstTimeStamp }
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
If (Test-Path $OutFile) { # Check if LP generated CSV.
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
		$iCSV = $ScriptPath+'\'+$TimeStamp+'-'+$OutTitle+'.csv'
		Remove-Item $iCSV
	$Excel.visible = $true
	$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
	# Stop-process -Name Excel 
} else {
	Write-Host 'No LogParser CSV found. Did we copied LSP logs? If so, did we converted LSP.log/bak to ANSI format using NotePad or NotePad++?' -ForegroundColor Red
}
