# This script will generate an IP/SiD/User/Process summary Excel sheet from LSP log using LogParser and Excel, for top talkers analysis. 
#		1. To enable LSP logging, save below to .REG and run, logging starts as soon as REG set. 
#			Windows Registry Editor Version 5.00 
#			[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa] 
#			"LspDbgInfoLevel"=dword:40000800 
#			"LspDbgTraceOptions"=dword:00000001 
#		2. Output in %windir%\debug\lsp.log & lsp.bak
#		3. To stop, delete REG in (1) and files in (2)
#		4. More info http://technet.microsoft.com/en-us/library/ff428139(v=WS.10).aspx#BKMK_LsaLookupNames 
#
# LogParserWrapper_LSP.ps1 v0.5 10/31
#		Steps: (From toolbox with Excel installed.)
#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
#     	Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
#   	2. Copy LSP.log & LSP.bak from target's %windir%\debug directory to same directory as this script.
#     	Note1: Script will rename LSP.bak to LSP_bak.log.
#				Note2: LogParser will process *.log(s) in script directory.
#			3. IMPORTANT: Open each .log and .bak file with notepad, insert a line (or any edit), save the change to convert log format LogParser can read.
#				**In NotePad++, select 'Encoding' > 'Convert to ANSI', file size should reduce by half. (Consider creating a Micros if you review LSP logs often.)
#   	4. Run script
# 
#------Main---------------------------------
	$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
	Get-ChildItem -Path $ScriptDirectory -Filter '*.bak' | Rename-Item -NewName {$_.name -replace '\.bak$', "_bak.log"} -ErrorAction Stop | Out-Null
		$InFiles = $ScriptPath+'\*.log'
		$InputFormat = New-Object -ComObject MSUtil.LogQuery.TextLineInputFormat
		$CurrentTime = Get-Date
		$OutFile = $ScriptPath+'\LSP_Combo-'+[string]$CurrentTime.Year+'-'+([string]$CurrentTime.Month).PadLeft(2,'0')+'-'+([string]$CurrentTime.Day).PadLeft(2,'0')+'-'+([string]$CurrentTime.Hour).PadLeft(2,'0')+'-'+([string]$CurrentTime.Minute).PadLeft(2,'0')+'-'+([string]$CurrentTime.Second).PadLeft(2,'0')+'.csv'
		$OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
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
		$LPQuery = New-Object -ComObject MSUtil.LogQuery
		$LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)| Out-Null
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery) | Out-Null
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) | Out-Null
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat) | Out-Null
	#---------Excel--------------------------------
If (Test-Path $OutFile) { # Check if LP generated CSV.
	$Excel = New-Object -ComObject excel.application  # https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model
		$Excel.visible = $true
		$Excel.Workbooks.OpenText("$OutFile")
		$Sheet = $Excel.Workbooks[1].Worksheets[1]
			$Sheet.Range("A1").AutoFilter() | Out-Null
			$Sheet.Application.ActiveWindow.SplitRow=1  
			$Sheet.Application.ActiveWindow.FreezePanes = $true
			$Sheet.Columns.Item(1).columnwidth = $Sheet.Columns.Item(2).columnwidth = $Sheet.Columns.Item(3).columnwidth = $Sheet.Columns.Item(4).columnwidth = $Sheet.Columns.Item(5).columnwidth = 15
			$Sheet.Columns.Item(5).numberformat = "###,###,###,###,###"
		$Chart = $Sheet.shapes.addChart().chart # https://codewala.net/2016/09/20/how-to-create-excel-chart-using-powershell-part-1/, https://codewala.net/2016/09/23/how-to-create-excel-chart-using-powershell-part-2/, https://codewala.net/2016/09/27/how-to-create-excel-chart-using-powershell-part-3/
			$Chart.chartType = 5 # ChartType https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.core.xlcharttype?view=office-pia
			$Chart.SizeWithWindow = $Chart.HasTitle=$true  
			$Chart.ChartTitle.text = "Top 20"
			$Chart.setSourceData($Sheet.range('e2:e22'))
			$Chart.ChartArea.Left = $Chart.ChartArea.Width = $Chart.ChartArea.Height = 450
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
} else {
	Write-Host 'No LogParser CSV found. Did we copied LSP logs? If so, did we converted LSP.log/bak to ANSI format using NotePad or NotePad++?' -ForegroundColor Red
}
