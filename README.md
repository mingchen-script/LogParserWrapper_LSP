# LogParserWrapper_LSP
Convert LSP logs to Excel for more insight toward LsarLookup* workload. 
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
