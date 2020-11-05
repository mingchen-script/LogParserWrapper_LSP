# LogParserWrapper_LSP
LSP logs to Excel chart. 
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
