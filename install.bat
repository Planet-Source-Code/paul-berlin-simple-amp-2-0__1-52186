@echo off
echo This will register the ccrpftv6.ocx control. 
echo Press any key to continue or Ctrl+C to exit.
pause
copy ccrpftv6.ocx %windir%\system32 /Y
regsvr32 ccrpftv6.ocx /s
echo Done.
pause