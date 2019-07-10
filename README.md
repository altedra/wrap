# wrap
This wrapper allows hidden execution and elevation of any command in Windows. Originally designed to help managing Chocolatey packages over Microsoft Intune

## Usage
This script neeeds to be started by using cscript.exe otherwise the script is dispatched from caller by windows and can not wait for subprocess to complete
Usage is:
```
cscript wrap.vbs [/elevate] [/maxwait:3600] <command_to_run> [<some> <more> <optional> <params>]

Parameters:
/log              Writes log to %temp%\scriptname.log"
/elevate          Run in elevated admin mode. Default is not elevated"
/maxwait:<sec>    Maximum seconds to wait for a program to finish before returning the"
                  controll to the caller process. A value of 0 (zero) distaches the called"
                  process and returns immediately. Default is 3600 seconds"
```
## Known problems
- [ ] Returncode of target command will not be reflected by the return code of wrap.vbs
