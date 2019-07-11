' =====================================================
' This script neeeds to be started by using cscript.exe
' Otherwise the script is dispatched from caller by
' windows and can not wait for subprocess to complete
' =====================================================
Option Explicit
Const forAppending = 8
dim fs, namedArgs, comspec, wShell, tempDir, lockFile, strID, maxWait, logFileAbsPath, logLine, debug, v
v = "0.0.1.001"

set namedArgs = wScript.arguments.Named
set wShell = CreateObject("wScript.Shell" )
set fs 		= CreateObject("Scripting.FileSystemObject")

comspec 	= wShell.expandEnvironmentStrings("%comspec%")
tempDir 	= wShell.expandEnvironmentStrings("%temp%")
logFileAbsPath = fs.buildPath(tempDir, wScript.scriptName & ".log") ' need to be set before the first call to log()

log("Starting version: " & v & " " & wScript.scriptName)
logLine =  ">> """ & logFileAbsPath & """ 2>&1"

Randomize()
strID		= Rnd() & "_" & unixTime()
lockFile 	= fs.buildPath(tempDir,strID) & ".lck"
if namedArgs.exists("lock") then lockFile = namedArgs.item("lock")
if namedArgs.exists("debug") then debug = 1 else debug = 0

if namedArgs.exists("?") then printUsage()
if namedArgs.exists("h") then printUsage()
if namedArgs.exists("maxwait") then maxWait = namedArgs.item("maxwait") else maxWait = 3600
if namedArgs.exists("elevate") then elevate()

runCmd()
log("Ending " & wScript.scriptName)

' =====================================================
' Script Functions
' =====================================================
function prepareArgs()
	dim args, oArg	
	for each oArg in wScript.arguments 
		if (inStr(oArg, "/log") = 0) and _
			(inStr(oArg, "/lock") = 0) and _
			(inStr(oArg, "/maxwait") = 0) and _
			(inStr(oArg, "/debug") = 0) and _
			(inStr(oArg, "/elevate") = 0) then
			if inStr(oArg, " ") then oArg = """" & oArg & """"
			args = args & " " & oArg
		end if
	next
	if namedArgs.exists("log") then args = args & " " & logLine
	prepareArgs = Trim(args)
end function

sub runCmd()
	dim args, cORk
	args = prepareArgs()
	if debug then cORk = " /k" else cORk = " /c"
	log("Executing " & args)
	getLock()
	CreateObject("Shell.Application").shellExecute comspec _
		, cORk & " title " & strID & " && " & args, "", "", debug
	if maxWait <> 0 then wScript.sleep 1000 ' Give cmd some ms to set the name
	waitForProcess strID, maxWait
	freeLock()
end sub

sub elevate()
	dim myArgs : myArgs = "/lock:""" & lockFile & """ " & prepareArgs()
	if debug then myArgs = "/debug " & myArgs
	on error resume next
	CreateObject("wScript.Shell").RegRead("HKEY_USERS\s-1-5-19\") ' check if already elevated
	If err.number <> 0 Then ' check if already elevated --> is err in case of not elevated
		log("Elevating " & wScript.scriptName)
		CreateObject("Shell.Application").shellExecute wScript.fullName _
			, "" & wScript.scriptFullName & " " & myArgs, "", "runas", debug
		if maxWait <> 0 then wScript.sleep 1000 ' Give elevated process some ms to set the lock
		waitForFreeLock maxWait
		log("Ending after elevation" & wScript.scriptName)
		wScript.quit
	end if
end sub

function getLock()
	getLock = not isLocked()
	if getLock then fs.createTextFile(lockFile)
end function

function isLocked()
	isLocked = fs.fileExists(lockFile)
end function

function freeLock()
	freeLock = isLocked()
	if freeLock then fs.deleteFile lockFile, true
end function

function waitForFreeLock(maxSec)
	dim locked, start: start = unixTime()
	waitForFreeLock = true
	do while start + maxSec > unixTime()
		if not isLocked() then exit do
		wScript.sleep 200
	loop
	waitForFreeLock = false
end function

function waitForProcess(strID, maxSec)
	dim colItems, WMIService, start
	start = unixTime()
	
	set WMIService = GetObject( "winmgmts://./root/cimv2" )
	do while start + maxSec > unixTime()
		set colItems = WMIService.execQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & strID & "%'")
		if colItems.count = 0 then exit do ' Cmd is not or no more running. job done... exit
		wScript.sleep 200
	loop
end function

function log(logText)
	dim logFile
	if namedArgs.exists("log") then
		Set logFile = fs.openTextFile(logFileAbsPath, forAppending, True)
		logFile.writeLine getDateTime & ": " & logText
		logFile.close
	end if
end function

function unixTime()
    unixTime = DateDiff("S", "1/1/1970", Now())
end function

Function getDateTime()
	dim s, dt: dt = Now()
    s = datePart("yyyy", dt)
    s = s & "-" & right("0" & datePart("m",dt),2)
    s = s & "-" & right("0" & datePart("d",dt),2)
    s = s & " " & right("0" & datePart("h",dt),2)
    s = s & ":" & right("0" & datePart("n",dt),2)
    s = s & ":" & right("0" & datePart("s",dt),2)
    getDateTime = s
End Function

sub printUsage()
	if LCase(right(wScript.fullName, 12)) = "\cscript.exe" then 
		wScript.echo "This script neeeds to be started by using cscript.exe"
		wScript.echo "Otherwise the script is dispatched from caller by"
		wScript.echo "windows and can not wait for subprocess to complete"
		wScript.echo " "
		wScript.echo "Usage is:"
		wScript.echo "cscript " & wScript.scriptName & " [/elevate] [/maxwait:3600] <command_to_run> [<some> <more> <optional> <params>]"
		wScript.echo " "
		wScript.echo "Parameters:"
		wScript.echo "/log              Writes log to %temp%\scriptname.log"
		wScript.echo "/elevate          Run in elevated admin mode. Default is not elevated"
		wScript.echo "/maxwait:<sec>    Maximum seconds to wait for a program to finish before returning the"
		wScript.echo "                  controll to the caller process. A value of 0 (zero) distaches the called"
		wScript.echo "                  process and returns immediately. Default is 3600 seconds"
	else 
		MsgBox "Please start this script by using cscript.exe"
	end if
	wScript.quit
end sub
