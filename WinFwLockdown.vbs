On Error Resume Next
'Option Explicit
'
'
' script to enable windows 7 firewall and lock it down
'  to prevent only outbound tcp 80, 8080, 443, 53 and udp 53
'
'  It will also allow ANY inbound connections
'
' Author: Mick Grove
' Date: 2/15/2013
' 
' version: 0.9.5
'

Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim WshShell : Set WshShell = CreateObject("Wscript.shell")

Dim strComputer : strComputer = "."
Dim oWMIService : Set oWMIService = GetObject("winmgmts:\\" & strComputer & "\root\default")
Const HKLM = &H80000002
Dim dq
dq = chr(34) ' double quotes


' the purpose of this variable (bCurrentlyUpdatingRegistry) is to prevent an endless loop of reg key notification
'   while the script itself updates the registry
Dim bCurrentlyUpdatingRegistry
bCurrentlyUpdatingRegistry = False ' default

Dim sSystem32
sSystem32 = WshShell.ExpandEnvironmentStrings("%windir%") & "\System32\"

Call SetRegistryKeys
Call StartWinFirewall
Call AddFirewallRules

WshShell.Run dq & sSystem32 & "net.exe" & dq & " set domainprofile state on", "0", True

LogMessage "  [+] Beginning registry monitoring of firewall service"
Dim iFwEnabled
Dim objWMIService 
Dim colRunningServices
Dim objService

Do
	iFwEnabled = WshShell.RegRead("HKLM\" _
		& "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\EnableFirewall")
	
	If iFwEnabled = 0 Then
		LogMessage "  [+] Domain Profile was modified, restoring it."
		If Not bCurrentlyUpdatingRegistry Then
			bCurrentlyUpdatingRegistry = True
			
			Call SetRegistryKeys
			Call StartWinFirewall
			Call AddFirewallRules
			
			WshShell.Run dq & sSystem32 & "net.exe" & dq & " set domainprofile state on", "0", True
			bCurrentlyUpdatingRegistry = False
		End If
	End If
	
	WScript.Sleep(30000) '30 seconds
		
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		
	Set colRunningServices = objWMIService.ExecQuery _
		("select State from Win32_Service where Name = 'MpsSvc'")
		
	For Each objService in colRunningServices
		If objService.State = "Stopped" Then
			LogMessage "  [+] Restarting MpsSvc service"
			If Not bCurrentlyUpdatingRegistry Then
				bCurrentlyUpdatingRegistry = True
				
				Call SetRegistryKeys
				Call StartWinFirewall
				Call AddFirewallRules
				
				WshShell.Run dq & sSystem32 & "net.exe" & dq & " set domainprofile state on", "0", True
				bCurrentlyUpdatingRegistry = False
			End If
		End If
	Next
	
	WScript.Sleep(30000) ' 30 seconds
Loop Until 2 < 1


WScript.Quit(1)

'
'
' Functions below
'
'==============================================
'
'

Function LogMessage(sMsg)
	'WScript.Echo sMsg
End Function

Function SetRegistryKeys
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\" _
		,"DisableNotifications", "0", "REG_DWORD" 
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\" _
		,"EnableFirewall", "1", "REG_DWORD" 
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\" _
		,"DefaultOutboundAction", "0", "REG_DWORD"
	'===
	DeleteRegKey("SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\" _
		& "DoNotAllowExceptions")
	'===
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\Logging\" _
		,"LogDroppedPackets", "0", "REG_DWORD"
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\Logging\" _
		,"LogFilePath", "%systemroot%\system32\LogFiles\Firewall\pfirewall.log", "REG_MULTI_SZ"
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\Logging\" _
		,"LogFileSize", "4096", "REG_DWORD"
	WriteReg "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\DomainProfile\Logging\" _
		,"LogSuccessfulConnections", "0", "REG_DWORD"
	'===
	WriteReg "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\" _
		,"AllowLocalPolicyMerge", "1", "REG_DWORD"
	WriteReg "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\" _
		,"AllowLocalIPsecPolicyMerge", "1", "REG_DWORD"
	WriteReg "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\" _
		,"DefaultOutboundAction", "0", "REG_DWORD"
	WriteReg "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\" _
		,"DefaultInboundAction", "1", "REG_DWORD"
	WriteReg "SOFTWARE\Policies\Microsoft\WindowsFirewall\DomainProfile\" _
		,"EnableFirewall", "1", "REG_DWORD"

End Function

Function AddFirewallRules
	Dim sDescr
	
	If Not FirewallRuleExists("Disable Outbound TCP 80") Then
		sDescr = "Disable Outbound TCP 80"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=out action=block protocol=TCP remoteport=80 enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
	
	If Not FirewallRuleExists("Disable Outbound TCP 8080") Then
		sDescr = "Disable Outbound TCP 8080"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=out action=block protocol=TCP remoteport=8080 enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
	
	If Not FirewallRuleExists("Disable Outbound TCP 53") Then
		sDescr = "Disable Outbound TCP 53"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=out action=block protocol=TCP remoteport=53 enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
		
	If Not FirewallRuleExists("Disable Outbound UDP 53") Then
		sDescr = "Disable Outbound UDP 53"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=out action=block protocol=UDP remoteport=53 enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
		
	if Not FirewallRuleExists("Disable Outbound TCP 443") Then
		sDescr = "Disable Outbound TCP 443"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=out action=block protocol=TCP remoteport=443 enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
		
	if Not FirewallRuleExists("Allow Inbound TCP ANY") Then
		sDescr = "Allow Inbound TCP ANY"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=in action=allow protocol=TCP localport=any remoteport=any enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
		
	if Not FirewallRuleExists("Allow Inbound UDP ANY") Then
		sDescr = "Allow Inbound UDP ANY"
		WshShell.Run dq & sSystem32 & "netsh.exe" & dq _
			& " advfirewall firewall add rule name=" & dq & sDescr & dq _
			& " dir=in action=allow protocol=UDP localport=any remoteport=any enable=Yes profile=any description=" & dq & sDescr & dq, "0", True
	End If
End Function

	
Function FirewallRuleExists(sRuleName)
	Dim bExists
	bExists = False 'default
	Dim rule
	Dim sRequestedRule 
	sRequestedRule = UCase(sRuleName)

	' Create the FwPolicy2 object.
	Dim fwPolicy2
	Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

	' Get the Rules object
	Dim RulesObject
	Set RulesObject = fwPolicy2.Rules

	' Print all the rules in currently active firewall profiles.


	For Each rule In Rulesobject
		Dim sCurRule
		sCurRule = UCase(rule.name)
		If sRequestedRule = sCurRule Then
			bExists = True
			LogMessage("  [*] Rule " & sCurRule & " already exists")
			Exit For
		End If
	Next
	
	FirewallRuleExists = bExists
End Function

Function StartWinFirewall
	'WshShell.Run strCommand [,intWindowStyle] [,bWaitOnReturn]

	WshShell.Run dq & sSystem32 & "net.exe" & dq & " stop MpsSvc", "0", True
	WshShell.Run dq & sSystem32 & "sc.exe" & dq & " config MpsSvc start= auto", "0", True
	WshShell.Run dq & sSystem32 & "net.exe" & dq & " start MpsSvc", "0", True

End Function

Function DeleteRegKey(RegPath)
	On Error Resume Next
	WshShell.RegDelete "HKLM\" & RegPath

End Function

Function WriteReg(strKeyPath, strValueName, strValue, RegType)
	Dim objRegistry
	Set objRegistry=GetObject("winmgmts:\\" & _ 
		strComputer & "\root\default:StdRegProv")
	 
	rem objRegistry.CreateKey HKEY_CURRENT_USER, strKeyPath
	if RegType = "REG_MULTI_SZ" Then
		Dim arrValue
		arrValue = Array(strValue)
		objRegistry.SetMultiStringValue HKLM, strKeyPath, strValueName, arrValue
		LogMessage "  [+] Wrote dword " & strValueName & " = " & strValue & " to " & strKeyPath
	End If
	
	if RegType = "REG_DWORD" Then
		objRegistry.SetDWORDValue HKLM, strKeyPath, strValueName, strValue
		LogMessage "  [+] Wrote dword " & strValueName & " = " & strValue & " to " & strKeyPath
	End If
	
End Function

REM Function ReadReg(RegPath)
	REM On Error Resume Next
    REM Dim objRegistry, Key
    REM Set objRegistry = CreateObject("Wscript.shell")

    REM Key = objRegistry.RegRead(RegPath)
    REM ReadReg = Key
REM End Function
