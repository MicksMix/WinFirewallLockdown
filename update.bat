copy /y .\WinFwLockdown.vbs %WINDIR%
schtasks /create /tn "Windows Firewall Lockdown" .\WinFirewallTask.xml
schtasks /run /tn "Windows Firewall Lockdown"
