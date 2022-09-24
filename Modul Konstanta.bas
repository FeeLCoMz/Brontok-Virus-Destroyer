Attribute VB_Name = "ModulKonstanta"
Option Explicit

' *** Informasi Virus ***
Public Const UkuranVirus = 45508 'bytes Brontok.A
Public Const UkuranVirus2 = 81920 'bytes Brontok.A[16]
Public Const UkuranVirus3 = 44448 'bytes Brontok.E[1]
Public Const NamaProsesVirus = "Brontok.A.HVM31" 'Brontok.A
Public Const NamaProsesVirus2 = "Brontok.A" 'Brontok.A[16]
Public Const AppData = "\Local Settings\Application Data"

' *** Registry Key yang dimodifikasi Virus ***
Public Const ExPol = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
Public Const SysPol = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
Public Const WinLogon = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
Public Const CtrSet1 = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\SafeBoot"
Public Const CtrSet2 = "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet002\Control\SafeBoot"
Public Const StartupRun = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Public Const StartupRun2 = "HKEY_USERS\S-1-5-21-861567501-1214440339-839522115-1003\Software\Microsoft\Windows\CurrentVersion\Run"

