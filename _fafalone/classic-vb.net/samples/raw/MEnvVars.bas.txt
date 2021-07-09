Attribute VB_Name = "MEnvVars"
' *************************************************************************
'  Copyright ©2009 Karl E. Peterson
'  All Rights Reserved, http://vb.mvps.org/
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

Private Declare Function apiRegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function apiRegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function apiRegFlushKey Lib "advapi32.dll" Alias "RegFlushKey" (ByVal hKey As Long) As Long
Private Declare Function apiRegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SHDeleteValue Lib "shlwapi" Alias "SHDeleteValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String) As Long

' Constants used with registry calls
Private Const ERROR_SUCCESS As Long = 0&
Private Const REG_SZ As Long = 1&
Private Const REG_EXPAND_SZ As Long = 2&
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002

' Constants used to broadcast messages
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const WM_SETTINGCHANGE As Long = &H1A

' Reg Key Security Options
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private Const hklmSubKey As String = "System\CurrentControlSet\Control\Session Manager\Environment"
Private Const hkcuSubKey As String = "Environment"

Public Function eVarWrite(ByVal eVar As String, ByVal eVal As String, Optional ByVal HKLM As Boolean = False, Optional ByVal Expandable As Boolean = True) As Boolean
   Dim RootKey As Long
   Dim SubKey As String
   Dim dwType As Long
   Dim nRet As Long
   Dim hKey As Long
   
   ' Is this user-specific or machine-wide?
   If HKLM Then
      RootKey = HKEY_LOCAL_MACHINE
      SubKey = hklmSubKey
   Else
      RootKey = HKEY_CURRENT_USER
      SubKey = hkcuSubKey
   End If
   
   ' Allow for variable expansion, by default.
   If Expandable Then
      dwType = REG_EXPAND_SZ
   Else
      dwType = REG_SZ
   End If
   
   ' Open a key and set a value within it.
   If apiRegOpenKeyEx(RootKey, SubKey, 0&, KEY_ALL_ACCESS, hKey) = ERROR_SUCCESS Then
      ' Attempt to write data - Always a string.
      nRet = apiRegSetValueEx(hKey, eVar, 0&, dwType, ByVal eVal, Len(eVal))
      Call apiRegFlushKey(hKey)
      Call apiRegCloseKey(hKey)
      ' Return result of RegSetValueEx call.
      eVarWrite = (nRet = ERROR_SUCCESS)
   End If
End Function

Public Function eVarClear(ByVal eVar As String, Optional ByVal HKLM As Boolean = False) As Boolean
   Dim RootKey As Long
   Dim SubKey As String
   Dim nRet As Long
   
   ' Is this user-specific or machine-wide?
   If HKLM Then
      RootKey = HKEY_LOCAL_MACHINE
      SubKey = hklmSubKey
   Else
      RootKey = HKEY_CURRENT_USER
      SubKey = hkcuSubKey
   End If
   
   ' Just delete this single value.
   nRet = SHDeleteValue(RootKey, SubKey, eVar)
   ' Return result of SHDeleteValue call.
   eVarClear = (nRet = ERROR_SUCCESS)
End Function

Public Function eVarAlert() As Long
   ' This can take a few seconds, so it makes sense to have
   ' it in a separate routine and only call it after making
   ' all environment variable changes.
   Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, 0&, _
      "Environment", SMTO_ABORTIFHUNG, 5000, eVarAlert)
End Function
