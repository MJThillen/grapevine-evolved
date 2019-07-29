Attribute VB_Name = "basFileRegister"
Option Explicit

Public Const REG_SZ As Long = &H1
Public Const REG_DWORD As Long = &H4
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_USERS As Long = &H80000003

Public Const ERROR_SUCCESS As Long = 0
Public Const ERROR_BADDB As Long = 1009
Public Const ERROR_BADKEY As Long = 1010
Public Const ERROR_CANTOPEN As Long = 1011
Public Const ERROR_CANTREAD As Long = 1012
Public Const ERROR_CANTWRITE As Long = 1013
Public Const ERROR_OUTOFMEMORY As Long = 14
Public Const ERROR_INVALID_PARAMETER As Long = 87
Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_MORE_DATA As Long = 234
Public Const ERROR_NO_MORE_ITEMS As Long = 259

Public Const KEY_ALL_ACCESS As Long = &H3F
Public Const REG_OPTION_NON_VOLATILE As Long = 0

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Sub CreateGVAssociation(AppPath As String)

   Dim sPath As String
   
  'File Associations begin with a listing
  'of the default extension under HKEY_CLASSES_ROOT.
  'So the first step is to create that
  'root extension item
   CreateNewKey ".gv2", HKEY_CLASSES_ROOT
   CreateNewKey ".gv3", HKEY_CLASSES_ROOT
   CreateNewKey ".gex", HKEY_CLASSES_ROOT
   CreateNewKey ".gvm", HKEY_CLASSES_ROOT
   CreateNewKey ".gvu", HKEY_CLASSES_ROOT
   
   
  'To the extension just added, add a
  'subitem where the registry will look for
  'commands relating to the .xxx extension
  '("MyApp.Document"). Its type is String (REG_SZ)
   SetKeyValue ".gv2", "", "Grapevine.GameFile", REG_SZ
   SetKeyValue ".gv3", "", "Grapevine.GameFile", REG_SZ
   SetKeyValue ".gex", "", "Grapevine.ExchangeFile", REG_SZ
   SetKeyValue ".gvm", "", "Grapevine.MenuFile", REG_SZ
   SetKeyValue ".gvu", "", "Grapevine.MenuUpdate", REG_SZ
   
   
  'Create the 'MyApp.Document' item under
  'HKEY_CLASSES_ROOT. This is where you'll put
  'the command line to execute or other shell
  'statements necessary.
   CreateNewKey "Grapevine.GameFile\shell\open\command", HKEY_CLASSES_ROOT
   
   CreateNewKey "Grapevine.GameFile\DefaultIcon", HKEY_CLASSES_ROOT
   CreateNewKey "Grapevine.ExchangeFile\DefaultIcon", HKEY_CLASSES_ROOT
   CreateNewKey "Grapevine.MenuFile\DefaultIcon", HKEY_CLASSES_ROOT
   CreateNewKey "Grapevine.MenuUpdate\DefaultIcon", HKEY_CLASSES_ROOT
   
   
  'Set its default item to "MyApp Document".
  'This is what is displayed in Explorer against
  'for files with a xxx extension. Its type is
  'String (REG_SZ)
   SetKeyValue "Grapevine.GameFile", "", "Grapevine Game File", REG_SZ
   SetKeyValue "Grapevine.ExchangeFile", "", "Grapevine Exchange File", REG_SZ
   SetKeyValue "Grapevine.MenuFile", "", "Grapevine Menu File", REG_SZ
   SetKeyValue "Grapevine.MenuUpdate", "", "Grapevine Menu Update", REG_SZ
   
   
  'Finally, add the path to myapp.exe
  'Remember to add %1 as the final command
  'parameter to assure the app opens the passed
  'command line item.
  '(results in '"c:\LongPathname\Myapp.exe %1")
  'Again, its type is string.
   sPath = """" & AppPath & "Grapevine.exe"" %1"
   SetKeyValue "Grapevine.GameFile\shell\open\command", "", sPath, REG_SZ
        
   sPath = AppPath & "gv3.ico"
   SetKeyValue "Grapevine.GameFile\DefaultIcon", "", sPath, REG_SZ
   sPath = AppPath & "gex.ico"
   SetKeyValue "Grapevine.ExchangeFile\DefaultIcon", "", sPath, REG_SZ
   sPath = AppPath & "gvm.ico"
   SetKeyValue "Grapevine.MenuFile\DefaultIcon", "", sPath, REG_SZ
   SetKeyValue "Grapevine.MenuUpdate\DefaultIcon", "", sPath, REG_SZ

End Sub


Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, _
                           vValue As Variant) As Long

   Dim nValue As Long
   Dim sValue As String
   
   Select Case lType
      Case REG_SZ
         sValue = vValue & Chr$(0)
         SetValueEx = RegSetValueExString(hKey, _
                                          sValueName, _
                                          0&, _
                                          lType, _
                                          sValue, _
                                          Len(sValue))
         
      Case REG_DWORD
         nValue = vValue
         SetValueEx = RegSetValueExLong(hKey, _
                                        sValueName, _
                                        0&, _
                                        lType, _
                                        nValue, _
                                        4)
   
   End Select
   
End Function


Public Sub CreateNewKey(sNewKeyName As String, _
                        lPredefinedKey As Long)

  'handle to the new key
   Dim hKey As Long
   Dim Result As Long
   
   Call RegCreateKeyEx(lPredefinedKey, _
                       sNewKeyName, 0&, _
                       vbNullString, _
                       REG_OPTION_NON_VOLATILE, _
                       KEY_ALL_ACCESS, 0&, hKey, Result)
   
   Call RegCloseKey(hKey)

End Sub


Public Sub SetKeyValue(sKeyName As String, _
                       sValueName As String, _
                       vValueSetting As Variant, _
                       lValueType As Long)

  'handle of opened key
   Dim hKey As Long
   
  'open the specified key
   Call RegOpenKeyEx(HKEY_CLASSES_ROOT, _
                    sKeyName, 0, _
                    KEY_ALL_ACCESS, hKey)
                    
   Call SetValueEx(hKey, _
                   sValueName, _
                   lValueType, _
                   vValueSetting)
                  
   Call RegCloseKey(hKey)

End Sub

