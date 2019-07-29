Option Explicit

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Copyright ©1996-2000 VBnet, Randy Birch, All Rights Reserved.
  ' Some pages may also contain other copyrights by the author.
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' You are free to use this code within your own applications,
  ' but you are expressly forbidden from selling or otherwise 
  ' distributing this source code without prior written consent.
  ' This includes both posting free demo projects made from this 
  ' code as well as reproducing the code in text or html format. 
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

  Public Const KEY_ALL_ACCESS As Long= &H3F
  Public Const REG_OPTION_NON_VOLATILE As Long = 0

  Public Declare Function RegCloseKey Lib "advapi32.dll" _
     (ByVal hKey As Long) As Long

  Public Declare Function RegCreateKeyEx _
      Lib "advapi32.dll" Alias "RegCreateKeyExA" _
     (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal Reserved As Long, _
      ByVal lpClass As String, _
      ByVal dwOptions As Long, _
      ByVal samDesired As Long, _
      ByVal lpSecurityAttributes As Long, _
      phkResult As Long, _
      lpdwDisposition As Long) As Long

  Public Declare Function RegOpenKeyEx _
      Lib "advapi32.dll" Alias "RegOpenKeyExA" _
     (ByVal hKey As Long, _
      ByVal lpSubKey As String, _
      ByVal ulOptions As Long, _
      ByVal samDesired As Long, _
      phkResult As Long) As Long

  Public Declare Function RegSetValueExString _
      Lib "advapi32.dll" Alias "RegSetValueExA" _
     (ByVal hKey As Long, _
      ByVal lpValueName As String, _
      ByVal Reserved As Long, _
      ByVal dwType As Long, _
      ByVal lpValue As String, _
      ByVal cbData As Long) As Long

  Public Declare Function RegSetValueExLong _
     Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpValue As Long, _
     ByVal cbData As Long) As Long


  Public Sub CreateAssociation()

     Dim sPath As String
     
    'File Associations begin with a listing
    'of the default extension under HKEY_CLASSES_ROOT.
    'So the first step is to create that
    'root extension item
     CreateNewKey ".xxx", HKEY_CLASSES_ROOT
     
     
    'To the extension just added, add a
    'subitem where the registry will look for
    'commands relating to the .xxx extension
    '("MyApp.Document"). Its type is String (REG_SZ)
     SetKeyValue ".xxx", "", "MyApp.Document", REG_SZ
     
     
    'Create the 'MyApp.Document' item under
    'HKEY_CLASSES_ROOT. This is where you'll put
    'the command line to execute or other shell
    'statements necessary.
     CreateNewKey "MyApp.Document\shell\open\command", HKEY_CLASSES_ROOT
     
     
    'Set its default item to "MyApp Document".
    'This is what is displayed in Explorer against
    'for files with a xxx extension. Its type is
    'String (REG_SZ)
     SetKeyValue "MyApp.Document", "", "MyApp Document", REG_SZ
     
     
    'Finally, add the path to myapp.exe
    'Remember to add %1 as the final command
    'parameter to assure the app opens the passed
    'command line item.
    '(results in '"c:\LongPathname\Myapp.exe %1")
    'Again, its type is string.
     sPath = "c:\LongPathname\Myapp.exe %1"
     SetKeyValue "MyApp.Document\shell\open\command", "", sPath, REG_SZ
     
    'All done
     MsgBox "The file association has been made!"
     
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
     Dim result As Long
     
     Call RegCreateKeyEx(lPredefinedKey, _
                         sNewKeyName, 0&, _
                         vbNullString, _
                         REG_OPTION_NON_VOLATILE, _
                         KEY_ALL_ACCESS, 0&, hKey, result)
     
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