Attribute VB_Name = "Root"
'
' File:             Root.bas
' Author:           Adam Cerling
' Description:      The Main() function in which the program starts is here, as well as
'                   some functions it needs to load files and some other public functions.
'
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main()
'
' Name:         Main
' Description:  The starting point of the program.  Create all the needed objects and
'               initialize the global references.  Load and show the mdiMain window.
'

    Set Game = New GameClass
    Set OutputEngine = New OutputEngineClass
    
    Set PlayerList = Game.PlayerList
    Set CharacterList = Game.CharacterList
    Set ItemList = Game.ItemList
    Set RoteList = Game.RoteList
    Set LocationList = Game.LocationList
    Set ActionList = Game.APREngine.ActionList
    Set PlotList = Game.APREngine.PlotList
    Set RumorList = Game.APREngine.RumorList
    Set AllRumorLists = Game.AllRumorLists
    Set InfluenceUseList = Game.InfluenceUseList
    
    CharacterList.Name = "Characters"
    PlayerList.Name = "Players"
    
    StdHealth(0) = hlStdHealth0
    StdHealth(1) = hlStdHealth1
    StdHealth(2) = hlStdHealth2
    StdHealth(3) = hlStdHealth3
    StdHealth(4) = hlStdHealth4
    
    Randomize Timer
    
    mdiMain.Show

End Sub

Public Sub CleanUp()
'
' Name:         CleanUp
' Description:  Destroy all objects.  The last thing to call before ending the program.
'

    Set PlayerList = Nothing
    Set CharacterList = Nothing
    Set ItemList = Nothing
    Set RoteList = Nothing
    Set LocationList = Nothing
    Set ActionList = Nothing
    Set PlotList = Nothing
    Set RumorList = Nothing
    Set AllRumorLists = Nothing
    Set InfluenceUseList = Nothing

    Set OutputEngine = Nothing
    Set Game = Nothing
    
    If Dir(App.Path & "\~GVPrint.html") <> "" Then Kill App.Path & "\~GVPrint.html"
    
End Sub

Public Function DetectFileFormat(FileName As String) As FileFormatType
'
' Name:         DetectFileFormat
' Parameters:   FileName            File path & name to check
' Description:  Read the header of a file to tell what kind it is.
' Returns:      The format of the file.
'

    Dim I As Integer
    Dim FileNum As Integer
    Dim BHeader As String
    Dim SHeader As String

    DetectFileFormat = gvInvalid
    FileNum = FreeFile
    BHeader = String(BinHeaderLen, " ")
    SHeader = String(80, " ")
    
    Open FileName For Binary As #FileNum
    
    Get #FileNum, , I
    Get #FileNum, , BHeader
    Get #FileNum, 1, SHeader

    Select Case BHeader
        Case BinHeaderGame
            DetectFileFormat = gvBinaryGame
        Case BinHeaderMenu
            DetectFileFormat = gvBinaryMenu
        Case BinHeaderExchange
            DetectFileFormat = gvBinaryExchange
        Case Else
            If Left(SHeader, 5) = "<?xml" Then
                DetectFileFormat = gvXML
            Else
                SHeader = Left(SHeader, InStr(SHeader, ">"))
                If SHeader = GameFileVersionTag0 Or _
                        SHeader = GameFileVersionTag1 Or _
                        SHeader = GameFileVersionTag2 Or _
                        SHeader = GameFileVersionTag3 Or _
                        SHeader = GameFileVersionTag4 Or _
                        SHeader = GameFileVersionTag5 Then
                    DetectFileFormat = gv23Game
                ElseIf SHeader = ExchangeFileVersionTag0 Or _
                        SHeader = ExchangeFileVersionTag1 Then
                    DetectFileFormat = gv23Exchange
                End If
            End If
    End Select

    Close #FileNum

End Function

Public Function TrimTabs(ByVal Str As String) As String
'
' Name:         TrimTabs
' Parameters:   Str         the string to strip of tabs
' Description:  Strip the tabs off the left and right of a string.  Used to help parse
'               menu files.
' Returns:      the stripped-down string.
'

    Str = Trim(Str)
    Do While Left(Str, 1) = Chr(vbKeyTab)
        Str = Right(Str, Len(Str) - 1)
    Loop
    Do While Right(Str, 1) = Chr(vbKeyTab)
        Str = Left(Str, Len(Str) - 1)
    Loop
    TrimTabs = Str
    
End Function

Public Function TrimWhiteSpace(ByVal Str As String) As String
'
' Name:         TrimWhiteSpace
' Parameters:   Str         the string to strip of white space
' Description:  Strip the white space off the beginning and end of a string, including tabs,
'               carriage returns, linefeeds and spaces.
' Returns:      the stripped-down string.
'

    Dim Char As String * 1

    Str = Trim(Str)
    Char = Left(Str, 1)
    Do While (Char = " " Or Char = vbCr Or Char = vbLf Or Char = vbTab) And Not Str = ""
        Str = Right(Str, Len(Str) - 1)
        Char = Left(Str, 1)
    Loop
    Char = Right(Str, 1)
    Do While (Char = " " Or Char = vbCr Or Char = vbLf Or Char = vbTab) And Not Str = ""
        Str = Left(Str, Len(Str) - 1)
        Char = Right(Str, 1)
    Loop
    TrimWhiteSpace = Str
    
End Function

Public Function OutsideQuotes(QuoteStr As String, Loc As Long) As Boolean
'
' Name:         OutsideQuotes
' Parameters:   QuoteStr        String to examine
'               Loc             Position in string
' Description:  Return TRUE iff the given position in the string is not between quote marks.
'

    Dim I As Integer
    
    I = InStr(QuoteStr, """")
    OutsideQuotes = True
    
    Do Until I >= Loc Or I = 0
        OutsideQuotes = Not OutsideQuotes
        I = InStr(I + 1, QuoteStr, """")
    Loop

End Function

Public Function GetRelativeName(FileName As String, GameFile As String) As String
'
' Name:         FindFile
' Parameters:   FileName        Full path of the file to get the relative path for
'               GameFile        Full path of the game file
' Description:  Try to get a relative path from the application path.
'               Failing that, get one from the game file path.
' Return:       A relative path, or a full path if no relative is found.
'

    Dim AppFolder As String
    Dim GameFolder As String
    
    AppFolder = SlashPath(App.Path)
    GameFolder = Left(GameFile, Len(GameFile) - Len(ShortFile(GameFile)))
    
    If InStr(1, FileName, AppFolder, vbTextCompare) = 1 Then
        FileName = Mid(FileName, Len(AppFolder) + 1)
    ElseIf InStr(1, FileName, GameFolder, vbTextCompare) = 1 Then
        FileName = Mid(FileName, Len(GameFolder) + 1)
    End If
    
    GetRelativeName = FileName
    
End Function

Public Function FindFile(TryFile As String, Optional Default As String = "") As String
'
' Name:         FindFile
' Parameters:   TryFile     name of the file to find on the disks
'               Default     name of a default file to find if all else fails
' Description:  Look for the presence of a given file, correcting for absolute
'               or relative paths.  Look for a default file if the first one isn't
'               found.
' Returns:      A path to the given file, or "" if it's not found.
'

    Dim BaseFile As String
    Dim Found As Boolean
    Dim DefaultLoop As Boolean
    
    On Error Resume Next
    
    Do
    
        Found = False
        If TryFile = "" Then Exit Do
        
        If Mid(TryFile, 2, 1) = ":" Then
            Found = (Dir(TryFile) <> "")
            BaseFile = ShortFile(TryFile)
        Else
            BaseFile = TryFile
        End If
    
        If Not Found Then
        
            TryFile = SlashPath(App.Path) & BaseFile
            Found = (Dir(TryFile) <> "")
            
            If Not Found And InStr(Game.GameFile, "\") > 0 Then
                TryFile = Left(Game.GameFile, InStrRev(Game.GameFile, "\")) & BaseFile
                Found = (Dir(TryFile) <> "")
            End If
            
            If Not Found Then
                TryFile = Default
                DefaultLoop = Not DefaultLoop
            End If
            
        End If

        DefaultLoop = DefaultLoop And Not Found

    Loop While DefaultLoop

    On Error GoTo 0

    FindFile = TryFile

End Function

Public Function ShortFile(FileName As String) As String
'
' Name:         ShortFile
' Parameters:   A Filename to shorten.
' Description:  Clip the given filename from its last "\".
'

    ShortFile = Mid(FileName, InStrRev(FileName, "\") + 1)

End Function

Public Function SlashPath(PathName As String) As String
'
' Name:         ShortFile
' Parameters:   Path name to edit
' Description:  Add a backslash to the end of a path if it's not already there.
'

    SlashPath = PathName & IIf(Right(PathName, 1) = "\", "", "\")

End Function

Public Sub SelectText(BoxofText As TextBox)
'
' Name:         SelectText
' Parameters:   BoxofText   the text box
' Description:  Select (highlight) the contents of a text box.  Usually called when the box
'               recieves focus.
'

    BoxofText.SelStart = 0
    BoxofText.SelLength = Len(BoxofText)

End Sub

Public Function ConvertToFileName(ByVal Value As String) As String
'
' Name:         ConvertToFileName
' Parameters:   Value       the string to convert
' Description:  Convert a string to a legal filename by stripping out or changing reserved
'               characters.
' Returns:      the new filename
'

    Value = Replace(Value, "\", "")
    Value = Replace(Value, "/", "")
    Value = Replace(Value, """", "`")
    Value = Replace(Value, "*", "-")
    Value = Replace(Value, ">", ")")
    Value = Replace(Value, "<", "(")
    Value = Replace(Value, ":", "-")
    Value = Replace(Value, "|", "-")
    ConvertToFileName = Replace(Value, "?", "!")

End Function

Public Function ReadLongField(FileNum As Integer, Delimiter As String) As String
'
' Name:         ReadLongField
' Parameters:   FileNum         the number of the open file
'               Delimiter       the string at which to stop reading
' Description:  Read several lines of a file until an expected delimiter is reached.
'               Used by most classes that load data from a file.
' Returns:      the data read
'

    Dim Read As String
    Dim Result As String

    Result = ""
    Line Input #FileNum, Read
    Do Until Read = Delimiter
        Result = Result & Read & vbCrLf
        Line Input #FileNum, Read
    Loop
    If Result <> "" Then Result = Left(Result, Len(Result) - 2)
    
    ReadLongField = Result

End Function

Public Function XORScramble(Key As String, Text As String) As String
'
' Name:         XORScramble
' Parameters:   Key         key by which to scramble the input
'               Text        string to scramble
' Description:  Simple XOR encryption of two strings.
'

    Dim XPos As Long
    Dim InChar As Integer
    Dim KeyChar As Integer

    For XPos = 1 To Len(Text)
        InChar = Asc(Mid(Text, XPos, 1))
        KeyChar = Asc(Mid(Key, (XPos Mod Len(Key)) + 1, 1))
        XORScramble = XORScramble & Chr(InChar Xor KeyChar)
    Next XPos
   
End Function

Public Sub PutStrB(FileNum As Integer, ByVal StrVal As String)
'
' Name:         PutStrB
' Parameters:   FileNum             binary file number to write to
'               StrVal              String value to write
' Description:  Write a string's length before the string, when writing to a binary file.
'               Only by knowing the string's length can it later be read back in.
'
    
    Dim L As Long

    If HideSTFromFile Then StrVal = Game.STFilter(StrVal)

    L = Len(StrVal)
    If L > 32767 Then
        Put #FileNum, , CInt(32767)
        Put #FileNum, , Left(StrVal, 32767)
    Else
        Put #FileNum, , CInt(L)
        Put #FileNum, , StrVal
    End If

End Sub

Public Sub GetStrB(FileNum As Integer, ByRef StrVal As String)
'
' Name:         GetStrB
' Parameters:   FileNum             binary file number to read from
'               StrVal              String value to read
' Description:  Read a string's length, prepare the space to store it, and then read
'               in the string.
'
    Dim L As Integer

    Get #FileNum, , L
    StrVal = String(L, " ")
    Get #FileNum, , StrVal
    
End Sub

