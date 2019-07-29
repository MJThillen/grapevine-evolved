VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryClauseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Key As String                                    'The query key to compare by.
Public Comparison As QueryCompareType                   'The type of comparison.
Public Find As String                                   'a string value to look for.
Public Number As Double                                 'a number to compare to.
Public CompNot As Boolean                               'whether this is a NOT comparison.

Public Function Description() As String
'
' Name:         Description
' Description:  Return a string with an english description of this clause.
'               It's not perfect, but it gets the idea across.
'

    Dim SNum As String
    
    SNum = CStr(Number)
            
    Select Case Comparison
        Case qcContains:        Description = "contain " & Find
        Case qcEquals:          Description = "is equal to " & IIf(Find = "", SNum, Find)
        Case qcAtLeast:         Description = "is at least " & SNum
        Case qcGreater:         Description = "is greater than " & SNum
        Case qcLess:            Description = "is less than " & SNum
        Case qcNoMore:          Description = "is no more than " & SNum
        Case qcContainsExactly: Description = "contain exactly " & Find & " x" & SNum
        Case qcContainsAtLeast: Description = "contain at least " & Find & " x" & SNum
        Case qcContainsMore:    Description = "contain more than " & Find & " x" & SNum
        Case qcContainsLess:    Description = "contain less than " & Find & " x" & SNum
        Case qcContainsNoMore:  Description = "contain no more than " & Find & " x" & SNum
        Case qcTotals:          Description = "total " & SNum
        Case qcTotalsAtLeast:   Description = "total at least " & SNum
        Case qcTotalsMore:      Description = "total more than " & SNum
        Case qcTotalsNoMore:    Description = "total no more than " & SNum
        Case qcTotalsLess:      Description = "total less than " & SNum
        Case qcContainsNote:    Description = "contain the note " & Find
        Case qcIsTrue:          Description = "is true"
        Case qcIsFalse:         Description = "is false"
    End Select
    
    If CompNot Then
        If Left(Description, 2) = "is" Then
            Description = "is not" & Mid(Description, 3)
        Else
            Description = "do not " & Description
        End If
    End If
    
    On Error Resume Next
    Description = Game.QueryEngine.KeysToTitles(Key) & " " & Description
    On Error GoTo 0
        
End Function

Public Sub OutputToFile(XML As XMLWriterClass)
'
' Name:         OutputToFile
' Parameters:   XML     the XMLWriterClass
' Description:  Write this object to a file.
'

    Dim CompStr As String
    
    Select Case Comparison
        Case qcContains:        CompStr = "[="
        Case qcEquals:          CompStr = "=="
        Case qcAtLeast:         CompStr = "+="
        Case qcGreater:         CompStr = "+"
        Case qcLess:            CompStr = "-"
        Case qcNoMore:          CompStr = "-="
        Case qcContainsExactly: CompStr = "[=="
        Case qcContainsAtLeast: CompStr = "[+="
        Case qcContainsMore:    CompStr = "[+"
        Case qcContainsLess:    CompStr = "[-"
        Case qcContainsNoMore:  CompStr = "[-="
        Case qcTotals:          CompStr = "#="
        Case qcTotalsAtLeast:   CompStr = "#+="
        Case qcTotalsMore:      CompStr = "#+"
        Case qcTotalsLess:      CompStr = "#-"
        Case qcTotalsNoMore:    CompStr = "#-="
        Case qcContainsNote:    CompStr = "[N"
        Case qcIsTrue:          CompStr = "=T"
        Case qcIsFalse:         CompStr = "=F"
    End Select
    
    CompStr = IIf(CompNot, "!", "") & CompStr
    
    With XML
    
        .BeginTag "clause"
            .WriteAttribute "field", Key
            .WriteAttribute "comp", CompStr
            .WriteAttribute "find", Find
            .WriteAttribute "val", Number, 0
        .EndTag
    
    End With

End Sub

Public Sub InputFromFile(XML As XMLReaderClass, Version As Double)
'
' Name:         InputFromFile
' Parameters:   XML             XMLReaderClass that has read the tag
'               Version         file format version
' Description:  Reads the object in from an XML file.
'

    Dim CompStr As String
    
    With XML
    
        If .Tag = "clause" Then
            
            Key = .GetString("field")
            CompStr = .GetString("comp")
            Find = .GetString("find")
            Number = .GetDouble("val")

            CompNot = (Left(CompStr, 1) = "!")
            If CompNot Then CompStr = Mid(CompStr, 2)
            Select Case CompStr
                Case "==":  Comparison = qcEquals
                Case "+=":  Comparison = qcAtLeast
                Case "+":   Comparison = qcGreater
                Case "-":   Comparison = qcLess
                Case "-=":  Comparison = qcNoMore
                Case "[==": Comparison = qcContainsExactly
                Case "[+=": Comparison = qcContainsAtLeast
                Case "[+":  Comparison = qcContainsMore
                Case "[-":  Comparison = qcContainsLess
                Case "[-=": Comparison = qcContainsNoMore
                Case "#=":  Comparison = qcTotals
                Case "#+=": Comparison = qcTotalsAtLeast
                Case "#+":  Comparison = qcTotalsMore
                Case "#-":  Comparison = qcTotalsLess
                Case "#-=": Comparison = qcTotalsNoMore
                Case "[N":  Comparison = qcContainsNote
                Case "=T":  Comparison = qcIsTrue
                Case "=F":  Comparison = qcIsFalse
                Case Else:  Comparison = qcContains
            End Select
            
            Do Until .Tag = "clause" And .Terminal
                .ReadTag
            Loop
            
        End If
        
    End With

End Sub

Public Sub OutputToBinary(FileNum As Integer)
'
' Name:         OutputToBinary
' Parameters:   FileNum     the number of the open file to write to
' Description:  Write this character's data to a binary file.
'

    PutStrB FileNum, Key
    Put #FileNum, , Comparison
    Put #FileNum, , CompNot
    PutStrB FileNum, Find
    Put #FileNum, , Number

End Sub

Public Sub InputFromBinary(FileNum As Integer, Version As Double)
'
' Name:         InputFromBinary
' Parameters:   FileNum     the number of the open file to read from
'               Version     the version of the file format
' Description:  Read this character's data from a binary file.
'

    GetStrB FileNum, Key
    Get #FileNum, , Comparison
    Get #FileNum, , CompNot
    GetStrB FileNum, Find
    Get #FileNum, , Number

End Sub
