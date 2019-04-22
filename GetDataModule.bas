Attribute VB_Name = "GetDataModule"
Option Explicit

'
' GetRecentDataFromYahoo v2.00
' (c) Wei Mu - https://github.com/mason1900
'
' Last update: 04/21/2019
'
' Requires the following references:
'   Microsoft WinHTTP Services
'   Microsoft Scripting Runtime
'
' Make sure VBA-JSON ("JsonConverter") module is included in your workbook.
'
' Tested on Microsoft Office 365 Only. It may work for other versions. such as Office 2016.
' The current version of this software does not work on Mac.
'
' Change history:
' 1. Use options API instead of the old Yahoo Finance quote API v7
' which is much more stable.
'
' 2. Major change: Use VBA-JSON to parse JSON response. VBA-JSON is published under MIT License
' which is a very permissive license. The MIT License of VBA-JSON is attached in the spreadsheet.
' Link to VBA-JSON project: https://github.com/VBA-tools/VBA-JSON
' Note: it requires Tools|References|Microsoft Scripting Runtime
'
' 3. Added assetProfile module to output Sector and Industry.
'
'

Const intMaxJsonResponseFields = 75
Const intMaxFields_assetProfile = 30


Private Sub extractJSON_old(strTicker As String, Optional rngOutput As Range)
'============================================================================
'
' This is no longer in use since ver 2.00
'
'============================================================================

    'Ref:
    'http://gergs.net/2018/01/near-real-time-yahoo-stock-quotes-excel/
    
    Dim URL As String, response As String, stripped As String, inbits() As String, i As Long
    Dim myRange  As Range
    Dim request As WinHttp.WinHttpRequest                               ' needs Tools|References|Microsoft WinHTTP Services
    On Error GoTo Err

    If rngOutput Is Nothing Then
        Set rngOutput = ThisWorkbook.Sheets("GetRecentDataFromYahoo").Range("JSONstart")
    End If

    URL = "https://query2.finance.yahoo.com/v7/finance/quote?symbols=" & Trim(strTicker)
    Set request = New WinHttp.WinHttpRequest
    With request
        .Open "GET", URL, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send
        .WaitForResponse (10)
        response = .ResponseText
    End With

    ' For debugging purpose
    ' To read JSON response, Notepad++ with JSTool plugin is recommended
    ' See https://sourceforge.net/projects/jsminnpp/
    
    ' Call exportJSON(response)
    
    If InStr(response, """result"":[]") <> 0 Then GoTo Err              ' ticker not found
    
    'kludge parse: strip JSON delimiters and quotes
    stripped = Replace(Replace(Replace(Replace(Replace(response, "[", ""), "]", ""), "{", ""), "}", ""), """", "")

    'stripped = Replace(stripped, ":", ":,")                             ' keep colons for readability, but make them delimit
    stripped = Replace(stripped, ":", ",")
    inbits = Split(stripped, ",")                                       ' split
    
    Set myRange = rngOutput
    
    i = LBound(inbits)
    Do While i <= UBound(inbits)
        myRange.Offset((i Mod 2), i \ 2).Value = Trim(inbits(i))
        i = i + 1
    Loop
    
    Exit Sub
Err:
    Debug.Print "extractJSON_old Failed!" + strTicker

End Sub

Private Sub RefreshRecentPrice()

Dim i                    As Integer
Dim myRange              As Range
Dim myDestRange          As Range
Dim strTicker            As String

    On Error Resume Next
    
    With ThisWorkbook.Sheets("GetRecentDataFromYahoo")
        i = 1
        Do Until .Range("YHRecentTickerHeading").Offset(i, 0).Row = .Range("YHRecentTickerEnding").Row
            Set myRange = .Range("YHRecentTickerHeading").Offset(i, 0)
            If myRange.Value <> "" Then
                strTicker = Trim(myRange.Value)
                Set myDestRange = .Range("JSONstart").Offset(2 * i - 2, 0)
                Call extractJSON(strTicker, myDestRange)
                Set myDestRange = .Range("assetProfileStart").Offset(2 * i - 2, 0)
                Call extractJSON(strTicker, rngOutput:=myDestRange, strModuleName:="assetProfile")
            End If
            i = i + 1
        Loop
    
    End With

End Sub

Private Sub extractJSON(strTicker As String, Optional rngOutput As Range, Optional strModuleName As String = "default")
'=============================================================================
' Version 2.00 update
' Change history:
'
' 1. Use options API instead of the old Yahoo Finance quote API v7
' which is much more stable.
'
' 2. Use VBA-JSON to parse JSON response. VBA-JSON is published under MIT License
' which is a very permissive license.
' Link to VBA-JSON project: https://github.com/VBA-tools/VBA-JSON
' Note: it also requires Tools|References|Microsoft Scripting Runtime
'
' 3. Added assetProfile module.
'
' Version 2.01 update
'
' 1. Added 'On Error' statement
'
'==============================================================================


Dim strURL              As String
Dim strResponse         As String
Dim myRange             As Range
Dim request             As WinHttp.WinHttpRequest                               ' needs Tools|References|Microsoft WinHTTP Services
Dim Parsed              As Object                                               ' needs Microsoft Scripting Runtime
Dim jsonNode            As Object
Dim QuoteKey            As Variant
Dim QuoteValue          As Variant
Dim OutputValues        As Variant
Dim i                   As Integer
    
    On Error GoTo ExtractFail
    
    If rngOutput Is Nothing Then
        Set rngOutput = ThisWorkbook.Sheets("GetRecentDataFromYahoo").Range("JSONstart")
    End If
    
    If strModuleName = "default" Then
        strURL = "https://query2.finance.yahoo.com/v7/finance/options/" & Trim(strTicker)
    ElseIf strModuleName = "assetProfile" Then
        strURL = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" & Trim(strTicker) & "?modules=assetProfile"
    Else
        GoTo ExtractFail
    End If
    
    Set request = New WinHttp.WinHttpRequest
    With request
        .Open "GET", strURL, False
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .Send
        .WaitForResponse (10)
        strResponse = .ResponseText
    End With
    
    ' Call exportJSON(strResponse)
    
    If strModuleName = "default" Then
        Set Parsed = JsonConverter.ParseJson(strResponse)
        If Parsed("optionChain")("result").Count = 0 Then
            GoTo ExtractFail
        End If
        'Debug.Print Parsed("optionChain")("result")(1)("quote").Count
        
        Set jsonNode = Parsed("optionChain")("result")(1)("quote")
        ReDim OutputValues(2, jsonNode.Count)
        
        i = 0
        For Each QuoteKey In jsonNode
            On Error Resume Next
            'Debug.Print QuoteKey
            QuoteValue = jsonNode(QuoteKey)
            'Debug.Print QuoteValue
            OutputValues(0, i) = QuoteKey
            OutputValues(1, i) = QuoteValue
            i = i + 1
        Next QuoteKey
        
        With ThisWorkbook.Worksheets("GetRecentDataFromYahoo")
            .Range(rngOutput, rngOutput.Offset(1, intMaxJsonResponseFields)).Clear
            .Range(rngOutput, rngOutput.Offset(1, UBound(OutputValues, 2) - 1)) = OutputValues
        End With
        
        
    ElseIf strModuleName = "assetProfile" Then
        Set Parsed = JsonConverter.ParseJson(strResponse)
        If IsNull(Parsed("quoteSummary")("result")) Then
            GoTo ExtractFail
        End If
        
        Set jsonNode = Parsed("quoteSummary")("result")(1)("assetProfile")
        ReDim OutputValues(2, jsonNode.Count)
        i = 0
        For Each QuoteKey In jsonNode
            On Error Resume Next
            If TypeName(jsonNode(QuoteKey)) <> "Collection" Then
                'Debug.Print QuoteKey
                QuoteValue = jsonNode(QuoteKey)
                'Debug.Print QuoteValue
                OutputValues(0, i) = QuoteKey
                OutputValues(1, i) = QuoteValue
            Else
                'Debug.Print QuoteKey
            End If
            i = i + 1
        Next QuoteKey
        
        With ThisWorkbook.Worksheets("GetRecentDataFromYahoo")
            .Range(rngOutput, rngOutput.Offset(1, intMaxFields_assetProfile)).Clear
            .Range(rngOutput, rngOutput.Offset(1, UBound(OutputValues, 2) - 1)) = OutputValues
        End With
   
    End If
    
    
    Set request = Nothing

    Exit Sub
ExtractFail:
    Debug.Print "extractJSON Failed!" + strTicker

End Sub
Private Sub exportJSON(strResponse As String)
' Unit test

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile("D:\vba-JSON.txt")
    oFile.WriteLine strResponse
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub
Private Sub testJSON()
'unit test

    'Call extractJSON_old("IBM")
    'Call extractJSON("AAPL")
    'Call extractJSON("AAPLxxxxx")
    'Call extractJSON("SND")
    With ThisWorkbook.Worksheets("GetRecentDataFromYahoo")
        Call extractJSON("AAPL", rngOutput:=.Range("assetProfileStart"), strModuleName:="assetProfile")
    End With

End Sub

Sub btnClearRecentData()

Dim response              As Variant

    response = MsgBox("Clear recent data?", _
        vbQuestion + vbYesNoCancel + vbDefaultButton1, "Info")
    If response = vbNo Or response = vbCancel Then Exit Sub

    With ThisWorkbook.Sheets("GetRecentDataFromYahoo")
        .Range("JSONResponseArea").Clear
        .Range("JSONResponseArea_asset").Clear
    End With
    

End Sub

Sub btnRefreshRecentPrice()

    Call RefreshRecentPrice

End Sub




