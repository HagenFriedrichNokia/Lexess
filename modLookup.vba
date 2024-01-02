'-------------------------------------------------------------------------------
' Copyright (C)2013 by Hagen FRIEDRICH. All rights reserved.
'-------------------------------------------------------------------------------
'
' FILE: Lexess.accdb
' AUTHOR: Hagen FRIEDRICH, Vogelbeerweg 2, 71287 Weissach
' DATE: (C)2013
'
'-------------------------------------------------------------------------------
' DESCRIPTION:
' Replacement Functions for DLookup, DCount & DSum , DMax & DMin
'
' Notes:
' Any spaces in field names or table names will probably result in an error
' If this is the case then provide the brackets yourselfs, e.g.
' RstLookup("My field","My table name with spaces in") will blow big time
' RstLookup("[My field]","[My table name with spaces in]") will be ok
' These functions will not bracket the field/table names for you so as to
' remain as flexible as possible, e.g. you can call tSum() to add or multiply or
' whatever along the way, e.g. tSum("Price * Qty","Table","criteria") or if you're
' feeling adventurous, specify joins and the like in the table name.
'-------------------------------------------------------------------------------
' HISTORY :
'
'   12/02/02 - H.FRIEDRICH : new code
'
'-------------------------------------------------------------------------------
' Public Function RstCount(pstrField As String, pstrTable As String, Optional pstrCriteria As String) As Long
' Public Function RstAvg(pstrField As String, pstrTable As String, Optional pstrCriteria) As Variant
' Public Function RstMax(pstrField As String, pstrTable As String, Optional pstrCriteria As String) As Variant
' Public Function RstMin(pstrField As String, pstrTable As String, Optional pstrCriteria As String) As Variant
' Public Function RstSum(pstrField As String, pstrTable As String, Optional pstrCriteria As String) As Double
' Public Function RstLookup(Expression As String, Domain As String, Optional Criteria) As Variant
' Public Function RstAryLookup(Expression As String, Domain As String, Optional Criteria) As String()
' Public Function RstMultiLookup(Expression As String, Domain As String, Optional Criteria, Optional sep As String = ",") As String()
'
'-------------------------------------------------------------------------------
Option Compare Database   'Use database order for string comparisons
Option Explicit

Public Function RstCount(pstrField As String, pstrTable As String, Optional pstrCriteria) As Long
    Dim rc As Variant
    
    rc = RstLookup("COUNT (" & pstrField & ")", pstrTable, pstrCriteria)
    If Not IsNull(rc) Then RstCount = Int(rc)
End Function

Public Function RstAvg(pstrField As String, pstrTable As String, Optional pstrCriteria) As Variant
    RstAvg = RstLookup("AVG (" & pstrField & ")", pstrTable, pstrCriteria)
End Function

Public Function RstMax(pstrField As String, pstrTable As String, Optional pstrCriteria) As Variant
    RstMax = RstLookup("MAX (" & pstrField & ")", pstrTable, pstrCriteria)
End Function

Public Function RstMin(pstrField As String, pstrTable As String, Optional pstrCriteria) As Variant
    RstMin = RstLookup("MIN (" & pstrField & ")", pstrTable, pstrCriteria)
End Function

Public Function RstSum(pstrField As String, pstrTable As String, Optional pstrCriteria) As Double
    RstSum = Nz(RstLookup("SUM (" & pstrField & ")", pstrTable, pstrCriteria), 0)
End Function

Public Function RstLookup(Expression As String, Domain As String, Optional Criteria) As Variant
    On Error GoTo HandleErr
    Dim db As Object, rst As Object, fld As Object
    Dim strSql As String
    
    strSql = "SELECT " & Expression & " FROM " & Domain
    If Not IsMissing(Criteria) Then strSql = strSql & " WHERE " & Criteria
    'MsgBox strSQL
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSql)
    With rst
        If .EOF Then
            RstLookup = Null
        Else
            RstLookup = .Fields(0)
        End If
    End With

ExitHere:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function
 
HandleErr:
    RstLookup = Null
    MsgBox Err.Description
    Resume ExitHere
End Function

Public Function RstAryLookup(Expression As String, Domain As String, Optional Criteria) As String()
    On Error GoTo HandleErr
    Dim db As Object, rst As Object, fld As Object
    Dim strSql As String
    Dim cnt As Integer
    Dim strAry() As String
    ReDim Preserve strAry(255) 'hier vordimensioniert auf max 255 expressions
    
    strSql = "SELECT " & Expression & " FROM " & Domain
    If Not IsMissing(Criteria) Then strSql = strSql & " WHERE " & Criteria
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSql)
    cnt = 0
    With rst
        If .EOF Then
            strAry = Null
        Else
            For Each fld In .Fields
                strAry(cnt) = fld.Value
                cnt = cnt + 1
            Next fld
        End If
    End With
    RstAryLookup = strAry

ExitHere:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function
 
HandleErr:
    Resume Next
    Dim errStr As String
    
    Select Case Err.Number
        Case 3021  'kein Datensatz gefunden
            errStr = Null
        Case 3061  'einer der Feldnamen (Ausdruck oder Kriterium) stimmt nicht
            errStr = "#Ausdruck/Kriterium"
        Case 3078  'Name der Tabelle oder Abfrage stimmt nicht
            errStr = "#Domäne"
        Case 3464  'Datentyp im Kriterium ist falsch
            errStr = "#Kriterium"
        Case Else  'Sonstige Fehler
            errStr = "#Fehler"
      End Select
      MsgBox errStr
    Resume ExitHere
 End Function

Public Function RstMultiLookup(Expression As String, Domain As String, Optional Criteria, Optional sep As String = ",") As String()
    On Error GoTo HandleErr
    Dim db As Object, rst As Object, fld As Object
    Dim strSql As String
    Dim cnt As Integer
    Dim strAry() As String: ReDim strAry(1)
    
    strSql = "SELECT " & Expression & " FROM " & Domain
    If Not IsMissing(Criteria) Then strSql = strSql & " WHERE " & Criteria
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSql)
    cnt = 0
    While Not rst.EOF
        ReDim Preserve strAry(cnt + 1)
        For Each fld In rst.Fields
            strAry(cnt) = strAry(cnt) & sep & fld.Value
        Next fld
        If strAry(cnt) <> "" Then strAry(cnt) = Mid(strAry(cnt), Len(sep) + 1) 'delete first seperator
        cnt = cnt + 1
        rst.MoveNext
    Wend
    If strAry(0) = "" Then
        ReDim strAry(0)
    Else
        ReDim Preserve strAry(cnt - 1)
    End If
    RstMultiLookup = strAry

ExitHere:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Function
 
HandleErr:
    RstMultiLookup = strAry
    'MsgBox Err.Description
    Resume ExitHere
 End Function
