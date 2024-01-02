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
'
'-------------------------------------------------------------------------------
' HISTORY :
'
'   10/08/25 - H.FRIEDRICH : new code
'
'-------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'----------------------------------------------
' TESTAREA
'------------
Private Sub btnProp_Click()
    'MigrateMasterDb VER_TXT
End Sub

'----------------------------------------------
' THIS IS THE GLOBAL ENTRY POINT - CALLED MAIN
'------------
Private Sub Form_Load()
    Dim rc As Boolean, yar As Integer
    
    rc = False
    VBA_DEBUG = False
    If Command = "debug" Then VBA_DEBUG = True
    
    'set Name and Icon of application
    SetDbPropery CurrentDb, "AppIcon", CurrentProject.Path & "/Lexess.ico"
    yar = Year(Date)
    SetDbPropery CurrentDb, "AppTitle", "LEXESS - Warenwirtschaft V" & VER_TXT & " - © " & yar
    SetDbPropery CurrentDb, "UseAppIconForFrmRpt", 1
    RefreshTitleBar
    
    ' Test if registry is writable or if the current user have restricted access
    If RegTest(hkLocalMachine, "SOFTWARE\HFR\Lexess") = 0 Then _
        MsgBox "Achtung: " & vbCrLf & _
            "Die Einstellungen können mit beschänkten Benutzerrechten in der Registry" & vbCrLf & _
            "nicht gespeichert werden. Lexess kann nicht ordnungsgemäß arbeiten." & vbCrLf & _
            "Bitte als Admin einloggen und folgenden Key freischalten (s.a. Lexess Hilfe): " & vbCrLf & _
            "HKEY_LOCAL_MACHINE\SOFTWARE\HFR", vbCritical
    
    ' check if tblPasswd exist und open the login Dialog
    Dim usrId As Variant
    Dim masterDb As String
    masterDb = SrcMasterDb
    
    usrId = RstLookup("tblPasswd.PeNr", _
        "tblPasswd IN '" & masterDb & "'[MS Access;DATABASE=" & masterDb & ";PWD=" & MASTER_DB_PWD & "]", _
        "tblPasswd.PeNr = 5000")
    If usrId <> "" And Not IsNull(usrId) Then _
        DoCmd.OpenForm "frmUsrLogin", , , , , acDialog
    
    ProgressBar "Lexess wird gestartet:", 0, False

    ' function is now obsolete as performed by INNO Setup (20.06.2014)
    '
    ' ProgressBar "Barcode Schriftarten werden geprüft:", 5, False
    ' ChkBarCodeFonts
        
    ProgressBar "Lexess Master Datenbank wird geladen:", 10, False
    TracePrint "Table initializing, Debug Value: " & VBA_DEBUG
    
    'search for Master Table and link it
    If LinkMasterTables = True Then
        'migrate Master DB if necessary
        ProgressBar "Datenbank Migration wird geprüft:", 15, False
        If MigrationNeeded(VER_TXT) = True Then MigrateMasterDb VER_TXT
        
        'check if new version found in Lexess Master DB and write if front end ver is newer
        ProgressBar "Auf neue Version wird geprüft:", 20, False
        ChkNewVer VER_TXT
        
        'delete last user of BK form
        RegSaveSetting hkLocalMachine, "SOFTWARE\HFR\Lexess", "lastUsedPeNr", "", True

        rc = True
    End If
    
    If rc = True Then
        Dim runOnSrv As Integer
        ProgressBar "Client/Server Mode wird geprüft:", 30, False
        runOnSrv = ChkRunOnSrv                'Do I run on a Server/Client WS
        RegSaveSetting hkLocalMachine, "SOFTWARE\HFR\Lexess", "RUN_ON_SRV", runOnSrv, False
        If runOnSrv = True Then TracePrint "Lexess is running on the Lexware Server."
        If runOnSrv = False Then TracePrint "Lexess is running on the Lexware Client."
       
        ProgressBar "Formular Firmenauswahl wird geladen:", 80, False
        'ChkDsnEntry                               'check if ODBC entries exist and create if not
        LXDBSRV = ODBCDriverExists(SYBASE)

        'minimize the ribbon bar if not done already
        If Not RibbonState Then MinimizeRibbon
    
        DoCmd.OpenForm "frmSelectCmp"
    Else
        DoCmd.OpenForm "frmRprtMain"
    End If
    
    ProgressBar "Programm wird gestartet:", 101, False
    
    'user has to login to get additional admin tables if needed
    Dim grp As String
    'grp = UserLogin
    
    
    If VBA_DEBUG = False Then DoCmd.Close acForm, Me.Name
End Sub

Private Function ChkMasterTables() As Boolean
    Dim db As Object
    Dim i As Integer

    Set db = CurrentDb()
    Dim masterDbFn As String
    masterDbFn = ReverseString(db.Name)
    masterDbFn = ReverseString(Split(masterDbFn, "\", 2)(1)) & "\LexessMaster.mdb"

    For i = 0 To db.TableDefs.Count - 1
        If db.TableDefs(i).Connect <> "" Then
            If InStr(1, db.TableDefs(i).Connect, "LexessMaster") > 0 Then
                If Mid(db.TableDefs(i).Connect, 11) <> masterDbFn Then     ' remove str ";database="
                    db.TableDefs(i).Connect = ";database=" & masterDbFn
                    db.TableDefs(i).RefreshLink
                End If
            End If
        End If
    Next i
    
    ChkMasterTables = True
End Function

