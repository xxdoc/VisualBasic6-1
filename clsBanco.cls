VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function gfbAbrir_Banco() As Boolean

    On Error GoTo err_gfAbrir_Banco:

    gfbAbrir_Banco = False

    Screen.MousePointer = vbHourglass

    pConString = ""
    pConString = "Driver=" & csPrvMySQL & ";"
    pConString = pConString & "Server=" & pServidor & ";"
    pConString = pConString & "Charset=latin1;"
    pConString = pConString & "Port=3306;"
    pConString = pConString & "Database=" & pBaseDados & ";"
    pConString = pConString & "User=" & pUsuario & ";"
    pConString = pConString & "Password=" & pSenha & ";"
    pConString = pConString & "Option=3;"

    Set cn = New ADODB.Connection
    With cn

        .ConnectionString = pConString
        .CommandTimeout = 360000000
        .Open

        If .State = adStateOpen Then gfbAbrir_Banco = True

    End With

    Screen.MousePointer = vbDefault

    Exit Function

err_gfAbrir_Banco:
    Call objSystem.gsExibeErros(Err, "gfAbrir_Banco()", "Classe Banco")
    Screen.MousePointer = vbDefault

End Function

Public Sub gsFecharBancos()

    On Error GoTo err_gsFecharBancos:

    Call gsFecharRecordset
    Call gsFecharBancos_Aux

    If Not cn Is Nothing Then

        If cn.State = adStateOpen Then cn.Close: Set cn = Nothing

    End If

    Exit Sub

err_gsFecharBancos:
    Call objSystem.gsExibeErros(Err, "gsFecharBancos()", "Classe Banco")

End Sub

Public Function gfsSaveFloat(ByVal pText As String)

    Dim nLoop

    If Len(Trim(pText)) = 0 Then
        gfsSaveFloat = " 0.00 "
        Exit Function
    End If

    For nLoop = 1 To Len(pText)
        If Mid(pText, nLoop, 1) = Chr(46) Then
            pText = " " & Left(pText, nLoop - 1) & Mid(pText, nLoop + 1, Len(pText))
        End If
    Next

    For nLoop = 1 To Len(pText)
        If Mid(pText, nLoop, 1) = Chr(44) Then
            pText = Left(pText, nLoop - 1) & "." & Mid(pText, nLoop + 1, Len(pText))
        End If
    Next

    gfsSaveFloat = " " & pText & " "

End Function

Public Function gfsSaveInt(ByVal pText As String)

    If Len(Trim(pText)) = 0 Then
        gfsSaveInt = " 0 "
        Exit Function
    End If

    gfsSaveInt = " " & pText & " "

End Function

Public Function gfsSaveChar(ByVal pText As Variant)

    Dim nLoop

    If IsNull(pText) Then
        gfsSaveChar = "''"
        Exit Function
    End If

    If Len(Trim(pText)) = 0 Then
        gfsSaveChar = "''"
        Exit Function
    End If

    For nLoop = 1 To Len(pText)

        If Mid(pText, nLoop, 1) = Chr(39) Or Mid(pText, nLoop, 1) = Chr(34) Then

            pText = Left(pText, nLoop - 1) & " " & Mid(pText, nLoop + 1, Len(pText))

        End If

    Next

    gfsSaveChar = " '" & Trim(pText) & "' "

End Function

Public Function gfsReadChar(ByVal pText As Variant)

    Dim nLoop

    If IsNull(pText) Then
        gfsReadChar = ""
        Exit Function
    End If

    gfsReadChar = Trim(pText)

End Function

Public Function gfsReadDate(ByVal pText As Variant)

    Dim nLoop
    Dim ndia, nMes, nano

    If IsNull(pText) Then
        gfsReadDate = ""
        Exit Function
    End If

    If IsEmpty(pText) Then
        gfsReadDate = ""
        Exit Function
    End If

    Select Case UCase(Mid(pText, 4, 3))

        Case "JAN"
            pText = Left(pText, 3) & "01" & Mid(pText, 7, Len(pText))

        Case "FEB"
            pText = Left(pText, 3) & "02" & Mid(pText, 7, Len(pText))

        Case "MAR"
            pText = Left(pText, 3) & "03" & Mid(pText, 7, Len(pText))

        Case "APR"
            pText = Left(pText, 3) & "04" & Mid(pText, 7, Len(pText))

        Case "MAI"
            pText = Left(pText, 3) & "05" & Mid(pText, 7, Len(pText))

        Case "JUN"
            pText = Left(pText, 3) & "06" & Mid(pText, 7, Len(pText))

        Case "JUL"
            pText = Left(pText, 3) & "07" & Mid(pText, 7, Len(pText))

        Case "AUG"
            pText = Left(pText, 3) & "08" & Mid(pText, 7, Len(pText))

        Case "SEP"
            pText = Left(pText, 3) & "09" & Mid(pText, 7, Len(pText))

        Case "OCT"
            pText = Left(pText, 3) & "10" & Mid(pText, 7, Len(pText))

        Case "NOV"
            pText = Left(pText, 3) & "11" & Mid(pText, 7, Len(pText))

        Case "DEC"
            pText = Left(pText, 3) & "12" & Mid(pText, 7, Len(pText))

    End Select

    ndia = Format(Day(pText), "00")
    nMes = Format(Month(pText), "00")
    nano = Format(Year(pText), "0000")

    If Val(ndia) = 1 And Val(nMes) = 1 And Val(nano) = 1800 Then

        gfsReadDate = ""
        Exit Function

    End If

    gfsReadDate = Trim(ndia & "/" & nMes & "/" & nano)

End Function

Public Function gfsReadInt(ByVal pText As Variant)

    Dim nLoop

    If IsNull(pText) Then
        gfsReadInt = 0
        Exit Function
    End If

    gfsReadInt = pText

End Function

Public Function gfsSaveDate(ByVal vData As Variant, sFlag As String, vTipoData As Variant) As String

    Dim sTData As String

    If Not IsDate(vData) Then
        gfsSaveDate = " Null "
        Exit Function
    End If

    If vTipoData = "" Or Len(Trim(vData)) = 0 Then
        gfsSaveDate = " Null "
        Exit Function
    End If

    If Len(Trim(vData)) <> 10 And sFlag = "D" Then
        gfsSaveDate = " Null "
        Exit Function
    End If

    sTData = UCase(Left(vTipoData, 2))
    If sFlag = "DH" Then
        vData = Format(vData, "dd-mm-yyyy hh:mm:ss")
    Else
        vData = Format(vData, "dd-mm-yyyy")
    End If

    If Left(sTData, 2) = "MM" Then
        If sFlag = "DH" Then
            gfsSaveDate = " '" & Mid(vData, 7, 4) & "-" & Left(vData, 2) & "-" & Mid(vData, 4, 2) & " " & Mid(vData, 12) & "' "
        Else
            gfsSaveDate = " '" & Mid(vData, 7, 4) & "-" & Left(vData, 2) & "-" & Mid(vData, 4, 2) & "' "
        End If
    Else
        If sFlag = "DH" Then
            gfsSaveDate = " '" & Mid(vData, 7, 4) & "-" & Mid(vData, 4, 2) & "-" & Left(vData, 2) & " " & Mid(vData, 12) & "' "
        Else
            gfsSaveDate = " '" & Mid(vData, 7, 4) & "-" & Mid(vData, 4, 2) & "-" & Left(vData, 2) & "' "
        End If
    End If

End Function

Public Function gfiExecuteSql(pSql As String) As Integer

    Dim nTentativa As Integer

    On Error GoTo err_gfiExecuteSql:

    cn.Execute pSql
    gfiExecuteSql = "1"

    Exit Function

err_gfiExecuteSql:
    If cn.Errors(0).Number = 0 Then
        nTentativa = nTentativa + 1
        If nTentativa < 13 Then Resume
    End If

    gfiExecuteSql = -1

    Call objSystem.gsExibeErros(Err, "gfiExecuteSql()", "Classe Banco", pSql)

End Function

Public Function gfdDataSistema() As Date

    On Error GoTo err_gfdDataSistema:

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000
        .CommandType = adCmdText

        pSql = "": pSql = "Select " & csSystemDate & " Data "

        .CommandText = pSql

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If Not .EOF Then gfdDataSistema = .Fields("Data")
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfdDataSistema:
    Call objSystem.gsExibeErros(Err, "gfdDataSistema()", "Classe Banco")

End Function

Public Sub gsBackup_Log()

    On Error GoTo err_gsBackup_Log:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandType = adCmdStoredProc
        .CommandText = "SP_BACKUP_LOG"
        .Execute

    End With
    Set cmd = Nothing

    Exit Sub

err_gsBackup_Log:
    Call objSystem.gsExibeErros(Err, "gsBackup_Log()", "Classe Banco")

End Sub

Public Function gfbProcedure(pIns_SQL As String, pProcedure As String) As Boolean

    On Error GoTo err_gfbProcedure:

    If pfbDrop_Procedure(pProcedure) = False Then gfbProcedure = False: Exit Function
    If pfbCreate_Procedure(pIns_SQL, pProcedure) = False Then gfbProcedure = False: Exit Function
    gfbProcedure = True

    Exit Function

err_gfbProcedure:
    Call objSystem.gsExibeErros(Err, "gfbProcedure()", "Classe Banco")
    gfbProcedure = False

End Function

Private Function pfbDrop_Procedure(pProcedure As String) As Boolean

    Dim pSql1 As String

    On Error GoTo err_pfbDrop_Procedure:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandType = adCmdText

        pSql1 = ""
        pSql1 = "IF EXISTS (SELECT * FROM SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].["
        pSql1 = pSql1 & pProcedure
        pSql1 = pSql1 & "]') AND OBJECTPROPERTY(ID, N'ISPROCEDURE') = 1)"
        pSql1 = pSql1 & "DROP PROCEDURE [DBO].["
        pSql1 = pSql1 & pProcedure & "]"

        .CommandText = pSql1
        .Execute

    End With
    pfbDrop_Procedure = True

    Set cmd = Nothing

    Exit Function

err_pfbDrop_Procedure:
    Call objSystem.gsExibeErros(Err, "pfbDrop_Procedure()", "Classe Banco", pSql)
    pfbDrop_Procedure = False

End Function

Private Function pfbCreate_Procedure(pIns_SQL As String, pProcedure As String) As Boolean

    Dim pSql1 As String

    On Error GoTo err_pfbCreate_Procedure:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CREATE PROCEDURE "
        pSql1 = pSql1 & pProcedure & vbNewLine & vbNewLine
        pSql1 = pSql1 & "AS" & vbNewLine & vbNewLine
        pSql1 = pSql1 & pIns_SQL

        .CommandText = pSql1
        .Execute

    End With
    pfbCreate_Procedure = True

    Set cmd = Nothing

    Exit Function

err_pfbCreate_Procedure:
    Call objSystem.gsExibeErros(Err, "pfbCreate_Procedure()", "Classe Banco", pSql)
    pfbCreate_Procedure = False

End Function

Public Function gflProximoRegistro(pTabela As String) As Double

    Dim pNm_Table As ADODB.Parameter

    On Error GoTo err_gflProximoRegistro:

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandText = "sp_upd_numerador"
        .CommandType = adCmdStoredProc

        Set pNm_Table = .CreateParameter("pnm_tabela", adVarChar, adParamInput, 30)
        .Parameters.Append pNm_Table: pNm_Table.Value = UCase(pTabela)

        .Execute

    End With
    Set cmd3 = Nothing

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandText = "sp_con_numerador"
        .CommandType = adCmdStoredProc

        Set pNm_Table = .CreateParameter("pnm_tabela", adVarChar, adParamInput, 30)
        .Parameters.Append pNm_Table: pNm_Table.Value = UCase(pTabela)

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gflProximoRegistro = 0 Else gflProximoRegistro = .Fields("nr_ultimoregistro")
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gflProximoRegistro:
    Call objSystem.gsExibeErros(Err, "gflProximoRegistro", "Classe Banco", pSql)

End Function

Public Function gflCodigo_Transacao(pcn As ADODB.Connection, pTransacao As String) As Long

    Dim prTransacao As ADODB.Parameter

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = pcn
        .CommandText = "SP_SEL_TABTRANSACAO_ESPECIFICA"
        .CommandType = adCmdStoredProc

        Set prTransacao = .CreateParameter("Transacao", adVarChar, adParamInput, 20)
        .Parameters.Append prTransacao: prTransacao.Value = UCase(pTransacao)

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gflCodigo_Transacao = 0 Else gflCodigo_Transacao = .Fields("Cd_Transacao")
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

End Function

Public Function gfbAbrir_Banco_Aux(pDataBase As String) As Boolean

    Dim pConString As String

    On Error GoTo err_gfbAbrir_Banco_Aux:

    gfbAbrir_Banco_Aux = False

    Screen.MousePointer = vbHourglass

    pConString = ""
    pConString = "Driver=" & csPrvMySQL & ";"
    pConString = pConString & "Server=" & pServidor & ";"
    pConString = pConString & "Charset=latin1;"
    pConString = pConString & "Port=3306;"
    pConString = pConString & "Database=" & pDataBase & ";"
    pConString = pConString & "User=" & pUsuario & ";"
    pConString = pConString & "Password=" & pSenha & ";"
    pConString = pConString & "Option=3;"

    Set cn_Aux = New ADODB.Connection
    With cn_Aux

        .ConnectionString = pConString
        .CommandTimeout = 360000000
        .Open

        If .State = adStateOpen Then gfbAbrir_Banco_Aux = True

    End With

    Screen.MousePointer = vbDefault

    Exit Function

err_gfbAbrir_Banco_Aux:
    Call objSystem.gsExibeErros(Err, "gfbAbrir_Banco_Aux()", "Classe Banco")
    Screen.MousePointer = vbDefault

End Function

Public Sub gsFecharBancos_Aux()

    On Error GoTo err_gsFecharBancos_Aux:

    If Not cn_Aux Is Nothing Then

        If cn_Aux.State = adStateOpen Then cn_Aux.Close: Set cn_Aux = Nothing

    End If

    Exit Sub

err_gsFecharBancos_Aux:
    Call objSystem.gsExibeErros(Err, "gsFecharBancos_Aux()", "Classe Banco")

End Sub

Private Function pfbDrop_View(pView As String) As Boolean

    Dim pSql1 As String

    On Error GoTo err_pfbDrop_View:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandType = adCmdText

        pSql1 = ""
        pSql1 = "IF EXISTS (SELECT * FROM SYSOBJECTS WHERE ID = OBJECT_ID(N'[DBO].["
        pSql1 = pSql1 & pView
        pSql1 = pSql1 & "]') AND OBJECTPROPERTY(ID, N'ISVIEW') = 1)"
        pSql1 = pSql1 & "DROP VIEW [DBO].["
        pSql1 = pSql1 & pView & "]"

        .CommandText = pSql1
        .Execute

    End With
    pfbDrop_View = True

    Set cmd = Nothing

    Exit Function

err_pfbDrop_View:
    Call objSystem.gsExibeErros(Err, "pfbDrop_View()", "Classe Banco", pSql1)
    pfbDrop_View = False

End Function

Private Function pfbCreate_View(pIns_SQL As String, pView As String) As Boolean

    Dim pSql1 As String

    On Error GoTo err_pfbCreate_View:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CREATE VIEW "
        pSql1 = pSql1 & pView & vbNewLine & vbNewLine
        pSql1 = pSql1 & "AS" & vbNewLine & vbNewLine
        pSql1 = pSql1 & pIns_SQL

        .CommandText = pSql1
        .Execute

    End With
    pfbCreate_View = True

    Set cmd = Nothing

    Exit Function

err_pfbCreate_View:
    Call objSystem.gsExibeErros(Err, "pfbCreate_View()", "Classe Banco", pSql1)
    pfbCreate_View = False

End Function

Public Function gfbView(pIns_SQL As String, pView As String) As Boolean

    On Error GoTo err_gfbView:

    If pfbDrop_View(pView) = False Then gfbView = False: Exit Function
    If pfbCreate_View(pIns_SQL, pView) = False Then gfbView = False: Exit Function
    gfbView = True

    Exit Function

err_gfbView:
    Call objSystem.gsExibeErros(Err, "gfbView()", "Classe Banco")
    gfbView = False

End Function

Public Sub gsFecharRecordset()

    On Error GoTo err_gsFecharRecordset:

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    End If

    If Not rs2 Is Nothing Then
        If rs2.State = adStateOpen Then rs2.Close: Set rs2 = Nothing
    End If

    If Not rs3 Is Nothing Then
        If rs3.State = adStateOpen Then rs3.Close: Set rs3 = Nothing
    End If

    Exit Sub

err_gsFecharRecordset:
    Call objSystem.gsExibeErros(Err, "gsFecharRecordset()", "Classe Banco")

End Sub

Public Function gflTotalRegistros(pNm_Tabela As String)

    On Error GoTo err_gflTotalRegistros:

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000
        .CommandType = adCmdText

        pSql = "": pSql = "SELECT COUNT(*) qtde"
        pSql = pSql & " FROM " & pNm_Tabela

        .CommandText = pSql

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If Not .EOF Then gflTotalRegistros = objBanco.gfsReadInt(.Fields("QTDE"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gflTotalRegistros:
    Call objSystem.gsExibeErros(Err, "gflTotalRegistros()", "Classe Banco")

End Function
