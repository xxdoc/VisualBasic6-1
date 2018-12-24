Attribute VB_Name = "modPrincipal"
Option Explicit

' Declaracoes
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'--- Contantes ---
' Arquivo INI
Public Const csArquivoINI = "Eletropaulo.ini"

' Banco de Dados - ADO
Public Const csPrvMySQL = "{MySQL ODBC 3.51 Driver}"
Public Const csType = adOpenKeyset
Public Const csLocation = adUseServer
Public Const csSystemDate = "current_date()"
Public Const csProcedure = "SP_EXEC_PROC"

' Criptografia
Public Const ciEncrypt = 1
Public Const ciDecrypt = 2
Public Const csChave = "Kratos"

' Digitação
Public Const ciUpper = 1
Public Const ciChar = 2
Public Const ciInt = 3
Public Const ciFloat = 4
Public Const ciDate = 5
Public Const ciCGCCPF = 6
Public Const ciLower = 7

' E-Mail
Public Const csAdreess_Destinatario = "marcelo.oliveira@abacos.com.br"

'--- Variaveis ---
' ADODB
Public cn                   As ADODB.Connection
Public cn_Aux               As ADODB.Connection
Public cmd                  As ADODB.Command
Public cmd2                 As ADODB.Command
Public cmd3                 As ADODB.Command
Public rs                   As ADODB.Recordset
Public rs2                  As ADODB.Recordset
Public rs3                  As ADODB.Recordset

' Classes
Public objBanco             As New clsBanco
Public objSystem            As New clsSystem

' Boolean
Public pAutoNumeracao       As Boolean
Public pAutorizaCadastro    As Boolean
Public pClique              As Boolean
Public pInserirCadastro     As Boolean
Public pIs_Err              As Boolean
Public pLeituraColetor      As Boolean

' Integer
Public pArquivo             As Integer
Public pCd_Evento           As Integer
Public pCd_User             As Integer
Public pResp                As Integer
Public pRecords             As Integer
Public pTp_User             As Integer
Public pVetContador         As Integer
Public pVia                 As Integer

' Long
Public pCodigoAtual         As Long
Public pNumeracaoMaxima     As Long
Public pNumeracaoMinima     As Long

' String
Public pBaseDados           As String
Public pConString           As String
Public pDNS                 As String
Public pFileName            As String
Public pImpressora          As String
Public pLg_User             As String
Public pMsg                 As String
Public pPath                As String
Public pServidor            As String
Public pSenha               As String
Public pSql                 As String
Public pUsuario             As String

' Variant
Public pConnectSrv          As Variant

' Vetor
Public pVet_Sql(1 To 300000)    As String

Sub Main()

    Screen.MousePointer = vbHourglass

    If App.PrevInstance = True Then

        Screen.MousePointer = vbDefault
        MsgBox "Este aplicativo já está em uso.", vbCritical + vbOKOnly, "Acesso Negado."
        End

    Else

        pPath = App.Path: objSystem.sPath_File = App.Path

        Call gsBuscar_Dados_INI

        If Not objBanco.gfbAbrir_Banco() Then End

        Screen.MousePointer = vbDefault

        frmLogin.Show

    End If

End Sub

Public Function gfvTipoData() As Variant

    Dim pBuffer As String
    Dim pFlag   As Integer
    Dim pTemp   As Variant

    pBuffer = String$(255, 0)
    pFlag = GetPrivateProfileString("Intl", "sShortDate", "", pBuffer, Len(pBuffer), "Win.Ini")

    If pFlag = 0 Then

        MsgBox "Win.Ini não encontrado!"
        gfvTipoData = ""

    Else

        pTemp = UCase(Left(pBuffer, pFlag))
        pTemp = IIf(Right(pTemp, 4) = "YYYY", pTemp, pTemp & "YY")
        gfvTipoData = pTemp

    End If

End Function

Public Sub gsBuscar_Dados_INI()

    Dim pFileName_Ini As String

    pFileName_Ini = pPath & "\" & csArquivoINI
    pBaseDados = objSystem.gfsGetIni(pFileName_Ini, "DataBase", "BaseDados", "")
    pCd_Evento = objSystem.gfsGetIni(pFileName_Ini, "Event", "Evento", "")
    pImpressora = objSystem.gfsGetIni(pFileName_Ini, "System", "Impressora", "")
    pServidor = objSystem.gfsGetIni(pFileName_Ini, "Server", "Servidor", "")
    pSenha = objSystem.gfsGetIni(pFileName_Ini, "Password", "Senha", "")
    pUsuario = objSystem.gfsGetIni(pFileName_Ini, "User", "Usuario", "")

End Sub

Public Function gfiBuscarCodigoCategoria(pDe_Categoria As String) As Integer

    On Error GoTo err_gfiBuscarCodigoCategoria:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_categoria"
        pSql1 = pSql1 & " ( NULL"
        pSql1 = pSql1 & ", " & objBanco.gfsSaveChar(pDe_Categoria)
        pSql1 = pSql1 & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gfiBuscarCodigoCategoria = 0 Else gfiBuscarCodigoCategoria = objBanco.gfsReadInt(.Fields("ID_CATEGORIA"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfiBuscarCodigoCategoria:
    Call objSystem.gsExibeErros(Err, "gfiBuscarCodigoCategoria()", "Módulo Principal")

End Function

Public Function gfiBuscarCodigoEstadoCivil(pDe_EstadoCivil As String) As Integer

    On Error GoTo err_gfiBuscarCodigoEstadoCivil:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_estadocivil"
        pSql1 = pSql1 & " ( NULL"
        pSql1 = pSql1 & ", " & objBanco.gfsSaveChar(pDe_EstadoCivil)
        pSql1 = pSql1 & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gfiBuscarCodigoEstadoCivil = 0 Else gfiBuscarCodigoEstadoCivil = objBanco.gfsReadInt(.Fields("ID_ESTADOCIVIL"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfiBuscarCodigoEstadoCivil:
    Call objSystem.gsExibeErros(Err, "gfiBuscarCodigoEstadoCivil", "Módulo Principal")

End Function

Public Function gfiBuscarCodigoSexo(pDe_Sexo As String) As Integer

    On Error GoTo err_gfiBuscarCodigoSexo:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_sexo"
        pSql1 = pSql1 & " ( NULL"
        pSql1 = pSql1 & ", " & objBanco.gfsSaveChar(pDe_Sexo)
        pSql1 = pSql1 & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gfiBuscarCodigoSexo = 0 Else gfiBuscarCodigoSexo = objBanco.gfsReadInt(.Fields("ID_SEXO"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfiBuscarCodigoSexo:
    Call objSystem.gsExibeErros(Err, "gfiBuscarCodigoSexo()", "Módulo Principal")

End Function

Public Function gfiBuscarCodigoParentesco(pDe_Parentesco As String) As Integer

    On Error GoTo err_gfiBuscarCodigoParentesco:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_parentesco"
        pSql1 = pSql1 & " ( NULL"
        pSql1 = pSql1 & ", " & objBanco.gfsSaveChar(pDe_Parentesco)
        pSql1 = pSql1 & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gfiBuscarCodigoParentesco = 0 Else gfiBuscarCodigoParentesco = objBanco.gfsReadInt(.Fields("ID_PARENTESCO"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfiBuscarCodigoParentesco:
    Call objSystem.gsExibeErros(Err, "gfiBuscarCodigoParentesco()", "Módulo Principal")

End Function

Public Function gflBuscarCodigoPessoa() As Long

    On Error GoTo err_gflBuscarCodigoPessoa:

    Dim pId_Convite As Long
    Dim pDg_Convite As String
    Dim pNumero     As String
    Dim pNr_Convite As String
    Dim pSql1       As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "SELECT a.id_convite, a.nr_convite, a.dg_convite, a.fl_utilizado"
        pSql1 = pSql1 & " FROM tb_convite a"
        pSql1 = pSql1 & " WHERE a.id_convite = ( SELECT min(id_convite) FROM tb_convite WHERE fl_utilizado = 'N' )"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then

                gflBuscarCodigoPessoa = 0

            Else

                pId_Convite = objBanco.gfsReadInt(.Fields("ID_CONVITE"))
                pDg_Convite = objBanco.gfsReadChar(.Fields("DG_CONVITE"))
                pNr_Convite = objBanco.gfsReadChar(.Fields("NR_CONVITE"))
                pNumero = pNr_Convite & pDg_Convite

                gflBuscarCodigoPessoa = CLng(pNumero)

            End If
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    pSql1 = "": pSql1 = "UPDATE tb_convite"
    pSql1 = pSql1 & " SET fl_utilizado = " & objBanco.gfsSaveChar("S")
    pSql1 = pSql1 & " WHERE id_convite = " & objBanco.gfsSaveInt(pId_Convite)

    cn.BeginTrans
    If objBanco.gfiExecuteSql(pSql1) = -1 Then cn.RollbackTrans Else cn.CommitTrans

    Exit Function

err_gflBuscarCodigoPessoa:
    Call objSystem.gsExibeErros(Err, "gflBuscarCodigoPessoa()", "Módulo Principal")

End Function

Public Function gfiBuscarCodigoTipoLog(pDe_TipoLog As String) As Integer

    On Error GoTo err_gfiBuscarCodigoTipoLog:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_tipolog"
        pSql1 = pSql1 & " ( NULL"
        pSql1 = pSql1 & ", " & objBanco.gfsSaveChar(pDe_TipoLog)
        pSql1 = pSql1 & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gfiBuscarCodigoTipoLog = 0 Else gfiBuscarCodigoTipoLog = objBanco.gfsReadInt(.Fields("ID_TIPOLOG"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gfiBuscarCodigoTipoLog:
    Call objSystem.gsExibeErros(Err, "gfiBuscarCodigoTipoLog()", "Módulo Principal")

End Function

Public Function gflBuscarCodigoTitular(pNr_Matricula As Long) As Long

    On Error GoTo err_gflBuscarCodigoTitular:

    Dim pSql1 As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_pessoa_titular"
        pSql1 = pSql1 & " ( " & objBanco.gfsSaveInt(pNr_Matricula) & " );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then gflBuscarCodigoTitular = 0 Else gflBuscarCodigoTitular = objBanco.gfsReadInt(.Fields("CD_PESSOA"))
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Function

err_gflBuscarCodigoTitular:
    Call objSystem.gsExibeErros(Err, "gflBuscarCodigoTitular()", "Módulo Principal")

End Function

Public Function gfbImprimir(pRptSql As String, pRptFile As String, pRptTitle As String, pObj As Object, pDestino As Integer) As Boolean

    On Error GoTo err_gfbImprimir:

    gfbImprimir = True

    'Verifica se o rpt está no caminho correto
    If Dir(pRptFile) = "" Then

        pMsg = "": pMsg = "Arquivo de impressão não localizado"
        MsgBox pMsg, vbCritical, "Atenção."
        gfbImprimir = False: Exit Function

    End If

    'Inicio a impressão
    pObj.ReportFileName = pRptFile
    pObj.Connect = pConString
    pObj.SQLQuery = pRptSql
    pObj.WindowState = 2
    pObj.WindowTitle = pRptTitle
    pObj.ProgressDialog = False
    pObj.Destination = pDestino
    pObj.Action = 1

    Exit Function

err_gfbImprimir:
    Call objSystem.gsExibeErros(Err, "gfbImprimir()", "Módulo de Acordos.")
    gfbImprimir = False

End Function

Public Sub gsCarregaDadosPessoa(pCd_Pessoa As Long)

    On Error GoTo err_gsCarregaDadosPessoa:

    Dim pSql1   As String

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "CALL sp_con_pessoa"
        pSql1 = pSql1 & " ( " & objBanco.gfsSaveInt(pCd_Pessoa)
        pSql1 = pSql1 & ", NULL, NULL, NULL, NULL );"

        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            If .EOF Then Exit Sub

            pInserirCadastro = False

            frmCredenciamento.txtCodigo.Text = objBanco.gfsReadChar(.Fields("CD_PESSOA"))
            frmCredenciamento.txtVia.Text = objBanco.gfsReadInt(.Fields("NR_VIA"))
            Call objSystem.gsBuscaCombo(frmCredenciamento.cmbCategoria, objBanco.gfsReadInt(.Fields("ID_CATEGORIA")))
            frmCredenciamento.txtMatricula.Text = objBanco.gfsReadInt(.Fields("NR_MATRICULA"))
            frmCredenciamento.txtCPF.Text = objBanco.gfsReadChar(.Fields("NR_CPF"))
            frmCredenciamento.txtNome.Text = objBanco.gfsReadChar(.Fields("NM_PESSOA"))
            frmCredenciamento.txtNomeCracha.Text = objBanco.gfsReadChar(.Fields("NM_CRACHA"))
            frmCredenciamento.mkeDataNascimento.Text = objBanco.gfsReadDate(.Fields("DT_NASCIMENTO"))
            Call objSystem.gsBuscaCombo(frmCredenciamento.cmbSexo, objBanco.gfsReadInt(.Fields("ID_SEXO")))
            If Not IsNull(.Fields("ID_PARENTESCO")) Then Call objSystem.gsBuscaCombo(frmCredenciamento.cmbParentesco, objBanco.gfsReadInt(.Fields("ID_PARENTESCO")))
            Call objSystem.gsBuscaCombo(frmCredenciamento.cmbEstadoCivil, objBanco.gfsReadInt(.Fields("ID_ESTADOCIVIL")))
            Call objSystem.gsBuscaCombo(frmCredenciamento.cmbFilial, objBanco.gfsReadInt(.Fields("CD_FILIAL")))
            Call objSystem.gsBuscaCombo(frmCredenciamento.cmbCentroCusto, objBanco.gfsReadInt(.Fields("CD_CENTROCUSTO")))
            If IsNull(.Fields("DH_LEITURA")) Then pLeituraColetor = False Else pLeituraColetor = True
            If Not IsNull(.Fields("CD_PESSOATITULAR")) Then Call objSystem.gsBuscaCombo(frmCredenciamento.cmbTitular, objBanco.gfsReadInt(.Fields("CD_PESSOATITULAR")))
            frmCredenciamento.txtObservacao.Text = objBanco.gfsReadChar(.Fields("TE_OBSERVACAO"))

            If Not IsNull(.Fields("DH_LEITURA")) Then frmCredenciamento.mkeDataLeitura.Text = Format(.Fields("DH_LEITURA"), "dd/mm/yyyy hh:mm:ss")

            If Not IsNull(.Fields("DH_IMPRESSAOCUPOM")) Then frmCredenciamento.mkeDataImpressaoTicket.Text = Format(.Fields("DH_IMPRESSAOCUPOM"), "dd/mm/yyyy hh:mm:ss")

            .Close

            Call frmCredenciamento.gsHabilitarCampos(False, 2)

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    Exit Sub

err_gsCarregaDadosPessoa:
    Call objSystem.gsExibeErros(Err, "gsCarregaDadosPessoa()", "Módulo Principal")

End Sub

Public Function gfiCalcularIdade(pDt_Nascimento As Date) As Integer

    On Error GoTo err_gfiCalcularIdade:

    Dim pDt_Evento As Date

    pDt_Evento = CDate("10/12/2011")

    If Month(pDt_Evento) < Month(pDt_Nascimento) Or (Month(pDt_Evento) = Month(pDt_Nascimento) And Day(pDt_Evento) < Day(pDt_Nascimento)) Then

        gfiCalcularIdade = Year(pDt_Evento) - Year(pDt_Nascimento) - 1

    Else

       gfiCalcularIdade = Year(pDt_Evento) - Year(pDt_Nascimento)

    End If

    Exit Function

err_gfiCalcularIdade:
    Call objSystem.gsExibeErros(Err, "gfiCalcularIdade()", "Módulo Principal")

End Function
