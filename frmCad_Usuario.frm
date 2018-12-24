VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCad_Usuario 
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   4
         Left            =   5400
         Picture         =   "frmCad_Usuario.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   3
         Left            =   4440
         Picture         =   "frmCad_Usuario.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   2
         Left            =   3480
         Picture         =   "frmCad_Usuario.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   2520
         Picture         =   "frmCad_Usuario.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   1560
         Picture         =   "frmCad_Usuario.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtSenha 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5640
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Text            =   "WWWWWWWWWW"
         Top             =   1770
         Width           =   2175
      End
      Begin VB.TextBox txtLogin 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "WWWWWWWWWW"
         Top             =   1770
         Width           =   2175
      End
      Begin VB.ComboBox cmbTipoUsuario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1300
         Width           =   6015
      End
      Begin VB.Frame fraDados 
         Caption         =   "Relação de Usuários:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2535
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   7575
         Begin MSFlexGridLib.MSFlexGrid mfgUsuarios 
            Height          =   2055
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   3625
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   12648447
            SelectionMode   =   1
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox txtNome 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   830
         Width           =   6015
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "9999"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   4432
         TabIndex        =   9
         Top             =   1830
         Width           =   750
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1830
         Width           =   660
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   890
         Width           =   705
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   825
      End
      Begin VB.Label lblTipoUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Usuário:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   1360
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmCad_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pAcao As Integer

Private Sub cmbTipoUsuario_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbTipoUsuario, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0 ' Novo
            Call psNovo

        Case 1 ' Salva
            Call psSalvar

        Case 2 ' Excluir
            Call psExcluir

        Case 3 ' Limpar Tela
            Call psLimparCampos

        Case 4 ' Sair
            Unload Me

    End Select

End Sub

Private Sub Form_Load()

    Call psCarregarComboTipoUsuario
    Call psLimparCampos

End Sub

Private Sub mfgUsuarios_Click()

    With mfgUsuarios

        .Col = 0
        If Len(Trim(.Text)) > 0 Then Call psMostrarUsuario(CInt(Trim(.Text)))

    End With

End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtNome, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtLogin, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtSenha, KeyAscii, ciUpper)
    
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdAcao(1).SetFocus
    End If

End Sub

Private Sub psCarregarComboTipoUsuario()

    On Error GoTo err_psCarregarComboTipoUsuario:

    cmbTipoUsuario.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_tipousuario ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbTipoUsuario.AddItem objBanco.gfsReadChar(.Fields("DE_TIPOUSUARIO"))
                cmbTipoUsuario.ItemData(cmbTipoUsuario.NewIndex) = objBanco.gfsReadInt(.Fields("ID_TIPOUSUARIO"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboTipoUsuario:
    Call objSystem.gsExibeErros(Err, "psCarregarComboTipoUsuario()", CStr(Me.Name))

End Sub

Private Sub psCarregarUsuarios()

    On Error GoTo err_psCarregarUsuarios:

    Dim pLinha As Integer

    pLinha = 1

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_usuario ( NULL, NULL, NULL, NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                While Not .EOF

                    If pLinha > mfgUsuarios.Rows - 1 Then mfgUsuarios.Rows = mfgUsuarios.Rows + 1

                    mfgUsuarios.Row = pLinha

                    mfgUsuarios.Col = 0: mfgUsuarios.CellAlignment = 1
                    mfgUsuarios.Text = objBanco.gfsReadInt(.Fields("ID_USUARIO"))
                    mfgUsuarios.Col = 1: mfgUsuarios.CellAlignment = 1
                    mfgUsuarios.Text = objBanco.gfsReadChar(.Fields("NM_USUARIO"))
                    mfgUsuarios.Col = 2: mfgUsuarios.CellAlignment = 1
                    mfgUsuarios.Text = objBanco.gfsReadInt(.Fields("DE_TIPOUSUARIO"))
                    mfgUsuarios.Col = 3: mfgUsuarios.CellAlignment = 1
                    mfgUsuarios.Text = IIf(objBanco.gfsReadChar(.Fields("FL_ATIVO")) = "S", "Sim", "Não")

                    pLinha = pLinha + 1
                    .MoveNext

                Wend

                mfgUsuarios.Row = 1: mfgUsuarios.Col = 0
                mfgUsuarios.ColSel = mfgUsuarios.Cols - 1

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarUsuarios:
    Call objSystem.gsExibeErros(Err, "psCarregarUsuarios()", CStr(Me.Name))

End Sub

Private Sub psExcluir()

    On Error GoTo err_psExcluir:

    With Me

        pSql = "": pSql = "CALL sp_del_usuario"
        pSql = pSql & " ("
        pSql = pSql & objBanco.gfsSaveInt(.txtCodigo.Text)
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao excluir o usuário!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Usuário excluído com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psExcluir:
    Call objSystem.gsExibeErros(Err, "psExcluir()", CStr(Me.Name))

End Sub

Private Sub psFormatarGrid()

    With Me

        With .mfgUsuarios

            .Clear
            .Cols = 4: .Rows = 20
            .Col = 0: .Row = 0

            .Col = 0: .CellAlignment = 3: .Text = "Código"
            .Col = 1: .CellAlignment = 3: .Text = "Nome"
            .Col = 2: .CellAlignment = 3: .Text = "Tipo"
            .Col = 3: .CellAlignment = 3: .Text = "Ativo"

            .ColWidth(0) = 900: .ColWidth(1) = 4200
            .ColWidth(2) = 900: .ColWidth(3) = 900

            .Row = 1: .Col = 0: .ColSel = .Cols - 1

        End With

    End With

End Sub

Private Sub psHabilitarCampos(pHabilita As Boolean)

    With Me

        .txtNome.Enabled = pHabilita
        .cmbTipoUsuario.Enabled = pHabilita
        .txtLogin.Enabled = pHabilita
        .txtSenha.Enabled = pHabilita

        With .cmdAcao

            .Item(0).Enabled = Not pHabilita
            .Item(1).Enabled = pHabilita
            .Item(2).Enabled = pHabilita
            .Item(3).Enabled = pHabilita

        End With

    End With

End Sub

Private Sub psLimparCampos()

    With Me

        Call objSystem.gsLimparText(.Name)
        .cmbTipoUsuario.ListIndex = -1

        Call psHabilitarCampos(False)
        Call psFormatarGrid
        Call psCarregarUsuarios

        pAcao = 0

    End With

End Sub

Private Sub psMostrarUsuario(pCodigo As Integer)

    On Error GoTo err_psMostrarUsuario:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_usuario"
        pSql = pSql & " ( " & objBanco.gfsSaveInt(pCodigo)
        pSql = pSql & ", NULL"
        pSql = pSql & ", NULL"
        pSql = pSql & ", NULL"
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                pAcao = 2
                txtCodigo.Text = objBanco.gfsReadInt(.Fields("ID_USUARIO"))
                txtNome.Text = objBanco.gfsReadChar(.Fields("NM_USUARIO"))
                Call objSystem.gsBuscaCombo(cmbTipoUsuario, objBanco.gfsReadInt(.Fields("ID_TIPOUSUARIO")))
                txtLogin.Text = objBanco.gfsReadChar(.Fields("LG_USUARIO"))
                txtSenha.Text = objSystem.gfsEncryptString(objBanco.gfsReadChar(.Fields("PW_USUARIO")), ciDecrypt)

                Call psHabilitarCampos(True)

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psMostrarUsuario:
    Call objSystem.gsExibeErros(Err, "psMostrarUsuario()", CStr(Me.Name))

End Sub

Private Sub psNovo()

    On Error GoTo err_psNovo:

    With Me

        pAcao = 1

        .txtCodigo.Text = objBanco.gflProximoRegistro("tb_usuario")

        Call psHabilitarCampos(True)

        .txtNome.SetFocus

    End With

    Exit Sub

err_psNovo:
    Call objSystem.gsExibeErros(Err, "psNovo()", CStr(Me.Name))

End Sub

Private Sub psSalvar()

    On Error GoTo err_psSalvar:

    If Not pfbValidarCampos Then Exit Sub

    With Me

        If pAcao = 1 Then

            pSql = "": pSql = "CALL sp_ins_usuario"

        Else

            pSql = "": pSql = "CALL sp_upd_usuario"

        End If

        pSql = pSql & " (" & objBanco.gfsSaveInt(.txtCodigo.Text)

        If Len(Trim(.txtNome.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(Trim(.txtNome.Text))

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtLogin.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(Trim(.txtLogin.Text))

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtSenha.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(objSystem.gfsEncryptString(Trim(.txtSenha.Text), ciEncrypt))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbTipoUsuario.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbTipoUsuario.ItemData(.cmbTipoUsuario.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar("S")
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao salvar o usuário!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Usuário salvo com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psSalvar:
    Call objSystem.gsExibeErros(Err, "psSalvar()", CStr(Me.Name))

End Sub

Private Function pfbValidarCampos() As Boolean

    On Error GoTo err_pfbValidarCampos:

    pfbValidarCampos = False

    With Me

        If Len(Trim(.txtNome.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar o nome do usuário!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtNome.SetFocus: Exit Function

        End If

        If .cmbTipoUsuario.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o tipo de usuário!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbTipoUsuario.SetFocus: Exit Function
        
        End If

        If Len(Trim(.txtLogin.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar o login do usuário!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtLogin.SetFocus: Exit Function

        End If

        If Len(Trim(.txtSenha.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar a senha do usuário!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtSenha.SetFocus: Exit Function

        End If
        
    End With

    pfbValidarCampos = True

    Exit Function

err_pfbValidarCampos:
    Call objSystem.gsExibeErros(Err, "pfbValidarCampos()", CStr(Me.Name))

End Function

