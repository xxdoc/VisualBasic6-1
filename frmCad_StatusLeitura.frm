VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCad_StatusLeitura 
   Caption         =   "Cadastro de Status de Leitura"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   4
         Left            =   4995
         Picture         =   "frmCad_StatusLeitura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   3
         Left            =   4035
         Picture         =   "frmCad_StatusLeitura.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   2
         Left            =   3075
         Picture         =   "frmCad_StatusLeitura.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   2115
         Picture         =   "frmCad_StatusLeitura.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   1155
         Picture         =   "frmCad_StatusLeitura.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4080
         Width           =   975
      End
      Begin VB.Frame fraDados 
         Caption         =   "Relação de Status de Leitura:"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   6615
         Begin MSFlexGridLib.MSFlexGrid mfgStatusLeitura 
            Height          =   2055
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   6375
            _ExtentX        =   11245
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
      Begin VB.TextBox txtDescricao 
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
         MaxLength       =   25
         TabIndex        =   4
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   810
         Width           =   5055
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
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
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
         Top             =   870
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmCad_StatusLeitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pAcao As Integer

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0
            Call psNovo

        Case 1
            Call psSalvar

        Case 2
            Call psExcluir

        Case 3
            Call psLimparCampos

        Case 4
            Unload Me

    End Select

End Sub

Private Sub Form_Load()

    Call psLimparCampos

End Sub

Private Sub mfgStatusLeitura_Click()

    With mfgStatusLeitura

        .Col = 0
        If Len(Trim(.Text)) > 0 Then Call psMostrarStatusLeitura(CInt(Trim(.Text)))

    End With

End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtDescricao, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: cmdAcao(1).SetFocus
    End If

End Sub

Private Sub psCarregarStatusLeitura()

    On Error GoTo err_psCarregarStatusLeitura:

    Dim pLinha As Integer

    pLinha = 1

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_statusleitura ( NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                While Not .EOF

                    If pLinha > mfgStatusLeitura.Rows - 1 Then mfgStatusLeitura.Rows = mfgStatusLeitura.Rows + 1

                    mfgStatusLeitura.Row = pLinha

                    mfgStatusLeitura.Col = 0: mfgStatusLeitura.CellAlignment = 1
                    mfgStatusLeitura.Text = objBanco.gfsReadInt(.Fields("ID_STATUSLEITURA"))
                    mfgStatusLeitura.Col = 1: mfgStatusLeitura.CellAlignment = 1
                    mfgStatusLeitura.Text = objBanco.gfsReadChar(.Fields("DE_STATUSLEITURA"))
                    mfgStatusLeitura.Col = 2: mfgStatusLeitura.CellAlignment = 1
                    mfgStatusLeitura.Text = IIf(objBanco.gfsReadChar(.Fields("FL_ATIVO")) = "S", "Sim", "Não")

                    pLinha = pLinha + 1
                    .MoveNext

                Wend

                mfgStatusLeitura.Row = 1: mfgStatusLeitura.Col = 0
                mfgStatusLeitura.ColSel = mfgStatusLeitura.Cols - 1

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarStatusLeitura:
    Call objSystem.gsExibeErros(Err, "psCarregarStatusLeitura()", CStr(Me.Name))

End Sub

Private Sub psExcluir()

    On Error GoTo err_psExcluir:

    With Me

        pSql = "": pSql = "CALL sp_del_statusleitura"
        pSql = pSql & " ("
        pSql = pSql & objBanco.gfsSaveInt(.txtCodigo.Text)
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao excluir o status de leitura!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Status de leitura excluído com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psExcluir:
    Call objSystem.gsExibeErros(Err, "psExcluir()", CStr(Me.Name))

End Sub

Private Sub psFormatarGrid()

    With Me

        With .mfgStatusLeitura

            .Clear
            .Cols = 3: .Rows = 20
            .Col = 0: .Row = 0

            .Col = 0: .CellAlignment = 3: .Text = "Código"
            .Col = 1: .CellAlignment = 3: .Text = "Descrição"
            .Col = 2: .CellAlignment = 3: .Text = "Ativo"

            .ColWidth(0) = 900: .ColWidth(1) = 4200: .ColWidth(2) = 900

            .Row = 1: .Col = 0: .ColSel = .Cols - 1

        End With

    End With

End Sub

Private Sub psHabilitarCampos(pHabilita As Boolean)

    With Me

        .txtDescricao.Enabled = pHabilita

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
        Call psHabilitarCampos(False)
        Call psFormatarGrid
        Call psCarregarStatusLeitura

        pAcao = 0

    End With

End Sub

Private Sub psMostrarStatusLeitura(pCodigo As Integer)

    On Error GoTo err_psMostrarStatusLeitura:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_statusleitura"
        pSql = pSql & " ( "
        pSql = pSql & objBanco.gfsSaveInt(pCodigo)
        pSql = pSql & " );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                pAcao = 2
                txtCodigo.Text = objBanco.gfsReadInt(.Fields("ID_STATUSLEITURA"))
                txtDescricao.Text = objBanco.gfsReadChar(.Fields("DE_STATUSLEITURA"))

                Call psHabilitarCampos(True)

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psMostrarStatusLeitura:
    Call objSystem.gsExibeErros(Err, "psMostrarStatusLeitura()", CStr(Me.Name))

End Sub

Private Sub psNovo()

    On Error GoTo err_psNovo:

    With Me

        pAcao = 1

        .txtCodigo.Text = objBanco.gflProximoRegistro("tb_statusleitura")

        Call psHabilitarCampos(True)

        .txtDescricao.SetFocus

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

            pSql = "": pSql = "CALL sp_ins_statusleitura"

        Else

            pSql = "": pSql = "CALL sp_upd_statusleitura"

        End If

        pSql = pSql & " (" & objBanco.gfsSaveInt(.txtCodigo.Text)

        If Len(Trim(.txtDescricao.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(Trim(.txtDescricao.Text))

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar("S")
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao salvar o status de leitura!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Status de leitura salvo com sucesso!"
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

        If Len(Trim(.txtDescricao.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar a descrição do status de leitura!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtDescricao.SetFocus: Exit Function

        End If

    End With

    pfbValidarCampos = True

    Exit Function

err_pfbValidarCampos:
    Call objSystem.gsExibeErros(Err, "pfbValidarCampos()", CStr(Me.Name))

End Function

