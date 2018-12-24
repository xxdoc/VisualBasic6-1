VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCad_Categoria 
   Caption         =   "Cadastro de Categorias"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   4
         Left            =   5603
         Picture         =   "frmCad_Categoria.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4700
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   3
         Left            =   4643
         Picture         =   "frmCad_Categoria.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4700
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   2
         Left            =   3683
         Picture         =   "frmCad_Categoria.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4700
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   2723
         Picture         =   "frmCad_Categoria.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4700
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   1763
         Picture         =   "frmCad_Categoria.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4700
         Width           =   975
      End
      Begin VB.TextBox txtOrdem 
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
         TabIndex        =   6
         Text            =   "9999"
         Top             =   1320
         Width           =   615
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
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   810
         Width           =   6015
      End
      Begin VB.Frame fraDados 
         Caption         =   "Relação de Categorias:"
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
         TabIndex        =   7
         Top             =   1920
         Width           =   7575
         Begin MSFlexGridLib.MSFlexGrid mfgCategorias 
            Height          =   2055
            Left            =   120
            TabIndex        =   8
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
      Begin VB.Label lblOrdem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
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
         Top             =   1380
         Width           =   810
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
Attribute VB_Name = "frmCad_Categoria"
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

Private Sub mfgCategorias_Click()

    With mfgCategorias

        .Col = 0
        If Len(Trim(.Text)) > 0 Then Call psMostrarCategoria(CInt(Trim(.Text)))

    End With

End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtDescricao, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtOrdem_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtOrdem, KeyAscii, ciInt)
    
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdAcao(1).SetFocus
    End If

End Sub

Private Sub psCarregarCategorias()

    On Error GoTo err_psCarregarCategorias:

    Dim pLinha As Integer

    pLinha = 1

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_categoria ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                While Not .EOF

                    If pLinha > mfgCategorias.Rows - 1 Then mfgCategorias.Rows = mfgCategorias.Rows + 1

                    mfgCategorias.Row = pLinha

                    mfgCategorias.Col = 0: mfgCategorias.CellAlignment = 1
                    mfgCategorias.Text = objBanco.gfsReadInt(.Fields("ID_CATEGORIA"))
                    mfgCategorias.Col = 1: mfgCategorias.CellAlignment = 1
                    mfgCategorias.Text = objBanco.gfsReadChar(.Fields("DE_CATEGORIA"))
                    mfgCategorias.Col = 2: mfgCategorias.CellAlignment = 1
                    mfgCategorias.Text = objBanco.gfsReadInt(.Fields("NR_ORDEM"))
                    mfgCategorias.Col = 3: mfgCategorias.CellAlignment = 1
                    mfgCategorias.Text = IIf(objBanco.gfsReadChar(.Fields("FL_ATIVO")) = "S", "Sim", "Não")

                    pLinha = pLinha + 1
                    .MoveNext

                Wend

                mfgCategorias.Row = 1: mfgCategorias.Col = 0
                mfgCategorias.ColSel = mfgCategorias.Cols - 1

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarCategorias:
    Call objSystem.gsExibeErros(Err, "psCarregarCategorias()", CStr(Me.Name))

End Sub

Private Sub psExcluir()

    On Error GoTo err_psExcluir:

    With Me

        pSql = "": pSql = "CALL sp_del_categoria"
        pSql = pSql & " ("
        pSql = pSql & objBanco.gfsSaveInt(.txtCodigo.Text)
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao excluir a categoria!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Categoria excluída com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psExcluir:
    Call objSystem.gsExibeErros(Err, "psExcluir()", CStr(Me.Name))

End Sub

Private Sub psFormatarGrid()

    With Me

        With .mfgCategorias

            .Clear
            .Cols = 4: .Rows = 20
            .Col = 0: .Row = 0

            .Col = 0: .CellAlignment = 3: .Text = "Código"
            .Col = 1: .CellAlignment = 3: .Text = "Descrição"
            .Col = 2: .CellAlignment = 3: .Text = "Ordem"
            .Col = 3: .CellAlignment = 3: .Text = "Ativo"

            .ColWidth(0) = 900: .ColWidth(1) = 4200
            .ColWidth(2) = 900: .ColWidth(3) = 900

            .Row = 1: .Col = 0: .ColSel = .Cols - 1

        End With

    End With

End Sub

Private Sub psHabilitarCampos(pHabilita As Boolean)

    With Me

        .txtDescricao.Enabled = pHabilita
        .txtOrdem.Enabled = pHabilita

    End With

End Sub

Private Sub psLimparCampos()

    With Me

        Call objSystem.gsLimparText(.Name)

        Call psHabilitarCampos(False)

        With .cmdAcao

            .Item(0).Enabled = True
            .Item(1).Enabled = False
            .Item(2).Enabled = False
            .Item(3).Enabled = True

        End With

        Call psFormatarGrid
        Call psCarregarCategorias

        pAcao = 0

    End With

End Sub

Private Sub psMostrarCategoria(pCodigo As Integer)

    On Error GoTo err_psMostrarCategoria:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_categoria"
        pSql = pSql & " ( " & objBanco.gfsSaveInt(pCodigo)
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                pAcao = 2
                txtCodigo.Text = objBanco.gfsReadInt(.Fields("ID_CATEGORIA"))
                txtDescricao.Text = objBanco.gfsReadChar(.Fields("DE_CATEGORIA"))
                txtOrdem.Text = objBanco.gfsReadInt(.Fields("NR_ORDEM"))

                Call psHabilitarCampos(True)

                With cmdAcao

                    .Item(0).Enabled = False
                    .Item(1).Enabled = True
                    .Item(2).Enabled = True
                    .Item(3).Enabled = True

                End With

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psMostrarCategoria:
    Call objSystem.gsExibeErros(Err, "psMostrarCategoria()", CStr(Me.Name))

End Sub

Private Sub psNovo()

    On Error GoTo err_psNovo:

    With Me

        pAcao = 1

        .txtCodigo.Text = objBanco.gflProximoRegistro("tb_categoria")

        Call psHabilitarCampos(True)

        .txtDescricao.SetFocus

        With .cmdAcao

            .Item(0).Enabled = False
            .Item(1).Enabled = True
            .Item(2).Enabled = False
            .Item(3).Enabled = True

        End With

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

            pSql = "": pSql = "CALL sp_ins_categoria"

        Else

            pSql = "": pSql = "CALL sp_upd_categoria"

        End If

        pSql = pSql & " (" & objBanco.gfsSaveInt(.txtCodigo.Text)

        If Len(Trim(.txtDescricao.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(Trim(.txtDescricao.Text))

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtOrdem.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(Trim(.txtOrdem.Text))

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar("S")
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao salvar a categoria!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Categoria salva com sucesso!"
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

            pMsg = "": pMsg = "É necessário informar a descrição da categoria!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtDescricao.SetFocus: Exit Function

        End If

    End With

    pfbValidarCampos = True

    Exit Function

err_pfbValidarCampos:
    Call objSystem.gsExibeErros(Err, "pfbValidarCampos()", CStr(Me.Name))

End Function
