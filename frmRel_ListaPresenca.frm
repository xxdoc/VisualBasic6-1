VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRel_ListaPresenca 
   Caption         =   "Lista de Presença"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   Icon            =   "frmRel_ListaPresenca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin Crystal.CrystalReport rptRelatorios 
         Left            =   10680
         Top             =   8400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   2
         Left            =   10200
         Picture         =   "frmRel_ListaPresenca.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Sair"
         Top             =   8880
         Width           =   975
      End
      Begin VB.Frame fraLista 
         Caption         =   "Participantes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   7095
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11055
         Begin VB.OptionButton optSituacao 
            Caption         =   "Ausentes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton optSituacao 
            Caption         =   "Presentes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin MSComctlLib.ListView ltwParticipantes 
            Height          =   6015
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   10610
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "¤"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Filial"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Centro de Custo"
               Object.Width           =   3351
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Matrícula"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Nome"
               Object.Width           =   7937
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Categoria"
               Object.Width           =   2822
            EndProperty
         End
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   8040
         Picture         =   "frmRel_ListaPresenca.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Atualizar"
         Top             =   8880
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   9120
         Picture         =   "frmRel_ListaPresenca.frx":730E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Sair"
         Top             =   8880
         Width           =   975
      End
      Begin VB.Frame fraTotalGeral 
         Caption         =   "Totais:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   7440
         Width           =   7815
         Begin MSComctlLib.ListView ltwTotais 
            Height          =   1935
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "¤"
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Categoria"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Quantidade"
               Object.Width           =   3528
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmRel_ListaPresenca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pContador As Long

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0 ' Atualizar dados
            Call psAtualizar

        Case 1 ' Imprimir
            Call psImprimir

        Case 2 ' Sair
            Unload Me

    End Select

End Sub

Private Sub optSituacao_Click(Index As Integer)

    Call psAtualizar

End Sub

Private Sub psAtualizar()

    On Error GoTo err_psAtualizar:

    MousePointer = vbHourglass

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_lista_presenca ( "

        If optSituacao(0).Value Then

            pSql = pSql & objBanco.gfsSaveChar("N")

        ElseIf optSituacao(1).Value Then

            pSql = pSql & objBanco.gfsSaveChar("S")

        End If

        pSql = pSql & " );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            pContador = 1
            ltwParticipantes.ListItems.Clear

            Do While Not .EOF
                Call Carregar_Dados
                pContador = pContador + 1
                rs.MoveNext
            Loop
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Call psCarregarTotais

    If ltwParticipantes.ListItems.Count = 0 Then cmdAcao(1).Enabled = False Else cmdAcao(1).Enabled = True

    MousePointer = vbDefault

    Exit Sub

err_psAtualizar:
    MousePointer = vbDefault
    Call objSystem.gsExibeErros(Err, "psAtualizar()", CStr(Me.Name))

End Sub

Private Function Carregar_Dados()

    Dim pText_Aux As String

    On Error GoTo err_Carregar_Dados:

    With ltwParticipantes

        .ListItems.Add pContador
        .ListItems(pContador).ListSubItems.Add 1, , objBanco.gfsReadInt(rs.Fields("CD_FILIAL"))
        .ListItems(pContador).ListSubItems.Add 2, , objBanco.gfsReadInt(rs.Fields("CD_CENTROCUSTO"))
        .ListItems(pContador).ListSubItems.Add 3, , objBanco.gfsReadInt(rs.Fields("NR_MATRICULA"))
        .ListItems(pContador).ListSubItems.Add 4, , objBanco.gfsReadChar(rs.Fields("NM_PESSOA"))
        .ListItems(pContador).ListSubItems.Add 5, , objBanco.gfsReadChar(rs.Fields("DE_CATEGORIA"))

    End With

    Exit Function

err_Carregar_Dados:
    Call objSystem.gsExibeErros(Err, "Carregar_Dados()", CStr(Me.Caption))

End Function

Private Sub Form_Load()

    Call psAtualizar

End Sub

Private Sub psCarregarTotais()

    On Error GoTo err_psCarregarTotais:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_total_lista_presenca ( "

        If optSituacao(0).Value Then

            pSql = pSql & objBanco.gfsSaveChar("N")

        ElseIf optSituacao(1).Value Then

            pSql = pSql & objBanco.gfsSaveChar("S")

        End If

        pSql = pSql & " );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            pContador = 1
            ltwTotais.ListItems.Clear

            While Not .EOF

                ltwTotais.ListItems.Add pContador
                ltwTotais.ListItems(pContador).ListSubItems.Add 1, , objBanco.gfsReadChar(rs.Fields("DE_CATEGORIA"))
                ltwTotais.ListItems(pContador).ListSubItems.Add 2, , FormatNumber(objBanco.gfsReadInt(rs.Fields("QT_PESSOA")), 0)

                If LCase(objBanco.gfsReadChar(rs.Fields("DE_CATEGORIA"))) = LCase("TOTAL GERAL") Then

                    ltwTotais.ListItems(pContador).ListSubItems(1).ForeColor = &HC0&
                    ltwTotais.ListItems(pContador).ListSubItems(1).Bold = True
                    ltwTotais.ListItems(pContador).ListSubItems(2).ForeColor = &HC0&
                    ltwTotais.ListItems(pContador).ListSubItems(2).Bold = True

                End If

                pContador = pContador + 1
                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarTotais:
    Call objSystem.gsExibeErros(Err, "psCarregarTotais()", CStr(Me.Name))

End Sub

Private Sub psImprimir()

    Dim pRptSql    As String
    Dim pRptFile   As String
    Dim pRptTitulo As String

    On Error GoTo err_psImprimir:

    pRptSql = ""
    pRptSql = "SELECT tb_pessoa1.`nr_matricula`, tb_pessoa1.`nm_pessoa`, tb_pessoa1.`cd_filial`, tb_pessoa1.`cd_centrocusto`, tb_categoria1.`de_categoria`"
    pRptSql = pRptSql & " FROM eletropaulo11.tb_pessoa tb_pessoa1 INNER JOIN eletropaulo11.tb_categoria tb_categoria1 ON ( tb_pessoa1.id_categoria = tb_categoria1.id_categoria )"
    pRptSql = pRptSql & " WHERE tb_pessoa1.fl_ativo = 'S'"

    If optSituacao(0).Value Then

        pRptSql = pRptSql & " AND tb_pessoa1.fl_presente = " & objBanco.gfsSaveChar("N")

    Else

        pRptSql = pRptSql & " AND tb_pessoa1.fl_presente = " & objBanco.gfsSaveChar("S")

    End If

    pRptSql = pRptSql & " ORDER BY tb_pessoa1.cd_filial, tb_pessoa1.cd_centrocusto, tb_pessoa1.nr_matricula, tb_pessoa1.id_categoria, tb_pessoa1.nm_pessoa;"

    pRptFile = pPath & "\Relatórios\Rel_Lista_Presenca.rpt"

    pRptTitulo = "Lista de Presença."

    If gfbImprimir(pRptSql, pRptFile, pRptTitulo, rptRelatorios, 0) = False Then

        pMsg = "": pMsg = "Erro ao Gerar Relatório."
        MsgBox pMsg, vbCritical, "Atenção."
        Exit Sub

    End If

    Exit Sub

err_psImprimir:
    Call objSystem.gsExibeErros(Err, "psImprimir()", CStr(Me.Name))

End Sub

