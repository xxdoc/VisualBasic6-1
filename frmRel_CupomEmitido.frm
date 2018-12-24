VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmRel_CupomEmitido 
   Caption         =   "Quantidade de Cupons Emitidos"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11070
   Icon            =   "frmRel_CupomEmitido.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabPainel 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados"
      TabPicture(0)   =   "frmRel_CupomEmitido.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gráfico"
      TabPicture(1)   =   "frmRel_CupomEmitido.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGrafico"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraGrafico 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   10575
         Begin MSChart20Lib.MSChart Graf1 
            Height          =   6495
            Left            =   120
            OleObjectBlob   =   "frmRel_CupomEmitido.frx":0902
            TabIndex        =   5
            Top             =   120
            Width           =   10335
         End
      End
      Begin VB.Frame fraDados 
         Height          =   6855
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   10575
         Begin MSFlexGridLib.MSFlexGrid mfgDados 
            Height          =   6495
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   11456
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":2C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":3532
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":3E0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":46E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":4FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRel_CupomEmitido.frx":511A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrBarra 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   1482
      ButtonWidth     =   1588
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Pesquisar"
            Key             =   "Pesquisar"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A"
                  Text            =   "Categoria"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B"
                  Text            =   "Centro de Custo"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "C"
                  Text            =   "Estado Civil"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "D"
                  Text            =   "Faixa Etária"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "E"
                  Text            =   "Sexo"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Atualizar"
            Key             =   "Atualizar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Imprimir"
            Key             =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Fechar"
            Key             =   "Fechar"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin Crystal.CrystalReport rptRelatorios 
         Left            =   9720
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuCriterio 
         Caption         =   "Critérios"
         Begin VB.Menu mnuCategoria 
            Caption         =   "Categoria"
         End
         Begin VB.Menu mnuCentroCusto 
            Caption         =   "Centro de Custo"
         End
         Begin VB.Menu mnuEstadoCivil 
            Caption         =   "Estado Civil"
         End
         Begin VB.Menu mnuFaixaEtaria 
            Caption         =   "Faixa Etária"
         End
         Begin VB.Menu mnuSexo 
            Caption         =   "Sexo"
         End
      End
      Begin VB.Menu mnuAtualizar 
         Caption         =   "Atualizar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFechar 
         Caption         =   "Fechar"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmRel_CupomEmitido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Call mnuCategoria_Click

End Sub

Private Sub mnuAtualizar_Click()

    Call psCarregarDados

End Sub

Private Sub mnuCategoria_Click()

    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    mnuCategoria.Checked = True: mnuCentroCusto.Checked = False
    mnuEstadoCivil.Checked = False: mnuFaixaEtaria.Checked = False: mnuSexo.Checked = False
    Call psCarregarDados

End Sub

Private Sub mnuCentroCusto_Click()

    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    mnuCategoria.Checked = False: mnuCentroCusto.Checked = True
    mnuEstadoCivil.Checked = False: mnuFaixaEtaria.Checked = False: mnuSexo.Checked = False
    Call psCarregarDados

End Sub

Private Sub mnuEstadoCivil_Click()

    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    mnuCategoria.Checked = False: mnuCentroCusto.Checked = False
    mnuEstadoCivil.Checked = True: mnuFaixaEtaria.Checked = False: mnuSexo.Checked = False
    Call psCarregarDados

End Sub

Private Sub mnuFaixaEtaria_Click()

    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    mnuCategoria.Checked = False: mnuCentroCusto.Checked = False
    mnuEstadoCivil.Checked = False: mnuFaixaEtaria.Checked = True: mnuSexo.Checked = False
    Call psCarregarDados

End Sub

Private Sub mnuFechar_Click()

    Unload Me

End Sub

Private Sub mnuImprimir_Click()

    Call psImprimir

End Sub

Private Sub mnuSexo_Click()

    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    mnuCategoria.Checked = False: mnuCentroCusto.Checked = False
    mnuEstadoCivil.Checked = False: mnuFaixaEtaria.Checked = False: mnuSexo.Checked = True
    Call psCarregarDados

End Sub

Private Sub tbrBarra_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Mid(Button.Key, 1, 1)
        Case "A" ' Atualizar
            Call mnuAtualizar_Click

        Case "I" ' Imprimir
            Call mnuImprimir_Click

        Case "F" ' Fechar
            Call mnuFechar_Click

    End Select

End Sub

Private Sub tbrBarra_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    Select Case ButtonMenu.Index
        Case "1" ' Categoria
            Call mnuCategoria_Click

        Case "2" ' Centro de Custo
            Call mnuCentroCusto_Click

        Case "3" ' Estado Civil
            Call mnuEstadoCivil_Click

        Case "4" ' Faixa Etaria
            Call mnuFaixaEtaria_Click

        Case "5" ' Sexo
            Call mnuSexo_Click

    End Select

End Sub

Private Sub psCarregarDados()

    On Error GoTo err_psCarregarDados:

    Me.MousePointer = vbHourglass

    Call psFormatarGrid
    Call psCarregarGrid
    Call psGerarGrafico

    Me.MousePointer = vbDefault

    Exit Sub

err_psCarregarDados:
    Call objSystem.gsExibeErros(Err, "psCarregarDados()", CStr(Me.Name))

End Sub

Private Sub psCarregarGrid()

    On Error GoTo err_psCarregarGrid:

    Dim pLinha As Integer

    pLinha = 1

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        If mnuCategoria.Checked Then

            pSql = "": pSql = "CALL sp_con_cupom_emitido_categoria"

        ElseIf mnuCentroCusto.Checked Then

            pSql = "": pSql = "CALL sp_con_cupom_emitido_centrocusto"

        ElseIf mnuEstadoCivil.Checked Then

            pSql = "": pSql = "CALL sp_con_cupom_emitido_estadocivil"

        ElseIf mnuFaixaEtaria.Checked Then

            pSql = "": pSql = "CALL sp_con_cupom_emitido_faixaetaria"

        Else

            pSql = "": pSql = "CALL sp_con_cupom_emitido_sexo"

        End If

        pSql = pSql & " ( " & objBanco.gfsSaveInt(pCd_User) & " );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                While Not .EOF

                    If pLinha > mfgDados.Rows - 1 Then mfgDados.Rows = mfgDados.Rows + 1

                    mfgDados.Row = pLinha

                    mfgDados.Col = 0: mfgDados.CellAlignment = 1
                    mfgDados.Text = objBanco.gfsReadChar(.Fields("DE_CRITERIO"))
                    mfgDados.Col = 1: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadInt(.Fields("QT_CERVEJA")), 0)
                    mfgDados.Col = 2: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadChar(.Fields("QT_ESPETOCARNE")), 0)
                    mfgDados.Col = 3: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadInt(.Fields("QT_ESPETOFRANGO")), 0)
                    mfgDados.Col = 4: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadChar(.Fields("QT_ESPETOLINGUICA")), 0)
                    mfgDados.Col = 5: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadInt(.Fields("QT_LANCHE")), 0)
                    mfgDados.Col = 6: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadChar(.Fields("QT_PAPINHA")), 0)
                    mfgDados.Col = 7: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadInt(.Fields("QT_PIPOCA")), 0)
                    mfgDados.Col = 8: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadChar(.Fields("QT_SORVETE")), 0)
                    mfgDados.Col = 9: mfgDados.CellAlignment = 4
                    mfgDados.Text = FormatNumber(objBanco.gfsReadChar(.Fields("QT_TOTAL")), 0)

                    pLinha = pLinha + 1
                    .MoveNext

                Wend

                mfgDados.Row = 1: mfgDados.Col = 0
                mfgDados.ColSel = mfgDados.Cols - 1

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarGrid:
    Call objSystem.gsExibeErros(Err, "psCarregarGrid()", CStr(Me.Name))

End Sub

Private Sub psFormatarGrid()

    On Error GoTo err_psFormatarGrid:

    With Me

        .mfgDados.Clear
        .mfgDados.Cols = 10: .mfgDados.Rows = 20
        .mfgDados.Col = 0: .mfgDados.Row = 0

        If .mnuCategoria.Checked Then

            .mfgDados.Col = 0: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Categoria"

        ElseIf .mnuCentroCusto.Checked Then

            .mfgDados.Col = 0: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Centro de Custo"

        ElseIf .mnuEstadoCivil.Checked Then

            .mfgDados.Col = 0: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Estado Civil"

        ElseIf .mnuFaixaEtaria.Checked Then

            .mfgDados.Col = 0: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Faixa Etária"

        Else

            .mfgDados.Col = 0: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Sexo"

        End If

        .mfgDados.Col = 1: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Cerveja"
        .mfgDados.Col = 2: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Espeto Carne"
        .mfgDados.Col = 3: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Espeto Frango"
        .mfgDados.Col = 4: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Espeto Linguiça"
        .mfgDados.Col = 5: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Lanche"
        .mfgDados.Col = 6: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Papinha"
        .mfgDados.Col = 7: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Pipoca"
        .mfgDados.Col = 8: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Sorvete"
        .mfgDados.Col = 9: .mfgDados.CellAlignment = 3: .mfgDados.Text = "Total"

        If .mnuCategoria.Checked Then

            .mfgDados.ColWidth(0) = 2000

        ElseIf .mnuCentroCusto.Checked Then

            .mfgDados.ColWidth(0) = 2000

        ElseIf .mnuEstadoCivil.Checked Then

            .mfgDados.ColWidth(0) = 2500

        ElseIf .mnuFaixaEtaria.Checked Then

            .mfgDados.ColWidth(0) = 2500

        Else

            .mfgDados.ColWidth(0) = 2500

        End If

        .mfgDados.ColWidth(1) = 1200
        .mfgDados.ColWidth(2) = 1500
        .mfgDados.ColWidth(3) = 1700
        .mfgDados.ColWidth(4) = 1800
        .mfgDados.ColWidth(5) = 1000
        .mfgDados.ColWidth(6) = 1100
        .mfgDados.ColWidth(7) = 1000
        .mfgDados.ColWidth(8) = 1000
        .mfgDados.ColWidth(9) = 1200

        .mfgDados.Row = 1: .mfgDados.Col = 0: .mfgDados.ColSel = .mfgDados.Cols - 1

    End With

    Exit Sub

err_psFormatarGrid:
    Call objSystem.gsExibeErros(Err, "psFormatarGrid()", CStr(Me.Name))

End Sub

Private Sub psGerarGrafico()

    On Error GoTo err_psGerarGrafico

    Dim pTotal_Reg      As Long
    Dim pCont_Coluna    As Integer
    Dim pCont_Linha     As Integer

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "SELECT *"
        pSql = pSql & " FROM tb_relatoriocupom"
        pSql = pSql & " WHERE id_usuario = " & objBanco.gfsSaveInt(pCd_User)

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            pTotal_Reg = pflTotalRegistro(pSql)

            Graf1.ShowLegend = True
            Graf1.ColumnCount = 8: Graf1.RowCount = pTotal_Reg

            Graf1.Visible = True

            While Not .EOF

                For pCont_Linha = 1 To pTotal_Reg

                    Graf1.Row = pCont_Linha
                    Graf1.RowLabel = .Fields("DE_CRITERIO")

                    For pCont_Coluna = 1 To 8

                        Graf1.Column = pCont_Coluna

                        Select Case pCont_Coluna
                            Case 1 ' Cerveja
                                Graf1.ColumnLabel = "Cerveja"
                                Graf1.Data = .Fields("QT_CERVEJA")

                            Case 2 ' Espeto Carne
                                Graf1.ColumnLabel = "Espeto Carne"
                                Graf1.Data = .Fields("QT_ESPETOCARNE")

                            Case 3 ' Espeto Frango
                                Graf1.ColumnLabel = "Espeto Frango"
                                Graf1.Data = .Fields("QT_ESPETOFRANGO")

                            Case 4 ' Espeto Linguiça
                                Graf1.ColumnLabel = "Espeto Linguiça"
                                Graf1.Data = .Fields("QT_ESPETOLINGUICA")

                            Case 5 ' Lanche
                                Graf1.ColumnLabel = "Lanche"
                                Graf1.Data = .Fields("QT_LANCHE")

                            Case 6 ' Papinha
                                Graf1.ColumnLabel = "Papinha"
                                Graf1.Data = .Fields("QT_PAPINHA")

                            Case 7 ' Pipoca
                                Graf1.ColumnLabel = "Pipoca"
                                Graf1.Data = .Fields("QT_PIPOCA")

                            Case 8 ' Sorvete
                                Graf1.ColumnLabel = "Sorvete"
                                Graf1.Data = .Fields("QT_SORVETE")

                        End Select

                    Next pCont_Coluna

                    .MoveNext

                Next pCont_Linha

            Wend

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psGerarGrafico:
    Call objSystem.gsExibeErros(Err, "psGerarGrafico()", CStr(Me.Name))

End Sub

Private Sub psImprimir()

    Dim pRptSql    As String
    Dim pRptFile   As String
    Dim pRptTitulo As String

    On Error GoTo err_psImprimir:

    pRptSql = ""
    pRptSql = "SELECT tb_relatoriocupom1.`de_criterio`, tb_relatoriocupom1.`de_subtitulo`, tb_relatoriocupom1.`qt_cerveja`,"
    pRptSql = pRptSql & " tb_relatoriocupom1.`qt_espetocarne`, tb_relatoriocupom1.`qt_espetofrango`, tb_relatoriocupom1.`qt_espetolinguica`,"
    pRptSql = pRptSql & " tb_relatoriocupom1.`qt_lanche`, tb_relatoriocupom1.`qt_papinha`, tb_relatoriocupom1.`qt_pipoca`,"
    pRptSql = pRptSql & " tb_relatoriocupom1.`qt_sorvete`, tb_relatoriocupom1.`qt_total`"
    pRptSql = pRptSql & " FROM `eletropaulo11`.`tb_relatoriocupom` tb_relatoriocupom1"
    pRptSql = pRptSql & " WHERE tb_relatoriocupom1.id_usuario = " & objBanco.gfsSaveInt(pCd_User)

    pRptFile = pPath & "\Relatórios\Rel_CupomEmitido.rpt"

    pRptTitulo = "Quantidade de Acessos."

    If gfbImprimir(pRptSql, pRptFile, pRptTitulo, rptRelatorios, 0) = False Then

        pMsg = "": pMsg = "Erro ao Gerar Relatório."
        MsgBox pMsg, vbCritical, "Atenção."
        Exit Sub

    End If

    Exit Sub

err_psImprimir:
    Call objSystem.gsExibeErros(Err, "psImprimir()", CStr(Me.Name))

End Sub

Private Function pflTotalRegistro(pSql1 As String) As Long

    On Error GoTo err_pflTotalRegistro:

    Dim pContador As Integer

    pContador = 0

    Set cmd3 = New ADODB.Command
    With cmd3

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText
        .CommandText = pSql1

        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.CursorType = adOpenForwardOnly

        Set rs3 = .Execute
        With rs3

            While Not .EOF

                pContador = pContador + 1
                .MoveNext

            Wend
            .Close

        End With
        Set rs3 = Nothing

    End With
    Set cmd3 = Nothing

    pflTotalRegistro = pContador

    Exit Function

err_pflTotalRegistro:
    Call objSystem.gsExibeErros(Err, "pflTotalRegistro()", CStr(Me.Name))

End Function
