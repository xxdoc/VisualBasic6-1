VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Begin VB.Form frmConf_Etiqueta 
   Caption         =   "Configuração de Etiquetas"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin BarcodLib.Barcod Barcod1 
         Height          =   855
         Left            =   7560
         TabIndex        =   30
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
         _Version        =   65543
         _ExtentX        =   1720
         _ExtentY        =   1508
         _StockProps     =   75
         BackColor       =   16777215
         BarWidth        =   0
         Direction       =   0
         Style           =   18
         UPCNotches      =   3
         Alignment       =   0
         Extension       =   ""
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   5
         Left            =   6240
         Picture         =   "frmConf_Etiqueta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   4
         Left            =   5280
         Picture         =   "frmConf_Etiqueta.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   3
         Left            =   4320
         Picture         =   "frmConf_Etiqueta.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   2
         Left            =   3360
         Picture         =   "frmConf_Etiqueta.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   2400
         Picture         =   "frmConf_Etiqueta.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   6720
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   1440
         Picture         =   "frmConf_Etiqueta.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6720
         Width           =   975
      End
      Begin VB.ComboBox cmbTipoEtiqueta 
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
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   6255
      End
      Begin VB.Frame fraDados 
         Caption         =   "Relação de Etiquetas:"
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
         TabIndex        =   22
         Top             =   3960
         Width           =   8295
         Begin MSFlexGridLib.MSFlexGrid mfgEtiquetas 
            Height          =   2055
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   8055
            _ExtentX        =   14208
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
      Begin VB.CheckBox chkAlinhamento 
         Caption         =   "Centralizado"
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
         Left            =   6300
         TabIndex        =   21
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox chkItalico 
         Caption         =   "Itálico"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3570
         TabIndex        =   20
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkNegrito 
         Caption         =   "Negrito"
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
         TabIndex        =   19
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox txtPosicaoY 
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
         Left            =   4995
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "9999"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtPosicaoX 
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
         Left            =   2265
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "9999"
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbFonte 
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
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1872
         Width           =   6255
      End
      Begin VB.TextBox txtLargura 
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
         Left            =   7905
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "9999"
         Top             =   2376
         Width           =   615
      End
      Begin VB.TextBox txtAltura 
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
         Left            =   4995
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "9999"
         Top             =   2376
         Width           =   615
      End
      Begin VB.TextBox txtTamanho 
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
         Left            =   2265
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "9999"
         Top             =   2376
         Width           =   615
      End
      Begin VB.ComboBox cmbCodigoBarra 
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
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1368
         Width           =   6255
      End
      Begin VB.ComboBox cmbCampoEtiqueta 
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
         Left            =   2265
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   864
         Width           =   6255
      End
      Begin VB.Label lblTipoEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Etiqueta:"
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
         Width           =   1830
      End
      Begin VB.Label lblPosicaoY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posição Y:"
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
         Left            =   3570
         TabIndex        =   17
         Top             =   2940
         Width           =   1125
      End
      Begin VB.Label lblPosicaoX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posição X:"
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
         TabIndex        =   15
         Top             =   2940
         Width           =   1125
      End
      Begin VB.Label lblLargura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Largura:"
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
         Left            =   6300
         TabIndex        =   13
         Top             =   2436
         Width           =   915
      End
      Begin VB.Label lblAltura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Altura:"
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
         Left            =   3570
         TabIndex        =   11
         Top             =   2436
         Width           =   735
      End
      Begin VB.Label lblFonte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fonte:"
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
         Top             =   1932
         Width           =   705
      End
      Begin VB.Label lblTamanho 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tamanho:"
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
         TabIndex        =   9
         Top             =   2436
         Width           =   1080
      End
      Begin VB.Label lblCodigoBarra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Barras:"
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
         Top             =   1428
         Width           =   1935
      End
      Begin VB.Label lblCampo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
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
         Top             =   924
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmConf_Etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pNr_Linha As Integer

Private Sub cmbCampoEtiqueta_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbCampoEtiqueta, KeyAscii, ciUpper)

    If KeyAscii = 13 Then

        KeyAscii = 0
        If cmbCampoEtiqueta.ListIndex <> -1 Then

            If LCase(cmbCampoEtiqueta.Text) = LCase("CODIGOBARRA") Then

                cmbCodigoBarra.Enabled = True
                If cmbFonte.Enabled Then cmbFonte.Enabled = False
                If txtTamanho.Enabled Then txtTamanho.Enabled = False
                If Not txtAltura.Enabled Then txtAltura.Enabled = True
                If Not txtLargura.Enabled Then txtLargura.Enabled = True
                If cmbCodigoBarra.Enabled Then cmbCodigoBarra.SetFocus

            Else

                If cmbCodigoBarra.Enabled Then cmbCodigoBarra.Enabled = False
                If Not cmbFonte.Enabled Then cmbFonte.Enabled = True
                If Not txtTamanho.Enabled Then txtTamanho.Enabled = True
                If txtAltura.Enabled Then txtAltura.Enabled = False
                If txtLargura.Enabled Then txtLargura.Enabled = False
                If cmbFonte.Enabled Then cmbFonte.SetFocus

            End If

        End If

    End If

End Sub

Private Sub cmbCodigoBarra_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbCodigoBarra, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0
    End If

End Sub

Private Sub cmbFonte_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbFonte, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmbTipoEtiqueta_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbTipoEtiqueta, KeyAscii, ciUpper)

    If KeyAscii = 13 Then

        KeyAscii = 0
        If cmbTipoEtiqueta.ListIndex <> -1 Then Call psCarregarEtiquetas(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))
        SendKeys "{TAB}"

    End If

End Sub

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0
            Call psNovo

        Case 1
            Call psSalvar

        Case 2
            Call psExcluir

        Case 3
            Call psImprimir

        Case 4
            Call psLimparCampos

        Case 5
            Unload Me

    End Select

End Sub

Private Sub Form_Load()

    DoEvents: Call psCarregarComboTipoEtiqueta
    DoEvents: Call psCarregarComboCampoEtiqueta
    DoEvents: Call psCarregarComboCodigoBarra
    DoEvents: Call psCarregarComboFonte
    Call psLimparCampos

End Sub

Private Sub mfgEtiquetas_Click()

    With mfgEtiquetas

        .Col = 0
        If Len(Trim(.Text)) > 0 Then Call psMostrarEtiqueta(CInt(Trim(.Text)))

    End With

End Sub

Private Sub txtAltura_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtAltura, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtLargura_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtLargura, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPosicaoX_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtPosicaoX, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtPosicaoY_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtPosicaoY, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtTamanho_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtTamanho, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub psCarregarComboCampoEtiqueta()

    On Error GoTo err_psCarregarComboCampoEtiqueta:

    cmbCampoEtiqueta.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_campoetiqueta ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbCampoEtiqueta.AddItem objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))
                cmbCampoEtiqueta.ItemData(cmbCampoEtiqueta.NewIndex) = objBanco.gfsReadInt(.Fields("ID_CAMPOETIQUETA"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboCampoEtiqueta:
    Call objSystem.gsExibeErros(Err, "psCarregarComboCampoEtiqueta()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboCodigoBarra()

    On Error GoTo err_psCarregarComboCodigoBarra:

    cmbCodigoBarra.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_codigobarra ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbCodigoBarra.AddItem objBanco.gfsReadChar(.Fields("DE_CODIGOBARRA"))
                cmbCodigoBarra.ItemData(cmbCodigoBarra.NewIndex) = objBanco.gfsReadInt(.Fields("ID_CODIGOBARRA"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboCodigoBarra:
    Call objSystem.gsExibeErros(Err, "psCarregarComboCodigoBarra()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboFonte()

    On Error GoTo err_psCarregarComboFonte:

    Dim pContador As Integer

    For pContador = 0 To Screen.FontCount - 1

        cmbFonte.AddItem Screen.Fonts(pContador)
        cmbFonte.ItemData(cmbFonte.NewIndex) = pContador

    Next

    Exit Sub

err_psCarregarComboFonte:
    Call objSystem.gsExibeErros(Err, "psCarregarComboFonte()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboTipoEtiqueta()

    On Error GoTo err_psCarregarComboTipoEtiqueta:

    cmbTipoEtiqueta.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_tipoetiqueta ( NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbTipoEtiqueta.AddItem objBanco.gfsReadChar(.Fields("DE_TIPOETIQUETA"))
                cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.NewIndex) = objBanco.gfsReadInt(.Fields("ID_TIPOETIQUETA"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboTipoEtiqueta:
    Call objSystem.gsExibeErros(Err, "psCarregarComboTipoEtiqueta()", CStr(Me.Name))

End Sub

Private Sub psCarregarEtiquetas(pTipoEtiqueta As Integer)

    On Error GoTo err_psCarregarEtiquetas:

    Dim pLinha As Integer

    pLinha = 1

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_etiqueta"
        pSql = pSql & " ( NULL"
        pSql = pSql & ", " & objBanco.gfsSaveInt(pTipoEtiqueta)
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                While Not .EOF

                    If pLinha > mfgEtiquetas.Rows - 1 Then mfgEtiquetas.Rows = mfgEtiquetas.Rows + 1

                    mfgEtiquetas.Row = pLinha

                    mfgEtiquetas.Col = 0: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_LINHA"))
                    mfgEtiquetas.Col = 1: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadChar(.Fields("DE_TIPOETIQUETA"))
                    mfgEtiquetas.Col = 2: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))
                    mfgEtiquetas.Col = 3: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadChar(.Fields("DE_CODIGOBARRA"))
                    mfgEtiquetas.Col = 4: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadChar(.Fields("DE_FONTE"))
                    mfgEtiquetas.Col = 5: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_TAMANHO"))
                    mfgEtiquetas.Col = 6: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_ALTURA"))
                    mfgEtiquetas.Col = 7: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_LARGURA"))
                    mfgEtiquetas.Col = 8: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_POSICAOX"))
                    mfgEtiquetas.Col = 9: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadInt(.Fields("NR_POSICAOY"))
                    mfgEtiquetas.Col = 10: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = IIf(objBanco.gfsReadChar(.Fields("FL_NEGRITO")) = "S", "Sim", "Não")
                    mfgEtiquetas.Col = 11: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = IIf(objBanco.gfsReadChar(.Fields("FL_ITALICO")) = "S", "Sim", "Não")
                    mfgEtiquetas.Col = 12: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = objBanco.gfsReadChar(.Fields("FL_ALINHAMENTO"))
                    mfgEtiquetas.Col = 13: mfgEtiquetas.CellAlignment = 1
                    mfgEtiquetas.Text = IIf(objBanco.gfsReadChar(.Fields("FL_ATIVO")) = "S", "Sim", "Não")

                    pLinha = pLinha + 1
                    .MoveNext

                Wend

                mfgEtiquetas.Row = 1: mfgEtiquetas.Col = 0
                mfgEtiquetas.ColSel = mfgEtiquetas.Cols - 1

                With cmdAcao

                    .Item(0).Enabled = True
                    .Item(1).Enabled = True
                    .Item(2).Enabled = True
                    .Item(3).Enabled = True
                    .Item(4).Enabled = True

                End With

            Else

                With cmdAcao

                    .Item(0).Enabled = True
                    .Item(1).Enabled = False
                    .Item(2).Enabled = False
                    .Item(3).Enabled = False
                    .Item(4).Enabled = True

                End With

            End If
            .Close

            Call psHabilitarCampos(True)

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarEtiquetas:
    Call objSystem.gsExibeErros(Err, "psCarregarEtiquetas()", CStr(Me.Name))

End Sub

Private Sub psExcluir()

    On Error GoTo err_psExcluir:

    With Me

        pSql = "": pSql = "CALL sp_del_etiqueta"
        pSql = pSql & " ("
        pSql = pSql & objBanco.gfsSaveInt(pNr_Linha)
        pSql = pSql & ", " & objBanco.gfsSaveInt(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao excluir a categoria!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Etiqueta excluída com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psExcluir:
    Call objSystem.gsExibeErros(Err, "psExcluir()", CStr(Me.Name))

End Sub

Private Sub psFormatarGrid()

    With Me

        With .mfgEtiquetas

            .Clear
            .Cols = 14: .Rows = 20
            .Col = 0: .Row = 0

            .Col = 0: .CellAlignment = 3: .Text = "Linha"
            .Col = 1: .CellAlignment = 3: .Text = "Tipo Etiqueta"
            .Col = 2: .CellAlignment = 3: .Text = "Campo"
            .Col = 3: .CellAlignment = 3: .Text = "Código de Barras"
            .Col = 4: .CellAlignment = 3: .Text = "Fonte"
            .Col = 5: .CellAlignment = 3: .Text = "Tamanho"
            .Col = 6: .CellAlignment = 3: .Text = "Altura"
            .Col = 7: .CellAlignment = 3: .Text = "Largura"
            .Col = 8: .CellAlignment = 3: .Text = "Posição X"
            .Col = 9: .CellAlignment = 3: .Text = "Posição Y"
            .Col = 10: .CellAlignment = 3: .Text = "Negrito"
            .Col = 11: .CellAlignment = 3: .Text = "Itálico"
            .Col = 12: .CellAlignment = 3: .Text = "Alinhamento"
            .Col = 13: .CellAlignment = 3: .Text = "Ativo"

            .ColWidth(0) = 900: .ColWidth(1) = 4200
            .ColWidth(2) = 4200: .ColWidth(3) = 4200
            .ColWidth(4) = 4200: .ColWidth(5) = 1000
            .ColWidth(6) = 900: .ColWidth(7) = 900
            .ColWidth(8) = 1100: .ColWidth(9) = 1100
            .ColWidth(10) = 900: .ColWidth(11) = 900
            .ColWidth(12) = 1300: .ColWidth(13) = 900

            .Row = 1: .Col = 0: .ColSel = .Cols - 1

        End With

    End With

End Sub

Private Sub psHabilitarCampos(pHabilita As Boolean)

    With Me

        .cmbCampoEtiqueta.Enabled = pHabilita
        If .cmbCodigoBarra.Enabled Then .cmbCodigoBarra.Enabled = False
        .cmbFonte.Enabled = pHabilita
        .txtTamanho.Enabled = pHabilita
        .txtAltura.Enabled = pHabilita
        .txtLargura.Enabled = pHabilita
        .txtPosicaoX.Enabled = pHabilita
        .txtPosicaoY.Enabled = pHabilita
        .chkNegrito.Enabled = pHabilita
        .chkItalico.Enabled = pHabilita
        .chkAlinhamento.Enabled = pHabilita

    End With

End Sub

Private Sub psImprimir()

    On Error GoTo err_psImprimir:

    Dim pPrinter        As Printer
    Dim pCodigoBarra    As String
    Dim pImprime        As String

    Screen.MousePointer = vbHourglass

    For Each pPrinter In Printers
        If pPrinter.DeviceName = pImpressora Then
            Set Printer = pPrinter
            Exit For
        End If
    Next

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_etiqueta"
        pSql = pSql & " ( NULL"
        pSql = pSql & ", " & objBanco.gfsSaveInt(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            '=== Coloca a escala em MM
            Printer.ScaleMode = 7

            pCodigoBarra = Format(1, "000000") & "1"

            While Not .EOF

                If LCase(objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))) <> LCase("CODIGOBARRA") Then

                    If LCase(objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))) = LCase("CD_PESSOA") Then

                        pImprime = "Nº do convite: " & pCodigoBarra

                    Else

                        pImprime = objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))

                    End If
                    
                    Printer.FontName = objBanco.gfsReadChar(.Fields("DE_FONTE"))
                    Printer.FontBold = IIf(objBanco.gfsReadChar(.Fields("FL_NEGRITO")) = "S", True, False)
                    Printer.FontItalic = IIf(objBanco.gfsReadChar(.Fields("FL_ITALICO")) = "S", True, False)
                    Printer.FontSize = objBanco.gfsReadInt(.Fields("NR_TAMANHO"))

                    '=== Verifica se cabe, senão diminui a letra
                    Do While Printer.TextWidth(pImprime) > 8

                        Printer.FontSize = Printer.FontSize - 1

                    Loop

                    '=== Verifica o alinhamento
                    If objBanco.gfsReadChar(.Fields("FL_ALINHAMENTO")) = "E" Then ' esquerda

                        Printer.CurrentX = objBanco.gfsReadInt(.Fields("NR_POSICAOX")) * 0.1

                    Else 'centralizado

                        Printer.CurrentX = ((8 - Printer.TextWidth(pImprime)) / 2)

                    End If

                    Printer.CurrentY = objBanco.gfsReadInt(.Fields("NR_POSICAOY")) * 0.1
                    Printer.Print pImprime

                Else

                    Barcod1.PrinterScaleMode = 1
                    Barcod1.Caption = pCodigoBarra
                    Barcod1.PrinterLeft = objBanco.gfsReadInt(.Fields("NR_POSICAOX"))
                    Barcod1.PrinterTop = objBanco.gfsReadInt(.Fields("NR_POSICAOY"))
                    Barcod1.PrinterHeight = objBanco.gfsReadInt(.Fields("NR_ALTURA"))
                    Barcod1.PrinterWidth = objBanco.gfsReadInt(.Fields("NR_LARGURA"))
                    Barcod1.Style = objBanco.gfsReadInt(.Fields("ID_CODIGOBARRA"))
                    Barcod1.PrinterHDC = Printer.hDC

                End If

                .MoveNext
                DoEvents

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Printer.EndDoc

    Screen.MousePointer = vbDefault

    Exit Sub

err_psImprimir:
    Screen.MousePointer = vbDefault
    Call objSystem.gsExibeErros(Err, "psImprimir()", CStr(Me.Name))

End Sub

Private Sub psLimparCampos()

    With Me

        .cmbTipoEtiqueta.ListIndex = -1
        .cmbCampoEtiqueta.ListIndex = -1
        .cmbCodigoBarra.ListIndex = -1
        .cmbFonte.ListIndex = -1

        Call objSystem.gsLimparText(.Name)

        .chkNegrito.Value = False
        .chkItalico.Value = False
        .chkAlinhamento.Value = False

        Call psHabilitarCampos(False)

        With .cmdAcao

            .Item(0).Enabled = False
            .Item(1).Enabled = False
            .Item(2).Enabled = False
            .Item(3).Enabled = False
            .Item(4).Enabled = False

        End With

        Call psFormatarGrid

        pNr_Linha = 0

    End With

End Sub

Private Sub psMostrarEtiqueta(pCodigo As Integer)

    On Error GoTo err_psMostrarEtiqueta:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_etiqueta"
        pSql = pSql & " ( "
        pSql = pSql & objBanco.gfsSaveInt(pCodigo)
        pSql = pSql & ", " & objBanco.gfsSaveInt(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If Not .EOF Then

                pNr_Linha = objBanco.gfsReadInt(.Fields("NR_LINHA"))
                Call objSystem.gsBuscaCombo(cmbCampoEtiqueta, objBanco.gfsReadInt(.Fields("ID_CAMPOETIQUETA")))
                If objBanco.gfsReadInt(.Fields("ID_CODIGOBARRA")) <> 0 Then Call objSystem.gsBuscaCombo(cmbCodigoBarra, objBanco.gfsReadInt(.Fields("ID_CODIGOBARRA")))
                cmbFonte.Text = objBanco.gfsReadChar(.Fields("DE_FONTE"))
                If objBanco.gfsReadInt(.Fields("NR_TAMANHO")) <> 0 Then txtTamanho.Text = objBanco.gfsReadInt(.Fields("NR_TAMANHO"))
                If objBanco.gfsReadInt(.Fields("NR_ALTURA")) <> 0 Then txtAltura.Text = objBanco.gfsReadInt(.Fields("NR_ALTURA"))
                If objBanco.gfsReadInt(.Fields("NR_LARGURA")) <> 0 Then txtLargura.Text = objBanco.gfsReadInt(.Fields("NR_LARGURA"))
                If objBanco.gfsReadInt(.Fields("NR_POSICAOX")) <> 0 Then txtPosicaoX.Text = objBanco.gfsReadInt(.Fields("NR_POSICAOX"))
                If objBanco.gfsReadInt(.Fields("NR_POSICAOY")) <> 0 Then txtPosicaoY.Text = objBanco.gfsReadInt(.Fields("NR_POSICAOY"))
                If objBanco.gfsReadChar(.Fields("FL_NEGRITO")) = "S" Then chkNegrito.Value = 1 Else chkNegrito.Value = 0
                If objBanco.gfsReadChar(.Fields("FL_ITALICO")) = "S" Then chkItalico.Value = 1 Else chkItalico.Value = 0
                If objBanco.gfsReadChar(.Fields("FL_ALINHAMENTO")) = "C" Then chkAlinhamento.Value = 1 Else chkAlinhamento.Value = 0

                Call psHabilitarCampos(True)

                If LCase(objBanco.gfsReadChar(.Fields("DE_CAMPOETIQUETA"))) = LCase("CODIGOBARRA") Then

                    cmbCodigoBarra.Enabled = True
                    If cmbFonte.Enabled Then cmbFonte.Enabled = False
                    If txtTamanho.Enabled Then txtTamanho.Enabled = False
                    If Not txtAltura.Enabled Then txtAltura.Enabled = True
                    If Not txtLargura.Enabled Then txtLargura.Enabled = True

                Else

                    If cmbCodigoBarra.Enabled Then cmbCodigoBarra.Enabled = False
                    If Not cmbFonte.Enabled Then cmbFonte.Enabled = True
                    If Not txtTamanho.Enabled Then txtTamanho.Enabled = True
                    If txtAltura.Enabled Then txtAltura.Enabled = False
                    If txtLargura.Enabled Then txtLargura.Enabled = False

                End If

            End If
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psMostrarEtiqueta:
    Call objSystem.gsExibeErros(Err, "psMostrarEtiqueta()", CStr(Me.Name))

End Sub

Private Sub psNovo()

    On Error GoTo err_psNovo:

    If Not pfbValidarCampos Then Exit Sub

    With Me

        pSql = "": pSql = "CALL sp_ins_etiqueta"

        pSql = pSql & " (" & objBanco.gfsSaveInt(pflObterNumeroLinha)
        pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbTipoEtiqueta.ItemData(.cmbTipoEtiqueta.ListIndex))
        pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCampoEtiqueta.ItemData(.cmbCampoEtiqueta.ListIndex))

        If .cmbCodigoBarra.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCodigoBarra.ItemData(.cmbCodigoBarra.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbFonte.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.cmbFonte.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtTamanho.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtTamanho.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtAltura.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtAltura.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtLargura.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtLargura.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtPosicaoX.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtPosicaoX.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtPosicaoY.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtPosicaoY.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkNegrito.Value = 1, "S", "N"))
        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkItalico.Value = 1, "S", "N"))
        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkAlinhamento.Value = 1, "C", "E"))
        pSql = pSql & ", " & objBanco.gfsSaveChar("S")
        pSql = pSql & " );"

    End With

    cn.BeginTrans

    If objBanco.gfiExecuteSql(pSql) = -1 Then

        cn.RollbackTrans
        pMsg = "": pMsg = "Erro ao incluir a etiqueta!"
        MsgBox pMsg, vbOKOnly + vbCritical, "Atenção:"

    Else

        cn.CommitTrans
        pMsg = "": pMsg = "Etiqueta foi inclusa com sucesso!"
        MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
        Call psLimparCampos

    End If

    Exit Sub

err_psNovo:
    Call objSystem.gsExibeErros(Err, "psNovo()", CStr(Me.Name))

End Sub

Private Sub psSalvar()

    On Error GoTo err_psSalvar:

    If Not pfbValidarCampos Then Exit Sub

    With Me

        pSql = "": pSql = "CALL sp_upd_etiqueta"

        pSql = pSql & " (" & objBanco.gfsSaveInt(pNr_Linha)
        pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbTipoEtiqueta.ItemData(.cmbTipoEtiqueta.ListIndex))
        pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCampoEtiqueta.ItemData(.cmbCampoEtiqueta.ListIndex))

        If .cmbCodigoBarra.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCodigoBarra.ItemData(.cmbCodigoBarra.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbFonte.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.cmbFonte.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtTamanho.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtTamanho.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtAltura.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtAltura.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtLargura.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtLargura.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtPosicaoX.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtPosicaoX.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtPosicaoY.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtPosicaoY.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkNegrito.Value = 1, "S", "N"))
        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkItalico.Value = 1, "S", "N"))
        pSql = pSql & ", " & objBanco.gfsSaveChar(IIf(.chkAlinhamento.Value = 1, "C", "E"))
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
        pMsg = "": pMsg = "Etiqueta salva com sucesso!"
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

        If .cmbTipoEtiqueta.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o tipo de etiqueta!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbTipoEtiqueta.SetFocus: Exit Function

        End If

        If .cmbCampoEtiqueta.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o campo da etiqueta!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbCampoEtiqueta.SetFocus: Exit Function

        End If

        If .cmbCodigoBarra.Enabled Then

            If .cmbCodigoBarra.ListIndex = -1 Then

                pMsg = "": pMsg = "É necessário informar o código de barras da etiqueta!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .cmbCodigoBarra.SetFocus: Exit Function

            End If

        Else

            If .cmbFonte.ListIndex = -1 Then

                pMsg = "": pMsg = "É necessário informar a fonte da etiqueta!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .cmbFonte.SetFocus: Exit Function

            End If

        End If

        If .txtTamanho.Enabled Then

            If Len(Trim(.txtTamanho.Text)) = 0 Then

                pMsg = "": pMsg = "É necessário informar o tamanho!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .txtTamanho.SetFocus: Exit Function

            End If

        End If

        If .txtAltura.Enabled Then

            If Len(Trim(.txtAltura.Text)) = 0 Then

                pMsg = "": pMsg = "É necessário informar a altura!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .txtAltura.SetFocus: Exit Function

            End If

        End If

        If .txtLargura.Enabled Then

            If Len(Trim(.txtLargura.Text)) = 0 Then

                pMsg = "": pMsg = "É necessário informar a largura!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .txtLargura.SetFocus: Exit Function

            End If

        End If

        If .txtPosicaoX.Enabled Then

            If Len(Trim(.txtPosicaoX.Text)) = 0 Then

                pMsg = "": pMsg = "É necessário informar a posição X!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .txtPosicaoX.SetFocus: Exit Function

            End If

        End If

        If .txtPosicaoY.Enabled Then

            If Len(Trim(.txtPosicaoY.Text)) = 0 Then

                pMsg = "": pMsg = "É necessário informar a posição Y!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .txtPosicaoY.SetFocus: Exit Function

            End If

        End If

    End With

    pfbValidarCampos = True

    Exit Function

err_pfbValidarCampos:
    Call objSystem.gsExibeErros(Err, "pfbValidarCampos()", CStr(Me.Name))

End Function

Private Function pflObterNumeroLinha() As Long

    On Error GoTo err_pflObterNumeroLinha:

    Dim pSql1 As String

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql1 = "": pSql1 = "SELECT MAX(nr_linha) nr_linha"
        pSql1 = pSql1 & " FROM tb_etiqueta "
        pSql1 = pSql1 & " WHERE id_tipoetiqueta = " & objBanco.gfsSaveInt(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))

        .CommandText = pSql1

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If .EOF Then pflObterNumeroLinha = 1 Else pflObterNumeroLinha = objBanco.gfsReadInt(.Fields("NR_LINHA")) + 1
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Function

err_pflObterNumeroLinha:
    Call objSystem.gsExibeErros(Err, "pflObterNumeroLinha()", CStr(Me.Name))

End Function

