VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Begin VB.Form frmRel_Etiq_Envelope 
   Caption         =   "Emissão de Etiquetas"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   Icon            =   "frmRel_Etiq_Envelope.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   1
         Left            =   6000
         Picture         =   "frmRel_Etiq_Envelope.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdAcao 
         Height          =   855
         Index           =   0
         Left            =   4920
         Picture         =   "frmRel_Etiq_Envelope.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2400
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pbrBarra 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtQuantidade 
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "9999"
         Top             =   810
         Width           =   615
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   4815
      End
      Begin BarcodLib.Barcod Barcod1 
         Height          =   855
         Left            =   240
         TabIndex        =   9
         Top             =   2400
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
      Begin VB.Label lblProgresso 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
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
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   6735
      End
      Begin VB.Label lbltQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Etiquetas:"
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
         Width           =   1695
      End
      Begin VB.Label lblTipoEtiqueta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Etiqueta:"
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
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRel_Etiq_Envelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbTipoEtiqueta_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbTipoEtiqueta, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0 ' Imprimir
            Call psImprimirEtiqueta

        Case 1 ' Fechar
            Unload Me

    End Select

End Sub

Private Sub Form_Load()

    DoEvents: Call psCarregarComboTipoEtiqueta
    Call psLimparCampos

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtQuantidade, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: cmdAcao(0).SetFocus
    End If

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

Private Sub psAtualizarControleImpressao(pCd_Pessoa As Long)

    On Error GoTo err_psAtualizarControleImpressao:

    Set cmd2 = New ADODB.Command
    With cmd2

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "UPDATE tb_pessoa"
        pSql = pSql & " SET fl_etiquetaenvelope = " & objBanco.gfsSaveChar("S")
        pSql = pSql & " WHERE cd_pessoa = " & objBanco.gfsSaveInt(pCd_Pessoa)

        .CommandText = pSql
        .Execute

    End With
    Set cmd2 = Nothing

    Exit Sub

err_psAtualizarControleImpressao:
    Call objSystem.gsExibeErros(Err, "psAtualizarControleImpressao()", CStr(Me.Name))

End Sub

Private Sub psImprimirEtiqueta()

    On Error GoTo err_psImprimirEtiqueta:

    Dim pContador       As Integer
    Dim pPrinter        As Printer
    Dim pCodigoBarra    As String
    Dim pImprime        As String

    If cmbTipoEtiqueta.ListIndex <> -1 Then

        Screen.MousePointer = vbHourglass

        For Each pPrinter In Printers
            If pPrinter.DeviceName = pImpressora Then
                Set Printer = pPrinter
                Exit For
            End If
        Next
        
        pContador = 0: pbrBarra.Max = txtQuantidade.Text
        
        Set cmd = New ADODB.Command
        With cmd

            .ActiveConnection = cn
            .CommandTimeout = 360000000
            .CommandType = adCmdText

            pSql = "": pSql = "SELECT cd_pessoa, nr_via, nr_matricula, nm_pessoa, cd_filial, cd_centrocusto, id_categoria"
            pSql = pSql & " FROM tb_pessoa"
            pSql = pSql & " WHERE fl_ativo = " & objBanco.gfsSaveChar("S")
            pSql = pSql & " AND id_categoria = " & objBanco.gfsSaveInt(1)
            pSql = pSql & " AND fl_etiquetaenvelope = " & objBanco.gfsSaveChar("N")
            pSql = pSql & " ORDER BY cd_filial, cd_centrocusto, nr_matricula, nm_pessoa"
            If Len(Trim(txtQuantidade.Text)) > 0 Then pSql = pSql & " LIMIT " & objBanco.gfsSaveInt(txtQuantidade.Text)

            .CommandText = pSql

            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly

            Set rs = .Execute
            With rs

                While Not .EOF

                    Set cmd2 = New ADODB.Command

                    cmd2.ActiveConnection = cn
                    cmd2.CommandTimeout = 360000000
                    cmd2.CommandType = adCmdText

                    pSql = "": pSql = "CALL sp_con_etiqueta"
                    pSql = pSql & " ( NULL"
                    pSql = pSql & ", " & objBanco.gfsSaveInt(cmbTipoEtiqueta.ItemData(cmbTipoEtiqueta.ListIndex))
                    pSql = pSql & ", NULL );"

                    cmd2.CommandText = pSql

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.CursorType = adOpenForwardOnly

                    Set rs2 = cmd2.Execute

                    '=== Coloca a escala em MM
                    Printer.ScaleMode = 7

                    While Not rs2.EOF

                        If LCase(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))) <> LCase("CODIGOBARRA") Then

                            Select Case LCase(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA")))
                                Case LCase("NM_PESSOA")
                                    pImprime = objBanco.gfsReadChar(.Fields(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))))

                                Case LCase("NR_MATRICULA")
                                    pImprime = "Matrícula: " & objBanco.gfsReadChar(.Fields(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))))

                                Case LCase("CD_CENTROCUSTO")
                                    pImprime = "Centro de Custo: " & objBanco.gfsReadChar(.Fields(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))))

                            End Select

                            Printer.FontName = objBanco.gfsReadChar(rs2.Fields("DE_FONTE"))
                            Printer.FontBold = IIf(objBanco.gfsReadChar(rs2.Fields("FL_NEGRITO")) = "S", True, False)
                            Printer.FontItalic = IIf(objBanco.gfsReadChar(rs2.Fields("FL_ITALICO")) = "S", True, False)
                            Printer.FontSize = objBanco.gfsReadInt(rs2.Fields("NR_TAMANHO"))

                            '=== Verifica se cabe, senão diminui a letra
                            Do While Printer.TextWidth(pImprime) > 8

                                Printer.FontSize = Printer.FontSize - 1

                            Loop

                            '=== Verifica o alinhamento
                            If objBanco.gfsReadChar(rs2.Fields("FL_ALINHAMENTO")) = "E" Then ' esquerda

                                Printer.CurrentX = objBanco.gfsReadInt(rs2.Fields("NR_POSICAOX")) * 0.1

                            Else 'centralizado

                                Printer.CurrentX = ((8 - Printer.TextWidth(pImprime)) / 2)

                            End If

                            Printer.CurrentY = objBanco.gfsReadInt(rs2.Fields("NR_POSICAOY")) * 0.1
                            Printer.Print pImprime

                        Else

                            Barcod1.PrinterScaleMode = 1
                            Barcod1.Caption = pCodigoBarra
                            Barcod1.PrinterLeft = objBanco.gfsReadInt(rs2.Fields("NR_POSICAOX"))
                            Barcod1.PrinterTop = objBanco.gfsReadInt(rs2.Fields("NR_POSICAOY"))
                            Barcod1.PrinterHeight = objBanco.gfsReadInt(rs2.Fields("NR_ALTURA"))
                            Barcod1.PrinterWidth = objBanco.gfsReadInt(rs2.Fields("NR_LARGURA"))
                            Barcod1.Style = objBanco.gfsReadInt(rs2.Fields("ID_CODIGOBARRA"))
                            Barcod1.PrinterHDC = Printer.hDC

                        End If

                        rs2.MoveNext
                        DoEvents

                    Wend
                    rs2.Close

                    Set rs2 = Nothing
                    Set cmd2 = Nothing

                    Printer.EndDoc

                    Call psAtualizarControleImpressao(objBanco.gfsReadInt(.Fields("CD_PESSOA")))
                    
                    DoEvents: pContador = pContador + 1
                    DoEvents: lblProgresso.Caption = Format((pContador / txtQuantidade.Text) * 100, "###") & " %"
                    DoEvents: pbrBarra.Value = pContador
                    
                    .MoveNext

                Wend
                .Close

            End With
            Set rs = Nothing

        End With
        Set cmd = Nothing

        Call psLimparCampos
        
        Screen.MousePointer = vbDefault

    End If

    Exit Sub

err_psImprimirEtiqueta:
    Screen.MousePointer = vbDefault
    Call objSystem.gsExibeErros(Err, "psImprimirEtiqueta()", CStr(Me.Name))

End Sub

Private Sub psLimparCampos()

    With Me

        .cmbTipoEtiqueta.ListIndex = -1

        Call objSystem.gsLimparText(.Name)

        .lblProgresso.Caption = "0%"

        With .pbrBarra
            .Min = 0: .Value = 0
        End With

    End With

End Sub
