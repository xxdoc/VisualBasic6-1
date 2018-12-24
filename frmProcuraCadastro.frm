VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcuraCadastro 
   Caption         =   "Pesquisa de Cadastro"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11400
   Icon            =   "frmProcuraCadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ltwResultado 
      Height          =   6375
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11245
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "¤"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Matrícula"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nome"
         Object.Width           =   12876
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Centro de Custo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Cód. Pessoa"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame fraOpcao 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   11415
      Begin VB.TextBox txtPesquisa 
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
         Left            =   120
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   510
         Width           =   11175
      End
      Begin VB.Label lblOrdem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "##"
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
         Left            =   2115
         TabIndex        =   4
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblPor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1530
         TabIndex        =   3
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblPesquisar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesquisar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1140
      End
   End
   Begin MSComctlLib.Toolbar tbrBarra 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1429
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Pesquisar"
            Key             =   "Pesquisar"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "A"
                  Text            =   "Matrícula"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "B"
                  Text            =   "Nome"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "C"
                  Text            =   "CPF"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "D"
                  Text            =   "Convite"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Localizar"
            Key             =   "Localizar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " Fechar"
            Key             =   "Fechar"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcuraCadastro.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcuraCadastro.frx":0A24
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcuraCadastro.frx":12FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcuraCadastro.frx":1BD8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbBarra 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   8190
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12568
            Text            =   "Consulta: <CTRL+M> p/ Matrícula, <CTRL+N> p/ Nome ou <CTRL+P> p/ CPF"
            TextSave        =   "Consulta: <CTRL+M> p/ Matrícula, <CTRL+N> p/ Nome ou <CTRL+P> p/ CPF"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Total de Registro(s):"
            TextSave        =   "Total de Registro(s):"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "Arquivo"
      Begin VB.Menu mnuPesquisar 
         Caption         =   "Pesquisar"
         Begin VB.Menu mnuMatricula 
            Caption         =   "Matrícula"
            Checked         =   -1  'True
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuNome 
            Caption         =   "Nome"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuCPF 
            Caption         =   "CPF"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuConvite 
            Caption         =   "Convite"
            Shortcut        =   ^C
         End
      End
      Begin VB.Menu mnuLocalizar 
         Caption         =   "Localizar"
         Shortcut        =   {F3}
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
Attribute VB_Name = "frmProcuraCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pClique     As Boolean
Dim pContador   As Long

Private Sub Form_Activate()

    txtPesquisa.SetFocus

End Sub

Private Sub Form_Load()

    pClique = False
    objSystem.gsLimparText (Me.Name)
    Call mnuMatricula_Click

End Sub

Private Sub ltwResultado_DblClick()

    On Error GoTo err_ltwResultado:

    If ltwResultado.ListItems.Count > 0 Then

        Call gsCarregaDadosPessoa(ltwResultado.SelectedItem.ListSubItems(4).Text)
        Unload Me

    End If

    Exit Sub

err_ltwResultado:
    Call objSystem.gsExibeErros(Err, "ltwResultado_DblClick()", CStr(Me.Name))

End Sub

Private Sub ltwResultado_KeyPress(KeyAscii As Integer)

    On Error GoTo err_ltwResultado:

    If KeyAscii = 13 Then

        If ltwResultado.ListItems.Count > 0 Then

            Call gsCarregaDadosPessoa(ltwResultado.SelectedItem.ListSubItems(4).Text)
            Unload Me

        End If

    End If

    Exit Sub

err_ltwResultado:
    Call objSystem.gsExibeErros(Err, "ltwResultado_KeyPress()", CStr(Me.Name))

End Sub

Private Sub mnuConvite_Click()

    lblOrdem.Caption = "Convite"
    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(1).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(2).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(3).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(4).Enabled = True
    mnuCPF.Checked = False: mnuMatricula.Checked = False: mnuNome.Checked = False: mnuConvite.Checked = True

End Sub

Private Sub mnuFechar_Click()

    Unload Me

End Sub

Private Sub mnuLocalizar_Click()

    Call Listagem

End Sub

Private Sub mnuMatricula_Click()

    lblOrdem.Caption = "Matrícula"
    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(1).Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(2).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(3).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(4).Enabled = False
    mnuMatricula.Checked = True: mnuNome.Checked = False: mnuCPF.Checked = False: mnuConvite.Checked = False

End Sub

Private Sub mnuNome_Click()

    lblOrdem.Caption = "Nome"
    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(1).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(2).Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(3).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(4).Enabled = False
    mnuNome.Checked = True: mnuMatricula.Checked = False: mnuCPF.Checked = False: mnuConvite.Checked = False

End Sub

Private Sub mnuCPF_Click()

    lblOrdem.Caption = "CPF"
    If Not tbrBarra.Enabled Then tbrBarra.Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(1).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(2).Enabled = False
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(3).Enabled = True
    tbrBarra.Buttons.Item(1).ButtonMenus.Item(4).Enabled = False
    mnuCPF.Checked = True: mnuMatricula.Checked = False: mnuNome.Checked = False: mnuConvite.Checked = False

End Sub

Private Sub tbrBarra_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Mid(Button.Key, 1, 1)

        Case "L" ' Localizar
            mnuLocalizar_Click

        Case "F" ' Fechar
            mnuFechar_Click

    End Select

End Sub

Private Sub tbrBarra_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    Select Case ButtonMenu.Index

        Case "1"    ' Matrícula
            mnuMatricula_Click

        Case "2"    ' Nome
            mnuNome_Click

        Case "3"    ' CPF
            mnuCPF_Click

        Case "4"    ' Convite
            mnuConvite_Click

    End Select

End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)

    Select Case True

        Case mnuMatricula.Checked
            Call objSystem.gsKeyAscii(txtPesquisa, KeyAscii, ciInt)

        Case mnuNome.Checked
            Call objSystem.gsKeyAscii(txtPesquisa, KeyAscii, ciUpper)

        Case mnuCPF.Checked
            Call objSystem.gsKeyAscii(txtPesquisa, KeyAscii, ciCGCCPF)

        Case mnuConvite.Checked
            Call objSystem.gsKeyAscii(txtPesquisa, KeyAscii, ciInt)

    End Select

    If KeyAscii = 13 Then

        KeyAscii = 0: Call Listagem
        If stbBarra.Panels(3).Text = 0 Then txtPesquisa.Text = "": txtPesquisa.SetFocus

    End If

End Sub

Private Function Listagem()

    On Error GoTo err_Listagem:

    If txtPesquisa.Text = "" Then

        MsgBox "Informe ao menos um caracter para pesquisa!", vbInformation + vbOKOnly, "Atenção:"
        fraOpcao.Enabled = True: Exit Function

    End If

    tbrBarra.Enabled = False: fraOpcao.Enabled = False
    Me.Refresh

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000
        .CommandType = adCmdText
        .CommandText = Retorna_Consulta

        Set rs = .Execute
        With rs

            pContador = 1
            ltwResultado.ListItems.Clear

            Do While Not .EOF
                Call Carregar_Dados
                pContador = pContador + 1
                rs.MoveNext
            Loop

            stbBarra.Panels(3).Text = ltwResultado.ListItems.Count

            If stbBarra.Panels(3).Text = 0 Then MsgBox "Não foram Localizadas Registros para a Condição", vbInformation, Me.Caption

            tbrBarra.Enabled = True: fraOpcao.Enabled = True: Me.Refresh

            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Function

err_Listagem:
    Call objSystem.gsExibeErros(Err, "Listagem()", CStr(Me.Caption))

End Function

Private Function Retorna_Consulta() As String

    On Error GoTo err_Retorna_Consulta:

    Dim pText_Aux As String

    Select Case True

        ' Consulta por Matricula
        Case mnuMatricula.Checked

            pSql = "": pSql = "call sp_con_pessoa ("
            pSql = pSql & " NULL"
            pSql = pSql & ", " & objBanco.gfsSaveInt(Trim(txtPesquisa.Text))
            pSql = pSql & ", NULL, NULL, NULL );"

        ' Consulta por Nome
        Case mnuNome.Checked

            pText_Aux = "%" & Trim(txtPesquisa.Text) & "%"

            pSql = "": pSql = "call sp_con_pessoa ("
            pSql = pSql & " NULL, NULL, "
            pSql = pSql & objBanco.gfsSaveChar(Trim(pText_Aux))
            pSql = pSql & ", NULL, NULL );"

        ' Consulta por CPF
        Case mnuCPF.Checked

            pText_Aux = Replace(Replace(Replace(Trim(txtPesquisa.Text), ".", ""), "-", ""), "/", "")

            pSql = "": pSql = "call sp_con_pessoa ("
            pSql = pSql & " NULL, NULL, NULL, "
            pSql = pSql & objBanco.gfsSaveInt(Trim(pText_Aux))
            pSql = pSql & ", NULL );"

        ' Consulta por Convite
        Case mnuConvite.Checked

            pSql = "": pSql = "call sp_con_pessoa ("
            pSql = pSql & objBanco.gfsSaveInt(Trim(txtPesquisa.Text))
            pSql = pSql & ", NULL, NULL, NULL, NULL );"

    End Select

    Retorna_Consulta = pSql

    Exit Function

err_Retorna_Consulta:
    Call objSystem.gsExibeErros(Err, "Retorna_Consulta()", CStr(Me.Caption))

End Function

Private Function Carregar_Dados()

    On Error GoTo err_Carregar_Dados:

    With ltwResultado

        .ListItems.Add pContador
        .ListItems(pContador).ListSubItems.Add 1, , IIf(IsNull(rs.Fields("NR_MATRICULA")), "", rs.Fields("NR_MATRICULA"))
        .ListItems(pContador).ListSubItems.Add 2, , IIf(IsNull(rs.Fields("NM_PESSOA")), "", rs.Fields("NM_PESSOA"))
        .ListItems(pContador).ListSubItems.Add 3, , IIf(IsNull(rs.Fields("CD_CENTROCUSTO")), "", Trim(rs.Fields("CD_CENTROCUSTO")))
        .ListItems(pContador).ListSubItems.Add 4, , IIf(IsNull(rs.Fields("CD_PESSOA")), "", Trim(rs.Fields("CD_PESSOA")))

    End With

    Exit Function

err_Carregar_Dados:
    Call objSystem.gsExibeErros(Err, "Carregar_Dados()", CStr(Me.Caption))

End Function

