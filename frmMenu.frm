VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H8000000C&
   Caption         =   "..:: Abacos Automação de Eventos ::.."
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10440
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbBarra 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3075
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8273
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.ToolTipText     =   "Usuário Conectado no Momento."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "15:00"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuCredenciamento 
      Caption         =   "&Credenciamento"
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "C&adastros"
      Begin VB.Menu mnuCad_Acesso 
         Caption         =   "Acesso"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCad_Campos 
         Caption         =   "Campos"
         Begin VB.Menu mnuCad_CamposEtiqueta 
            Caption         =   "Etiqueta"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu mnuCad_Categoria 
         Caption         =   "Categoria"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCad_CodigoBarra 
         Caption         =   "Código de Barras"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuCad_EstadoCivil 
         Caption         =   "Estado Civil"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuCad_Evento 
         Caption         =   "Evento"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuCad_Parentesco 
         Caption         =   "Parentesco"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCad_Sexo 
         Caption         =   "Sexo"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuCad_StatusLeitura 
         Caption         =   "Status Leitura"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuCad_TipoEtiqueta 
         Caption         =   "Tipo Etiqueta"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuCad_TipoLog 
         Caption         =   "Tipo Log"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCad_TipoUsuario 
         Caption         =   "Tipo Usuário"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuCad_Usuario 
         Caption         =   "Usuário"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuConfiguracao 
      Caption         =   "C&onfiguração"
      Begin VB.Menu mnuConf_Etiqueta 
         Caption         =   "&Etiqueta"
      End
   End
   Begin VB.Menu mnuRelatorio 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnuRel_Acessos 
         Caption         =   "Acessos"
      End
      Begin VB.Menu mnuRel_CupomEmitido 
         Caption         =   "Cupons Emitidos"
      End
      Begin VB.Menu mnuRel_Etiqueta 
         Caption         =   "Etiqueta"
         Begin VB.Menu mnuRel_Eti_Envelope 
            Caption         =   "Envelope"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu mnuRel_ListaPresenca 
         Caption         =   "Lista de Presença"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With Me

        If .Tag = "LOGOFF" Then

            .Tag = "LOGOFF": Unload Me

        Else

            Unload frmLogin

        End If

        If pTp_User = 3 Then

            mnuCadastro.Enabled = False
            mnuConfiguracao.Enabled = False
            mnuRelatorio.Enabled = False

        Else

            mnuCadastro.Enabled = True
            mnuConfiguracao.Enabled = True
            mnuRelatorio.Enabled = True

        End If

    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Deseja sair do Aplicativo?", vbInformation + vbYesNo + vbDefaultButton2, "Atenção:") = vbYes Then

        Call objBanco.gsFecharBancos: End

    Else

        Cancel = 1

    End If

End Sub

Private Sub mnuCad_Acesso_Click()

    frmCad_Acesso.Show vbModal

End Sub

Private Sub mnuCad_CamposEtiqueta_Click()

    frmCad_CampoEtiqueta.Show vbModal

End Sub

Private Sub mnuCad_Categoria_Click()

    frmCad_Categoria.Show vbModal

End Sub

Private Sub mnuCad_CodigoBarra_Click()

    frmCad_CodigoBarra.Show vbModal

End Sub

Private Sub mnuCad_EstadoCivil_Click()

    frmCad_EstadoCivil.Show vbModal

End Sub

Private Sub mnuCad_Parentesco_Click()

    frmCad_Parentesco.Show vbModal

End Sub

Private Sub mnuCad_Sexo_Click()

    frmCad_Sexo.Show vbModal

End Sub

Private Sub mnuCad_StatusLeitura_Click()

    frmCad_StatusLeitura.Show vbModal

End Sub

Private Sub mnuCad_TipoEtiqueta_Click()

    frmCad_TipoEtiqueta.Show vbModal

End Sub

Private Sub mnuCad_TipoLog_Click()

    frmCad_TipoLog.Show vbModal

End Sub

Private Sub mnuCad_TipoUsuario_Click()

    frmCad_TipoUsuario.Show vbModal

End Sub

Private Sub mnuCad_Usuario_Click()

    frmCad_Usuario.Show vbModal

End Sub

Private Sub mnuConf_Etiqueta_Click()

    frmConf_Etiqueta.Show vbModal

End Sub

Private Sub mnuCredenciamento_Click()

    frmCredenciamento.Show vbModal

End Sub

Private Sub mnuRel_Acessos_Click()

    frmRel_Acessos.Show vbModal

End Sub

Private Sub mnuRel_CupomEmitido_Click()

    frmRel_CupomEmitido.Show vbModal

End Sub

Private Sub mnuRel_Eti_Envelope_Click()

    frmRel_Etiq_Envelope.Show vbModal

End Sub

Private Sub mnuRel_ListaPresenca_Click()

    frmRel_ListaPresenca.Show vbModal

End Sub

Private Sub mnuSair_Click()

    Unload Me

End Sub

