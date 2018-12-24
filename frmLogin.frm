VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbImpressora 
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2200
      Width           =   4095
   End
   Begin VB.Frame fraPainel 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4095
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   840
         Width           =   1150
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   1155
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
         Left            =   1162
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "9999999999"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtUsuario 
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
         Left            =   1162
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "WWWWWWWWWW"
         Top             =   240
         Width           =   1455
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
         Left            =   270
         TabIndex        =   6
         Top             =   780
         Width           =   750
      End
      Begin VB.Label lblUsuario 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.PictureBox picLabel 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   4080
      Begin VB.Image Image1 
         Height          =   720
         Left            =   3240
         Picture         =   "frmLogin.frx":0CCA
         Top             =   150
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Entre com seu usuário e senha para acesso ao sistema..."
         Height          =   435
         Left            =   480
         TabIndex        =   2
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conexão"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   90
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()

    Call objBanco.gsFecharBancos: End

End Sub

Private Sub cmdOK_Click()

    With Me

        pImpressora = Trim(.cmbImpressora.Text)

        If pfbValidar_User(.txtUsuario.Text, .txtSenha.Text) Then

            With frmMenu

                With .stbBarra

                    .Panels(2).Text = "Usuário: " & pLg_User
                    .Panels(3).Text = Format(objBanco.gfdDataSistema, "dd/mm/yyyy")

                End With

                .Tag = "LOGON": .Show

            End With

        Else

            pMsg = "": pMsg = "Usuário ou Senha, inválidos. Tente novamente!!" & Chr(13) & "Ou Procure Administrador do Sistema!"
            MsgBox pMsg, vbInformation + vbOKOnly, "Atenção:"
            .txtUsuario.SetFocus
            SendKeys "{Home}+{End}"

        End If

    End With

End Sub

Private Sub Form_Activate()

    txtUsuario.SetFocus

End Sub

Private Sub Form_Load()

    Call objSystem.gsLimparText(Me.Name)
    Call CarregarComboImpressora

End Sub

Private Sub txtSenha_GotFocus()

    SendKeys "{Home}+{End}"

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtSenha, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: Call cmdOK_Click
    End If

End Sub

Private Sub txtUsuario_GotFocus()

    SendKeys "{Home}+{End}"

End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtUsuario, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Function pfbValidar_User(pLg_Usuario As String, pPw_Usuario As String) As Boolean

    On Error GoTo err_pfbValidar_User:

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_usuario ( NULL, NULL"
        pSql = pSql & ", " & objBanco.gfsSaveChar(pLg_Usuario)
        pSql = pSql & ", " & objBanco.gfsSaveChar(objSystem.gfsEncryptString(pPw_Usuario, ciEncrypt))
        pSql = pSql & ", NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            If .EOF Then

                pfbValidar_User = False

            Else

                pCd_User = objBanco.gfsReadInt(.Fields("ID_USUARIO"))
                pLg_User = objBanco.gfsReadChar(.Fields("LG_USUARIO"))
                pTp_User = objBanco.gfsReadInt(.Fields("ID_TIPOUSUARIO"))
                pfbValidar_User = True

            End If

            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Function

err_pfbValidar_User:
    Call objSystem.gsExibeErros(Err, "pfbValidar_User()", CStr(Me.Name))

End Function

Private Sub CarregarComboImpressora()

    On Error GoTo err_CarregarComboImpressora:

    Dim pPrinter As Printer
    Dim pPrinterTemp As Printer

    Set pPrinterTemp = Printer

    With Me

        .cmbImpressora.Clear

        For Each pPrinter In Printers

            .cmbImpressora.AddItem pPrinter.DeviceName
            If pPrinter.DeviceName = pImpressora Then Set pPrinterTemp = pPrinter

        Next

        .cmbImpressora.Text = pPrinterTemp.DeviceName

    End With

    Exit Sub

err_CarregarComboImpressora:
   Call objSystem.gsExibeErros(Err, "CarregarComboImpressora()", CStr(Me.Name))

End Sub
