VERSION 5.00
Begin VB.Form frmAutorizacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorização"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmAutorizacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
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
         TabIndex        =   2
         Text            =   "WWWWWWWWWW"
         Top             =   240
         Width           =   1455
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
         TabIndex        =   4
         Text            =   "9999999999"
         Top             =   720
         Width           =   1455
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
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         TabIndex        =   6
         Top             =   720
         Width           =   1150
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
         TabIndex        =   1
         Top             =   300
         Width           =   900
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
         TabIndex        =   3
         Top             =   780
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmAutorizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelar_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    With Me

        If pfbValidar_User(.txtUsuario.Text, .txtSenha.Text) Then pAutorizaCadastro = True Else pAutorizaCadastro = False
        
        Call cmdCancelar_Click

    End With

End Sub

Private Sub Form_Activate()

    txtUsuario.SetFocus

End Sub

Private Sub Form_Load()

    Call objSystem.gsLimparText(Me.Name)

End Sub

Private Sub txtSenha_GotFocus()

    SendKeys "{Home}+{End}"

End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtSenha, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK.SetFocus
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

                If objBanco.gfsReadInt(.Fields("ID_TIPOUSUARIO")) <> 3 Then pfbValidar_User = True Else pfbValidar_User = False

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
