VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Begin VB.Form frmCredenciamento 
   Caption         =   "Credenciamento"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12735
   Icon            =   "frmCredenciamento.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPainel 
      Height          =   9735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12495
      Begin VB.ComboBox cmbTitular 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3960
         Width           =   9855
      End
      Begin VB.TextBox txtObservacao 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2400
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmCredenciamento.frx":0CCA
         Top             =   4440
         Width           =   9855
      End
      Begin VB.TextBox txtVia 
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
         Left            =   5534
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "99"
         Top             =   360
         Width           =   375
      End
      Begin BarcodLib.Barcod Barcod1 
         Height          =   360
         Left            =   9600
         TabIndex        =   53
         Top             =   413
         Visible         =   0   'False
         Width           =   2655
         _Version        =   65543
         _ExtentX        =   4683
         _ExtentY        =   635
         _StockProps     =   75
         BackColor       =   16777215
         BarWidth        =   0
         Direction       =   0
         Style           =   18
         UPCNotches      =   3
         Alignment       =   0
         Extension       =   ""
      End
      Begin VB.TextBox txtMatricula 
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "99999999999"
         Top             =   1388
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mkeDataNascimento 
         Height          =   360
         Left            =   2400
         TabIndex        =   16
         Top             =   2930
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Fechar"
         Height          =   1000
         Index           =   5
         Left            =   10800
         Picture         =   "frmCredenciamento.frx":0CFD
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Imprimir"
         Height          =   1000
         Index           =   4
         Left            =   8688
         Picture         =   "frmCredenciamento.frx":15C7
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Procurar (F2)"
         Height          =   1000
         Index           =   3
         Left            =   6576
         Picture         =   "frmCredenciamento.frx":1E91
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Cancelar"
         Height          =   1000
         Index           =   2
         Left            =   4464
         Picture         =   "frmCredenciamento.frx":275B
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Salvar (F3)"
         Height          =   1000
         Index           =   1
         Left            =   2352
         Picture         =   "frmCredenciamento.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton cmdAcao 
         Caption         =   "&Novo"
         Height          =   1000
         Index           =   0
         Left            =   240
         Picture         =   "frmCredenciamento.frx":38EF
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   8520
         Width           =   1455
      End
      Begin VB.Frame fraFilial 
         Caption         =   "Filial:"
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
         Height          =   2775
         Left            =   240
         TabIndex        =   26
         Top             =   5640
         Width           =   12015
         Begin VB.TextBox txtTelefone 
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
            Left            =   10355
            MaxLength       =   11
            TabIndex        =   46
            Text            =   "99999999999"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtCEP 
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
            Left            =   5779
            MaxLength       =   11
            TabIndex        =   44
            Text            =   "99999999999"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox txtEstado 
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
            Left            =   2555
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "WW"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox txtCidade 
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
            Left            =   7835
            MaxLength       =   20
            TabIndex        =   40
            Text            =   "WWWWWWWWWWWWWWWWWWWW"
            Top             =   1785
            Width           =   3975
         End
         Begin VB.TextBox txtBairro 
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
            Left            =   2555
            MaxLength       =   20
            TabIndex        =   38
            Text            =   "WWWWWWWWWWWWWWWWWWWW"
            Top             =   1785
            Width           =   3975
         End
         Begin VB.TextBox txtComplemento 
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
            Left            =   7835
            MaxLength       =   20
            TabIndex        =   36
            Text            =   "WWWWWWWWWWWWWWWWWWWW"
            Top             =   1290
            Width           =   3975
         End
         Begin VB.TextBox txtNumero 
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
            Left            =   2555
            MaxLength       =   4
            TabIndex        =   34
            Text            =   "9999"
            Top             =   1290
            Width           =   615
         End
         Begin VB.TextBox txtEndereco 
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
            Left            =   2555
            MaxLength       =   50
            TabIndex        =   32
            Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
            Top             =   795
            Width           =   9255
         End
         Begin VB.ComboBox cmbCentroCusto 
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
            Left            =   7835
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   300
            Width           =   2535
         End
         Begin VB.ComboBox cmbFilial 
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
            Left            =   2555
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label lblTelefone 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone:"
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
            Left            =   8291
            TabIndex        =   45
            Top             =   2340
            Width           =   1005
         End
         Begin VB.Label lblCEP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CEP:"
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
            Left            =   4227
            TabIndex        =   43
            Top             =   2340
            Width           =   495
         End
         Begin VB.Label lblEstado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
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
            TabIndex        =   41
            Top             =   2340
            Width           =   825
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
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
            Left            =   6720
            TabIndex        =   39
            Top             =   1845
            Width           =   825
         End
         Begin VB.Label lblBairro 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro:"
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
            TabIndex        =   37
            Top             =   1845
            Width           =   735
         End
         Begin VB.Label lblComplemento 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento:"
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
            Left            =   5940
            TabIndex        =   35
            Top             =   1350
            Width           =   1605
         End
         Begin VB.Label lblNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
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
            TabIndex        =   33
            Top             =   1350
            Width           =   930
         End
         Begin VB.Label lblEndereco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço:"
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
            TabIndex        =   31
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label lblCentroCusto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Custo:"
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
            Left            =   6045
            TabIndex        =   29
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblFilial 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filial:"
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
            TabIndex        =   27
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.ComboBox cmbEstadoCivil 
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
         Left            =   2400
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3444
         Width           =   3135
      End
      Begin VB.ComboBox cmbParentesco 
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
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2930
         Width           =   2655
      End
      Begin VB.ComboBox cmbSexo 
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
         Left            =   5534
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2930
         Width           =   1815
      End
      Begin VB.TextBox txtNomeCracha 
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   2416
         Width           =   9855
      End
      Begin VB.TextBox txtNome 
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
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW"
         Top             =   1902
         Width           =   9855
      End
      Begin VB.TextBox txtCPF 
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
         Left            =   5534
         MaxLength       =   11
         TabIndex        =   10
         Text            =   "99999999999"
         Top             =   1388
         Width           =   1455
      End
      Begin VB.ComboBox cmbCategoria 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   874
         Width           =   9855
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "99999999999"
         Top             =   360
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mkeDataLeitura 
         Height          =   360
         Left            =   2400
         TabIndex        =   54
         Top             =   5160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   128
         Enabled         =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mkeDataImpressaoTicket 
         Height          =   360
         Left            =   9600
         TabIndex        =   58
         Top             =   5160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         ForeColor       =   128
         Enabled         =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/#### ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Impressão do Ticket:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   5940
         TabIndex        =   57
         Top             =   5220
         Width           =   3180
      End
      Begin VB.Label lblDataLeitura 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Leitura:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   5220
         Width           =   1740
      End
      Begin VB.Label lblObservacao 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
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
         TabIndex        =   55
         Top             =   4500
         Width           =   1380
      End
      Begin VB.Label lblTitular 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Titular:"
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
         TabIndex        =   23
         Top             =   4020
         Width           =   765
      End
      Begin VB.Label lblVia 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Via:"
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
         Left            =   4440
         TabIndex        =   3
         Top             =   420
         Width           =   420
      End
      Begin VB.Label lblMatricula 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matrícula:"
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
         Top             =   1448
         Width           =   1080
      End
      Begin VB.Label lblEstadoCivil 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil:"
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
         TabIndex        =   21
         Top             =   3504
         Width           =   1350
      End
      Begin VB.Label lblParentesco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parentesco:"
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
         Left            =   7815
         TabIndex        =   19
         Top             =   2990
         Width           =   1305
      End
      Begin VB.Label lblSexo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
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
         Left            =   4440
         TabIndex        =   17
         Top             =   2990
         Width           =   615
      End
      Begin VB.Label lblDtNascimento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Nascimento:"
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
         Top             =   2990
         Width           =   1935
      End
      Begin VB.Label lblNomeCracha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Crachá:"
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
         TabIndex        =   13
         Top             =   2476
         Width           =   1530
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
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
         TabIndex        =   11
         Top             =   1962
         Width           =   705
      End
      Begin VB.Label lblCPF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CPF:"
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
         Left            =   4440
         TabIndex        =   9
         Top             =   1448
         Width           =   495
      End
      Begin VB.Label lblCategoria 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
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
         Top             =   934
         Width           =   1140
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
   End
   Begin VB.Menu mnuAcao 
      Caption         =   "Ações"
      Begin VB.Menu mnuAcao_Procurar 
         Caption         =   "Procurar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAcao_Salvar 
         Caption         =   "Salvar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuAcao_Imprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "frmCredenciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pContador As Long

Private Sub cmbCategoria_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbCategoria, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmbCategoria_LostFocus()

    If cmbCategoria.ListIndex <> -1 Then

        If cmbCategoria.ItemData(cmbCategoria.ListIndex) = 2 Then cmbTitular.Enabled = True Else cmbTitular.Enabled = False

    End If

End Sub

Private Sub cmbCentroCusto_Click()

    On Error GoTo err_cmbCentroCusto:

    If cmbCentroCusto.ListIndex <> -1 Then

        Set cmd = New ADODB.Command
        With cmd

            .ActiveConnection = cn
            .CommandTimeout = 360000000
            .CommandType = adCmdText

            pSql = "": pSql = "CALL sp_con_filial"
            pSql = pSql & " (" & objBanco.gfsSaveInt(cmbFilial.ItemData(cmbFilial.ListIndex))
            pSql = pSql & ", " & objBanco.gfsSaveInt(cmbCentroCusto.ItemData(cmbCentroCusto.ListIndex))
            pSql = pSql & " );"

            .CommandText = pSql

            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly

            Set rs = .Execute
            With rs

                If Not .EOF Then

                    txtEndereco.Text = objBanco.gfsReadChar(.Fields("DE_ENDERECO"))
                    txtNumero.Text = objBanco.gfsReadInt(.Fields("NR_ENDERECO"))
                    txtComplemento.Text = objBanco.gfsReadChar(.Fields("DE_COMPLEMENTO"))
                    txtBairro.Text = objBanco.gfsReadChar(.Fields("DE_BAIRRO"))
                    txtCidade.Text = objBanco.gfsReadChar(.Fields("DE_CIDADE"))
                    txtEstado.Text = objBanco.gfsReadChar(.Fields("DE_ESTADO"))
                    txtCEP.Text = objBanco.gfsReadChar(.Fields("NR_CEP"))
                    txtTelefone.Text = objBanco.gfsReadInt(.Fields("NR_TELEFONE"))

                End If
                .Close

            End With
            Set rs = Nothing

        End With
        Set cmd = Nothing

    End If

    Exit Sub

err_cmbCentroCusto:
    Call objSystem.gsExibeErros(Err, "cmbCentroCusto_Click()", CStr(Me.Name))

End Sub

Private Sub cmbFilial_Click()

    On Error GoTo err_cmbFilial:

    If cmbFilial.ListIndex <> -1 Then

        cmbCentroCusto.Clear

        Set cmd = New ADODB.Command
        With cmd

            .ActiveConnection = cn
            .CommandTimeout = 360000000
            .CommandType = adCmdText

            pSql = "": pSql = "CALL sp_con_filial"
            pSql = pSql & " ("
            pSql = pSql & objBanco.gfsSaveInt(cmbFilial.ItemData(cmbFilial.ListIndex))
            pSql = pSql & ", NULL );"

            .CommandText = pSql

            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly

            Set rs = .Execute
            With rs

                While Not .EOF

                    cmbCentroCusto.AddItem objBanco.gfsReadChar(.Fields("CD_CENTROCUSTO"))
                    cmbCentroCusto.ItemData(cmbCentroCusto.NewIndex) = objBanco.gfsReadInt(.Fields("CD_CENTROCUSTO"))

                    .MoveNext

                Wend
                .Close

            End With
            Set rs = Nothing

        End With
        Set cmd = Nothing

    End If

    Exit Sub

err_cmbFilial:
    Call objSystem.gsExibeErros(Err, "cmbFilial_Click()", CStr(Me.Name))

End Sub

Private Sub cmbEstadoCivil_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbEstadoCivil, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmbParentesco_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbParentesco, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmbSexo_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(cmbSexo, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub cmdAcao_Click(Index As Integer)

    Select Case Index
        Case 0  ' Novo
            Call psNovo

        Case 1  ' Salvar
            Call psSalvar

        Case 2  ' Cancelar
            Call psLimparCampos

        Case 3  ' Procurar
            frmProcuraCadastro.Show vbModal

        Case 4  ' Imprimir
            Call psImprimirEtiqueta
            Call cmdAcao_Click(2)

        Case 5  ' Sair
            Unload Me

    End Select

End Sub

Private Sub Form_Load()

    DoEvents: Call psCarregarComboCategoria
    DoEvents: Call psCarregarComboSexo
    DoEvents: Call psCarregarComboParentesco
    DoEvents: Call psCarregarComboEstadoCivil
    DoEvents: Call psCarregarComboTitular
    DoEvents: Call psCarregarComboFilial
    DoEvents: Call psLimparCampos

End Sub

Private Sub mkeDataNascimento_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(mkeDataNascimento, KeyAscii, ciDate)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub mnuAcao_Imprimir_Click()

    Call cmdAcao_Click(4)

End Sub

Private Sub mnuAcao_Procurar_Click()

    Call cmdAcao_Click(3)

End Sub

Private Sub mnuAcao_Salvar_Click()

    Call cmdAcao_Click(1)

End Sub

Private Sub txtBairro_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtBairro, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCEP_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtCEP, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCidade_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtCidade, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtCodigo, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtComplemento_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtComplemento, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtCPF_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtCPF, KeyAscii, ciInt)

    If KeyAscii = 13 Then

        KeyAscii = 0

        If CDbl(Trim(txtCPF.Text)) <> 0 Then

            If Not objSystem.gfbValidarCPF(Trim(txtCPF.Text)) Then

                pMsg = "": pMsg = "CPF inválido!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                txtCPF.SetFocus

            Else

                SendKeys "{TAB}"

            End If

                SendKeys "{TAB}"

        End If

    End If

End Sub

Private Sub txtEndereco_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtEndereco, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtEstado, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtMatricula_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtMatricula, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtMatricula_LostFocus()

    If Len(Trim(txtMatricula.Text)) > 0 Then

        If cmbCategoria.ItemData(cmbCategoria.ListIndex) = 2 Then Call objSystem.gsBuscaCombo(cmbTitular, gflBuscarCodigoTitular(CLng(txtMatricula.Text)))

    End If

End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtNome, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtNome_LostFocus()

    txtNomeCracha.Text = pfsNomeCracha(Trim(txtNome.Text))

End Sub

Private Sub txtNomeCracha_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtNomeCracha, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtNumero, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtObservacao_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtObservacao, KeyAscii, ciUpper)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Private Sub txtTelefone_KeyPress(KeyAscii As Integer)

    Call objSystem.gsKeyAscii(txtTelefone, KeyAscii, ciInt)

    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If

End Sub

Public Sub gsHabilitarCampos(pPar As Boolean, pSituacao As Integer)

    With Me

        .cmbCategoria.Enabled = pPar
        .txtMatricula.Enabled = pPar
        .txtCPF.Enabled = pPar
        .txtNome.Enabled = pPar
        .txtNomeCracha.Enabled = pPar
        .mkeDataNascimento.Enabled = pPar
        .cmbSexo.Enabled = pPar
        .cmbParentesco.Enabled = pPar
        .cmbEstadoCivil.Enabled = pPar
        .cmbTitular.Enabled = pPar
        If pSituacao = 2 Then .txtObservacao.Enabled = Not pPar Else .txtObservacao.Enabled = pPar
        .cmbFilial.Enabled = pPar
        .cmbCentroCusto.Enabled = pPar
        .txtEndereco.Enabled = pPar
        .txtNumero.Enabled = pPar
        .txtComplemento.Enabled = pPar
        .txtBairro.Enabled = pPar
        .txtCidade.Enabled = pPar
        .txtEstado.Enabled = pPar
        .txtCEP.Enabled = pPar
        .txtTelefone.Enabled = pPar

        With .cmdAcao

            Select Case pSituacao
                Case 1
                    .Item(0).Enabled = Not pPar
                    mnuAcao_Salvar.Enabled = pPar
                    .Item(1).Enabled = pPar
                    .Item(2).Enabled = pPar
                    mnuAcao_Procurar.Enabled = Not pPar
                    .Item(3).Enabled = Not pPar

                    If Not pLeituraColetor Then

                        mnuAcao_Imprimir.Enabled = pPar
                        .Item(4).Enabled = pPar

                    Else

                        Select Case pTp_User
                            Case 1, 2
                                mnuAcao_Imprimir.Enabled = pPar
                                .Item(4).Enabled = pPar

                            Case Else
                                mnuAcao_Imprimir.Enabled = False
                                .Item(4).Enabled = False

                        End Select

                    End If

                Case 2
                    .Item(0).Enabled = pPar
                    mnuAcao_Salvar.Enabled = Not pPar
                    .Item(1).Enabled = Not pPar
                    .Item(2).Enabled = pPar
                    mnuAcao_Procurar.Enabled = pPar
                    .Item(3).Enabled = pPar

                    If Not pLeituraColetor Then

                        mnuAcao_Imprimir.Enabled = Not pPar
                        .Item(4).Enabled = Not pPar

                    Else

                        Select Case pTp_User
                            Case 1, 2
                                mnuAcao_Imprimir.Enabled = Not pPar
                                .Item(4).Enabled = Not pPar

                            Case Else
                                mnuAcao_Imprimir.Enabled = False
                                .Item(4).Enabled = False

                        End Select

                    End If

                Case 3
                    .Item(0).Enabled = pPar
                    mnuAcao_Salvar.Enabled = pPar
                    .Item(1).Enabled = pPar
                    .Item(2).Enabled = Not pPar
                    mnuAcao_Procurar.Enabled = pPar
                    .Item(3).Enabled = pPar

                    If Not pLeituraColetor Then

                        mnuAcao_Imprimir.Enabled = Not pPar
                        .Item(4).Enabled = Not pPar

                    Else

                        Select Case pTp_User
                            Case 1, 2
                                mnuAcao_Imprimir.Enabled = Not pPar
                                .Item(4).Enabled = Not pPar

                            Case Else
                                mnuAcao_Imprimir.Enabled = False
                                .Item(4).Enabled = False

                        End Select

                    End If

            End Select

        End With

    End With

End Sub

Private Sub psAtualizarDadosPessoa(pCd_Pessoa As Long, pNr_Via As Integer)

    On Error GoTo err_psAtualizarDadosPessoa:

    Set cmd2 = New ADODB.Command
    With cmd2

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_upd_pessoa_impressao ("
        pSql = pSql & objBanco.gfsSaveInt(pCd_Pessoa)
        pSql = pSql & ", " & objBanco.gfsSaveInt(pNr_Via) & " );"

        .CommandText = pSql
        .Execute

    End With
    Set cmd2 = Nothing

    Exit Sub

err_psAtualizarDadosPessoa:
    Call objSystem.gsExibeErros(Err, "psAtualizarDadosPessoa()", CStr(Me.Name), pSql)

End Sub

Private Sub psCarregarComboCategoria()

    On Error GoTo err_psCarregarComboCategoria:

    cmbCategoria.Clear

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

            While Not .EOF

                cmbCategoria.AddItem objBanco.gfsReadChar(.Fields("DE_CATEGORIA"))
                cmbCategoria.ItemData(cmbCategoria.NewIndex) = objBanco.gfsReadInt(.Fields("ID_CATEGORIA"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboCategoria:
    Call objSystem.gsExibeErros(Err, "psCarregarComboCategoria()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboEstadoCivil()

    On Error GoTo err_psCarregarComboEstadoCivil:

    cmbEstadoCivil.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_estadocivil ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbEstadoCivil.AddItem objBanco.gfsReadChar(.Fields("DE_ESTADOCIVIL"))
                cmbEstadoCivil.ItemData(cmbEstadoCivil.NewIndex) = objBanco.gfsReadInt(.Fields("ID_ESTADOCIVIL"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboEstadoCivil:
    Call objSystem.gsExibeErros(Err, "psCarregarComboEstadoCivil()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboFilial()

    On Error GoTo err_psCarregarComboFilial:

    cmbFilial.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_filial ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbFilial.AddItem objBanco.gfsReadChar(.Fields("CD_FILIAL"))
                cmbFilial.ItemData(cmbFilial.NewIndex) = objBanco.gfsReadInt(.Fields("CD_FILIAL"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboFilial:
    Call objSystem.gsExibeErros(Err, "psCarregarComboFilial()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboParentesco()

    On Error GoTo err_psCarregarComboParentesco:

    cmbParentesco.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_parentesco ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbParentesco.AddItem objBanco.gfsReadChar(.Fields("DE_PARENTESCO"))
                cmbParentesco.ItemData(cmbParentesco.NewIndex) = objBanco.gfsReadInt(.Fields("ID_PARENTESCO"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboParentesco:
    Call objSystem.gsExibeErros(Err, "psCarregarComboParentesco()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboSexo()

    On Error GoTo err_psCarregarComboSexo:

    cmbSexo.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_sexo ( NULL, NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbSexo.AddItem objBanco.gfsReadChar(.Fields("DE_SEXO"))
                cmbSexo.ItemData(cmbSexo.NewIndex) = objBanco.gfsReadInt(.Fields("ID_SEXO"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboSexo:
    Call objSystem.gsExibeErros(Err, "psCarregarComboSexo()", CStr(Me.Name))

End Sub

Private Sub psCarregarComboTitular()

    On Error GoTo err_psCarregarComboTitular:

    cmbTitular.Clear

    Set cmd = New ADODB.Command
    With cmd

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_con_pessoa_titular ( NULL );"

        .CommandText = pSql

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly

        Set rs = .Execute
        With rs

            While Not .EOF

                cmbTitular.AddItem objBanco.gfsReadChar(.Fields("NM_PESSOA"))
                cmbTitular.ItemData(cmbTitular.NewIndex) = objBanco.gfsReadInt(.Fields("CD_PESSOA"))

                .MoveNext

            Wend
            .Close

        End With
        Set rs = Nothing

    End With
    Set cmd = Nothing

    Exit Sub

err_psCarregarComboTitular:
    Call objSystem.gsExibeErros(Err, "psCarregarComboTitular()", CStr(Me.Name))

End Sub

Private Sub psGerarCupons()

    On Error GoTo err_psGerarCupons:

    Dim pDt_Evento      As Date
    Dim pDt_Nascimento  As Date
    Dim pNr_Idade       As Integer
    Dim pQt_Cupom       As Integer
    Dim pCd_Pessoa      As Long
    Dim pNm_Tabela      As String

    With Me

        pNm_Tabela = "tb_pessoacupom": pDt_Evento = CDate("10/12/2011")

        pCd_Pessoa = CLng(.txtCodigo.Text)
        pDt_Nascimento = CDate(.mkeDataNascimento.Text)

        pNr_Idade = DateDiff("yyyy", pDt_Nascimento, pDt_Evento)

        Set cmd = New ADODB.Command
        With cmd

            .ActiveConnection = cn
            .CommandTimeout = 360000000
            .CommandType = adCmdText

            pSql = "": pSql = "CALL sp_con_cupom ( NULL, NULL );"

            .CommandText = pSql

            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly

            Set rs = .Execute
            With rs

                While Not .EOF

                    Select Case pNr_Idade
                        Case Is < 1
                            If LCase(objBanco.gfsReadChar(.Fields("DE_CUPOM"))) = LCase("PAPINHA") Then

                                For pQt_Cupom = 1 To objBanco.gfsReadInt(.Fields("QT_CUPOM"))

                                    pSql = "": pSql = "CALL sp_ins_pessoacupom"
                                    pSql = pSql & " (" & objBanco.gfsSaveInt(objBanco.gflProximoRegistro(pNm_Tabela))
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(pCd_Pessoa)
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(objBanco.gfsReadInt(.Fields("ID_CUPOM")))
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("S")
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("N")
                                    pSql = pSql & " );"

                                    pVet_Sql(pContador) = pSql: pContador = pContador + 1

                                Next pQt_Cupom

                            End If

                        Case 1 To 21
                            If LCase(objBanco.gfsReadChar(.Fields("DE_CUPOM"))) <> LCase("CERVEJA") And LCase(objBanco.gfsReadChar(.Fields("DE_CUPOM"))) <> LCase("PAPINHA") Then

                                For pQt_Cupom = 1 To objBanco.gfsReadInt(.Fields("QT_CUPOM"))

                                    pSql = "": pSql = "CALL sp_ins_pessoacupom"
                                    pSql = pSql & " (" & objBanco.gfsSaveInt(objBanco.gflProximoRegistro(pNm_Tabela))
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(pCd_Pessoa)
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(objBanco.gfsReadInt(.Fields("ID_CUPOM")))
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("S")
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("N")
                                    pSql = pSql & " );"

                                    pVet_Sql(pContador) = pSql: pContador = pContador + 1

                                Next pQt_Cupom

                            End If

                        Case Is > 21
                            If LCase(objBanco.gfsReadChar(.Fields("DE_CUPOM"))) <> LCase("PAPINHA") Then

                                For pQt_Cupom = 1 To objBanco.gfsReadInt(.Fields("QT_CUPOM"))

                                    pSql = "": pSql = "CALL sp_ins_pessoacupom"
                                    pSql = pSql & " (" & objBanco.gfsSaveInt(objBanco.gflProximoRegistro(pNm_Tabela))
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(pCd_Pessoa)
                                    pSql = pSql & ", " & objBanco.gfsSaveInt(objBanco.gfsReadInt(.Fields("ID_CUPOM")))
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("S")
                                    pSql = pSql & ", " & objBanco.gfsSaveChar("N")
                                    pSql = pSql & " );"

                                    pVet_Sql(pContador) = pSql: pContador = pContador + 1

                                Next pQt_Cupom

                            End If

                    End Select

                    DoEvents: .MoveNext: DoEvents

                Wend
                .Close

            End With
            Set rs = Nothing

        End With
        Set cmd = Nothing

    End With

    Exit Sub

err_psGerarCupons:
    Call objSystem.gsExibeErros(Err, "psGerarCupons()", CStr(Me.Name))

End Sub

Private Sub psGravarLogImpressao(pCd_Pessoa As Long)

    On Error GoTo err_psGravarLogImpressao:

    Set cmd2 = New ADODB.Command
    With cmd2

        .ActiveConnection = cn
        .CommandTimeout = 360000000
        .CommandType = adCmdText

        pSql = "": pSql = "CALL sp_ins_log ( "
        pSql = pSql & objBanco.gfsSaveInt(gfiBuscarCodigoTipoLog("IMPRESSAO"))
        pSql = pSql & ", " & objBanco.gfsSaveInt(pCd_Pessoa)
        pSql = pSql & ", " & objBanco.gfsSaveInt(pCd_User)
        pSql = pSql & ", NULL );"

        .CommandText = pSql
        .Execute

    End With
    Set cmd2 = Nothing

    Exit Sub

err_psGravarLogImpressao:
    Call objSystem.gsExibeErros(Err, "psGravarLogImpressao()", CStr(Me.Name), pSql)

End Sub

Private Sub psImprimirEtiqueta()

    On Error GoTo err_psImprimirEtiqueta:

    Dim pNr_Via         As Integer
    Dim pPrinter        As Printer
    Dim pCodigoBarra    As String
    Dim pImprime        As String

    If Len(Trim(txtCodigo.Text)) > 0 Then

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

            pSql = "": pSql = "CALL sp_con_pessoa"
            pSql = pSql & " ( " & objBanco.gfsSaveInt(txtCodigo.Text)
            pSql = pSql & ", NULL, NULL, NULL, NULL );"

            .CommandText = pSql

            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            rs.CursorType = adOpenForwardOnly

            Set rs = .Execute
            With rs

                If Not .EOF Then

                    pCodigoBarra = Format(objBanco.gfsReadInt(.Fields("CD_PESSOA")), "000000")
                    pNr_Via = objBanco.gfsReadInt(.Fields("NR_VIA"))

                    If Not IsNull(.Fields("DH_IMPRESSAO")) Then

                        pNr_Via = pNr_Via + 1

                    Else

                        If pNr_Via = 1 Then pNr_Via = pNr_Via + 1

                    End If

                    pCodigoBarra = pCodigoBarra & CStr(pNr_Via)

                    Set cmd2 = New ADODB.Command

                    cmd2.ActiveConnection = cn
                    cmd2.CommandTimeout = 360000000
                    cmd2.CommandType = adCmdText

                    pSql = "": pSql = "CALL sp_con_etiqueta"
                    pSql = pSql & " ( NULL"
                    pSql = pSql & ", " & objBanco.gfsSaveInt("1")
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

                            If LCase(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))) = LCase("CD_PESSOA") Then

                                pImprime = "Nº do convite: " & pCodigoBarra

                            Else

                                pImprime = objBanco.gfsReadChar(.Fields(objBanco.gfsReadChar(rs2.Fields("DE_CAMPOETIQUETA"))))

                            End If

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

                    Call psAtualizarDadosPessoa(objBanco.gfsReadInt(.Fields("CD_PESSOA")), pNr_Via)

                    Call psGravarLogImpressao(objBanco.gfsReadInt(.Fields("CD_PESSOA")))

                End If
                .Close

            End With
            Set rs = Nothing

        End With
        Set cmd = Nothing

        Screen.MousePointer = vbDefault

    End If

    Exit Sub

err_psImprimirEtiqueta:
    Screen.MousePointer = vbDefault
    Call objSystem.gsExibeErros(Err, "psImprimirEtiqueta()", CStr(Me.Name))

End Sub

Private Sub psLimparCampos()

    With Me

        Call psCarregarComboTitular

        pAutorizaCadastro = False: pInserirCadastro = False: pLeituraColetor = False

        .cmbCategoria.ListIndex = -1
        .cmbSexo.ListIndex = -1
        .cmbParentesco.ListIndex = -1
        .cmbEstadoCivil.ListIndex = -1
        .cmbTitular.ListIndex = -1
        .cmbFilial.ListIndex = -1
        .cmbCentroCusto.ListIndex = -1

        Call objSystem.gsLimparText(.Name)

        .mkeDataNascimento.Mask = ""
        .mkeDataNascimento.Text = ""
        .mkeDataNascimento.Mask = "##/##/####"
        .mkeDataLeitura.Mask = ""
        .mkeDataLeitura.Text = ""
        .mkeDataLeitura.Mask = "##/##/#### ##:##:##"
        .mkeDataImpressaoTicket.Mask = ""
        .mkeDataImpressaoTicket.Text = ""
        .mkeDataImpressaoTicket.Mask = "##/##/#### ##:##:##"

        Call gsHabilitarCampos(False, 1)

    End With

End Sub

Private Sub psNovo()

    On Error GoTo err_psNovo:

    If pTp_User = 3 Then

        frmAutorizacao.Show vbModal

        If pAutorizaCadastro Then

            With Me

                pInserirCadastro = True

                .txtCodigo.Text = gflBuscarCodigoPessoa
                .txtVia.Text = 1

                Call gsHabilitarCampos(True, 2)

                .cmbCategoria.SetFocus

            End With

        Else

            pMsg = "": pMsg = "Usuário sem permissão para realizar esta operação."
            MsgBox pMsg, vbCritical + vbOKOnly, "Atenção:"

            Call psLimparCampos

        End If

    Else

        With Me

            pInserirCadastro = True

            .txtCodigo.Text = gflBuscarCodigoPessoa
            .txtVia.Text = 1

            Call gsHabilitarCampos(True, 2)

            .cmbCategoria.SetFocus

        End With

    End If

    Exit Sub

err_psNovo:
    Call objSystem.gsExibeErros(Err, "psNovo()", CStr(Me.Name))

End Sub

Private Sub psSalvar()

    On Error GoTo err_psSalvar:

    Dim pTamanho As Long

    If Not pfbValidarCampos Then Exit Sub

    pContador = 1: Call objSystem.gsLimparVetor(pVet_Sql)

    With Me

        pSql = ""

        If pInserirCadastro Then pSql = "CALL sp_ins_pessoa" Else pSql = "CALL sp_upd_pessoa"

        pSql = pSql & " (" & objBanco.gfsSaveInt(.txtCodigo.Text)
        pSql = pSql & ", " & objBanco.gfsSaveInt(.txtVia.Text)

        If Len(Trim(.txtMatricula.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.txtMatricula.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtNome.Text)) Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.txtNome.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtNomeCracha.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.txtNomeCracha.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        If Len(Trim(.txtCPF.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.txtCPF.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveDate(.mkeDataNascimento.Text, "DH", gfvTipoData)

        If .cmbSexo.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbSexo.ItemData(.cmbSexo.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbParentesco.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbParentesco.ItemData(.cmbParentesco.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbEstadoCivil.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbEstadoCivil.ItemData(.cmbEstadoCivil.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbCategoria.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCategoria.ItemData(.cmbCategoria.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbFilial.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbFilial.ItemData(.cmbFilial.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If .cmbCentroCusto.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbCentroCusto.ItemData(.cmbCentroCusto.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & ", " & objBanco.gfsSaveChar("S")

        If pInserirCadastro Then pSql = pSql & ", " & objBanco.gfsSaveChar("N")

        If .cmbTitular.ListIndex <> -1 Then

            pSql = pSql & ", " & objBanco.gfsSaveInt(.cmbTitular.ItemData(.cmbTitular.ListIndex))

        Else

            pSql = pSql & ", NULL"

        End If

        If pInserirCadastro Then pSql = pSql & ", " & objBanco.gfsSaveChar("SISTEMA DE CREDENCIAMENTO")

        If Len(Trim(.txtObservacao.Text)) > 0 Then

            pSql = pSql & ", " & objBanco.gfsSaveChar(.txtObservacao.Text)

        Else

            pSql = pSql & ", NULL"

        End If

        pSql = pSql & " );"

        pVet_Sql(pContador) = pSql: pContador = pContador + 1

        If pInserirCadastro Then Call psGerarCupons

    End With

    pTamanho = UBound(pVet_Sql, 1)

    cn.BeginTrans
    For pContador = 1 To pTamanho

        pSql = pVet_Sql(pContador)
        If Len(Trim(pSql)) = 0 Then Exit For
        If objBanco.gfiExecuteSql(pSql) = -1 Then cn.RollbackTrans: Exit Sub

    Next pContador
    cn.CommitTrans

    Call gsHabilitarCampos(False, 3)

    Exit Sub

err_psSalvar:
    Call objSystem.gsExibeErros(Err, "psSalvar()", CStr(Me.Name), pSql)

End Sub

Private Function pfbValidarCampos() As Boolean

    On Error GoTo err_pfbValidarCampos:

    pfbValidarCampos = False

    With Me

        If .cmbCategoria.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar a categoria!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbCategoria.SetFocus: Exit Function

        End If

        If Len(Trim(.txtMatricula.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar a matrícula do titular!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtMatricula.SetFocus: Exit Function

        End If

        If Len(Trim(.txtCPF.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar o CPF do titular!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtCPF.SetFocus: Exit Function

        End If

        If Len(Trim(.txtNome.Text)) = 0 Then

            pMsg = "": pMsg = "É necessário informar o Nome!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .txtNome.SetFocus: Exit Function

        End If

        If Trim(.mkeDataNascimento.Text) = "__/__/____" Then

            pMsg = "": pMsg = "É necessário informar a Data de Nascimento!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .mkeDataNascimento.SetFocus: Exit Function

        Else

            If Not IsDate(Trim(.mkeDataNascimento.Text)) Then

                pMsg = "": pMsg = "Data inválida!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .mkeDataNascimento.SetFocus: Exit Function

            End If

        End If

        If .cmbSexo.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o sexo!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbSexo.SetFocus: Exit Function

        End If

        If .cmbEstadoCivil.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o estado civil!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbEstadoCivil.SetFocus: Exit Function

        End If

        If .cmbCategoria.ItemData(.cmbCategoria.ListIndex) = 2 Then

            If .cmbTitular.ListIndex = -1 Then

                pMsg = "": pMsg = "É necessário informar o titular do dependente!"
                MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
                .cmbTitular.SetFocus: Exit Function

            End If

        End If

        If .cmbFilial.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar a filial!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbFilial.SetFocus: Exit Function

        End If

        If .cmbCentroCusto.ListIndex = -1 Then

            pMsg = "": pMsg = "É necessário informar o centro de custo!"
            MsgBox pMsg, vbOKOnly + vbInformation, "Atenção:"
            .cmbCentroCusto.SetFocus: Exit Function

        End If

    End With

    pfbValidarCampos = True

    Exit Function

err_pfbValidarCampos:
    Call objSystem.gsExibeErros(Err, "pfbValidarCampos()", CStr(Me.Name))

End Function

Private Function pfsNomeCracha(pNome As String) As String

    On Error GoTo err_pfsNomeCracha:

    Dim pTamanho As Integer
    Dim pVet_Nome As Variant

    If Len(Trim(pNome)) > 22 Then

        pVet_Nome = Split(pNome, " ", -1)
        pTamanho = UBound(pVet_Nome, 1)
        pfsNomeCracha = pVet_Nome(0) + " " + pVet_Nome(pTamanho)

    Else

        pfsNomeCracha = pNome

    End If

    Exit Function

err_pfsNomeCracha:
    Call objSystem.gsExibeErros(Err, "pfsNomeCracha()", CStr(Me.Name))

End Function
