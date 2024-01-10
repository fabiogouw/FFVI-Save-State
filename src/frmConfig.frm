VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2535
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1215
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame fracaminhoIni 
      Caption         =   "Arquivo de inicialização"
      Height          =   1335
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCaminhoIni 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.OptionButton optCaminhoIni 
         Caption         =   "Especificar caminho"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optCaminhoIni 
         Caption         =   "Usar último diretório acessado"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub gravaConfiguracoes()
    If (optCaminhoIni(1).Value = True) Then
    If (Trim(config.Path) <> "") Then
        WriteINIFile "PathSaveState", "Path", txtCaminhoIni.Text, App.Path & "\" & ARQUIVO_INI
    End If
    End If
End Sub

Private Sub cmdOk_Click()
    gravaConfiguracoes
End Sub

Private Sub Form_Load()
    If (config.TipoPath = 0) Then
        optCaminhoIni(1).Value = True
        txtCaminhoIni.Enabled = True
    Else
        optCaminhoIni(0).Value = True
        txtCaminhoIni.Enabled = False
    End If
    If (Trim(config.Path) = "") Then
        txtCaminhoIni = App.Path & "\"
    Else
        txtCaminhoIni = config.Path
    End If
End Sub

Private Sub optCaminhoIni_Click(Index As Integer)
    If (Index = 0) Then
        txtCaminhoIni.Enabled = False
        txtCaminhoIni.BackColor = vbButtonFace
    Else
        txtCaminhoIni.Enabled = True
        txtCaminhoIni.BackColor = vbWindowBackground
    End If
End Sub
