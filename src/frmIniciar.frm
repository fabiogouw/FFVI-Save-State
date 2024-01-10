VERSION 5.00
Begin VB.Form frmIniciar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3600
      Top             =   1440
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   450
      Picture         =   "frmIniciar.frx":0000
      Top             =   240
      Width           =   3225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Iniciando aplicativo, aguarde..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "frmIniciar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub criaConfiguracoes()
    Dim hArquivo As Long
    Dim temp As String
    'como o arquivo não existe, ele é criado
    hArquivo = FreeFile
    Open (App.Path & "\" & ARQUIVO_INI) For Output Access Write As #hArquivo
    Print #hArquivo, "[ArquivosRecentes]"
    Print #hArquivo, "Arq1="
    Print #hArquivo, "Arq2="
    Print #hArquivo, "Arq3="
    Print #hArquivo, "Arq4="
    Print #hArquivo, "[PathSaveState]"
    Print #hArquivo, "TipoPath=0"
    Print #hArquivo, "Path="
    Close #hArquivo
    'o ED de configuração é preenchida
    config.Arq(0) = ""
    config.Arq(1) = ""
    config.Arq(2) = ""
    config.Arq(3) = ""
    config.Path = App.Path
    config.TipoPath = 0
End Sub

Private Sub carregaConfiguracoes()
    config.Arq(0) = ReadINIFile("ArquivosRecentes", "Arq1", App.Path & "\" & ARQUIVO_INI)
    config.Arq(1) = ReadINIFile("ArquivosRecentes", "Arq2", App.Path & "\" & ARQUIVO_INI)
    config.Arq(2) = ReadINIFile("ArquivosRecentes", "Arq3", App.Path & "\" & ARQUIVO_INI)
    config.Arq(3) = ReadINIFile("ArquivosRecentes", "Arq4", App.Path & "\" & ARQUIVO_INI)
    config.TipoPath = CInt(ReadINIFile("PathSaveState", "TipoPath", App.Path & "\" & ARQUIVO_INI))
    config.Path = ReadINIFile("PathSaveState", "Path", App.Path & "\" & ARQUIVO_INI)
    frmPrincipal.ajustaMenuReabrir
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    If (Dir(App.Path & "\" & ARQUIVO_INI) = ARQUIVO_INI) Then
        carregaConfiguracoes
    Else
        criaConfiguracoes
    End If
    frmPrincipal.preencheCampos
    arquivoAberto = False
    frmPrincipal.travarControles (arquivoAberto)
    frmPrincipal.Show
    Unload Me
End Sub
