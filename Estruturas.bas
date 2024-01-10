Attribute VB_Name = "Estruturas"
Public Const OFFSET_BASE = &H2213
Public Const OFFSET_BASE_MAGIA = &H2681
Public Const OFFSET_ELMO = &H69
Public Const OFFSET_ARMADURA = &H84
Public Const OFFSET_RELIC = &HB0
Public Const OFFSET_SWDTECH = &H2A97
Public Const OFFSET_BLITZ = &H293C
Public Const OFFSET_DANCE = &H295F
Public Const OFFSET_BASE_RAGE = &H293F
Public Const OFFSET_BASE_LORE = &H293C
Public Const OFFSET_BASE_GP = &H2473
Public Const OFFSET_BASE_STEP = &H2479
Public Const OFFSET_BASE_TIPO_ITEM = &H247D
Public Const OFFSET_BASE_QTD_ITEM = &H257D
Public Const OFFSET_BASE_ESPER = &H267D
Public Const CARACTER_NOME = 63
Public Const NOME_APP As String = "FFVI - Editor de SaveState (ZSNES)"
Public Const ARQUIVO_INI As String = "ffvi.ini"
'---------ESTRUTURA DE DADOS DO SAVE STATE---------
Type tElemento
    valor As Byte
    offset As Integer
End Type

Type tNome
    Caracter(5) As tElemento
End Type

Type tPersonagem
    nome As tNome
    HP(1) As tElemento
    HPmax(1) As tElemento
    MP(1) As tElemento
    MPmax(1) As tElemento
    experiencia(2) As tElemento
    level As tElemento
    status As tElemento
    float As tElemento
    comando(3) As tElemento
    vigor As tElemento
    speed As tElemento
    stamina As tElemento
    magic As tElemento
    esper As tElemento
    maoEsq As tElemento
    maoDir As tElemento
    cabeca As tElemento
    corpo As tElemento
    relic1 As tElemento
    relic2 As tElemento
    magias(54) As tElemento
End Type

Type tEspeciais
    lore(2) As tElemento
    rage(31) As tElemento
    blitz As tElemento
    dance As tElemento
    swdTech As tElemento
End Type

Type tItens
    tipoItem(254) As tElemento
    qtdItem(254) As tElemento
    ctrlItem(254) As Integer
End Type

Type tState
    personagens(13) As tPersonagem
    especiais As tEspeciais
    itens As tItens
    espers(3) As tElemento
    gil(2) As tElemento
    steps(2) As tElemento
End Type
'---------ESTRUTURA DE DADOS DE CONFIGURAÇÃO---------
Type tConfig
    Arq(4) As String
    TipoPath As Integer
    Path As String
End Type
