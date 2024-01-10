Attribute VB_Name = "Funcoes"
Option Explicit
Public mapaMemoria As tState
Public config As tConfig
Public arquivoAberto As Boolean
Public arquivoAlterado As Boolean

'Declaração de API´s
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnedString$, ByVal nSize%, ByVal lpFileName$) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function ReadINIFile(sSecao As String, sItem As String, sArquivo As String) As String
   
   Dim sLinha As String
   Dim iBytes As Integer

   sLinha = Space(255)
   iBytes = GetPrivateProfileString(sSecao, sItem, "", sLinha, 255, sArquivo)
   ReadINIFile = Nulo(Left(sLinha, iBytes))
   
End Function

Public Function WriteINIFile(sSecao As String, sItem As String, sValor As String, sArquivo As String)
   
   Dim iRet As Integer
   iRet = WritePrivateProfileString(sSecao, sItem, sValor, sArquivo)
   
End Function

Public Function Nulo(sNulo As Variant) As String
   
   'Ajusta a string de dados
   On Error GoTo Erros
   Nulo = "" & Trim(sNulo)
   On Error GoTo 0
   
   Exit Function

Erros:
   Nulo = ""
   
End Function

Public Function converteDeLetra(letra As Byte) As String
    If (letra >= 65 And letra <= 90) Then
        ' letra maiúscula
        converteDeLetra = Chr(letra)
    Else
        If (letra >= 91 And letra <= 116) Then
            ' letra minúscula
            converteDeLetra = Chr(letra + 6)
        Else
            If (letra >= 117 And letra <= 126) Then
                ' número
                converteDeLetra = Trim(Str(letra - 117))
            Else
                Select Case (letra)
                    Case 192
                        converteDeLetra = " "
                    Case 127
                        converteDeLetra = "!"
                    Case 128
                        converteDeLetra = "?"
                    Case 129
                        converteDeLetra = "/"
                    Case 130
                        converteDeLetra = ":"
                    Case 131
                        ' aspas duplas
                        converteDeLetra = Chr(34)
                    Case 132
                        ' aspas simples
                        converteDeLetra = Chr(39)
                    Case 133
                        converteDeLetra = "-"
                    Case 134
                        converteDeLetra = "."
                    Case Else
                        ' caso não seja nenhum caracter
                        ' padrão disponível para nomes
                        ' do jogo, insere um espaço
                        converteDeLetra = " "
                End Select
            End If
        End If
    End If
End Function

Public Function converteParaLetra(letra As String) As Byte
    If (letra >= 65 And letra <= 90) Then
        ' letra maiúscula
        converteDeLetra = CByte(letra)
    Else
        If (letra >= 97 And letra <= 122) Then
            ' letra minúscula
            converteDeLetra = CByte(letra + 6)
        Else
            If (letra >= 48 And letra <= 57) Then
                ' número
                converteDeLetra = CByte(letra) + 117
            Else
                Select Case (letra)
                    Case " "
                        converteDeLetra = 192
                    Case "!"
                        converteDeLetra = 127
                    Case "?"
                        converteDeLetra = 128
                    Case "/"
                        converteDeLetra = 129
                    Case ":"
                        converteDeLetra = 130
                    Case Chr(34)
                        ' aspas duplas
                        converteDeLetra = 131
                    Case Chr(39)
                        ' aspas simples
                        converteDeLetra = 132
                    Case "-"
                        converteDeLetra = 133
                    Case "."
                        converteDeLetra = 134
                    Case Else
                        ' caso não seja nenhum caracter
                        ' padrão disponível para nomes
                        ' do jogo, insere um espaço
                        converteDeLetra = 192
                End Select
            End If
        End If
    End If
End Function

Public Function Pow(ByVal base As Integer, ByVal expoente As Integer) As Integer
    Dim I, retorno As Integer
    retorno = 1
    For I = 1 To expoente
        retorno = retorno * base
    Next
    Pow = retorno
End Function

Public Sub msgErro(erro As ErrObject)
    MsgBox "Ocorreu um erro:" & Chr(13) & erro.Number & Chr(13) & erro.Description, vbOKOnly, NOME_APP
End Sub

Public Sub checkboxByte(elemento As tElemento, ByRef chkBox As Variant, base As Integer, limite As Integer)
    Dim bitTemp, I As Integer
    bitTemp = elemento.valor
    For I = 7 To 0 Step -1
        ' verifica se o índice está
        ' dentro dos limites do array
        If ((I + base) <= limite) Then
            ' adicionar verificação de máximo
            If (bitTemp >= Pow(2, I)) Then
                chkBox(I + base).Value = 1
                bitTemp = bitTemp - Pow(2, I)
            Else
                chkBox(I + base).Value = 0
            End If
        End If
    Next
End Sub

Public Sub listboxByte(elemento As tElemento, ByRef listBox As Variant, base As Integer)
    Dim bitTemp, I As Integer
    bitTemp = elemento.valor
    For I = 7 To 0 Step -1
        If (I + base <= listBox.ListCount - 1) Then
            If (bitTemp >= Pow(2, I)) Then
                listBox.Selected(I + base) = True
                bitTemp = bitTemp - Pow(2, I)
            Else
                listBox.Selected(I + base) = False
            End If
        End If
    Next
End Sub

Public Function gravaByte(ByRef hArquivo As Long, elemento As tElemento) As Boolean
    On Error GoTo Erro1
    If (hArquivo <> 0) Then
        Seek #hArquivo, elemento.offset + OFFSET_BASE
        'Put #hArquivo, Binary, elemento.valor
        gravaByte = True
    Else
        gravaByte = False
    End If
    Exit Function
Erro1:
    gravaByte = False
End Function

Public Sub leByte(ByRef hArquivo As Long, ByRef elemento As tElemento)
    On Error GoTo Erro1
    If (hArquivo <> 0) Then
        Get #hArquivo, elemento.offset, elemento.valor
    End If
    Exit Sub
Erro1:
    msgErro Err
End Sub

Public Sub preencheMapaMemoria(arquivo As String)
    Dim I, J As Integer
    Dim offsetAtual As Integer
    Dim hArquivo As Long
    Dim erro As Boolean
    
    hArquivo = FreeFile
    Open arquivo For Binary Access Read As hArquivo
    
    offsetAtual = OFFSET_BASE
        ' nesse loop são preenchidos os
        ' atributos de todos os personagens
        For I = 0 To 13
            ' soma 2 para pular os offsets
            ' que identificam personagem
            offsetAtual = offsetAtual + 2
            
            For J = 0 To 5
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).nome.Caracter(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).nome.Caracter(J)
            Next
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).level.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).level
            
            For J = 0 To 1
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).HP(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).HP(J)
            Next
            
            For J = 0 To 1
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).HPmax(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).HPmax(J)
            Next
            
            For J = 0 To 1
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).MP(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).MP(J)
            Next
            
            For J = 0 To 1
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).MPmax(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).MPmax(J)
            Next
            
            For J = 0 To 2
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).experiencia(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).experiencia(J)
            Next
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).status.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).status
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).float.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).float
            
            For J = 0 To 3
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).comando(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).comando(J)
            Next
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).vigor.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).vigor
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).stamina.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).stamina
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).speed.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).speed
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).magic.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).magic
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).esper.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).esper
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).maoEsq.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).maoEsq
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).maoDir.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).maoDir
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).cabeca.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).cabeca
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).corpo.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).corpo
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).relic1.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).relic1
            
            offsetAtual = offsetAtual + 1
            mapaMemoria.personagens(I).relic2.offset = offsetAtual
            leByte hArquivo, mapaMemoria.personagens(I).relic2
        Next
        
        ' magias dos heróis que as podem usar
        offsetAtual = OFFSET_BASE_MAGIA - 1
        For I = 0 To 11
            For J = 0 To 53
                offsetAtual = offsetAtual + 1
                mapaMemoria.personagens(I).magias(J).offset = offsetAtual
                leByte hArquivo, mapaMemoria.personagens(I).magias(J)
            Next
        Next
        
        ' comandos especiais, como blitz, lore, dance, etc.
        mapaMemoria.especiais.blitz.offset = OFFSET_BLITZ
        leByte hArquivo, mapaMemoria.especiais.blitz
        
        mapaMemoria.especiais.swdTech.offset = OFFSET_SWDTECH
        leByte hArquivo, mapaMemoria.especiais.swdTech
        
        mapaMemoria.especiais.dance.offset = OFFSET_DANCE
        leByte hArquivo, mapaMemoria.especiais.dance
        
        offsetAtual = OFFSET_BASE_RAGE - 1
        For I = 0 To 31
            offsetAtual = offsetAtual + 1
            mapaMemoria.especiais.rage(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.especiais.rage(I)
        Next
        
        offsetAtual = OFFSET_BASE_LORE - 1
        For I = 0 To 2
            offsetAtual = offsetAtual + 1
            mapaMemoria.especiais.lore(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.especiais.lore(I)
        Next
        
        ' dinheiro (gil/GP) e steps
        offsetAtual = OFFSET_BASE_GP - 1
        For I = 0 To 2
            offsetAtual = offsetAtual + 1
            mapaMemoria.gil(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.gil(I)
        Next
        
        offsetAtual = OFFSET_BASE_STEP - 1
        For I = 0 To 2
            offsetAtual = offsetAtual + 1
            mapaMemoria.steps(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.steps(I)
        Next
        
        ' itens
        offsetAtual = OFFSET_BASE_TIPO_ITEM - 1
        For I = 0 To 254
            offsetAtual = offsetAtual + 1
            mapaMemoria.itens.tipoItem(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.itens.tipoItem(I)
        Next
        
        offsetAtual = OFFSET_BASE_QTD_ITEM - 1
        For I = 0 To 254
            offsetAtual = offsetAtual + 1
            mapaMemoria.itens.qtdItem(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.itens.qtdItem(I)
        Next
        
        'invocações (Espers)
        offsetAtual = OFFSET_BASE_ESPER - 1
        For I = 0 To 3
            offsetAtual = offsetAtual + 1
            mapaMemoria.espers(I).offset = offsetAtual
            leByte hArquivo, mapaMemoria.espers(I)
        Next
    Close hArquivo
End Sub

Public Function ZeroAEsquerda(valor As Byte) As String
    Dim retorno As String
    If (valor < 10) Then
        retorno = "00"
    Else
        If (valor < 100) Then
            retorno = "0"
        End If
    End If
    retorno = retorno & CStr(valor)
    ZeroAEsquerda = retorno
End Function
