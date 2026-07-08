Attribute VB_Name = "Mod7_IncluirCalculadora"
'==========================================================================
' ATUALIZADOR DE SWAPS - Modulo 7: Assistente Incluir Calculadora (v1.0)
'
' OBJETIVO
'   Cadastrar uma nova calculadora na aba Config sem precisar editar codigo
'   nem formatar a linha manualmente. Acionado pelo botao "Incluir
'   Calculadora" no Painel.
'
'   Arquitetura escalavel: a macro de atualizacao (Mod5) processa TODAS as
'   linhas da Config. Cadastrar = inserir uma linha. Para cadastrar
'   varias calculadoras de uma vez (mesmo em grande quantidade) NUNCA e
'   preciso mexer no codigo.
'
' FLUXO (assistente em MsgBox/InputBox)
'   1. Nome da calculadora                   (com checagem de duplicidade)
'   2. Curvas utilizadas (separadas por ';')
'   3. O .xlsx ja esta na pasta MTM mais recente?
'        Sim -> varre a pasta e confirma (OK + caminho) ou avisa
'        Nao -> mostra o caminho exato e o nome sugerido
'                 (padrao: Calculadora SWAP_<NN> - <Nome>.xlsx)
'   4. Loop p/ incluir varias na mesma sessao
'
' DEPENDENCIAS (funcoes publicas de outros modulos)
'   Mod1_Orquestrador : SHEET_CONFIG, GravarLog, LerParametro
'   Mod3_PastasMTM    : ColetarXLSX, BuscarArquivoPorNome
'
' INSTALACAO
'   Este modulo e sincronizado automaticamente pelo botao
'   "Atualizar Modulos" (Mod0_SyncModulos), bastando que a primeira versao
'   esteja importada no projeto VBA.
'==========================================================================
Option Explicit

'==========================================================================
' ENTRY POINT (acionado pelo botao do Painel)
'==========================================================================
Public Sub IncluirCalculadora()
    On Error GoTo TratarErro

    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets(SHEET_CONFIG)

    Dim adicionadas As Long: adicionadas = 0

    Do
        ' --- 1. Nome -----------------------------------------------------
        Dim nome As String
        nome = Trim(InputBox( _
            "Nome da calculadora (ex.: Banco Kaito V):" & vbCrLf & vbCrLf & _
            "Deixe em branco (ou Cancelar) para encerrar.", _
            "Incluir Calculadora  -  passo 1 de 3"))
        If Len(nome) = 0 Then Exit Do

        If NomeJaExisteConfig(wsConfig, nome) Then
            MsgBox "Ja existe uma calculadora chamada '" & nome & "' na Config." & vbCrLf & _
                   "Escolha outro nome.", vbExclamation, "Nome duplicado"
            GoTo Continua
        End If

        ' --- 2. Curvas ---------------------------------------------------
        Dim curvas As String
        curvas = Trim(InputBox( _
            "Curvas utilizadas por '" & nome & "', separadas por ';'." & vbCrLf & _
            "Ex.: FatorPre;FWDUSDBRL;CDI", _
            "Incluir Calculadora  -  passo 2 de 3"))

        ' --- 3. Insere a linha ------------------------------------------
        Application.ScreenUpdating = False
        Dim novaLinha As Long
        novaLinha = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row + 1
        If novaLinha < 2 Then novaLinha = 2

        FormatarLinhaConfig wsConfig, novaLinha
        wsConfig.Cells(novaLinha, 1).Value = nome                                    ' A: Contrato
        wsConfig.Cells(novaLinha, 3).Value = curvas                                  ' C: Curvas
        wsConfig.Cells(novaLinha, 5).Value = "(novo - ainda nao processado)"         ' E: Status
        ' B (caminho) e D (celula data) ficam vazios -> preenchimento automatico
        Application.ScreenUpdating = True

        adicionadas = adicionadas + 1
        GravarLog "INFO", nome, "Calculadora incluida na Config (linha " & novaLinha & _
                  ") | Curvas: " & curvas

        ' --- 4. Verifica presenca na pasta MTM mais recente -------------
        Dim msgArquivo As String
        msgArquivo = VerificarPresencaPastaMTM(nome)
        MsgBox "'" & nome & "' adicionada na linha " & novaLinha & "." & vbCrLf & vbCrLf & _
               msgArquivo, vbInformation, _
               "Incluir Calculadora  -  passo 3 de 3"

        ' --- 5. Loop p/ proxima calculadora -----------------------------
        If MsgBox("Deseja incluir outra calculadora?", _
                  vbQuestion + vbYesNo, "Incluir Calculadora") <> vbYes Then Exit Do
Continua:
    Loop

    If adicionadas > 0 Then
        MsgBox adicionadas & " calculadora(s) incluida(s) com sucesso." & vbCrLf & vbCrLf & _
               "Proximos passos:" & vbCrLf & _
               "  - O caminho (coluna B) sera preenchido automaticamente na" & vbCrLf & _
               "    proxima execucao do botao 'Atualizar'." & vbCrLf & _
               "  - A celula de data (coluna D) e detectada sozinha." & vbCrLf & _
               "  - Confira o nome e as curvas antes de rodar.", _
               vbInformation, "Incluir Calculadora"
    End If
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    GravarLog "ERRO", "IncluirCalculadora", Err.Description
    MsgBox "Erro ao incluir calculadora: " & Err.Description, vbExclamation, "Incluir Calculadora"
End Sub

'==========================================================================
' HELPERS PRIVADOS
'==========================================================================

' Verdadeiro se 'nome' ja existe na coluna A da Config (case-insensitive).
Private Function NomeJaExisteConfig(wsConfig As Worksheet, nome As String) As Boolean
    Dim ult As Long, i As Long
    ult = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ult
        If LCase(Trim(CStr(wsConfig.Cells(i, 1).Value))) = LCase(nome) Then
            NomeJaExisteConfig = True
            Exit Function
        End If
    Next i
End Function

' Aplica a formatacao padrao da tabela a uma linha recem-criada da Config:
' copia os formatos da linha anterior (bordas/fonte/alinhamento), corrige a
' faixa zebrada pela paridade da linha e recria a formula do farol (col F).
Private Sub FormatarLinhaConfig(wsConfig As Worksheet, linha As Long)
    Dim origem As Long: origem = linha - 1
    If origem >= 2 Then
        wsConfig.Range(wsConfig.Cells(origem, 1), wsConfig.Cells(origem, 6)).Copy
        wsConfig.Range(wsConfig.Cells(linha, 1), wsConfig.Cells(linha, 6)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        wsConfig.Range(wsConfig.Cells(linha, 1), wsConfig.Cells(linha, 6)).ClearContents
    End If

    ' Zebra: linhas impares recebem azul claro; pares ficam brancas.
    Dim corFundo As Long
    If linha Mod 2 = 1 Then corFundo = RGB(217, 225, 242) Else corFundo = RGB(255, 255, 255)
    wsConfig.Range(wsConfig.Cells(linha, 1), wsConfig.Cells(linha, 6)).Interior.Color = corFundo

    ' Farol (col F): mesma formula das demais linhas (ChrW p/ os simbolos).
    wsConfig.Cells(linha, 6).Formula = _
        "=IF(ISNUMBER(SEARCH(""ERRO"",E" & linha & ")),""" & ChrW(10005) & """," & _
        "IF(ISNUMBER(SEARCH(""AVISO"",E" & linha & ")),""" & ChrW(9650) & """," & _
        "IF(ISNUMBER(SEARCH(""OK"",E" & linha & ")),""" & ChrW(9679) & ""","""")))"
End Sub

' Pergunta ao usuario se a calculadora ja esta na pasta MTM mais recente.
' Se SIM: varre a pasta e confirma (OK + caminho) ou avisa que nao achou.
' Se NAO: mostra o caminho exato esperado + nome sugerido no padrao
'         "Calculadora SWAP_<NN> - <Nome>.xlsx" (NN = max+1 da pasta).
' Retorna a mensagem que sera mostrada ao usuario.
Private Function VerificarPresencaPastaMTM(nome As String) As String
    Dim pastaRecente As String
    pastaRecente = ObterPastaMTMMaisRecente()

    If Len(pastaRecente) = 0 Then
        VerificarPresencaPastaMTM = _
            "Cadastro salvo na Config. (Nao foi possivel localizar a pasta MTM" & vbCrLf & _
            "mais recente para verificar o arquivo .xlsx.)"
        GravarLog "AVISO", nome, "Inclusao: pasta MTM mais recente nao encontrada"
        Exit Function
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "O arquivo .xlsx desta calculadora ja esta na pasta MTM mais recente?" & vbCrLf & vbCrLf & _
        pastaRecente & vbCrLf & vbCrLf & _
        "  Sim -> a macro vai conferir se o arquivo esta la." & vbCrLf & _
        "  Nao -> a macro mostra onde colocar o arquivo antes de rodar.", _
        vbQuestion + vbYesNo, "Arquivo da calculadora '" & nome & "'")

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim col As Collection: Set col = New Collection
    ColetarXLSX pastaRecente, col, fso

    If resp = vbYes Then
        Dim caminho As String
        caminho = BuscarArquivoPorNome(col, nome)
        If Len(caminho) > 0 Then
            VerificarPresencaPastaMTM = _
                "Arquivo encontrado na pasta MTM mais recente:" & vbCrLf & vbCrLf & _
                "  " & caminho & vbCrLf & vbCrLf & _
                "No proximo 'Atualizar', sera copiado para a pasta D-1 e processado."
            GravarLog "OK", nome, "Inclusao: arquivo confirmado em " & caminho
        Else
            VerificarPresencaPastaMTM = _
                "AVISO: nao encontrei nenhum .xlsx com '" & nome & "' no nome em:" & vbCrLf & vbCrLf & _
                "  " & pastaRecente & vbCrLf & vbCrLf & _
                "Confira se o arquivo ja esta la e se o nome do contrato bate." & vbCrLf & _
                "Padrao esperado: Calculadora SWAP_<NN> - " & nome & ".xlsx"
            GravarLog "AVISO", nome, "Inclusao: arquivo nao encontrado em " & pastaRecente
        End If
    Else
        Dim proximoNN As Long: proximoNN = ProximoNumeroSWAP(col)
        Dim nomeSugerido As String
        nomeSugerido = "Calculadora SWAP_" & Format(proximoNN, "00") & " - " & nome & ".xlsx"
        VerificarPresencaPastaMTM = _
            "Coloque o arquivo .xlsx na pasta MTM mais recente antes do proximo" & vbCrLf & _
            "'Atualizar':" & vbCrLf & vbCrLf & _
            "  " & pastaRecente & vbCrLf & vbCrLf & _
            "Nome sugerido (mantem o padrao):" & vbCrLf & vbCrLf & _
            "  " & nomeSugerido & vbCrLf & vbCrLf & _
            "Sem o arquivo na pasta, o contrato vai aparecer com ERRO no proximo" & vbCrLf & _
            "fluxo (e o cadastro continua salvo aqui na Config)."
        GravarLog "INFO", nome, "Inclusao: arquivo ainda nao esta em " & pastaRecente & _
                  " | nome sugerido: " & nomeSugerido
    End If
End Function

' Retorna o caminho da pasta "MTM - DD.MM.YY" mais recente (data mais alta)
' dentro de PastaMTM_PU\MTM & PU (YYYY)\MTM & PU - MMM.YY\. Vazio se nao achar.
Private Function ObterPastaMTMMaisRecente() As String
    Dim pastaMTM As String: pastaMTM = LerParametro("PastaMTM_PU")
    If Len(pastaMTM) = 0 Then Exit Function

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pastaMTM) Then Exit Function

    Dim melhorData As Date: melhorData = DateSerial(1900, 1, 1)
    Dim melhorPasta As String

    Dim anoFo As Object, mesFo As Object, diaFo As Object
    On Error Resume Next
    For Each anoFo In fso.GetFolder(pastaMTM).SubFolders
        If InStr(LCase(anoFo.Name), "mtm & pu (") > 0 Then
            For Each mesFo In anoFo.SubFolders
                If InStr(LCase(mesFo.Name), "mtm & pu - ") > 0 Then
                    For Each diaFo In mesFo.SubFolders
                        Dim dt As Date
                        If TentarParseDataPastaMTM(diaFo.Name, dt) Then
                            If dt > melhorData Then
                                melhorData = dt
                                melhorPasta = diaFo.Path
                            End If
                        End If
                    Next diaFo
                End If
            Next mesFo
        End If
    Next anoFo
    On Error GoTo 0

    ObterPastaMTMMaisRecente = melhorPasta
End Function

' Espera nome no formato "MTM - DD.MM.YY". Retorna True + dt em sucesso.
Private Function TentarParseDataPastaMTM(nome As String, ByRef dt As Date) As Boolean
    On Error GoTo Falha
    Dim pos As Integer: pos = InStr(nome, "- ")
    If pos = 0 Then Exit Function
    Dim s As String: s = Trim(Mid(nome, pos + 2))
    Dim partes() As String: partes = Split(s, ".")
    If UBound(partes) < 2 Then Exit Function
    dt = DateSerial(2000 + CInt(partes(2)), CInt(partes(1)), CInt(partes(0)))
    TentarParseDataPastaMTM = True
    Exit Function
Falha:
End Function

' Varre os arquivos da pasta MTM mais recente atras de "Calculadora SWAP_NN"
' e retorna NN+1 (sugestao p/ proxima calculadora). Se nao achar nenhum, 1.
Private Function ProximoNumeroSWAP(col As Collection) As Long
    Dim item As Variant
    Dim maior As Long: maior = 0
    For Each item In col
        Dim base As String
        Dim p   As Integer: p = InStrRev(CStr(item), "\")
        base = Mid(CStr(item), p + 1)
        Dim ini As Integer: ini = InStr(1, base, "SWAP_", vbTextCompare)
        If ini > 0 Then
            Dim k As Integer: k = ini + 5
            Dim num As String: num = ""
            Do While k <= Len(base)
                Dim ch As String: ch = Mid(base, k, 1)
                If ch Like "[0-9]" Then num = num & ch Else Exit Do
                k = k + 1
            Loop
            If Len(num) > 0 Then
                If CLng(num) > maior Then maior = CLng(num)
            End If
        End If
    Next item
    ProximoNumeroSWAP = maior + 1
End Function
