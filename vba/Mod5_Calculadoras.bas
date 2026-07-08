Attribute VB_Name = "Mod5_Calculadoras"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 5: Calculadoras (v4.3)
'
' NOVIDADE v4.2 - Auto-deteccao da celula de data:
'   Cada calculadora pode ter a data em celulas diferentes:
'     Nortis, Agroverde, Rodalux II, Metalcast -> N6
'     Kaito I/II/III, Zenith I/II/III, Orbita II -> M6
'     Voxa, Rodalux I, Provento                  -> P6
'
'   A macro detecta automaticamente qual celula usar lendo a formula
'   da celula Check (que sempre comeca com a ref da celula de data):
'     =IF(AND(N6='Fator Pre'!A2,...  -> data em N6
'     =IF(AND(M6='Fator Pre'!A2,...  -> data em M6
'     =IF(AND(P6='Fator Pre'!A2,...  -> data em P6
'
'   A celula Check tambem e detectada automaticamente (N2, O2 ou P2).
'
'   Col D da aba Config: override manual por contrato (opcional).
'   Deixe vazio para auto-deteccao, ou coloque "M6" para forcar.
'
' ESTRUTURA ABA CONFIG:
'   A: Nome | B: Caminho (auto) | C: Curvas | D: Celula data (opcional)
'   E: Status execucao | F: Sinal (farol - formula automatica)
'
' OBS: a macro processa TODOS os contratos listados na Config.
'      Para deixar de processar um contrato, apague/limpe a linha dele.
'      Para cadastrar uma nova calculadora use o botao 'Incluir Calculadora'
'      no Painel (assistente vive em Mod7_IncluirCalculadora).
'
' MACRO UTILITARIA:
'   ScanearCelulasData -> escaneia todas as calculadoras e preenche col D
'   Execute uma vez apos instalar para revisar o que foi detectado.
'===========================================================================
Option Explicit

Public Const COL_STATUS   As Integer = 5   ' coluna E = status da ultima execucao
Public Const COL_OVERRIDE As Integer = 4   ' coluna D = override celula data

Private Function ABA_CDI() As String
    ABA_CDI = "Hist" & Chr(243) & "rico CDI B3"
End Function

'===========================================================================
' Loop principal
'===========================================================================
Public Sub AtualizarCalculadoras(dataD1 As Date, taxaCDI As Double, _
                                       caminhoCSS As String)
    Dim wsConfig  As Worksheet
    Dim wbCSS     As Workbook
    Dim ultLinha  As Long
    Dim i         As Long
    Dim totalOK   As Long
    Dim totalErro As Long

    On Error GoTo TratarErro

    ' Valida entrada basica
    If Len(Trim(caminhoCSS)) = 0 Then
        GravarLog "ERRO", "Calculadoras", "Parametro caminhoCSS vazio em AtualizarCalculadoras"
        Exit Sub
    End If
    If Dir(caminhoCSS) = "" Then
        GravarLog "ERRO", "Calculadoras", "Arquivo Curvas_Separadas nao encontrado: " & caminhoCSS
        Exit Sub
    End If

    Set wsConfig = ThisWorkbook.Sheets(SHEET_CONFIG)
    ultLinha = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row

    Set wbCSS = Workbooks.Open(Filename:=caminhoCSS, UpdateLinks:=0, ReadOnly:=True)
    GravarLog "INFO", "Calculadoras", "Curvas_Separadas aberto"

    For i = 2 To ultLinha
        ' Processa todos os contratos listados na Config.
        ' Linha sem nome (coluna A vazia) e ignorada.
        If Len(Trim(CStr(wsConfig.Cells(i, 1).Value))) = 0 Then GoTo ProximaLinha

        Dim nome     As String: nome     = Trim(wsConfig.Cells(i, 1).Value)
        Dim caminho  As String: caminho  = Trim(wsConfig.Cells(i, 2).Value)
        Dim curvas   As String: curvas   = Trim(wsConfig.Cells(i, 3).Value)
        Dim override As String: override = Trim(wsConfig.Cells(i, COL_OVERRIDE).Value)

        If InStr(LCase(caminho), "auto-preenchido") > 0 Then caminho = ""

        If Len(caminho) = 0 Then
            Dim msgVazio As String
            msgVazio = "ERRO: caminho nao encontrado. Re-execute o FluxoCompleto."
            wsConfig.Cells(i, COL_STATUS).Value = msgVazio
            GravarLog "ERRO", nome, msgVazio
            totalErro = totalErro + 1
            GoTo ProximaLinha
        End If

        Application.StatusBar = "[4/5] Atualizando: " & nome & "..."
        Dim resultado As String
        resultado = AtualizarUmaCalculadora(nome, caminho, curvas, dataD1, taxaCDI, wbCSS, override)

        wsConfig.Cells(i, COL_STATUS).Value = resultado
        GravarLog IIf(InStr(resultado, "ERRO") > 0, "ERRO", "OK"), nome, resultado

        If InStr(resultado, "ERRO") > 0 Then
            totalErro = totalErro + 1
        Else
            totalOK = totalOK + 1
        End If
ProximaLinha:
    Next i

    wbCSS.Close SaveChanges:=False
    GravarLog "INFO", "Calculadoras", "Calculadoras concluidas. OK=" & totalOK & " | ERRO=" & totalErro

ExitSub_AtualizarCalculadoras:
    On Error Resume Next
    If Not wbCSS Is Nothing Then wbCSS.Close SaveChanges:=False
    Application.StatusBar = False
    Exit Sub

TratarErro:
    Dim errMsg As String
    errMsg = "Erro " & Err.Number & ": " & Err.Description
    GravarLog "ERRO", "AtualizarCalculadoras", errMsg
    Resume ExitSub_AtualizarCalculadoras
End Sub
'===========================================================================
' Atualiza uma calculadora
'===========================================================================
Private Function AtualizarUmaCalculadora(nome As String, caminho As String, _
        listaCurvas As String, dataD1 As Date, taxaCDI As Double, _
        wbCSS As Workbook, override As String) As String
    On Error GoTo TratarErro

    If Dir(caminho) = "" Then
        AtualizarUmaCalculadora = "ERRO: arquivo nao encontrado: " & caminho
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=False)

    Dim nomeAbaMTM As String: nomeAbaMTM = EncontrarAbaMTM(wb)
    If Len(nomeAbaMTM) = 0 Then
        Dim listaAbas As String: listaAbas = ListarAbas(wb)
        wb.Close SaveChanges:=False
        AtualizarUmaCalculadora = "ERRO: aba MTM (PU/PMT) nao encontrada. Abas: " & listaAbas
        Exit Function
    End If

    Dim wsMTM As Worksheet: Set wsMTM = wb.Sheets(nomeAbaMTM)

    ' Auto-detecta celulas de data e check
    Dim celulaData  As String: celulaData  = DetectarCelulaData(wsMTM, override)
    Dim celulaCheck As String: celulaCheck = DetectarCelulaCheck(wsMTM)
    GravarLog "INFO", nome, "Aba=" & nomeAbaMTM & " | Data=" & celulaData & _
              " | Check=" & celulaCheck

    ' Processa curvas
    Dim arrCurvas() As String: arrCurvas = Split(listaCurvas, ";")
    Dim k As Integer

    ' Flags para evitar dupla execucao quando o Config lista uma curva
    ' que o fallback (abaixo do loop) tambem dispararia.
    Dim sofrFredJaProcessado As Boolean: sofrFredJaProcessado = False
    Dim sofr6mJaProcessado   As Boolean: sofr6mJaProcessado = False

    For k = LBound(arrCurvas) To UBound(arrCurvas)
        Dim nomeCurva As String: nomeCurva = Trim(arrCurvas(k))
        If Len(nomeCurva) = 0 Then GoTo ProxCurva

        Dim errMsg As String: errMsg = ""

        Select Case LCase(nomeCurva)
            Case "fatorpre", "pre"
                errMsg = CopiarCurvaDeCSS(wbCSS, "Pre", wb, "Fator Pre")
            Case "fwdusdbrl", "fwd_usd_brl"
                errMsg = CopiarCurvaDeCSS(wbCSS, "FWD_USD_BRL", wb, "FWD USDBRL")
            Case "cupomsofr", "cupom_sofr"
                errMsg = CopiarCurvaDeCSS(wbCSS, "Cupom_SOFR", wb, "Cupom SOFR")
            Case "sofr"
                ' Provento tem aba "SOFR" que recebe Daily Compound SOFR via
                ' FRED API (NAO eh Cupom SOFR do CSS/INETX como nas demais).
                ' Para as outras calculadoras, "sofr" continua sendo
                ' sinonimo de Cupom SOFR (compatibilidade retroativa).
                If LCase(nome) = "provento" Then
                    errMsg = AtualizarDailySOFR_FRED(wb, dataD1)
                    sofrFredJaProcessado = True
                Else
                    errMsg = CopiarCurvaDeCSS(wbCSS, "Cupom_SOFR", wb, "Cupom SOFR")
                End If
            Case "daily compound sofr", "dailycompoundsofr", "daily_compound_sofr"
                ' Voxa - taxa SOFR via FRED API (append no historico)
                errMsg = AtualizarDailySOFR_FRED(wb, dataD1)
                sofrFredJaProcessado = True
            Case "cdi"
                errMsg = AtualizarCDIHistorico(wb, dataD1, taxaCDI)
            Case "fwdjpybrl", "fwd_jpybrl"
                ' Agroverde, Orbita II - FWD JPYBRL (arquivo Bloomberg diario)
                errMsg = CopiarFWD_JPYBRL(wb, dataD1)
            Case "sofr6m"
                ' Rodalux I - SOFR6M (arquivo Bloomberg mensal).
                ' Provento NAO tem aba 'SOFR6M na BBG' - pular explicitamente
                ' (mesmo que o Config liste 'sofr6m' por inercia).
                If LCase(nome) = "provento" Then
                    GravarLog "INFO", nome, "Curva 'sofr6m' ignorada (Provento nao usa SOFR6M)"
                Else
                    errMsg = CopiarSOFR6M(wb, dataD1)
                    sofr6mJaProcessado = True
                End If
            Case "cupomusd"
                GravarLog "AVISO", nome, "Curva 'cupomusd' nao implementada -> ignorada"
            Case Else
                GravarLog "AVISO", nome, "Curva '" & nomeCurva & "' nao reconhecida"
        End Select

        If Len(errMsg) > 0 Then
            wb.Close SaveChanges:=False
            AtualizarUmaCalculadora = errMsg
            Exit Function
        End If
ProxCurva:
    Next k

    ' --- FALLBACK Provento: SOFR via FRED ---------------------------------
    ' A aba "Parametros" do Config nao precisa listar 'sofr' / 'daily
    ' compound sofr' para a Provento: se a calculadora for Provento e ela
    ' tiver aba "SOFR" (ou o que estiver em SOFR_ABA_PROVENTO), forcamos
    ' a chamada do FRED aqui. Isso evita o caso comum em que o Config
    ' da Provento so lista 'sofr6m' e a taxa diaria nunca era atualizada.
    If LCase(nome) = "provento" And Not sofrFredJaProcessado Then
        Dim abaProventoSofr As String
        abaProventoSofr = LerParametro("SOFR_ABA_PROVENTO")
        If Len(Trim(abaProventoSofr)) = 0 Then abaProventoSofr = "SOFR"

        If AbaExiste(wb, abaProventoSofr) Then
            Dim errMsgSofrAuto As String
            errMsgSofrAuto = AtualizarDailySOFR_FRED(wb, dataD1)
            If Len(errMsgSofrAuto) > 0 Then
                GravarLog "AVISO", nome, "Fallback SOFR FRED falhou: " & errMsgSofrAuto
                ' Nao aborta a calculadora por causa disso: a Provento ja
                ' processou Fator Pre / CDI / etc, melhor seguir e marcar
                ' o erro no log do que perder o trabalho ja feito.
            Else
                GravarLog "INFO", nome, _
                    "Fallback SOFR FRED executado (curva ausente no Config; aba='" & _
                    abaProventoSofr & "')"
            End If
        Else
            GravarLog "AVISO", nome, _
                "Fallback SOFR pulado: aba '" & abaProventoSofr & "' nao existe na Provento"
        End If
    End If

    ' --- FALLBACK Rodalux I: SOFR6M via arquivo Bloomberg mensal --------
    ' Mesma logica do fallback Provento: se o Config nao lista 'sofr6m',
    ' a curva nunca seria atualizada. Aqui forcamos a chamada se a
    ' calculadora for Rodalux I E tiver a aba 'SOFR6M na BBG'.
    If LCase(nome) = "rodalux i" And Not sofr6mJaProcessado Then
        If AbaExiste(wb, "SOFR6M na BBG") Then
            Dim errMsgSofr6mAuto As String
            errMsgSofr6mAuto = CopiarSOFR6M(wb, dataD1)
            If Len(errMsgSofr6mAuto) > 0 Then
                GravarLog "AVISO", nome, "Fallback SOFR6M falhou: " & errMsgSofr6mAuto
                ' Nao aborta: outras curvas ja foram processadas.
            Else
                GravarLog "INFO", nome, _
                    "Fallback SOFR6M executado (curva ausente no Config)"
            End If
        Else
            GravarLog "AVISO", nome, _
                "Fallback SOFR6M pulado: aba 'SOFR6M na BBG' nao existe na Rodalux I"
        End If
    End If

    ' Escreve data na celula correta
    wsMTM.Range(celulaData).Value = dataD1
    wsMTM.Range(celulaData).NumberFormat = "dd/mm/yyyy"
    GravarLog "INFO", nome, "Data gravada em " & nomeAbaMTM & "!" & celulaData & _
              " = " & Format(dataD1, "dd/mm/yyyy")

    Application.CalculateFull
    Dim checkValor As Variant
    If Len(celulaCheck) > 0 Then
        checkValor = wsMTM.Range(celulaCheck).Value
    Else
        checkValor = "Check nao localizado"
    End If

    wb.Save
    wb.Close SaveChanges:=False

    If UCase(CStr(checkValor)) = "OK" Then
        AtualizarUmaCalculadora = "OK | Check=OK | Data=" & celulaData & " | " & Format(Now, "dd/mm hh:mm")
    Else
        AtualizarUmaCalculadora = "AVISO: Check=" & CStr(checkValor) & " | Data=" & celulaData
    End If
    Exit Function

TratarErro:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    AtualizarUmaCalculadora = "ERRO: " & Err.Description
End Function

'===========================================================================
' Detecta automaticamente a celula da data-base em MTM & PMT
'
' Prioridade:
'   1. Override manual (col D da Config) - ex: "M6"
'   2. Formula da celula Check: AND(N6=... -> N6
'   3. Fallback: "N6"
'===========================================================================
Public Function DetectarCelulaData(ws As Worksheet, override As String) As String
    Dim ov As String: ov = Trim(UCase(override))
    If Len(ov) > 0 And ov <> "-" Then
        DetectarCelulaData = ov: Exit Function
    End If

    Dim checkAddr As String: checkAddr = DetectarCelulaCheck(ws)
    If Len(checkAddr) > 0 Then
        Dim formula As String: formula = ws.Range(checkAddr).Formula
        Dim extraida As String: extraida = ExtrairCelulaDeFormula(formula)
        If Len(extraida) > 0 Then
            DetectarCelulaData = extraida: Exit Function
        End If
    End If

    DetectarCelulaData = "N6"   ' fallback
End Function

'===========================================================================
' Detecta a celula Check varrendo a linha 2
' Retorna endereco (ex: "O2", "N2", "P2")
'===========================================================================
Public Function DetectarCelulaCheck(ws As Worksheet) As String
    Dim c As Long
    For c = 1 To 25
        On Error Resume Next
        Dim f As String: f = CStr(ws.Cells(2, c).Formula)
        On Error GoTo 0
        If Left(f, 1) = "=" And _
           InStr(UCase(f), "IF(") > 0 And _
           InStr(UCase(f), """OK""") > 0 Then
            DetectarCelulaCheck = ws.Cells(2, c).Address(False, False)
            Exit Function
        End If
    Next c
    DetectarCelulaCheck = ""
End Function

'===========================================================================
' Extrai a primeira referencia de celula INTERNA de uma formula
'
' Logica: busca AND( ou IF( e pega o que vem antes do primeiro '='
'   =IF(AND(N6='Fator Pre'!A2,... -> "N6"
'   =IF(AND(M6=...                -> "M6"
'   =IF(P6=...                    -> "P6"
'===========================================================================
Private Function ExtrairCelulaDeFormula(formula As String) As String
    Dim fU       As String: fU = UCase(formula)
    Dim posIni   As Integer
    Dim candidata As String
    Dim posEqual As Integer

    ' Busca AND( primeiro, depois IF(
    posIni = InStr(fU, "AND(")
    If posIni > 0 Then
        posIni = posIni + 4
    Else
        posIni = InStr(fU, "IF(")
        If posIni > 0 Then posIni = posIni + 3 Else Exit Function
    End If

    Dim resto As String: resto = Mid(fU, posIni)
    posEqual = InStr(resto, "=")
    If posEqual = 0 Then Exit Function

    candidata = Trim(Left(resto, posEqual - 1))
    candidata = Replace(candidata, "$", "")

    ' Valida: 1-2 letras + 1-2 digitos
    Dim i As Integer
    Dim letras As String: letras = ""
    Dim digitos As String: digitos = ""
    For i = 1 To Len(candidata)
        Dim ch As String: ch = Mid(candidata, i, 1)
        If ch >= "A" And ch <= "Z" Then
            letras = letras & ch
        ElseIf ch >= "0" And ch <= "9" Then
            digitos = digitos & ch
        End If
    Next i

    If Len(letras) >= 1 And Len(letras) <= 2 And _
       Len(digitos) >= 1 And Len(digitos) <= 2 Then
        ExtrairCelulaDeFormula = letras & digitos
    End If
End Function

'===========================================================================
' MACRO UTILITARIA: Escaneia calculadoras e preenche col D com
' a celula de data detectada. Execute uma vez para revisar.
'
' Comportamento do fallback:
'   Se o caminho em col B estiver vazio ou o arquivo nao existir,
'   a macro varre TODAS as subpastas "MTM - *" dentro do PastaMTM_PU
'   e usa a primeira copia encontrada de cada calculadora.
'   Isso permite rodar a macro independente do dia.
'===========================================================================
Public Sub ScanearCelulasData()
    Dim wsConfig As Worksheet
    Dim ultLinha As Long
    Dim i        As Long
    Dim total    As Long

    Set wsConfig = ThisWorkbook.Sheets(SHEET_CONFIG)
    ultLinha = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row

    ' Garante cabecalho na col D
    If Trim(wsConfig.Cells(1, COL_OVERRIDE).Value) = "" Then
        wsConfig.Cells(1, COL_OVERRIDE).Value = "Celula Data (auto/override)"
        wsConfig.Cells(1, COL_OVERRIDE).Font.Bold = True
    End If

    ' Pre-carrega colecao de todos os .xlsx nas pastas MTM (fallback)
    Dim colMTM As New Collection
    CarregarXLSX_TodasPastasMTM colMTM

    For i = 2 To ultLinha
        Dim nome    As String: nome    = Trim(wsConfig.Cells(i, 1).Value)
        Dim caminho As String: caminho = Trim(wsConfig.Cells(i, 2).Value)
        If Len(nome) = 0 Then GoTo Prox

        ' Limpa marcas de execucoes anteriores desta macro
        If InStr(LCase(caminho), "auto-preenchido") > 0 Then caminho = ""

        ' Se caminho invalido: tenta encontrar nas pastas MTM
        If Len(caminho) = 0 Or Dir(caminho) = "" Then
            caminho = BuscarArquivoPorNome(colMTM, nome)
        End If

        If Len(caminho) = 0 Or Dir(caminho) = "" Then
            wsConfig.Cells(i, COL_OVERRIDE).Value = "arquivo nao encontrado"
            GoTo Prox
        End If

        On Error Resume Next
        Dim wb As Workbook
        Set wb = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True)
        On Error GoTo 0
        If wb Is Nothing Then
            wsConfig.Cells(i, COL_OVERRIDE).Value = "nao foi possivel abrir"
            GoTo Prox
        End If

        Dim nomeAbaMTM_S As String: nomeAbaMTM_S = EncontrarAbaMTM(wb)
        If Len(nomeAbaMTM_S) > 0 Then
            Dim dataCell As String
            dataCell = DetectarCelulaData(wb.Sheets(nomeAbaMTM_S), "")
            wsConfig.Cells(i, COL_OVERRIDE).Value = dataCell
            total = total + 1
        Else
            wsConfig.Cells(i, COL_OVERRIDE).Value = "aba MTM nao encontrada (" & ListarAbas(wb) & ")"
        End If
        wb.Close SaveChanges:=False
Prox:
    Next i

    MsgBox "Escaneamento concluido: " & total & " calculadoras." & vbCrLf & vbCrLf & _
           "Revise a coluna D da aba Config." & vbCrLf & _
           "Para forcar uma celula especifica, edite o valor manualmente." & vbCrLf & _
           "Para usar auto-deteccao, deixe a celula em branco.", _
           vbInformation, "Celulas Data Detectadas"
End Sub

'===========================================================================
' Carrega em uma Collection todos os .xlsx de calculadoras encontrados
' em QUALQUER subpasta "MTM - *" dentro de PastaMTM_PU.
' Usada como fallback pelo ScanearCelulasData.
'===========================================================================
Private Sub CarregarXLSX_TodasPastasMTM(col As Collection)
    Dim pastaMTM As String: pastaMTM = LerParametro("PastaMTM_PU")
    If Len(pastaMTM) = 0 Then Exit Sub

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pastaMTM) Then Exit Sub

    ' Estrutura: pastaMTM\MTM & PU (YYYY)\MTM & PU - MMM.YY\MTM - DD.MM.YY\
    Dim pasta  As Object: Set pasta  = fso.GetFolder(pastaMTM)
    Dim anoFo  As Object
    Dim mesFo  As Object
    Dim diaFo  As Object
    For Each anoFo In pasta.SubFolders
        If InStr(LCase(anoFo.Name), "mtm & pu (") > 0 Then
            For Each mesFo In anoFo.SubFolders
                If InStr(LCase(mesFo.Name), "mtm & pu - ") > 0 Then
                    For Each diaFo In mesFo.SubFolders
                        If InStr(LCase(diaFo.Name), "mtm - ") > 0 Then
                            ColetarXLSX diaFo.Path, col, fso
                        End If
                    Next diaFo
                End If
            Next mesFo
        End If
    Next anoFo
End Sub

'===========================================================================
' Copia aba do Curvas_Separadas para a calculadora
'===========================================================================
Private Function CopiarCurvaDeCSS(wbCSS As Workbook, nomeAbaCSS As String, _
                                   wbCalc As Workbook, nomeAbaCalc As String) As String
    If Not AbaExiste(wbCSS, nomeAbaCSS) Then
        CopiarCurvaDeCSS = "ERRO: aba '" & nomeAbaCSS & "' nao encontrada no Curvas_Separadas"
        Exit Function
    End If
    If Not AbaExiste(wbCalc, nomeAbaCalc) Then
        CopiarCurvaDeCSS = "ERRO: aba '" & nomeAbaCalc & "' nao existe na calculadora"
        Exit Function
    End If

    Dim wsOrig As Worksheet: Set wsOrig = wbCSS.Sheets(nomeAbaCSS)
    Dim wsDest As Worksheet: Set wsDest = wbCalc.Sheets(nomeAbaCalc)
    Dim ultLinha As Long: ultLinha = wsOrig.Cells(wsOrig.Rows.Count, 1).End(xlUp).Row

    ' --- DETECCAO DE ultCol ----------------------------------------------
    ' Usar linha 2 (primeira de dados) em vez de linha 1 (cabecalho).
    ' Bug anterior: cabecalho de A1:D1 + colunas vazias E:H formatadas no
    ' CSS faziam End(xlToLeft) na linha 1 pegar ate H, e a limpeza acabava
    ' destruindo as colunas F:H da Provento (YieldDaily, Rentabilidade
    ' Diaria, Accrued Factor). Linha 2 so tem dados reais (A:D).
    Dim ultCol As Long
    If ultLinha >= 2 Then
        ultCol = wsOrig.Cells(2, wsOrig.Columns.Count).End(xlToLeft).Column
    Else
        ultCol = wsOrig.Cells(1, wsOrig.Columns.Count).End(xlToLeft).Column
    End If
    If ultCol < 1 Then ultCol = 4   ' fallback Fator Pre: A:D

    If ultLinha < 1 Then
        CopiarCurvaDeCSS = "ERRO: aba '" & nomeAbaCSS & "' esta vazia"
        Exit Function
    End If

    ' --- LIMPEZA SEGURA --------------------------------------------------
    ' Limpa apenas o retangulo que sera sobrescrito (A1:<ultCol,ultLinhaDest>),
    ' preservando colunas a direita (F:I da Provento, etc).
    Dim ultLinhaDest As Long
    ultLinhaDest = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    If ultLinhaDest < ultLinha Then ultLinhaDest = ultLinha
    wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(ultLinhaDest, ultCol)).ClearContents

    wsOrig.Range(wsOrig.Cells(1, 1), wsOrig.Cells(ultLinha, ultCol)).Copy
    wsDest.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsDest.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    CopiarCurvaDeCSS = ""
End Function

'===========================================================================
' Atualiza historico CDI
'===========================================================================
Private Function AtualizarCDIHistorico(wb As Workbook, dataD1 As Date, _
                                        taxaCDI As Double) As String
    If Not AbaExiste(wb, ABA_CDI) Then
        AtualizarCDIHistorico = "ERRO: aba '" & ABA_CDI & "' nao existe"
        Exit Function
    End If

    Dim ws As Worksheet: Set ws = wb.Sheets(ABA_CDI)
    Dim ultimaLinhaComCDI As Long: ultimaLinhaComCDI = 1
    Dim ultimaData        As Date
    Dim ultimoFatorAcum   As Double
    Dim ult As Long: ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Dim r   As Long

    For r = 2 To ult
        Dim cellB As Variant: cellB = ws.Cells(r, 2).Value
        If IsNumeric(cellB) And CDbl(cellB) > 0 Then
            ultimaLinhaComCDI = r
            ultimaData       = CDate(ws.Cells(r, 1).Value)
            ultimoFatorAcum  = CDbl(ws.Cells(r, 4).Value)
        End If
    Next r

    If ultimaLinhaComCDI = 1 Then
        AtualizarCDIHistorico = "ERRO: nenhum CDI em '" & ABA_CDI & "'"
        Exit Function
    End If
    If ultimaData = dataD1 Then AtualizarCDIHistorico = "": Exit Function

    Dim fatorDiario As Double: fatorDiario = (1 + taxaCDI / 100) ^ (1 / 252)
    Dim linhaD1 As Long: linhaD1 = 0
    For r = ultimaLinhaComCDI + 1 To ult
        If IsDate(ws.Cells(r, 1).Value) Then
            If CDate(ws.Cells(r, 1).Value) = dataD1 Then linhaD1 = r: Exit For
        End If
    Next r
    If linhaD1 = 0 Then
        linhaD1 = ultimaLinhaComCDI + 1
        ws.Cells(linhaD1, 1).Value = dataD1
        ws.Cells(linhaD1, 1).NumberFormat = "dd/mm/yyyy"
    End If

    ' --- PREENCHIMENTO POR DROPDOWN DA LINHA ANTERIOR ------------------
    ' Estrategia: copia A:<ultColAnt> da ultimaLinhaComCDI para a linha
    ' alvo via PasteSpecial xlPasteAll. Isso preserva formulas, formatos
    ' e bordas, e o Excel ajusta as referencias relativas automaticamente
    ' � equivale exatamente a um "dropdown" (arrastar a alca de
    ' preenchimento) da celula de cima.
    '
    ' Em seguida sobrescreve APENAS A (data) e B (taxa CDI). NAO mexemos
    ' em C, D, E, F: cada calculadora tem um padrao proprio nessas
    ' colunas, e o dropdown ja produz o resultado correto.
    '
    ' Padroes que o dropdown reproduz fielmente:
    '
    '   Nortis / Agroverde / Rodalux I / Rodalux II (A:D)
    '     C<r>: =(1+B<r>/100)^(1/252)
    '     D<r>: =D<r-1>*C<r-1>        <- "dropdown puro" da linha de cima
    '
    '   Provento (A:F)
    '     C<r>: =(1+B<r>/100)^(1/252)
    '     D<r>: =C<r>-1
    '     E<r>: =D<r>*$H$1
    '     F<r>: =F<r-1>*(1+E<r-1>)
    '
    ' Importante: a versao anterior reescrevia D<r> como =D<r-1>*C<r>,
    ' o que quebrava o padrao das calculadoras (que esperam *C<r-1>).
    ' Por isso removemos qualquer override de formula apos o paste.
    ' ------------------------------------------------------------------
    Dim ultColAnt As Long
    ultColAnt = ws.Cells(ultimaLinhaComCDI, ws.Columns.Count).End(xlToLeft).Column
    If ultColAnt < 4 Then ultColAnt = 4

    For r = ultimaLinhaComCDI + 1 To linhaD1
        Dim cdiAtual As Variant: cdiAtual = ws.Cells(r, 2).Value
        If Not IsNumeric(cdiAtual) Or CDbl(cdiAtual) = 0 Then
            ' 1. Dropdown: replica linha anterior inteira (A:<ultColAnt>)
            '    O Excel ajusta as referencias relativas automaticamente.
            ws.Range(ws.Cells(ultimaLinhaComCDI, 1), _
                     ws.Cells(ultimaLinhaComCDI, ultColAnt)).Copy
            ws.Cells(r, 1).PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False

            ' 2. Sobrescreve apenas A (data) e B (taxa CDI). C/D/E/F sao
            '    deixadas exatamente como o dropdown produziu.
            ws.Cells(r, 1).Value = dataD1
            ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy"
            ws.Cells(r, 2).Value = taxaCDI
        End If
    Next r

    GravarLog "INFO", "CDI", "Atualizado ate " & Format(dataD1, "dd/mm/yyyy") & _
              " (B=" & taxaCDI & "%, dropdown A:" & _
              ColLetra(ultColAnt) & " da linha anterior, sem override de formulas)"
    AtualizarCDIHistorico = ""
End Function

' Helper: numero da coluna -> letra (1=A, 2=B, ..., 26=Z, 27=AA)
Private Function ColLetra(n As Long) As String
    Dim s As String: s = ""
    Do While n > 0
        Dim m As Long: m = ((n - 1) Mod 26)
        s = Chr(65 + m) & s
        n = (n - 1) \ 26
    Loop
    ColLetra = s
End Function

'===========================================================================
' Detecta o nome real da aba MTM (PU, PMT ou variantes)
'
' Estrategia:
'   1. Tenta nomes conhecidos: "MTM & PU", "MTM & PMT", "MTM&PU", "MTM&PMT"
'   2. Tenta qualquer aba que comece com "MTM" e contenha "PU" ou "PMT"
'   3. Tenta qualquer aba que comece com "MTM "
'   4. Retorna "" se nada bate
'
' Compara case-insensitive e normaliza espacos.
'===========================================================================
Public Function EncontrarAbaMTM(wb As Workbook) As String
    ' 1. Tentativas exatas (cobre as variantes mais comuns)
    Dim candidatos As Variant
    candidatos = Array("MTM & PU", "MTM & PMT", "MTM&PU", "MTM&PMT", _
                       "MTM - PU", "MTM PU", "MTM PMT", "MTM")
    Dim c As Variant
    Dim ws As Worksheet
    For Each c In candidatos
        For Each ws In wb.Worksheets
            If LCase(Trim(ws.Name)) = LCase(CStr(c)) Then
                EncontrarAbaMTM = ws.Name
                Exit Function
            End If
        Next ws
    Next c

    ' 2. Fuzzy: aba que comeca com "MTM" e tem "PU" ou "PMT"
    For Each ws In wb.Worksheets
        Dim n As String: n = UCase(Trim(ws.Name))
        If Left(n, 3) = "MTM" And (InStr(n, "PU") > 0 Or InStr(n, "PMT") > 0) Then
            EncontrarAbaMTM = ws.Name
            Exit Function
        End If
    Next ws

    ' 3. Fallback: primeira aba que comece com "MTM "
    For Each ws In wb.Worksheets
        If Left(UCase(Trim(ws.Name)), 4) = "MTM " Then
            EncontrarAbaMTM = ws.Name
            Exit Function
        End If
    Next ws

    EncontrarAbaMTM = ""
End Function

'===========================================================================
' Lista todas as abas de um workbook (para diagnostico)
'===========================================================================
Public Function ListarAbas(wb As Workbook) As String
    Dim ws  As Worksheet
    Dim res As String
    For Each ws In wb.Worksheets
        res = res & ws.Name & " | "
    Next ws
    If Len(res) > 200 Then res = Left(res, 200) & "..."
    ListarAbas = res
End Function
