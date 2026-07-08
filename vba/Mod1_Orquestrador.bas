Attribute VB_Name = "Mod1_Orquestrador"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 1: Orquestrador (v4.1)
'
' CORRECAO v4.1:
'   ExecutarCopiarPastaMTM agora trata DOIS cenarios:
'
'   CENARIO A - Pasta D-1 NAO existe:
'     Copia a pasta D-2 inteira para D-1 (comportamento anterior)
'
'   CENARIO B - Pasta D-1 JA existe (criada pelo BaixaArquivosINETX):
'     NAO tenta copiar a pasta toda (falharia pois D-1 existe)
'     Em vez disso chama CopiarCalculadorasDeD2ParaD1 que copia
'     apenas os .xlsx das calculadoras arquivo por arquivo,
'     sem tocar no Curvas_Separadas que ja esta em D-1.
'
'   Em ambos os cenarios, AutoPreencherCaminhosConfig e chamado
'   para descobrir os caminhos e preencher a col B da Config.
'===========================================================================
Option Explicit

Public Const SHEET_CONFIG     As String = "Config"
Public Const SHEET_FONTES     As String = "Fontes"
Public Const SHEET_PARAMETROS As String = "Parametros"
Public Const SHEET_FERIADOS   As String = "Feriados_Conf"
Public Const SHEET_LOG        As String = "Log"

Public g_dataD1     As Date
Public g_dataD2     As Date
Public g_taxaCDI    As Double
Public g_caminhoCSS As String
Public g_pastaD1    As String
Public g_wsLog      As Worksheet

' --- Log fisico CSV (espelho da aba Log, gravado na pasta MTM do dia) -----
' g_arqLogCSV: path absoluto do CSV (vazio enquanto a pasta MTM do dia
'              ainda nao foi descoberta -> as linhas vao para o buffer).
' g_bufLogCSV: buffer de linhas CSV ate g_arqLogCSV ser definido.
' g_inicioRun: data/hora capturada em FluxoCompleto; usada no nome do CSV.
Private g_arqLogCSV As String
Private g_bufLogCSV As Collection
Private g_inicioRun As Date

'===========================================================================
' ENTRY POINT
'===========================================================================
Public Sub FluxoCompleto()
    On Error GoTo TratarErro

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Reset do log CSV (nao herda estado de runs anteriores).
    g_inicioRun = Now
    g_arqLogCSV = ""
    Set g_bufLogCSV = Nothing

    Set g_wsLog = ObterOuCriarLog()
    GravarLog "INICIO", "FluxoCompleto", "Iniciando fluxo de atualizacao"

    ' Garante que a query Power Query "MTM & PU" aponte para o caminho do
    ' usuario atual (substitui CaminhoBase = "..." pelo valor de PastaMTM_PU
    ' resolvido via %USERPROFILE%). Permite que toda a equipe use a mesma
    ' planilha sem copiar e colar para o proprio diretorio.
    AtualizarCaminhoPQ

    ' 1. Dias uteis
    Application.StatusBar = "[1/5] Calculando dias uteis..."
    Dim feriados() As Date
    feriados = CarregarFeriados()

    ' Le a data-base (D-1) da celula informada na aba Parametros (chave 'DataBase').
    ' Essa celula normalmente contem =DIATRABALHO(HOJE();-1;Feriados!A:A), ou seja,
    ' o valor JA e o D-1. Por isso D-1 = valor da celula e D-2 = 1 dia util antes.
    ' Se a celula estiver vazia/invalida, usa a data de hoje (comportamento antigo).
    Dim dataBase As Date
    If ObterDataBaseManual(dataBase) Then
        g_dataD1 = dataBase
        g_dataD2 = DiasUteisVoltar(g_dataD1, 1, feriados)
        GravarLog "INFO", "DiasUteis", "Data-base lida de Parametros!DataBase: " & _
                  Format(g_dataD1, "dd/mm/yyyy")
    Else
        g_dataD1 = DiasUteisVoltar(Date, 1, feriados)
        g_dataD2 = DiasUteisVoltar(Date, 2, feriados)
    End If
    GravarLog "INFO", "DiasUteis", "D-1=" & Format(g_dataD1, "dd/mm/yyyy") & _
              " | D-2=" & Format(g_dataD2, "dd/mm/yyyy")

    ' 2. Taxa CDI
    Application.ScreenUpdating = True
    If Not PedirTaxaCDI() Then GoTo Cancelado
    Application.ScreenUpdating = False

    ' 3. Copia calculadoras D-2->D-1 e mapeia caminhos
    Application.StatusBar = "[2/5] Copiando calculadoras D-2 para D-1..."
    ExecutarCopiarPastaMTM

    ' Materializa o CSV de log assim que g_pastaD1 esta resolvido.
    ' Flusha o buffer (linhas anteriores do INICIO/INFO ja gravadas na aba Log).
    MaterializarArquivoLogCSV

    ' 4. Baixa curvas INETX via WinSCP/SFTP e gera Curvas_Separadas
    '    (substitui a planilha externa BaixaArquivosINETX.xlsm).
    '    Continua mesmo se falhar: ExecutarINETX validara se o arquivo
    '    realmente foi gerado e abortara com mensagem clara se nao.
    Application.StatusBar = "[3/5] Baixando curvas INETX..."
    BaixarCurvasINETX

    ' 5. Localiza e valida Curvas_Separadas
    Application.StatusBar = "[4/5] Verificando Curvas_Separadas..."
    If Not ExecutarINETX() Then GoTo Finalizar

    ' 6. Atualiza calculadoras
    Application.StatusBar = "[5/6] Atualizando calculadoras..."
    ExecutarCalculadoras

    GravarLog "FIM", "FluxoCompleto", "Concluido com sucesso"
    Application.StatusBar = "[6/6] Concluido!"
    Application.ScreenUpdating = True

    MsgBox "Fluxo concluido!" & vbCrLf & vbCrLf & _
           "Data-base (D-1): " & Format(g_dataD1, "dd/mm/yyyy") & vbCrLf & _
           "Taxa CDI: " & g_taxaCDI & "% a.a." & vbCrLf & vbCrLf & _
           "Veja a aba 'Log' para detalhes.", vbInformation, "Atualizador SWAPs"
    GoTo Finalizar

Cancelado:
    GravarLog "CANCELADO", "FluxoCompleto", "Usuario cancelou"
    MsgBox "Fluxo cancelado.", vbExclamation, "Cancelado"

TratarErro:
    Dim errMsg As String
    errMsg = "Erro " & Err.Number & ": " & Err.Description
    GravarLog "ERRO FATAL", "FluxoCompleto", errMsg
    Application.ScreenUpdating = True
    MsgBox "Erro inesperado:" & vbCrLf & errMsg, vbCritical, "Erro"

Finalizar:
    ' Se o CSV nunca foi materializado (falha antes da pasta MTM ser
    ' identificada) grava o buffer numa pasta de fallback (Desktop).
    FlushFinalLogCSV

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'===========================================================================
' Pede taxa CDI com validacao
'===========================================================================
Private Function PedirTaxaCDI() As Boolean
    Dim resp As Variant
    Dim tentativas As Integer

TentarNovamente:
    resp = Application.InputBox( _
        "Data-base (D-1): " & Format(g_dataD1, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
        "Informe a taxa CDI (% a.a.):" & vbCrLf & _
        "Use virgula como separador (ex: 14,40)", _
        "Taxa CDI", "14,40", Type:=1)

    If resp = False Then PedirTaxaCDI = False: Exit Function

    If resp <= 0 Or resp > 100 Then
        tentativas = tentativas + 1
        If tentativas >= 3 Then PedirTaxaCDI = False: Exit Function
        MsgBox "Taxa invalida. Digite entre 0 e 100 usando virgula (ex: 14,40).", vbExclamation
        GoTo TentarNovamente
    End If

    g_taxaCDI = CDbl(resp)
    GravarLog "INFO", "CDI", "Taxa informada: " & g_taxaCDI & "% a.a."
    PedirTaxaCDI = True
End Function

'===========================================================================
' Helper: Retorna o caminho base do usuario (C:\Users\nome\)
'===========================================================================
Public Function ObterCaminhoBaseUsuario() As String
    Dim perfil As String
    perfil = Environ("USERPROFILE")
    If Right(perfil, 1) <> "\" Then perfil = perfil & "\"
    ObterCaminhoBaseUsuario = perfil
End Function

'===========================================================================
' Expande tokens dependentes do usuario em caminhos lidos da aba Parametros.
'   %USERPROFILE%  -> C:\Users\<login> (Environ("USERPROFILE"))
'   %USERNAME%     -> login do usuario
'   %ONEDRIVE%     -> Environ("OneDrive") (vazio se nao houver)
'   %DESKTOP%      -> %USERPROFILE%\Desktop
' Permite manter a planilha unica para o time inteiro sem caminhos fixos
' por usuario. Strings sem token sao devolvidas inalteradas.
'===========================================================================
Public Function ExpandirCaminho(s As String) As String
    Dim r As String
    r = s
    If Len(r) = 0 Then ExpandirCaminho = "": Exit Function
    Dim perfil As String: perfil = Environ("USERPROFILE")
    If Right(perfil, 1) = "\" Then perfil = Left(perfil, Len(perfil) - 1)
    Dim onedrive As String: onedrive = Environ("OneDrive")
    If Len(onedrive) > 0 And Right(onedrive, 1) = "\" Then onedrive = Left(onedrive, Len(onedrive) - 1)
    r = Replace(r, "%USERPROFILE%", perfil)
    r = Replace(r, "%UserProfile%", perfil)
    r = Replace(r, "%userprofile%", perfil)
    r = Replace(r, "%USERNAME%", Environ("USERNAME"))
    r = Replace(r, "%DESKTOP%", perfil & "\Desktop")
    r = Replace(r, "%ONEDRIVE%", onedrive)
    ' Prefixo ~\ (estilo unix) -> USERPROFILE\
    If Left(r, 2) = "~\" Then r = perfil & "\" & Mid(r, 3)
    ExpandirCaminho = r
End Function

'===========================================================================
' Reescreve a formula M da query "MTM & PU" trocando o CaminhoBase pelo
' valor expandido de PastaMTM_PU em tempo de execucao. Isso permite que a
' planilha funcione para qualquer usuario do time sem editar a M.
' Silencioso e idempotente: se a query nao existir ou nao houver linha
' CaminhoBase, nao faz nada.
'===========================================================================
Public Sub AtualizarCaminhoPQ()
    On Error GoTo Sair
    Dim q As Object 'WorkbookQuery
    Dim caminho As String
    caminho = ExpandirCaminho(LerParametro("PastaMTM_PU"))
    If Len(caminho) = 0 Then Exit Sub
    If Right(caminho, 1) <> "\" Then caminho = caminho & "\"

    Set q = Nothing
    On Error Resume Next
    Set q = ThisWorkbook.Queries("MTM & PU")
    On Error GoTo Sair
    If q Is Nothing Then Exit Sub

    Dim novaFormula As String
    novaFormula = q.Formula
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.Pattern = "CaminhoBase\s*=\s*""[^""]*"""
    If Not re.Test(novaFormula) Then Exit Sub
    novaFormula = re.Replace(novaFormula, "CaminhoBase = """ & caminho & """")
    If novaFormula <> q.Formula Then
        q.Formula = novaFormula
        GravarLog "OK", "PQ", "CaminhoBase atualizado para: " & caminho
    End If
Sair:
End Sub

'===========================================================================
' Etapa 3: Copiar calculadoras de D-2 para D-1 + mapear caminhos
'===========================================================================
Private Sub ExecutarCopiarPastaMTM()
    On Error GoTo ErrHandler

    Dim pastaMTM As String
    pastaMTM = LerParametro("PastaMTM_PU")

    If Len(pastaMTM) = 0 Then
        ' Fallback dinamico por usuario
        pastaMTM = ObterCaminhoBaseUsuario() & "Empresa\Bloomberg - Documentos\Sakura\MTM & PU"
        GravarLog "AVISO", "MTM", "Parametro 'PastaMTM_PU' nao encontrado. Usando fallback: " & pastaMTM
    End If

    ' Monta bases mensais para D-2 e D-1
    Dim baseMesD2 As String: baseMesD2 = ObterPastaBaseComMes(pastaMTM, g_dataD2)
    Dim baseMesD1 As String: baseMesD1 = ObterPastaBaseComMes(pastaMTM, g_dataD1)

    ' Localiza D-2
    Dim pastaD2 As String
    pastaD2 = EncontrarPastaData(baseMesD2, g_dataD2)

    If Len(pastaD2) = 0 Then
        GravarLog "AVISO", "MTM", "Pasta D-2 nao encontrada para " & Format(g_dataD2, "dd/mm/yyyy")
        MsgBox "Pasta MTM D-2 (" & Format(g_dataD2, "dd/mm/yyyy") & ") nao encontrada." & vbCrLf & _
               "Verifique o parametro 'PastaMTM_PU' e confirme que ele aponta para a raiz da pasta MTM." & vbCrLf & _
               "Ele deve ser o caminho para a pasta que contem 'MTM & PU (YYYY)'." & vbCrLf & vbCrLf & _
               "O fluxo continuara tentando mapear o que estiver em D-1.", _
               vbExclamation
        GoTo TentarMapear
    End If

    ' Monta caminho esperado para D-1
    g_pastaD1 = baseMesD1 & "\" & NomePastaData(g_dataD1)

    ' Garante que a pasta D-1 exista (vazia se for o caso). Depois SEMPRE
    ' usa a versao filtrada que pula Curvas_Separadas/Config/etc — o
    ' Mod8.BaixarCurvasINETX vai gerar o Curvas_Separadas correto do dia
    ' logo apos esta etapa. Se a versao anterior caia em "CENARIO A"
    ' (CopiarPasta bruto), copiava o Curvas_Separadas antigo da pasta D-2.
    If Dir(g_pastaD1, vbDirectory) = "" Then
        GravarLog "INFO", "MTM", "Pasta D-1 nao existe. Criando vazia em " & g_pastaD1
        CriarHierarquiaPastas baseMesD1
        On Error Resume Next
        MkDir g_pastaD1
        On Error GoTo ErrHandler
    Else
        GravarLog "INFO", "MTM", "Pasta D-1 ja existe."
    End If

    ' Copia somente as calculadoras (filtra Curvas_Separadas, Config, etc.)
    GravarLog "INFO", "MTM", "Copiando calculadoras de D-2 para D-1..."
    CopiarCalculadorasDeD2ParaD1 pastaD2, g_pastaD1

    ' Remove qualquer Curvas_Separadas obsoleto (de outra data) que tenha
    ' sobrado em D-1 de runs antigos com cenario A. O do dia atual sera
    ' regerado pelo BaixarCurvasINETX logo a seguir.
    LimparCurvasObsoletas g_pastaD1, g_dataD1

TentarMapear:
    ' Tenta encontrar D-1 caso g_pastaD1 ainda esteja vazio
    If Len(g_pastaD1) = 0 Then
        g_pastaD1 = EncontrarPastaData(baseMesD1, g_dataD1)
        If Len(g_pastaD1) = 0 Then
            GravarLog "AVISO", "MTM", "Pasta D-1 nao encontrada para auto-mapeamento"
        End If
    End If

    ' Auto-preenche coluna B da Config com os caminhos encontrados em D-1
    If Len(g_pastaD1) > 0 Then
        AutoPreencherCaminhosConfig g_pastaD1
    End If
    Exit Sub

ErrHandler:
    GravarLog "ERRO", "MTM", "Erro na copia MTM: " & Err.Description
    MsgBox "Erro ao copiar pasta MTM:" & vbCrLf & Err.Description, vbExclamation
End Sub

'===========================================================================
' Etapa 4: Localiza e valida Curvas_Separadas
'===========================================================================
Private Function ExecutarINETX() As Boolean
    On Error GoTo ErrHandler

    If Not ValidarCurvasSeparadas(g_dataD1) Then
        ExecutarINETX = False: Exit Function
    End If

    g_caminhoCSS = LocalizarCurvasSeparadas(g_dataD1)

    If Not ValidarConteudoCurvas(g_caminhoCSS) Then
        ExecutarINETX = False: Exit Function
    End If

    GravarLog "OK", "INETX", "Curvas_Separadas localizado: " & g_caminhoCSS
    ExecutarINETX = True
    Exit Function

ErrHandler:
    GravarLog "ERRO", "INETX", Err.Description
    MsgBox "Erro ao localizar Curvas_Separadas: " & Err.Description, vbCritical
    ExecutarINETX = False
End Function

'===========================================================================
' Etapa 5: Atualiza calculadoras
'===========================================================================
Private Sub ExecutarCalculadoras()
    On Error GoTo ErrHandler
    If Len(g_caminhoCSS) = 0 Or Dir(g_caminhoCSS) = "" Then
        GravarLog "ERRO", "Calculadoras", "Curvas_Separadas nao encontrado"
        MsgBox "Arquivo de curvas nao encontrado. A etapa INETX pode ter falhado.", vbExclamation
        Exit Sub
    End If
    AtualizarCalculadoras g_dataD1, g_taxaCDI, g_caminhoCSS
    Exit Sub

ErrHandler:
    GravarLog "ERRO", "Calculadoras", Err.Description
    MsgBox "Erro ao atualizar calculadoras: " & Err.Description, vbExclamation
End Sub

'===========================================================================
' HELPERS GLOBAIS
'===========================================================================

'---------------------------------------------------------------------------
' Le a data-base (D-1) da aba Parametros, linha cuja coluna A = "DataBase".
' O VALOR (coluna B) ja deve ser o D-1, normalmente vindo da formula
' =DIATRABALHO(HOJE();-1;Feriados!A:A).
'
' Retorna True e preenche 'resultado' quando a celula contem uma data valida.
' Retorna False quando a aba/chave nao existe ou a celula esta vazia/invalida
' (nesse caso o chamador deve cair no comportamento antigo baseado em Date).
'---------------------------------------------------------------------------
Public Function ObterDataBaseManual(ByRef resultado As Date) As Boolean
    Dim ws As Worksheet
    Dim i As Long, ult As Long
    Dim v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_PARAMETROS)
    On Error GoTo 0
    If ws Is Nothing Then ObterDataBaseManual = False: Exit Function

    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ult
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = "database" Then
            v = ws.Cells(i, 2).Value
            If IsDate(v) Then
                resultado = CDate(v)
                ObterDataBaseManual = True
            End If
            Exit Function
        End If
    Next i

    ObterDataBaseManual = False
End Function

Public Function LerParametro(chave As String) As String
    Dim ws As Worksheet
    Dim i As Long, ult As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_PARAMETROS)
    On Error GoTo 0
    If ws Is Nothing Then LerParametro = "": Exit Function

    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ult
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = LCase(Trim(chave)) Then
            LerParametro = ExpandirCaminho(Trim(CStr(ws.Cells(i, 2).Value)))
            Exit Function
        End If
    Next i
    LerParametro = ""
End Function

Public Function ObterOuCriarLog() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_LOG
        ws.Range("A1:E1").Value = Array("Timestamp", "Status", "Contexto", "Detalhe", "Analista")
        ws.Range("A1:E1").Font.Bold = True
        ws.Columns("A").ColumnWidth = 20
        ws.Columns("B").ColumnWidth = 12
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 90
        ws.Columns("E").ColumnWidth = 18
    End If
    Set ObterOuCriarLog = ws
End Function

Public Sub GravarLog(status As String, contexto As String, detalhe As String)
    If g_wsLog Is Nothing Then Set g_wsLog = ObterOuCriarLog()
    Dim ult As Long
    Dim agora As Date: agora = Now
    Dim analista As String: analista = Environ$("USERNAME")
    ult = g_wsLog.Cells(g_wsLog.Rows.Count, "A").End(xlUp).Row + 1
    With g_wsLog
        .Cells(ult, 1).Value = agora
        .Cells(ult, 1).NumberFormat = "dd/mm/yyyy hh:mm:ss"
        .Cells(ult, 2).Value = status
        .Cells(ult, 3).Value = contexto
        .Cells(ult, 4).Value = detalhe
        .Cells(ult, 5).Value = analista
        Select Case UCase(status)
            Case "ERRO", "ERRO FATAL": .Cells(ult, 2).Interior.Color = RGB(255, 200, 200)
            Case "OK":                 .Cells(ult, 2).Interior.Color = RGB(200, 255, 200)
            Case "AVISO":              .Cells(ult, 2).Interior.Color = RGB(255, 255, 200)
            Case "INICIO", "FIM":      .Cells(ult, 2).Interior.Color = RGB(200, 200, 255)
        End Select
    End With

    ' Espelho no CSV fisico (colateral; falha silenciosa nao quebra fluxo).
    On Error Resume Next
    Dim linhaCSV As String
    linhaCSV = MontarLinhaCSV(agora, status, analista, contexto, detalhe)
    If Len(g_arqLogCSV) = 0 Then
        If g_bufLogCSV Is Nothing Then Set g_bufLogCSV = New Collection
        g_bufLogCSV.Add linhaCSV
    Else
        AppendCSVLine g_arqLogCSV, linhaCSV
    End If
    On Error GoTo 0
End Sub

Public Function AbaExiste(wb As Workbook, nome As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = nome Then AbaExiste = True: Exit Function
    Next ws
End Function

'===========================================================================
' LOG CSV FISICO - helpers privados
'===========================================================================

' Constroi uma linha CSV (delimitador ';') escapando segundo RFC 4180:
'   - aspas duplas internas viram duas aspas duplas
'   - se o valor contiver ';', '"', CR ou LF, envolve em aspas
Private Function MontarLinhaCSV(agora As Date, status As String, _
                                analista As String, contexto As String, _
                                detalhe As String) As String
    Dim ts As String
    ts = Format$(agora, "yyyy-mm-dd hh:nn:ss")
    MontarLinhaCSV = EscaparCSV(ts) & ";" & _
                     EscaparCSV(status) & ";" & _
                     EscaparCSV(analista) & ";" & _
                     EscaparCSV(contexto) & ";" & _
                     EscaparCSV(detalhe)
End Function

Private Function EscaparCSV(s As String) As String
    Dim r As String: r = CStr(s)
    Dim precisaAspas As Boolean
    precisaAspas = (InStr(r, ";") > 0) Or (InStr(r, """") > 0) Or _
                   (InStr(r, vbCr) > 0) Or (InStr(r, vbLf) > 0)
    If InStr(r, """") > 0 Then r = Replace(r, """", """""")
    If precisaAspas Then r = """" & r & """"
    EscaparCSV = r
End Function

' Nome do CSV: "LOG_FluxoCompleto_YYYYMMDD_HHMMSS_<USERNAME>.csv".
' Usa g_inicioRun (fixo durante todo o run) para o stamp data/hora.
Private Function NomeArquivoLogCSV() As String
    Dim stamp As String, user As String
    stamp = Format$(g_inicioRun, "yyyymmdd_hhnnss")
    user = Environ$("USERNAME")
    If Len(user) = 0 Then user = "anon"
    NomeArquivoLogCSV = "LOG_FluxoCompleto_" & stamp & "_" & user & ".csv"
End Function

' Resolve a pasta do CSV:
'   1) g_pastaD1 (pasta MTM do dia, preferida)
'   2) Parametros!PastaOutputCurvas (geralmente Desktop)
'   3) Fallback: %USERPROFILE%\Desktop
Private Function ResolverPastaLogCSV() As String
    Dim p As String
    If Len(g_pastaD1) > 0 Then
        If Dir(g_pastaD1, vbDirectory) <> "" Then
            ResolverPastaLogCSV = AdicionarBarra(g_pastaD1)
            Exit Function
        End If
    End If

    On Error Resume Next
    p = LerParametro("PastaOutputCurvas")
    On Error GoTo 0
    If Len(p) > 0 Then
        If Dir(p, vbDirectory) <> "" Then
            ResolverPastaLogCSV = AdicionarBarra(p)
            Exit Function
        End If
    End If

    ResolverPastaLogCSV = AdicionarBarra(ObterCaminhoBaseUsuario() & "Desktop")
End Function

Private Function AdicionarBarra(p As String) As String
    If Len(p) = 0 Then AdicionarBarra = "": Exit Function
    If Right(p, 1) = "\" Then
        AdicionarBarra = p
    Else
        AdicionarBarra = p & "\"
    End If
End Function

' Cria o CSV em g_pastaD1 (ou fallback) e flusha o buffer de linhas que
' foram emitidas antes de g_pastaD1 ser conhecido.
Private Sub MaterializarArquivoLogCSV()
    On Error Resume Next
    If Len(g_arqLogCSV) > 0 Then Exit Sub      ' ja foi materializado

    Dim pasta As String, arquivo As String
    pasta = ResolverPastaLogCSV()
    If Len(pasta) = 0 Then Exit Sub

    arquivo = pasta & NomeArquivoLogCSV()

    ' Cria com BOM UTF-8 + cabecalho (ADODB.Stream)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                                ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText "Timestamp;Status;Analista;Contexto;Detalhe" & vbCrLf
    ' Flusha o buffer (linhas anteriores)
    If Not g_bufLogCSV Is Nothing Then
        Dim i As Long
        For i = 1 To g_bufLogCSV.Count
            stm.WriteText g_bufLogCSV(i) & vbCrLf
        Next i
    End If
    stm.SaveToFile arquivo, 2                   ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing

    g_arqLogCSV = arquivo
    Set g_bufLogCSV = Nothing
    On Error GoTo 0
End Sub

' Append em CSV ja existente. Le binario, anexa bytes UTF-8 da nova linha,
' reescreve. Robusto para poucas chamadas/run; evita locking complexo.
Private Sub AppendCSVLine(arquivo As String, linha As String)
    On Error Resume Next
    Dim leitura As Object, escrita As Object
    Dim conteudo As Variant

    ' Le binario atual
    Set leitura = CreateObject("ADODB.Stream")
    leitura.Type = 1                            ' adTypeBinary
    leitura.Open
    leitura.LoadFromFile arquivo
    conteudo = leitura.Read                     ' bytes existentes
    leitura.Close
    Set leitura = Nothing

    ' Codifica nova linha como UTF-8 (sem BOM)
    Dim txt As Object
    Set txt = CreateObject("ADODB.Stream")
    txt.Type = 2                                ' adTypeText
    txt.Charset = "utf-8"
    txt.Open
    txt.WriteText linha & vbCrLf
    txt.Position = 0
    txt.Type = 1                                ' adTypeBinary
    ' Pula o BOM (3 bytes) se presente para que o append nao injete BOM no meio
    Dim bom As Long: bom = 0
    Dim sniff As Variant: sniff = txt.Read(3)
    If LenB(sniff) = 3 Then
        If AscB(MidB(sniff, 1, 1)) = &HEF And _
           AscB(MidB(sniff, 2, 1)) = &HBB And _
           AscB(MidB(sniff, 3, 1)) = &HBF Then
            bom = 3
        End If
    End If
    txt.Position = bom
    Dim novosBytes As Variant: novosBytes = txt.Read
    txt.Close
    Set txt = Nothing

    ' Reescreve = bytes antigos + novos
    Set escrita = CreateObject("ADODB.Stream")
    escrita.Type = 1                            ' adTypeBinary
    escrita.Open
    escrita.Write conteudo
    escrita.Write novosBytes
    escrita.SaveToFile arquivo, 2               ' adSaveCreateOverWrite
    escrita.Close
    Set escrita = Nothing
    On Error GoTo 0
End Sub

' Garante que ainda exista CSV mesmo se o fluxo falhar antes do step 3
' (g_pastaD1 nao definido). Chamado em Finalizar do FluxoCompleto.
Private Sub FlushFinalLogCSV()
    On Error Resume Next
    If Len(g_arqLogCSV) > 0 Then Exit Sub                ' ja gravado
    If g_bufLogCSV Is Nothing Then Exit Sub              ' nada para gravar
    If g_bufLogCSV.Count = 0 Then Exit Sub
    MaterializarArquivoLogCSV
    On Error GoTo 0
End Sub
