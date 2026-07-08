Attribute VB_Name = "Mod6_Bloomberg"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 6: Bloomberg + FRED API (v1.0)
'
' Responsabilidades:
'   - Localizar arquivos Bloomberg JPYBRL (diario) e SOFR6M (mensal)
'   - Copiar dados para as abas das calculadoras
'   - Buscar taxa SOFR diaria via FRED API (Daily Compound SOFR - Voxa)
'
' FASE 2: Voxa  -> Daily Compound SOFR via FRED API
' FASE 3: Agroverde, Orbita II -> FWD JPYBRL via Bloomberg
' FASE 4: Rodalux I, Provento -> SOFR6M via Bloomberg
'===========================================================================
Option Explicit

' --- Defaults (usados quando a aba Parametros nao tem override) ----------
' Todos podem ser sobrescritos via aba "Parametros" do Config:
'   SOFR_API_URL    -> base da URL FRED (sem query string)
'   SOFR_API_KEY    -> chave de acesso
'   SOFR_SERIES_ID  -> identificador da serie (default "SOFR")
'   SOFR_FATOR_CONV -> multiplicador aplicado apos o parse (default 1.0)
'   SOFR_ABA_VOXA    -> nome da aba na Voxa     (default "Daily Compound SOFR " com espaco)
'   SOFR_ABA_PROVENTO-> nome da aba na Provento (default "SOFR")
' Sem default hardcoded por design: a chave da FRED API (gratuita, obtida em
' https://fred.stlouisfed.org/docs/api/api_key.html) deve ser configurada na
' aba Parametros do Config (chave "SOFR_API_KEY"), nunca no codigo-fonte.
Private Const FRED_API_KEY      As String = ""
Private Const FRED_SERIES_SOFR  As String = "SOFR"
Private Const FRED_API_URL      As String = "https://api.stlouisfed.org/fred/series/observations"
Private Const SOFR_FATOR_CONV_DEFAULT As Double = 1#
Private Const ABA_FWD_JPYBRL    As String = "FWD JPYBRL na BBG"
Private Const ABA_SOFR6M        As String = "SOFR6M na BBG"
Private Const ABA_DAILY_SOFR    As String = "Daily Compound SOFR "  ' espaco no final
Private Const ABA_SOFR_PROVENTO  As String = "SOFR"

'===========================================================================
' HELPERS DE DATA / PARSE / PARAMETROS
'===========================================================================
' Helper: numero da coluna -> letra (1=A, 2=B, ..., 26=Z, 27=AA)
Private Function ColLetraBBG(n As Long) As String
    Dim s As String: s = ""
    Do While n > 0
        Dim m As Long: m = ((n - 1) Mod 26)
        s = Chr(65 + m) & s
        n = (n - 1) \ 26
    Loop
    ColLetraBBG = s
End Function

' Parser de numero vindo de API JSON (que sempre usa "." como decimal).
'
' Necessario porque o VBA CDbl respeita o locale do sistema operacional:
' em pt-BR, CDbl("3.59") retorna 359 (interpreta o ponto como separador
' de milhares). Esse helper detecta o separador decimal do Excel e
' normaliza a string antes de converter, garantindo que o parser
' funcione em qualquer locale.
Private Function ParseNumeroAPI(s As String) As Double
    Dim sepDec As String
    sepDec = Application.International(xlDecimalSeparator)
    If sepDec = "." Then
        ParseNumeroAPI = CDbl(s)
    Else
        ParseNumeroAPI = CDbl(Replace(s, ".", sepDec))
    End If
End Function

' Le um parametro da aba "Parametros" do Config como Double, tolerando
' valores escritos com "." ou "," como separador decimal. Se o parametro
' nao existir / estiver vazio / for invalido, retorna defVal.
Private Function LerParamDouble(chave As String, defVal As Double) As Double
    Dim raw As String: raw = LerParametro(chave)
    If Len(Trim(raw)) = 0 Then LerParamDouble = defVal: Exit Function

    Dim sepDec As String
    sepDec = Application.International(xlDecimalSeparator)
    Dim t As String: t = Trim(raw)
    ' Normaliza para o separador decimal do locale (aceita "." ou ",")
    If sepDec = "." Then
        t = Replace(t, ",", ".")
    Else
        t = Replace(t, ".", sepDec)
    End If
    On Error Resume Next
    LerParamDouble = CDbl(t)
    If Err.Number <> 0 Then LerParamDouble = defVal
    On Error GoTo 0
End Function

' Wrapper sobre LerParametro com fallback para um default String.
Private Function LerParamStr(chave As String, defVal As String) As String
    Dim raw As String: raw = LerParametro(chave)
    If Len(Trim(raw)) = 0 Then
        LerParamStr = defVal
    Else
        LerParamStr = raw
    End If
End Function

Public Function NomeMesPT(mes As Integer) As String
    Select Case mes
        Case 1:  NomeMesPT = "Jan"
        Case 2:  NomeMesPT = "Fev"
        Case 3:  NomeMesPT = "Mar"
        Case 4:  NomeMesPT = "Abr"
        Case 5:  NomeMesPT = "Mai"
        Case 6:  NomeMesPT = "Jun"
        Case 7:  NomeMesPT = "Jul"
        Case 8:  NomeMesPT = "Ago"
        Case 9:  NomeMesPT = "Set"
        Case 10: NomeMesPT = "Out"
        Case 11: NomeMesPT = "Nov"
        Case 12: NomeMesPT = "Dez"
        Case Else: NomeMesPT = ""
    End Select
End Function

'===========================================================================
' LOCALIZADORES DE ARQUIVOS BLOOMBERG
'===========================================================================

'--- JPYBRL diario ----------------------------------------------------------
' Estrutura: JPYBRL (YYYY)\JPYBRL - Mes.YY\JPYBRL - DD.MM.YY.xlsx
Public Function LocalizarArquivoJPYBRL(dataD1 As Date) As String
    Dim base As String
    base = "C:\Users\" & Environ("Username") & _
           "\Empresa\Bloomberg - Documentos\Sakura\JPYBRL\"

    Dim pastaAno As String
    pastaAno = base & "JPYBRL (" & Format(dataD1, "YYYY") & ")\"

    If Dir(pastaAno, vbDirectory) = "" Then
        GravarLog "ERRO", "JPYBRL", "Pasta do ano nao encontrada: " & pastaAno
        LocalizarArquivoJPYBRL = ""
        Exit Function
    End If

    Dim pastaMes As String
    pastaMes = pastaAno & "JPYBRL - " & NomeMesPT(Month(dataD1)) & "." & _
               Format(dataD1, "YY") & "\"

    If Dir(pastaMes, vbDirectory) = "" Then
        GravarLog "ERRO", "JPYBRL", "Pasta do mes nao encontrada: " & pastaMes
        LocalizarArquivoJPYBRL = ""
        Exit Function
    End If

    ' Tentativa 1: nome canonico "JPYBRL - DD.MM.YY.xlsx"
    Dim caminho As String
    caminho = pastaMes & "JPYBRL - " & Format(dataD1, "DD.MM.YY") & ".xlsx"
    If Dir(caminho) <> "" Then
        GravarLog "OK", "JPYBRL", "Arquivo encontrado: " & caminho
        LocalizarArquivoJPYBRL = caminho
        Exit Function
    End If

    ' Tentativa 2: fuzzy - varre a pasta procurando qualquer .xlsx que
    ' contenha o stamp DD.MM.YY (cobre variacoes "JPYBRL DD.MM.YY",
    ' "JPYBRL_DD.MM.YY", espacos extras, etc.)
    Dim stamp As String: stamp = Format(dataD1, "DD.MM.YY")
    Dim arq   As String: arq = Dir(pastaMes & "*.xlsx")
    Do While Len(arq) > 0
        If InStr(arq, stamp) > 0 Then
            caminho = pastaMes & arq
            GravarLog "OK", "JPYBRL", "Arquivo encontrado (fuzzy): " & caminho
            LocalizarArquivoJPYBRL = caminho
            Exit Function
        End If
        arq = Dir
    Loop

    GravarLog "ERRO", "JPYBRL", "Arquivo do dia nao encontrado em " & pastaMes & _
              " (procurado: stamp=" & stamp & ")"
    LocalizarArquivoJPYBRL = ""
End Function

'--- SOFR6M mensal ----------------------------------------------------------
' Estrutura: Curves - TSOFR6M (YYYY)\6M Term Sofr BBG - Curves Mes.YY.xlsx
Public Function LocalizarArquivoSOFR6M(dataD1 As Date) As String
    Dim base As String
    base = Environ("USERPROFILE") & "\Empresa\Bloomberg - Documentos\Sakura\SOFR\Curves - TSOFR 6M\"

    Dim pastaAno As String
    pastaAno = base & "Curves - TSOFR6M (" & Format(dataD1, "YYYY") & ")\"

    If Dir(pastaAno, vbDirectory) = "" Then
        GravarLog "ERRO", "SOFR6M", "Pasta do ano nao encontrada: " & pastaAno
        LocalizarArquivoSOFR6M = ""
        Exit Function
    End If

    Dim caminho As String
    caminho = pastaAno & "6M Term Sofr BBG - Curves " & _
              NomeMesPT(Month(dataD1)) & "." & Format(dataD1, "YY") & ".xlsx"

    If Dir(caminho) <> "" Then
        GravarLog "OK", "SOFR6M", "Arquivo encontrado: " & caminho
        LocalizarArquivoSOFR6M = caminho
    Else
        GravarLog "ERRO", "SOFR6M", "Arquivo mensal nao encontrado: " & caminho
        LocalizarArquivoSOFR6M = ""
    End If
End Function

'===========================================================================
' FASE 3: Copia FWD JPYBRL para a calculadora
'   Origem : 1a aba do arquivo Bloomberg, intervalo A1:L39
'   Destino: aba "FWD JPYBRL na BBG", a partir de A2
'===========================================================================
Public Function CopiarFWD_JPYBRL(wbCalc As Workbook, dataD1 As Date) As String
    If Not AbaExiste(wbCalc, ABA_FWD_JPYBRL) Then
        CopiarFWD_JPYBRL = "ERRO: aba '" & ABA_FWD_JPYBRL & "' nao existe na calculadora"
        Exit Function
    End If

    Dim caminho As String: caminho = LocalizarArquivoJPYBRL(dataD1)
    If Len(caminho) = 0 Then
        CopiarFWD_JPYBRL = "ERRO: arquivo JPYBRL nao encontrado para " & _
                            Format(dataD1, "dd/mm/yyyy")
        Exit Function
    End If

    Dim wbJPY As Workbook
    On Error Resume Next
    Set wbJPY = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0
    If wbJPY Is Nothing Then
        CopiarFWD_JPYBRL = "ERRO: nao foi possivel abrir: " & caminho
        Exit Function
    End If

    Dim wsOrig As Worksheet: Set wsOrig = wbJPY.Sheets(1)

    ' Verifica se ha dados em A1:L39
    If IsEmpty(wsOrig.Range("A1").Value) And IsEmpty(wsOrig.Range("A2").Value) Then
        wbJPY.Close SaveChanges:=False
        CopiarFWD_JPYBRL = "ERRO: intervalo A1:L39 parece vazio no arquivo JPYBRL"
        Exit Function
    End If

    Dim wsDest As Worksheet: Set wsDest = wbCalc.Sheets(ABA_FWD_JPYBRL)

    wsOrig.Range("A1:L39").Copy
    wsDest.Range("A2").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    wbJPY.Close SaveChanges:=False
    GravarLog "OK", "JPYBRL", "FWD JPYBRL copiado (A1:L39 -> A2) em: " & wbCalc.Name
    CopiarFWD_JPYBRL = ""
End Function

'===========================================================================
' FASE 4: Copia SOFR6M para a calculadora
'   Origem : 1a aba do arquivo Bloomberg, intervalo usado completo
'   Destino: aba "SOFR6M na BBG", a partir de A1 (substitui tudo)
'===========================================================================
Public Function CopiarSOFR6M(wbCalc As Workbook, dataD1 As Date) As String
    If Not AbaExiste(wbCalc, ABA_SOFR6M) Then
        ' Algumas calculadoras (ex: Provento) tem 'sofr6m' no Config por
        ' engano mas nao possuem a aba. Em vez de quebrar o fluxo, apenas
        ' registra um AVISO e segue - assim a calculadora ainda atualiza
        ' as outras curvas (Fator Pre, CDI, SOFR, etc.).
        GravarLog "AVISO", "SOFR6M", "Aba '" & ABA_SOFR6M & _
                  "' nao existe em " & wbCalc.Name & " - curva ignorada"
        CopiarSOFR6M = ""
        Exit Function
    End If

    Dim caminho As String: caminho = LocalizarArquivoSOFR6M(dataD1)
    If Len(caminho) = 0 Then
        CopiarSOFR6M = "ERRO: arquivo SOFR6M nao encontrado para " & _
                        Format(dataD1, "dd/mm/yyyy")
        Exit Function
    End If

    Dim wbSOFR As Workbook
    On Error Resume Next
    Set wbSOFR = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0
    If wbSOFR Is Nothing Then
        CopiarSOFR6M = "ERRO: nao foi possivel abrir: " & caminho
        Exit Function
    End If

    Dim wsOrig As Worksheet: Set wsOrig = wbSOFR.Sheets(1)

    ' ---------- ORIGEM (template Bloomberg "6M Term Sofr BBG - Curves") ----
    ' Linha 6: D6, E6, F6, ... = datas de cada curva (uma coluna por dia)
    ' Linhas 8..17: dados (10 linhas)
    '   C8:C17  = Reset Date (fixos)
    '   <Col>8..<Col>17 = Reset Rate do dia correspondente a <Col>6
    '
    ' ---------- DESTINO (aba "SOFR6M na BBG" - Rodalux I) ----------------
    '   A1        = data D-1 (NAO TOCAR)
    '   A2/B2     = "Curve Date" / label (NAO TOCAR)
    '   A3/B3     = "Reset Date" label / data DD.MM.YYYY  <- ATUALIZAMOS AQUI
    '   A4:A13    = Reset Dates fixos (NAO TOCAR)
    '   B4:B13    = Reset Rates do dia (UNICA COLUNA QUE ATUALIZAMOS)
    ' ---------------------------------------------------------------------

    Const SRC_ROW_HEADER_DATE As Long = 6
    Const SRC_ROW_INI         As Long = 8
    Const SRC_ROW_FIM         As Long = 17
    Const DEST_ROW_INI        As Long = 4

    ' Localiza na linha 6 da origem a coluna cuja data corresponde a dataD1.
    ' Estrategia em 3 passos:
    '   1. Match EXATO com dataD1
    '   2. Se nao achar: pega a data MAIS RECENTE <= dataD1 (cobre feriados
    '      e arquivos Bloomberg atrasados em 1-2 dias)
    '   3. Se nao achar: loga TODAS as datas vistas + retorna erro
    Dim ultColOrig As Long
    ultColOrig = wsOrig.Cells(SRC_ROW_HEADER_DATE, wsOrig.Columns.Count).End(xlToLeft).Column

    Dim srcCol         As Long: srcCol = 0
    Dim srcColFallback As Long: srcColFallback = 0
    Dim dataFallback   As Date: dataFallback = #1/1/1900#
    Dim datasVistas    As String: datasVistas = ""
    Dim cc             As Long

    For cc = 4 To ultColOrig   ' a partir de D
        Dim v As Variant: v = wsOrig.Cells(SRC_ROW_HEADER_DATE, cc).Value
        Dim dataCol As Date: dataCol = #1/1/1900#
        Dim ehData  As Boolean: ehData = False

        ' Tenta interpretar como data (Double / Date / String)
        On Error Resume Next
        If IsNumeric(v) Then
            If CLng(CDate(v)) = CLng(dataD1) Then
                srcCol = cc
                Exit For
            End If
            dataCol = CDate(v)
            ehData = (Err.Number = 0)
        ElseIf IsDate(v) Then
            If CDate(v) = dataD1 Then
                srcCol = cc
                Exit For
            End If
            dataCol = CDate(v)
            ehData = True
        ElseIf VarType(v) = vbString And Len(Trim(CStr(v))) > 0 Then
            ' Tenta parse de string (formatos "13/05/2026", "2026-05-13", etc.)
            dataCol = CDate(CStr(v))
            ehData = (Err.Number = 0)
        End If
        Err.Clear
        On Error GoTo 0

        If ehData Then
            ' Acumula descricao para log de diagnostico (so primeiros 15)
            If Len(datasVistas) < 600 Then
                datasVistas = datasVistas & ColLetraBBG(cc) & "=" & _
                              Format(dataCol, "dd/mm/yyyy") & "  "
            End If

            ' 1. Match EXATO
            If CLng(dataCol) = CLng(dataD1) Then
                srcCol = cc
                Exit For
            End If
            ' 2. Fallback: data mais recente <= dataD1
            If dataCol < dataD1 And dataCol > dataFallback Then
                srcColFallback = cc
                dataFallback = dataCol
            End If
        End If
    Next cc

    ' Se nao achou match exato, usa o fallback (data mais recente <= D-1)
    If srcCol = 0 And srcColFallback > 0 Then
        srcCol = srcColFallback
        GravarLog "AVISO", "SOFR6M", "Match exato para " & _
                  Format(dataD1, "dd/mm/yyyy") & " nao encontrado; usando " & _
                  "data mais recente <=D-1: " & ColLetraBBG(srcCol) & "=" & _
                  Format(dataFallback, "dd/mm/yyyy")
    End If

    If srcCol = 0 Then
        ' Diagnostico: lista o que viu na linha 6
        GravarLog "DIAG", "SOFR6M", "Datas vistas na linha 6 (D:" & _
                  ColLetraBBG(ultColOrig) & "): " & _
                  IIf(Len(datasVistas) = 0, "(nenhuma data parseavel)", datasVistas)
        wbSOFR.Close SaveChanges:=False
        CopiarSOFR6M = "ERRO: nenhuma data <= " & Format(dataD1, "dd/mm/yyyy") & _
                       " na linha 6 da origem (varrido D:" & _
                       ColLetraBBG(ultColOrig) & "). Verifique log DIAG."
        Exit Function
    End If

    Dim wsDest As Worksheet: Set wsDest = wbCalc.Sheets(ABA_SOFR6M)

    ' Limpa apenas B3 e B4:B13 (preserva A1, A2:B2, A3, A4:A13)
    wsDest.Range("B3").ClearContents
    Dim DEST_ROW_FIM As Long
    DEST_ROW_FIM = DEST_ROW_INI + (SRC_ROW_FIM - SRC_ROW_INI)   ' = 13
    wsDest.Range("B" & DEST_ROW_INI & ":B" & DEST_ROW_FIM).ClearContents

    ' B3 = data da curva em DD.MM.YYYY (usamos dataD1, ja confirmada)
    wsDest.Range("B3").NumberFormat = "@"
    wsDest.Range("B3").Value = Format(dataD1, "dd.mm.yyyy")

    ' Copia rates: origem (SRC_ROW_INI..SRC_ROW_FIM, srcCol) -> destino B4..B13
    Dim r As Long, destR As Long
    For r = SRC_ROW_INI To SRC_ROW_FIM
        destR = DEST_ROW_INI + (r - SRC_ROW_INI)
        Dim vRate As Variant: vRate = wsOrig.Cells(r, srcCol).Value
        wsDest.Cells(destR, 2).NumberFormat = "0.00000%"
        wsDest.Cells(destR, 2).Value = vRate
    Next r

    wbSOFR.Close SaveChanges:=False
    GravarLog "OK", "SOFR6M", "SOFR6M copiado (col origem=" & ColLetraBBG(srcCol) & _
              " dia=" & Format(dataD1, "dd/mm/yyyy") & ", destino B3 + B4:B13) em: " & _
              wbCalc.Name
    CopiarSOFR6M = ""
End Function

'===========================================================================
' FASE 2: Busca taxa SOFR diaria via FRED API
'
' Serie: SOFR (Secured Overnight Financing Rate)
' Retorna o valor mais recente <= dataAlvo (cobre fins de semana / feriados)
' Retorna -1 em caso de falha
'===========================================================================
Public Function ObterSOFR_FRED(dataAlvo As Date) As Double
    On Error GoTo ErrHandler

    ' --- Parametros (com fallback para os defaults hardcoded) -----------
    Dim urlBase   As String: urlBase   = LerParamStr("SOFR_API_URL",    FRED_API_URL)
    Dim apiKey    As String: apiKey    = LerParamStr("SOFR_API_KEY",    FRED_API_KEY)
    Dim seriesId  As String: seriesId  = LerParamStr("SOFR_SERIES_ID",  FRED_SERIES_SOFR)
    Dim fatorConv As Double: fatorConv = LerParamDouble("SOFR_FATOR_CONV", SOFR_FATOR_CONV_DEFAULT)

    ' Busca os ultimos 10 dias para cobrir fins de semana e feriados
    Dim dtStart As String: dtStart = Format(DateAdd("d", -10, dataAlvo), "YYYY-MM-DD")
    Dim dtEnd   As String: dtEnd   = Format(dataAlvo, "YYYY-MM-DD")

    Dim url As String
    url = urlBase & _
          "?series_id=" & seriesId & _
          "&observation_start=" & dtStart & _
          "&observation_end=" & dtEnd & _
          "&sort_order=desc" & _
          "&limit=5" & _
          "&api_key=" & apiKey & _
          "&file_type=json"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/4.0"
    http.Send

    If http.Status <> 200 Then
        GravarLog "ERRO", "FRED_API", "HTTP " & http.Status & " | URL: " & url
        ObterSOFR_FRED = -1
        Exit Function
    End If

    Dim resp As String: resp = http.responseText

    ' Parseia JSON manualmente: localiza "value":"<numero>"
    ' sort_order=desc -> primeiro resultado e o mais recente
    ' FRED usa "." para dados ausentes -> pula e tenta o proximo
    Dim rate     As Double: rate = -1
    Dim pos      As Long
    Dim srchFrom As Long:   srchFrom = 1
    Dim rateRaw  As String: rateRaw = ""

    Do
        pos = InStr(srchFrom, resp, """value"":""")
        If pos = 0 Then Exit Do

        Dim vStart As Long: vStart = pos + 9
        Dim vEnd   As Long: vEnd   = InStr(vStart, resp, """")
        If vEnd = 0 Then Exit Do

        Dim valStr As String: valStr = Mid(resp, vStart, vEnd - vStart)

        ' FRED usa "." para indicar feriado / sem dado. Validacao manual
        ' (NAO usar IsNumeric, que e locale-sensitive: em pt-BR, "3.59"
        ' pode ser considerado numerico mas seria interpretado como 359).
        If valStr <> "." And valStr <> "" Then
            ' Parser locale-aware: converte "3.59" -> 3.59 em qualquer
            ' regional setting. CDbl puro converteria para 359 em pt-BR.
            rate = ParseNumeroAPI(valStr)
            rateRaw = valStr
            Exit Do
        End If

        srchFrom = vEnd + 1
    Loop

    ' Aplica fator de conversao (default 1.0). Permite ajustar caso a
    ' fonte mude para basis points (0.01) ou decimal puro (100) sem
    ' alterar codigo.
    If rate >= 0 Then
        Dim rateFinal As Double: rateFinal = rate * fatorConv
        GravarLog "OK", "FRED_API", "SOFR raw=""" & rateRaw & """ -> parse=" & rate & _
                  " x fator=" & fatorConv & " => " & rateFinal & _
                  "% (referencia ate " & dtEnd & ")"
        ObterSOFR_FRED = rateFinal
    Else
        GravarLog "ERRO", "FRED_API", "Valor nao encontrado na resposta FRED para " & dtEnd
        ObterSOFR_FRED = -1
    End If
    Exit Function

ErrHandler:
    GravarLog "ERRO", "FRED_API", "Excecao ao chamar FRED: " & Err.Description
    ObterSOFR_FRED = -1
End Function

'===========================================================================
' FASE 2: Atualiza aba "Daily Compound SOFR " com taxa do FRED
'
' Comportamento identico ao CDI:
'   - Verifica se a data ja existe -> atualiza o valor
'   - Se nao existe -> acrescenta nova linha (data, taxa)
'===========================================================================
Public Function AtualizarDailySOFR_FRED(wb As Workbook, dataD1 As Date) As String
    ' --- Resolucao do nome da aba ---------------------------------------
    ' Ordem de prioridade:
    '   1. Parametro explicito da aba "Parametros" do Config
    '        SOFR_ABA_VOXA     -> se a aba existir no wb, usa
    '        SOFR_ABA_PROVENTO -> se a aba existir no wb, usa
    '   2. Auto-deteccao por candidatos comuns (defaults conhecidos)
    '
    ' Os parametros sao OPCIONAIS: se a aba "Parametros" nao listar nenhum,
    ' a auto-deteccao cobre Voxa ("Daily Compound SOFR ") e Provento ("SOFR").
    Dim nomeAba As String: nomeAba = ""

    Dim abaTagParam     As String: abaTagParam     = LerParametro("SOFR_ABA_VOXA")
    Dim abaProventoParam As String: abaProventoParam = LerParametro("SOFR_ABA_PROVENTO")

    If Len(Trim(abaTagParam)) > 0 Then
        If AbaExiste(wb, abaTagParam) Then nomeAba = abaTagParam
    End If
    If Len(nomeAba) = 0 And Len(Trim(abaProventoParam)) > 0 Then
        If AbaExiste(wb, abaProventoParam) Then nomeAba = abaProventoParam
    End If

    If Len(nomeAba) = 0 Then
        Dim candidatos As Variant
        candidatos = Array(ABA_DAILY_SOFR, "Daily Compound SOFR", _
                           "DailyCompoundSOFR", ABA_SOFR_PROVENTO)
        Dim cand As Variant
        For Each cand In candidatos
            If AbaExiste(wb, CStr(cand)) Then
                nomeAba = CStr(cand)
                Exit For
            End If
        Next cand
    End If

    If Len(nomeAba) = 0 Then
        AtualizarDailySOFR_FRED = "ERRO: aba SOFR (Daily Compound SOFR / SOFR) nao existe na calculadora"
        Exit Function
    End If

    Dim taxa As Double: taxa = ObterSOFR_FRED(dataD1)
    If taxa < 0 Then
        AtualizarDailySOFR_FRED = "ERRO: nao foi possivel obter SOFR do FRED para " & _
                                   Format(dataD1, "dd/mm/yyyy")
        Exit Function
    End If

    Dim ws  As Worksheet: Set ws  = wb.Sheets(nomeAba)
    Dim ult As Long:      ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' --- Detecta o layout da aba ----------------------------------------
    ' Duas estruturas conhecidas:
    '
    '   LAYOUT COMPLETO (Voxa / "Daily Compound SOFR "):
    '     A=Effective Date | B="SOFR" (tipo) | C=Rate (%) | D:I=formulas
    '
    '   LAYOUT ENXUTO (Provento / aba "SOFR"):
    '     A=Date | B=Rate (%)  (sem coluna de "tipo")
    '
    ' Como saber qual eh? Heuristica:
    '   - Se a aba se chama exatamente "SOFR" (case-insensitive), eh Provento
    '   - Se o nome contem "daily compound" ou "compoundsofr", eh Voxa
    '   - Override: parametros SOFR_COL_TAXA / SOFR_COL_TIPO permitem
    '     forcar manualmente (uteis se a aba for renomeada)
    Dim colTaxa As Long: colTaxa = 3   ' default Voxa: C
    Dim colTipo As Long: colTipo = 2   ' default Voxa: B (gravar "SOFR")

    Dim nomeAbaLC As String: nomeAbaLC = LCase(Trim(nomeAba))
    If nomeAbaLC = "sofr" Then
        ' Layout enxuto Provento: B=taxa, sem coluna de tipo
        colTaxa = 2
        colTipo = 0
    End If

    ' Override por parametro (string "B", "C", "D" etc. ou numero 2, 3...)
    Dim ovTaxa As String: ovTaxa = LerParametro("SOFR_COL_TAXA")
    Dim ovTipo As String: ovTipo = LerParametro("SOFR_COL_TIPO")
    If Len(Trim(ovTaxa)) > 0 Then colTaxa = ConverterParaColuna(ovTaxa, colTaxa)
    If Len(Trim(ovTipo)) > 0 Then colTipo = ConverterParaColuna(ovTipo, colTipo)

    ' Verifica se a data ja existe -> atualiza a taxa
    Dim r As Long
    For r = 2 To ult
        Dim cel As Variant: cel = ws.Cells(r, 1).Value
        If IsDate(cel) Then
            If CDate(cel) = dataD1 Then
                ' Herda o formato da linha anterior (se existir) para
                ' manter o padrao visual da coluna ("3,60" etc).
                If r > 2 Then
                    ws.Cells(r, colTaxa).NumberFormat = _
                        ws.Cells(r - 1, colTaxa).NumberFormat
                Else
                    ws.Cells(r, colTaxa).NumberFormat = "0.00"
                End If
                ws.Cells(r, colTaxa).Value = taxa
                ' So grava "SOFR" se a aba tem coluna de tipo e ela esta vazia
                If colTipo > 0 Then
                    If Len(Trim(CStr(ws.Cells(r, colTipo).Value))) = 0 Then _
                        ws.Cells(r, colTipo).Value = "SOFR"
                End If
                GravarLog "DIAG", "SOFR_FRED", "Cel " & _
                          ws.Cells(r, colTaxa).Address(False, False) & _
                          " | Value=" & ws.Cells(r, colTaxa).Value & _
                          " | Format=" & ws.Cells(r, colTaxa).NumberFormat & _
                          " | Text=" & ws.Cells(r, colTaxa).Text
                GravarLog "INFO", "SOFR_FRED", "Taxa atualizada (" & nomeAba & _
                          "!" & ColLetraBBG(colTaxa) & "): " & _
                          Format(dataD1, "dd/mm/yyyy") & " = " & taxa & "%"
                AtualizarDailySOFR_FRED = ""
                Exit Function
            End If
        End If
    Next r

    ' Acrescenta nova linha replicando a linha anterior inteira e
    ' depois sobrescreve A=data, [colTipo]="SOFR", [colTaxa]=taxa.
    Dim nova As Long: nova = ult + 1
    Dim ultColSOFR As Long
    ultColSOFR = ws.Cells(ult, ws.Columns.Count).End(xlToLeft).Column
    If ultColSOFR < colTaxa Then ultColSOFR = colTaxa

    If ult >= 2 Then
        On Error Resume Next
        ws.Range(ws.Cells(ult, 1), ws.Cells(ult, ultColSOFR)).Copy
        ws.Cells(nova, 1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        On Error GoTo 0
    End If

    ws.Cells(nova, 1).Value = dataD1
    ws.Cells(nova, 1).NumberFormat = "dd/mm/yyyy"
    If colTipo > 0 Then ws.Cells(nova, colTipo).Value = "SOFR"
    ws.Cells(nova, colTaxa).Value = taxa
    ' Se a linha anterior nao tinha formato, garante 2 casas
    If ws.Cells(nova, colTaxa).NumberFormat = "General" Then _
        ws.Cells(nova, colTaxa).NumberFormat = "0.00"

    ' --- DIAGNOSTICO: registra exatamente o que o Excel armazenou ---
    GravarLog "DIAG", "SOFR_FRED", "Cel " & _
              ws.Cells(nova, colTaxa).Address(False, False) & _
              " | Value=" & ws.Cells(nova, colTaxa).Value & _
              " | Format=" & ws.Cells(nova, colTaxa).NumberFormat & _
              " | Text=" & ws.Cells(nova, colTaxa).Text

    Dim descLayout As String
    If colTipo > 0 Then
        descLayout = "A=data," & ColLetraBBG(colTipo) & "=SOFR," & _
                     ColLetraBBG(colTaxa) & "=" & taxa
    Else
        descLayout = "A=data," & ColLetraBBG(colTaxa) & "=" & taxa & _
                     " (layout enxuto, sem coluna de tipo)"
    End If
    GravarLog "INFO", "SOFR_FRED", "Nova linha (" & nomeAba & _
              " replica A:" & ColLetraBBG(ultColSOFR) & " | " & descLayout & _
              "): " & Format(dataD1, "dd/mm/yyyy")
    AtualizarDailySOFR_FRED = ""
End Function

' Converte "B" / "C" / "AA" / "2" / "3" para numero de coluna.
' Retorna defVal se a string nao for parseavel.
Private Function ConverterParaColuna(s As String, defVal As Long) As Long
    Dim t As String: t = UCase(Trim(s))
    If Len(t) = 0 Then ConverterParaColuna = defVal: Exit Function
    If IsNumeric(t) Then
        ConverterParaColuna = CLng(t)
        Exit Function
    End If
    ' Conversao letra -> numero (A=1, B=2, ..., Z=26, AA=27, ...)
    Dim i As Long, n As Long: n = 0
    For i = 1 To Len(t)
        Dim ch As String: ch = Mid(t, i, 1)
        If ch < "A" Or ch > "Z" Then
            ConverterParaColuna = defVal
            Exit Function
        End If
        n = n * 26 + (Asc(ch) - Asc("A") + 1)
    Next i
    If n = 0 Then n = defVal
    ConverterParaColuna = n
End Function
