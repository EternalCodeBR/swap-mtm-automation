Attribute VB_Name = "Mod8_DownloadINETX"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 8: Download das curvas INETX (v1.0)
'
' OBJETIVO
'   Integra na planilha principal o processo antes feito pelo workbook
'   externo BaixaArquivosINETX.xlsm. Roda dentro do FluxoCompleto (Mod1)
'   ou pode ser chamado standalone via Alt+F8 -> BaixarCurvasINETX.
'
' FLUXO
'   1. Cria pasta de download local "INETX - DD.MM.YY"
'   2. Gera script WinSCP com credenciais SFTP e roda winscp.com
'   3. Move arquivos baixados para a pasta de destino
'   4. Filtra o CSV gerando "Curvas_Separadas_YYYYMMDD.xlsx" na pasta MTM D-1
'   5. Loga cada passo via Mod1.GravarLog (que tambem espelha no CSV fisico)
'
' DEPENDENCIAS (Public, em Mod1)
'   g_dataD1, g_pastaD1, GravarLog, LerParametro, ExpandirCaminho,
'   ObterCaminhoBaseUsuario
'
' DEPENDENCIAS EXTERNAS
'   - WinSCP instalado em %USERPROFILE%\AppData\Local\Programs\WinSCP\
'     (caminho pode ser sobrescrito via Parametros!PastaWinSCP)
'   - VPN/rede com acesso ao servidor SFTP do provedor de curvas
'
' PARAMETROS OPCIONAIS (aba Parametros)
'   PastaINETX_Base  -> pasta raiz INETX. Default: %USERPROFILE%\Empresa\
'                       Riscos - Documentos\Rotinas\PU CRI\Fluxo de pagamento\INETX\
'   PastaWinSCP      -> pasta com winscp.com. Default: %USERPROFILE%\AppData\
'                       Local\Programs\WinSCP\
'   INETX_SFTP_URL   -> URL de conexao SFTP completa, no formato
'                       "sftp://usuario%40host:senha@host:porta".
'                       OBRIGATORIO informar via Parametros; nunca commitar
'                       a credencial no codigo-fonte (usar Windows Credential
'                       Manager ou planilha protegida por senha em producao).
'===========================================================================
Option Explicit

' Curvas que filtramos do CSV bruto e viram abas do Curvas_Separadas.xlsx
Private Const CURVAS_FILTRO As String = "Pre;Cupom_SOFR;FWD_USD_BRL;Daily Compound SOFR"

'===========================================================================
' Entry point standalone (Alt+F8 -> BaixarCurvasINETX) ou chamavel do Mod1
'===========================================================================
Public Sub BaixarCurvasINETX()
    On Error GoTo ErrHandler

    ' Se rodando standalone (sem FluxoCompleto), inicializa g_dataD1 e g_wsLog
    If g_dataD1 = 0 Then
        Dim dataBase As Date
        If ObterDataBaseManual(dataBase) Then
            g_dataD1 = dataBase
        Else
            g_dataD1 = Date
        End If
    End If
    If g_wsLog Is Nothing Then Set g_wsLog = ObterOuCriarLog()

    GravarLog "INICIO", "BaixaINETX", "Download de curvas INETX para D-1=" & _
              Format(g_dataD1, "dd/mm/yyyy")

    Dim pastaDestino As String, pastaTmp As String, scriptCmd As String
    pastaDestino = ResolverPastaINETXDestino(g_dataD1)
    pastaTmp = ObterCaminhoBaseUsuario() & "Downloads\WinSCPGet\"

    CriarPastaRecursivo pastaDestino
    On Error Resume Next: MkDir pastaTmp: On Error GoTo ErrHandler

    ' Credencial SFTP: nunca hardcoded. Deve ser configurada na aba
    ' Parametros (chave "INETX_SFTP_URL"), idealmente lida de um cofre
    ' local (Windows Credential Manager) em vez de texto puro na planilha.
    Dim sftpUrl As String: sftpUrl = LerParametro("INETX_SFTP_URL")
    If Len(sftpUrl) = 0 Then
        GravarLog "ERRO", "BaixaINETX", "Parametro 'INETX_SFTP_URL' nao configurado."
        MsgBox "Configure a URL de conexao SFTP na aba Parametros " & _
               "(chave INETX_SFTP_URL) antes de rodar esta rotina.", _
               vbCritical, "BaixarCurvasINETX"
        Exit Sub
    End If

    ' Gera script WinSCP e roda
    scriptCmd = pastaTmp & "WinSCPGet.txt"
    CriarArquivoComando scriptCmd, sftpUrl, _
                        Format(g_dataD1, "YYYYMMDD"), _
                        Format(g_dataD1, "YYYY-MM-DD")

    Dim winscpExe As String
    winscpExe = ResolverWinSCPExe()
    If Len(winscpExe) = 0 Then
        GravarLog "ERRO", "BaixaINETX", "winscp.com nao encontrado. " & _
                  "Configure Parametros!PastaWinSCP ou instale o WinSCP."
        MsgBox "WinSCP nao encontrado. Verifique a instalacao ou ajuste " & _
               "Parametros!PastaWinSCP.", vbCritical, "BaixarCurvasINETX"
        Exit Sub
    End If

    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim exitCode As Long
    exitCode = wsh.Run(Chr(34) & winscpExe & Chr(34) & " /script=" & _
                       Chr(34) & scriptCmd & Chr(34), 4, True)
    Set wsh = Nothing
    On Error Resume Next: Kill scriptCmd: On Error GoTo ErrHandler

    If exitCode <> 0 Then
        GravarLog "ERRO", "BaixaINETX", "winscp.com retornou codigo " & exitCode
        MsgBox "WinSCP falhou (codigo " & exitCode & "). Veja a aba Log e " & _
               "verifique VPN/rede.", vbCritical, "BaixarCurvasINETX"
        Exit Sub
    End If

    ' Move arquivos baixados para pastaDestino
    Dim arqMov As String, qtdMov As Long
    arqMov = Dir(pastaTmp & "*.*")
    Do While Len(arqMov) > 0
        On Error Resume Next
        Name pastaTmp & arqMov As pastaDestino & arqMov
        If Err.Number = 0 Then qtdMov = qtdMov + 1
        Err.Clear
        On Error GoTo ErrHandler
        arqMov = Dir
    Loop
    On Error Resume Next: RmDir pastaTmp: On Error GoTo ErrHandler

    GravarLog "OK", "BaixaINETX", qtdMov & " arquivo(s) baixado(s) -> " & pastaDestino

    ' Filtra o CSV e gera Curvas_Separadas na pasta MTM D-1
    Dim caminhoCSV As String
    caminhoCSV = LocalizarCSVInetx(pastaDestino, Format(g_dataD1, "YYYYMMDD"), _
                                                  Format(g_dataD1, "YYYY-MM-DD"))
    If Len(caminhoCSV) = 0 Then
        GravarLog "ERRO", "BaixaINETX", "CSV de curvas nao encontrado em " & pastaDestino
        MsgBox "CSV de curvas nao foi baixado. Veja a aba Log.", vbCritical, _
               "BaixarCurvasINETX"
        Exit Sub
    End If
    GravarLog "INFO", "BaixaINETX", "CSV localizado: " & caminhoCSV

    Dim caminhoCurvas As String
    caminhoCurvas = SepararCurvasEmPlanilhas(caminhoCSV, g_dataD1)
    If Len(caminhoCurvas) = 0 Then
        GravarLog "ERRO", "BaixaINETX", "Falha ao separar curvas"
        Exit Sub
    End If

    GravarLog "FIM", "BaixaINETX", "Curvas_Separadas salvo: " & caminhoCurvas
    Exit Sub

ErrHandler:
    GravarLog "ERRO", "BaixaINETX", "Erro " & Err.Number & ": " & Err.Description
    MsgBox "Erro no download das curvas: " & Err.Description, vbCritical, _
           "BaixarCurvasINETX"
End Sub

'===========================================================================
' Pasta de destino: PastaINETX_Base \ Download INETX \ INETX - DD.MM.YY \
'===========================================================================
Private Function ResolverPastaINETXDestino(d As Date) As String
    Dim base As String
    base = LerParametro("PastaINETX_Base")
    If Len(base) = 0 Then
        base = ObterCaminhoBaseUsuario() & _
               "Empresa\Riscos - Documentos\Rotinas\PU CRI\" & _
               "Fluxo de pagamento\INETX\"
    End If
    If Right(base, 1) <> "\" Then base = base & "\"
    ResolverPastaINETXDestino = base & "Download INETX\" & _
                                "INETX - " & Format(d, "DD.MM.YY") & "\"
End Function

'===========================================================================
' Resolve o caminho do winscp.com (executavel CLI do WinSCP).
'===========================================================================
Private Function ResolverWinSCPExe() As String
    Dim p As String
    p = LerParametro("PastaWinSCP")
    If Len(p) = 0 Then
        p = ObterCaminhoBaseUsuario() & "AppData\Local\Programs\WinSCP\"
    End If
    If Right(p, 1) <> "\" Then p = p & "\"
    If Dir(p & "winscp.com") <> "" Then
        ResolverWinSCPExe = p & "winscp.com"
        Exit Function
    End If
    ' Fallback: Program Files
    Dim alt As String
    alt = Environ$("ProgramFiles") & "\WinSCP\winscp.com"
    If Dir(alt) <> "" Then ResolverWinSCPExe = alt: Exit Function
    alt = Environ$("ProgramFiles(x86)") & "\WinSCP\winscp.com"
    If Dir(alt) <> "" Then ResolverWinSCPExe = alt: Exit Function
    ResolverWinSCPExe = ""
End Function

'===========================================================================
' Cria o script de comandos do WinSCP (texto puro)
'===========================================================================
Private Sub CriarArquivoComando(arquivoSaida As String, _
                                 sftpUrl As String, _
                                 dataYYYYMMDD As String, _
                                 dataYYYY_MM_DD As String)
    Dim fs As Object, ts As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set ts = fs.CreateTextFile(arquivoSaida, True)
    Dim pastaTmp As String: pastaTmp = Left(arquivoSaida, InStrRev(arquivoSaida, "\"))
    ts.WriteLine "open " & sftpUrl
    ts.WriteLine "cd ArquivosEnviados"
    ts.WriteLine "get *" & dataYYYYMMDD & "* " & pastaTmp
    ts.WriteLine "get *" & dataYYYY_MM_DD & "* " & pastaTmp
    ts.WriteLine "exit"
    ts.Close
    Set ts = Nothing: Set fs = Nothing
End Sub

'===========================================================================
' MkDir recursivo (mkdir -p) - igual ao BaixaArquivosINETX original
'===========================================================================
Private Sub CriarPastaRecursivo(sPath As String)
    If Len(sPath) = 0 Then Exit Sub
    Dim partes As Variant: partes = Split(sPath, "\")
    Dim ini As Integer: ini = IIf(Left(sPath, 2) = "\\", 3, 1)
    Dim acc As String: acc = Left(sPath, InStr(ini, sPath, "\"))
    Dim i As Integer
    For i = ini To UBound(partes)
        acc = acc & partes(i) & "\"
        If Dir(acc, vbDirectory) = "" Then
            On Error Resume Next
            MkDir acc
            On Error GoTo 0
        End If
    Next i
End Sub

'===========================================================================
' Acha o CSV de curvas dentro da pasta (testa varios padroes de nome)
'===========================================================================
Private Function LocalizarCSVInetx(pasta As String, _
                                    dataYYYYMMDD As String, _
                                    dataYYYY_MM_DD As String) As String
    Dim tentativa As String
    tentativa = pasta & "Curvas_Inetx_" & dataYYYYMMDD & ".csv"
    If Dir(tentativa) <> "" Then LocalizarCSVInetx = tentativa: Exit Function
    tentativa = pasta & "Curvas_Inetx_" & dataYYYY_MM_DD & ".csv"
    If Dir(tentativa) <> "" Then LocalizarCSVInetx = tentativa: Exit Function
    Dim varredura As String: varredura = Dir(pasta & "Curvas_Inetx_*.csv")
    If Len(varredura) > 0 Then LocalizarCSVInetx = pasta & varredura
End Function

'===========================================================================
' Filtra o CSV bruto do INETX por curva e salva Curvas_Separadas_<data>.xlsx
' na pasta MTM D-1 (g_pastaD1 se preenchido; senao tenta resolver via
' Mod3.NomePastaData/ObterPastaBaseComMes).
' Retorna o caminho do .xlsx gerado (vazio se falhar).
'===========================================================================
Private Function SepararCurvasEmPlanilhas(caminhoCSV As String, _
                                           dtData As Date) As String
    On Error GoTo ErrSep

    ' Determina pasta destino
    Dim pastaDestino As String
    If Len(g_pastaD1) > 0 And Dir(g_pastaD1, vbDirectory) <> "" Then
        pastaDestino = g_pastaD1
    Else
        ' Tenta resolver via Mod3
        Dim raiz As String: raiz = LerParametro("PastaMTM_PU")
        Dim baseMes As String: baseMes = ObterPastaBaseComMes(raiz, dtData)
        If Len(baseMes) > 0 Then
            pastaDestino = baseMes & "\" & NomePastaData(dtData)
            If Dir(pastaDestino, vbDirectory) = "" Then CriarPastaRecursivo pastaDestino & "\"
        Else
            pastaDestino = ObterCaminhoBaseUsuario() & "Desktop"
            GravarLog "AVISO", "BaixaINETX", "Pasta MTM nao localizada; salvando no Desktop"
        End If
    End If
    If Right(pastaDestino, 1) <> "\" Then pastaDestino = pastaDestino & "\"

    Dim caminhoSaida As String
    caminhoSaida = pastaDestino & "Curvas_Separadas_" & _
                   Format(dtData, "YYYYMMDD") & ".xlsx"

    Const COL_FILTRO As Integer = 3
    Dim curvas() As String
    curvas = Split(CURVAS_FILTRO, ";")

    Dim wbCSV As Workbook, wsCSV As Worksheet
    Set wbCSV = Workbooks.Open(Filename:=caminhoCSV, Local:=True)
    Set wsCSV = wbCSV.Sheets(1)

    Dim ultLin As Long, ultCol As Long
    ultLin = wsCSV.Cells(wsCSV.Rows.Count, COL_FILTRO).End(xlUp).Row
    ultCol = wsCSV.Cells(1, wsCSV.Columns.Count).End(xlToLeft).Column
    If ultLin < 2 Then
        wbCSV.Close SaveChanges:=False
        GravarLog "ERRO", "BaixaINETX", "CSV vazio: " & caminhoCSV
        SepararCurvasEmPlanilhas = ""
        Exit Function
    End If

    Dim rng As Range
    Set rng = wsCSV.Range(wsCSV.Cells(1, 1), wsCSV.Cells(ultLin, ultCol))
    If wsCSV.AutoFilterMode Then wsCSV.AutoFilterMode = False

    Dim wbNovo As Workbook: Set wbNovo = Workbooks.Add
    Dim i As Integer
    For i = LBound(curvas) To UBound(curvas)
        rng.AutoFilter Field:=COL_FILTRO, Criteria1:="*" & curvas(i) & "*"
        Dim wsNova As Worksheet
        Set wsNova = wbNovo.Worksheets.Add(After:=wbNovo.Worksheets(wbNovo.Worksheets.Count))
        On Error Resume Next
        wsNova.Name = curvas(i)
        On Error GoTo ErrSep
        Dim temDados As Boolean: temDados = False
        On Error Resume Next
        temDados = (rng.SpecialCells(xlCellTypeVisible).Count > ultCol)
        On Error GoTo ErrSep
        If temDados Then
            rng.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNova.Range("A1")
            wsNova.Columns.AutoFit
        Else
            GravarLog "AVISO", "BaixaINETX", "Curva '" & curvas(i) & "' nao encontrada no CSV"
        End If
    Next i
    wsCSV.AutoFilterMode = False

    ' Limpa as abas padrao (Plan1/Sheet1/Folha1)
    Application.DisplayAlerts = False
    Dim wsDel As Worksheet
    For Each wsDel In wbNovo.Worksheets
        If wsDel.Name Like "Plan*" Or wsDel.Name Like "Sheet*" Or wsDel.Name Like "Folha*" Then
            If wbNovo.Sheets.Count > 1 Then wsDel.Delete
        End If
    Next wsDel

    wbNovo.SaveAs Filename:=caminhoSaida, FileFormat:=51  ' xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wbNovo.Close SaveChanges:=False
    wbCSV.Close SaveChanges:=False

    SepararCurvasEmPlanilhas = caminhoSaida
    Exit Function

ErrSep:
    Application.DisplayAlerts = True
    GravarLog "ERRO", "BaixaINETX", "SepararCurvas: " & Err.Description
    On Error Resume Next
    If Not wbNovo Is Nothing Then wbNovo.Close SaveChanges:=False
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False
    SepararCurvasEmPlanilhas = ""
End Function

