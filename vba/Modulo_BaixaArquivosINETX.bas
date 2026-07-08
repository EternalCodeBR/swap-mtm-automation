Attribute VB_Name = "Módulo1"
'===========================================================================
' BaixaArquivosINETX - Módulo1 (v4)
'
' Versao standalone (workbook proprio) do download de curvas, anterior a
' integracao no Mod8_DownloadINETX do Atualizador de Swaps. Mantida no
' repositorio como referencia da evolucao do projeto (ver docs/roadmap.md).
'
' ALTERACOES v4:
'   1. Formato pasta MTM corrigido: DD.MM.YY (ano 2 digitos)
'   2. "Daily Compound SOFR" adicionado ao filtro de curvas (para Voxa)
'   3. Bug "Do sem Loop" corrigido (loop de movimentacao de arquivos)
'
' CREDENCIAL SFTP: le da celula nomeada "SftpUrl" na aba "Inicio" (nunca
' hardcoded no codigo-fonte). Formato esperado:
'   sftp://usuario%40host:senha@host:porta
'===========================================================================
Option Explicit

'===========================================================================
' Sub principal
'===========================================================================
Public Sub baixaArquivosINETX()
    Dim sAspas              As String
    Dim sComando            As String
    Dim objNetwork          As Object
    Dim sPastaBaseINETX     As String
    Dim sPastaDownloadINETX As String
    Dim sPastaDestino       As String
    Dim sPastaWinSCP        As String
    Dim sArquivoComando     As String
    Dim sArquivoDownload    As String
    Dim sUserName           As String
    Dim dtData              As Date
    Dim sData1              As String
    Dim sData2              As String
    Dim StrFile             As String
    Dim wsh                 As Object

    dtData = Worksheets("Inicio").Range("hj").Value
    sData1 = Format(dtData, "YYYYMMDD")
    sData2 = Format(dtData, "YYYY-MM-DD")

    Set objNetwork = CreateObject("WScript.Network")
    sUserName = objNetwork.UserName
    Set objNetwork = Nothing

    sPastaBaseINETX     = "C:\Users\" & sUserName & "\Empresa\Riscos - Documentos\" & _
                          "Rotinas\PU CRI\Fluxo de pagamento\INETX\"
    sPastaDownloadINETX = sPastaBaseINETX & "Download INETX\"
    sPastaDestino       = sPastaDownloadINETX & "INETX - " & Format(dtData, "DD.MM.YY") & "\"
    sPastaWinSCP        = "C:\Users\" & sUserName & "\AppData\Local\Programs\WinSCP\"
    sArquivoDownload    = "C:\Users\" & sUserName & "\Downloads\WinSCPGet\"

    MyMkDir sPastaDownloadINETX
    MyMkDir sPastaDestino
    On Error Resume Next
    MkDir sArquivoDownload
    On Error GoTo 0

    Dim sSftpUrl As String
    On Error Resume Next
    sSftpUrl = Worksheets("Inicio").Range("SftpUrl").Value
    On Error GoTo 0
    If Len(sSftpUrl) = 0 Then
        MsgBox "Configure a URL de conexao SFTP na celula nomeada 'SftpUrl' " & _
               "da aba Inicio antes de rodar esta rotina.", vbCritical, "Parametro faltando"
        Exit Sub
    End If

    sArquivoComando = sArquivoDownload & "WinSCPGet.txt"
    criaArquivoComando sArquivoDownload, sSftpUrl, sData1, sData2

    sAspas   = Chr(34)
    sComando = "/script=" & sArquivoComando

    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run sPastaWinSCP & "winscp.com " & sAspas & sComando & sAspas, 4, True

    ' Remove script e move arquivos para INETX
    On Error Resume Next
    Kill sArquivoComando
    On Error GoTo 0

    ' CORRECAO v4: "Do sem Loop" corrigido -- Loop adicionado
    ' CORRECAO v5: Dir() sem wildcard nao lista arquivos na pasta; corrigido para "*.*"
    StrFile = Dir(sArquivoDownload & "*.*")
    On Error Resume Next
    Do While Len(StrFile) > 0
        Name sArquivoDownload & StrFile As sPastaDestino & StrFile
        StrFile = Dir
    Loop
    RmDir sArquivoDownload
    On Error GoTo 0

    Call movarqu

    MsgBox "Arquivos baixados com sucesso!", vbInformation, "Download INETX"

    ' Filtra curvas e salva Curvas_Separadas na pasta MTM D-1
    ' Tenta YYYYMMDD, depois YYYY-MM-DD, depois qualquer Curvas_Inetx_*.csv
    Dim caminhoCSV As String
    caminhoCSV = sPastaDestino & "Curvas_Inetx_" & sData1 & ".csv"
    If Dir(caminhoCSV) = "" Then
        caminhoCSV = sPastaDestino & "Curvas_Inetx_" & sData2 & ".csv"
    End If
    If Dir(caminhoCSV) = "" Then
        Dim nomeCSV As String
        nomeCSV = Dir(sPastaDestino & "Curvas_Inetx_*.csv")
        If nomeCSV <> "" Then caminhoCSV = sPastaDestino & nomeCSV
    End If

    If Dir(caminhoCSV) <> "" Then
        Call SepararCurvasEmPlanilhas(caminhoCSV, sData1, dtData, sUserName)
    Else
        MsgBox "CSV de curvas nao encontrado." & vbCrLf & _
               "Caminho verificado: " & sPastaDestino & vbCrLf & _
               "Padroes testados: Curvas_Inetx_" & sData1 & ".csv  /  Curvas_Inetx_" & sData2 & ".csv", _
               vbCritical, "Arquivo Nao Encontrado"
    End If
End Sub

'===========================================================================
' Cria o script WinSCP para download SFTP
'===========================================================================
Sub criaArquivoComando(sFilePath As String, sSftpUrl As String, sData1 As String, sData2 As String)
    Dim fs As Object
    Dim a  As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a  = fs.CreateTextFile(sFilePath & "WinSCPGet.txt", True)
    a.WriteLine ("open " & sSftpUrl)
    a.WriteLine ("cd ArquivosEnviados")
    a.WriteLine ("get *" & sData1 & "* " & sFilePath)
    a.WriteLine ("get *" & sData2 & "* " & sFilePath)
    a.WriteLine ("exit")
    a.Close
End Sub

'===========================================================================
' Cria diretorios aninhados (mkdir -p)
'===========================================================================
Public Sub MyMkDir(sPath As String)
    Dim iStart  As Integer
    Dim aDirs   As Variant
    Dim sCurDir As String
    Dim i       As Integer

    If sPath <> "" Then
        aDirs   = Split(sPath, "\")
        iStart  = IIf(Left(sPath, 2) = "\\", 3, 1)
        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))
        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then MkDir sCurDir
        Next i
    End If
End Sub

'===========================================================================
' Filtra o CSV do INETX e gera Curvas_Separadas na pasta MTM D-1
'
' ALTERACAO v4:
'   - Formato pasta MTM: "MTM - DD.MM.YY" (ano 2 digitos)
'   - Adicionado "Daily Compound SOFR" ao filtro (necessario para Voxa)
'===========================================================================
Public Sub SepararCurvasEmPlanilhas(caminhoCompleto As String, _
                                     dataArquivo     As String, _
                                     dtData          As Date, _
                                     sUserName       As String)
    Dim wbCSV         As Workbook
    Dim wsCSV         As Worksheet
    Dim wbNovo        As Workbook
    Dim wsNova        As Worksheet
    Dim rngDados      As Range
    Dim ultimaLinha   As Long
    Dim ultimaColuna  As Long
    Dim arrCurvas     As Variant
    Dim i             As Integer
    Dim caminhoSalvar As String

    Const COLUNA_FILTRO As Integer = 3

    ' ALTERACAO v4: "Daily Compound SOFR" adicionado para atender Voxa
    arrCurvas = Array("Pre", "Cupom_SOFR", "FWD_USD_BRL", "Daily Compound SOFR")

    ' Determina destino: pasta MTM D-1 -> pergunta criar -> Desktop
    Dim pastaMTM As String
    pastaMTM = ObterPastaMTM_D1(dtData, sUserName)

    If Len(pastaMTM) > 0 Then
        caminhoSalvar = pastaMTM & "Curvas_Separadas_" & dataArquivo & ".xlsx"

    ElseIf Len(ObterBasePastaMTM(sUserName, dtData)) > 0 Then
        Dim nomeMesPasta As String
        nomeMesPasta = "MTM & PU - " & AbrevMesPT(Month(dtData)) & "." & Format(dtData, "YY")
        Dim nomePasta As String
        nomePasta = "MTM - " & Format(dtData, "DD.MM.YY")
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("A pasta '" & nomePasta & "' nao existe." & vbCrLf & vbCrLf & _
                          "Sim = Criar a pasta e salvar la" & vbCrLf & _
                          "Nao = Salvar no Desktop", _
                          vbQuestion + vbYesNo, "Pasta nao encontrada")

        If resposta = vbYes Then
            Dim novaPasta As String
            novaPasta = ObterBasePastaMTM(sUserName, dtData) & nomeMesPasta & "\" & nomePasta & "\"
            MyMkDir novaPasta
            If Dir(novaPasta, vbDirectory) <> "" Then
                caminhoSalvar = novaPasta & "Curvas_Separadas_" & dataArquivo & ".xlsx"
                GravarLogExterno "INFO", "BaixaINETX", "Pasta criada: " & novaPasta
            Else
                caminhoSalvar = ObterCaminhoDesktop() & "Curvas_Separadas_" & dataArquivo & ".xlsx"
                GravarLogExterno "AVISO", "BaixaINETX", "Falha ao criar pasta; salvo no Desktop"
            End If
        Else
            caminhoSalvar = ObterCaminhoDesktop() & "Curvas_Separadas_" & dataArquivo & ".xlsx"
            GravarLogExterno "INFO", "BaixaINETX", "Usuario optou pelo Desktop"
        End If
    Else
        caminhoSalvar = ObterCaminhoDesktop() & "Curvas_Separadas_" & dataArquivo & ".xlsx"
        GravarLogExterno "AVISO", "BaixaINETX", "Pasta MTM nao localizada; salvo no Desktop"
    End If

    ' Abre o CSV
    Set wbCSV = Workbooks.Open(Filename:=caminhoCompleto, Local:=True)
    Set wsCSV = wbCSV.Sheets(1)

    ultimaLinha  = wsCSV.Cells(wsCSV.Rows.Count, COLUNA_FILTRO).End(xlUp).Row
    ultimaColuna = wsCSV.Cells(1, wsCSV.Columns.Count).End(xlToLeft).Column

    If ultimaLinha < 2 Then
        MsgBox "O arquivo CSV esta vazio.", vbExclamation, "Sem Dados"
        wbCSV.Close SaveChanges:=False
        Exit Sub
    End If

    Set rngDados = wsCSV.Range(wsCSV.Cells(1, 1), wsCSV.Cells(ultimaLinha, ultimaColuna))
    If wsCSV.AutoFilterMode Then wsCSV.AutoFilterMode = False

    Set wbNovo = Workbooks.Add

    For i = LBound(arrCurvas) To UBound(arrCurvas)
        rngDados.AutoFilter Field:=COLUNA_FILTRO, Criteria1:="*" & arrCurvas(i) & "*"

        ' Verifica se ha celulas visiveis alem do cabecalho
        Dim temDados As Boolean
        temDados = False
        On Error Resume Next
        temDados = (rngDados.SpecialCells(xlCellTypeVisible).Count > ultimaColuna)
        On Error GoTo 0

        Set wsNova = wbNovo.Worksheets.Add(After:=wbNovo.Worksheets(wbNovo.Worksheets.Count))
        On Error Resume Next
        wsNova.Name = arrCurvas(i)
        On Error GoTo 0

        If temDados Then
            rngDados.SpecialCells(xlCellTypeVisible).Copy Destination:=wsNova.Range("A1")
            wsNova.Columns.AutoFit
        Else
            ' Curva nao encontrada no CSV: aba criada vazia (nao bloqueia)
            GravarLogExterno "AVISO", "BaixaINETX", "Curva '" & arrCurvas(i) & "' nao encontrada no CSV"
        End If
    Next i

    wsCSV.AutoFilterMode = False
    Application.DisplayAlerts = False

    ' Remove abas padrao (Plan1 / Sheet1 / Folha1)
    Dim wsDel As Worksheet
    For Each wsDel In wbNovo.Worksheets
        If wsDel.Name Like "Plan*" Or wsDel.Name Like "Sheet*" Or wsDel.Name Like "Folha*" Then
            If wbNovo.Sheets.Count > 1 Then wsDel.Delete
        End If
    Next wsDel

    wbNovo.SaveAs Filename:=caminhoSalvar, FileFormat:=51
    Application.DisplayAlerts = True

    wbNovo.Close  SaveChanges:=False
    wbCSV.Close   SaveChanges:=False

    GravarLogExterno "OK", "BaixaINETX", "Curvas_Separadas salvo: " & caminhoSalvar
    MsgBox "Arquivo gerado com sucesso!" & vbCrLf & vbCrLf & _
           "Local: " & caminhoSalvar, vbInformation, "Processo Concluido"
End Sub

'===========================================================================
' Move arquivos (placeholder para compatibilidade)
'===========================================================================
Public Sub movarqu()
    ' Conforme informado pelo usuario, esta sub nao influencia no fluxo.
    ' Mantida vazia para evitar erro de "Sub ou Function nao definida".
End Sub

'===========================================================================
' Filtra o CSV do INETX e gera Curvas_Separadas na pasta MTM D-1
' ... (rest of the code remains the same until ObterPastaMTM_D1)
'===========================================================================
' ...

'===========================================================================
' Retorna caminho da pasta MTM D-1 se existir, ou ""
' ALTERACAO v4: agora utiliza a logica centralizada
'===========================================================================
Private Function ObterPastaMTM_D1(dtData As Date, sUserName As String) As String
    Dim raiz As String: raiz = ObterRaizMTM(sUserName)
    Dim base As String: base = ObterPastaBaseComMes(raiz, dtData)
    
    If Len(base) = 0 Then ObterPastaMTM_D1 = "": Exit Function

    Dim nomeDia As String
    nomeDia = "MTM - " & Format(dtData, "DD.MM.YY")

    Dim caminho As String: caminho = base & "\" & nomeDia & "\"
    If Dir(caminho, vbDirectory) <> "" Then
        ObterPastaMTM_D1 = caminho
    Else
        ObterPastaMTM_D1 = ""
    End If
End Function

Private Function ObterBasePastaMTM(sUserName As String, dtData As Date) As String
    ' Esta funcao agora retorna a pasta do MES para compatibilidade com o fluxo antigo
    ObterBasePastaMTM = ObterPastaBaseComMes(ObterRaizMTM(sUserName), dtData)
End Function

Private Function ObterRaizMTM(sUserName As String) As String
    ' Tenta ler do orquestrador se possivel, senao usa fallback dinamico
    Dim raiz As String
    On Error Resume Next
    raiz = Application.Run("'Config_AtualizadorSWAPs.xlsm'!Mod1_Orquestrador.LerParametro", "PastaMTM_PU")
    On Error GoTo 0
    
    If Len(raiz) = 0 Then
        raiz = Environ("USERPROFILE") & "\Empresa\Bloomberg - Documentos\Sakura\MTM & PU\"
    End If
    ObterRaizMTM = raiz
End Function

Private Function ObterCaminhoDesktop() As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ObterCaminhoDesktop = wsh.SpecialFolders("Desktop") & "\"
    Set wsh = Nothing
End Function

Private Function AbrevMesPT(mes As Integer) As String
    Select Case mes
        Case 1:  AbrevMesPT = "JAN"
        Case 2:  AbrevMesPT = "FEV"
        Case 3:  AbrevMesPT = "MAR"
        Case 4:  AbrevMesPT = "ABR"
        Case 5:  AbrevMesPT = "MAI"
        Case 6:  AbrevMesPT = "JUN"
        Case 7:  AbrevMesPT = "JUL"
        Case 8:  AbrevMesPT = "AGO"
        Case 9:  AbrevMesPT = "SET"
        Case 10: AbrevMesPT = "OUT"
        Case 11: AbrevMesPT = "NOV"
        Case 12: AbrevMesPT = "DEZ"
    End Select
End Function

Private Sub GravarLogExterno(status As String, contexto As String, detalhe As String)
    On Error Resume Next
    Dim wb As Workbook
    For Each wb In Workbooks
        If InStr(LCase(wb.Name), "atualizadorswaps") > 0 Then
            Application.Run "'" & wb.Name & "'!Mod1_Orquestrador.GravarLog", _
                            status, contexto, detalhe
            Exit For
        End If
    Next wb
    On Error GoTo 0
    Debug.Print Now & " | " & status & " | " & contexto & " | " & detalhe
End Sub
