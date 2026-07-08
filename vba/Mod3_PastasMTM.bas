Attribute VB_Name = "Mod3_PastasMTM"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 3: Pastas MTM (v4.1)
'
' CORRECAO v4.1:
'   Adicionada CopiarCalculadorasDeD2ParaD1:
'   Quando a pasta D-1 ja existe (criada pelo BaixaArquivosINETX),
'   copia apenas os .xlsx de calculadoras de D-2 para D-1,
'   sem sobrescrever o Curvas_Separadas que ja esta la.
'===========================================================================
Option Explicit

Private Function NormalizarCaminho(caminho As String) As String
    caminho = Trim(caminho)
    If caminho = "" Then
        NormalizarCaminho = ""
        Exit Function
    End If
    caminho = Replace(caminho, "/", "\")
    Do While Right(caminho, 1) = "\"
        caminho = Left(caminho, Len(caminho) - 1)
    Loop
    NormalizarCaminho = caminho
End Function

'===========================================================================
' Procura subpasta correspondente a uma data no diretorio pastaMTM
' Testa formato curto (DD.MM.YY) e longo (DD.MM.YYYY)
' Aceita tambem o caminho direto para a pasta mes/ano ou para a pasta D-2/D-1.
'===========================================================================
Public Function EncontrarPastaData(pastaMTM As String, dataAlvo As Date) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    pastaMTM = NormalizarCaminho(pastaMTM)
    If Not fso.FolderExists(pastaMTM) Then
        GravarLog "ERRO", "MTM", "Pasta base MTM nao encontrada: " & pastaMTM
        EncontrarPastaData = ""
        Exit Function
    End If

    Dim nomeAlvo1 As String: nomeAlvo1 = "MTM - " & Format(dataAlvo, "DD.MM.YY")
    Dim nomeAlvo2 As String: nomeAlvo2 = "MTM - " & Format(dataAlvo, "DD.MM.YYYY")

    Dim pasta As Object
    Set pasta = fso.GetFolder(pastaMTM)

    Dim nomeAtual As String: nomeAtual = LCase(Dir(pastaMTM))
    If nomeAtual = LCase(nomeAlvo1) Or nomeAtual = LCase(nomeAlvo2) Then
        EncontrarPastaData = pastaMTM
        Exit Function
    End If

    Dim sub_ As Object
    For Each sub_ In pasta.SubFolders
        Dim n As String: n = Trim(sub_.Name)
        If LCase(n) = LCase(nomeAlvo1) Or LCase(n) = LCase(nomeAlvo2) Then
            EncontrarPastaData = sub_.Path
            Exit Function
        End If
    Next sub_

    ' Nao encontrou: loga pastas existentes para diagnostico
    Dim listaPastas As String
    Dim i As Integer: i = 0
    For Each sub_ In pasta.SubFolders
        listaPastas = listaPastas & sub_.Name & " | "
        i = i + 1
        If i >= 20 Then listaPastas = listaPastas & "...": Exit For
    Next sub_
    GravarLog "AVISO", "MTM", "Pasta nao encontrada para " & Format(dataAlvo, "dd/mm/yyyy") & _
              ". Existentes: " & Left(listaPastas, 250)
    EncontrarPastaData = ""
End Function

'===========================================================================
' Copia a pasta inteira de origem para destino (D-2 -> D-1 quando D-1 nao existe)
'===========================================================================
Public Sub CopiarPasta(origem As String, destino As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(origem) Then
        Err.Raise vbObjectError + 1, "CopiarPasta", "Pasta origem nao existe: " & origem
    End If
    If fso.FolderExists(destino) Then
        Err.Raise vbObjectError + 2, "CopiarPasta", "Pasta destino ja existe: " & destino
    End If

    fso.CopyFolder origem, destino, False
End Sub

'===========================================================================
' NOVA FUNCAO v4.1 - Copia apenas os .xlsx de calculadoras de D-2 para D-1
'
' Usada quando a pasta D-1 JA EXISTE (criada pelo BaixaArquivosINETX).
' Copia arquivo por arquivo, pulando:
'   - Curvas_Separadas (ja esta em D-1, gerado pelo BaixaArquivosINETX)
'   - Arquivos que ja existem em D-1 (evita sobrescrita acidental)
'
' Parametros:
'   pastaD2 -> caminho completo da pasta D-2 (sem barra no final)
'   pastaD1 -> caminho completo da pasta D-1 (sem barra no final)
'===========================================================================
Public Sub CopiarCalculadorasDeD2ParaD1(pastaD2 As String, pastaD1 As String)
    Dim fso       As Object
    Dim pasta     As Object
    Dim arq       As Object
    Dim copiados  As Integer
    Dim pulados   As Integer
    Dim naoCalc   As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(pastaD2) Then
        GravarLog "ERRO", "MTM", "Pasta D-2 nao encontrada para copia de calculadoras: " & pastaD2
        Exit Sub
    End If
    If Not fso.FolderExists(pastaD1) Then
        GravarLog "ERRO", "MTM", "Pasta D-1 nao encontrada: " & pastaD1
        Exit Sub
    End If

    Set pasta = fso.GetFolder(pastaD2)

    For Each arq In pasta.Files
        Dim ext      As String: ext      = LCase(fso.GetExtensionName(arq.Name))
        Dim nomeMinusc As String: nomeMinusc = LCase(arq.Name)

        ' So copia arquivos .xlsx e .xlsm que sejam calculadoras
        If ext = "xlsx" Or ext = "xlsm" Then
            ' Pula arquivos que nao sao calculadoras
            If InStr(nomeMinusc, "curvas_separadas")   > 0 Or _
               InStr(nomeMinusc, "config_atualizador") > 0 Or _
               InStr(nomeMinusc, "baixaarquivos")      > 0 Then
                naoCalc = naoCalc + 1
            Else
                Dim destArq As String
                destArq = pastaD1 & "\" & arq.Name

                If fso.FileExists(destArq) Then
                    ' Ja existe em D-1: pula para nao sobrescrever
                    pulados = pulados + 1
                    GravarLog "INFO", "MTM", "Ja existe em D-1 (pulado): " & arq.Name
                Else
                    ' Copia para D-1
                    On Error Resume Next
                    fso.CopyFile arq.Path, destArq, False
                    If Err.Number = 0 Then
                        copiados = copiados + 1
                        GravarLog "INFO", "MTM", "Copiado para D-1: " & arq.Name
                    Else
                        GravarLog "AVISO", "MTM", "Falha ao copiar '" & arq.Name & "': " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
    Next arq

    GravarLog "OK", "MTM", "Calculadoras copiadas de D-2 para D-1: " & _
              copiados & " copiadas | " & pulados & " ja existiam | " & naoCalc & " ignoradas"
End Sub

'===========================================================================
' Remove Curvas_Separadas_*.xlsx OBSOLETOS dentro de uma pasta MTM.
' Mantem apenas o do dia (dataAtual). Util para limpar Curvas_Separadas
' antigos que tenham sido herdados de uma copia bruta da pasta D-2.
'
' Parametros:
'   pasta      -> caminho da pasta MTM (com ou sem barra final)
'   dataAtual  -> data D-1 do run; arquivo cujo nome contem essa data fica.
'===========================================================================
Public Sub LimparCurvasObsoletas(pasta As String, dataAtual As Date)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pasta) Then Exit Sub

    Dim stampAtual As String: stampAtual = Format(dataAtual, "YYYYMMDD")
    Dim removidos As Long
    Dim arq As Object
    For Each arq In fso.GetFolder(pasta).Files
        Dim nomeMinusc As String: nomeMinusc = LCase(arq.Name)
        If InStr(nomeMinusc, "curvas_separadas") > 0 Then
            ' So remove se a data no nome NAO for a data atual
            If InStr(nomeMinusc, LCase(stampAtual)) = 0 Then
                On Error Resume Next
                arq.Delete True
                If Err.Number = 0 Then
                    removidos = removidos + 1
                    GravarLog "INFO", "MTM", "Curvas_Separadas obsoleto removido: " & arq.Name
                Else
                    GravarLog "AVISO", "MTM", "Falha ao remover '" & arq.Name & "': " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    Next arq
    If removidos > 0 Then
        GravarLog "OK", "MTM", removidos & " Curvas_Separadas obsoleto(s) removido(s)"
    End If
End Sub

'===========================================================================
' Retorna o nome padrao da pasta (formato confirmado: DD.MM.YY)
'===========================================================================
Public Function NomePastaData(d As Date) As String
    NomePastaData = "MTM - " & Format(d, "DD.MM.YY")
End Function

'===========================================================================
' Auto-preenche coluna B da aba Config varrendo a pasta D-1
'===========================================================================
Public Sub AutoPreencherCaminhosConfig(pastaD1 As String)
    Dim wsConfig As Worksheet
    Dim fso      As Object
    Dim ultLinha As Long
    Dim i        As Long
    Dim encontrados    As Long
    Dim naoEncontrados As Long

    If Len(pastaD1) = 0 Then
        GravarLog "AVISO", "MTM", "Pasta D-1 vazia; auto-preenchimento ignorado"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pastaD1) Then
        GravarLog "AVISO", "MTM", "Pasta D-1 nao existe: " & pastaD1
        Exit Sub
    End If

    Set wsConfig = ThisWorkbook.Sheets(SHEET_CONFIG)
    ultLinha = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row

    ' Coleta todos os .xlsx de calculadoras na pasta D-1
    Dim todosArquivos As Collection
    Set todosArquivos = New Collection
    ColetarXLSX pastaD1, todosArquivos, fso

    GravarLog "INFO", "MTM", "Arquivos .xlsx encontrados na pasta D-1: " & todosArquivos.Count

    If todosArquivos.Count = 0 Then
        GravarLog "AVISO", "MTM", "Nenhuma calculadora encontrada em D-1. " & _
                  "Verifique se CopiarCalculadorasDeD2ParaD1 rodou corretamente."
    End If

    For i = 2 To ultLinha
        Dim nomeContrato As String
        nomeContrato = Trim(wsConfig.Cells(i, 1).Value)
        If Len(nomeContrato) = 0 Then GoTo ProximaLinha

        Dim caminhoEncontrado As String
        caminhoEncontrado = BuscarArquivoPorNome(todosArquivos, nomeContrato)

        If Len(caminhoEncontrado) > 0 Then
            wsConfig.Cells(i, 2).Value = caminhoEncontrado
            encontrados = encontrados + 1
            GravarLog "INFO", "MTM", "Mapeado [" & nomeContrato & "] -> " & _
                      Mid(caminhoEncontrado, InStrRev(caminhoEncontrado, "\") + 1)
        Else
            ' Deixa col B vazia se nao encontrou (Mod5 tratara corretamente)
            If InStr(LCase(Trim(wsConfig.Cells(i, 2).Value)), "auto-preenchido") > 0 Or _
               Len(Trim(wsConfig.Cells(i, 2).Value)) = 0 Then
                wsConfig.Cells(i, 2).Value = ""
            End If
            naoEncontrados = naoEncontrados + 1
            GravarLog "AVISO", "MTM", "Nao encontrado em D-1: " & nomeContrato
        End If
ProximaLinha:
    Next i

    GravarLog "OK", "MTM", "Auto-preenchimento concluido: " & encontrados & _
              " mapeados | " & naoEncontrados & " nao encontrados"
End Sub

'===========================================================================
' Coleta recursivamente .xlsx de calculadoras em uma pasta
'===========================================================================
Public Sub ColetarXLSX(caminhoPasta As String, col As Collection, fso As Object)
    Dim pasta As Object
    Dim arq   As Object
    Dim sub_  As Object

    On Error Resume Next
    Set pasta = fso.GetFolder(caminhoPasta)
    On Error GoTo 0
    If pasta Is Nothing Then Exit Sub

    For Each arq In pasta.Files
        Dim ext As String: ext = LCase(fso.GetExtensionName(arq.Name))
        Dim nm  As String: nm  = LCase(arq.Name)
        If (ext = "xlsx" Or ext = "xlsm") And _
           InStr(nm, "curvas_separadas")   = 0 And _
           InStr(nm, "config_atualizador") = 0 And _
           InStr(nm, "baixaarquivos")      = 0 Then
            col.Add arq.Path
        End If
    Next arq

    For Each sub_ In pasta.SubFolders
        ColetarXLSX sub_.Path, col, fso
    Next sub_
End Sub

'===========================================================================
' Busca arquivo cujo nome contenha o nome do contrato (fuzzy match)
'===========================================================================
Public Function BuscarArquivoPorNome(col As Collection, nomeContrato As String) As String
    Dim item        As Variant
    Dim nomeNorm    As String
    Dim nomeNoSp    As String
    Dim arqBase     As String
    Dim arqNorm     As String
    Dim arqNoSp     As String

    nomeNorm = LCase(Trim(nomeContrato))
    nomeNoSp = Replace(nomeNorm, " ", "")

    For Each item In col
        Dim pos As Integer: pos = InStrRev(CStr(item), "\")
        arqBase = LCase(Mid(CStr(item), pos + 1))
        Dim pt  As Integer: pt  = InStrRev(arqBase, ".")
        If pt > 0 Then arqBase = Left(arqBase, pt - 1)

        arqNorm = Replace(Replace(arqBase, "_", " "), "-", " ")
        Do While InStr(arqNorm, "  ") > 0
            arqNorm = Replace(arqNorm, "  ", " ")
        Loop
        arqNoSp = Replace(arqNorm, " ", "")

        If InStr(arqNorm, nomeNorm) > 0 Or InStr(arqNoSp, nomeNoSp) > 0 Then
            BuscarArquivoPorNome = CStr(item)
            Exit Function
        End If
    Next item

    BuscarArquivoPorNome = ""
End Function

'===========================================================================
' Retorna a pasta base mensal: pastaMTM\MTM & PU (YYYY)\MTM & PU - MMM.YY
'
' Esta funcao agora eh robusta: ela verifica se a pastaMTM fornecida
' ja aponta para o ano ou para o mes, evitando duplicacao de subpastas.
'===========================================================================
Public Function ObterPastaBaseComMes(pastaMTM As String, dtData As Date) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    pastaMTM = NormalizarCaminho(pastaMTM)
    If Len(pastaMTM) = 0 Then
        ObterPastaBaseComMes = ""
        Exit Function
    End If

    Dim nomeAno As String: nomeAno = "MTM & PU (" & Year(dtData) & ")"
    Dim nomeMes As String: nomeMes = "MTM & PU - " & AbrevMesPT(Month(dtData)) & "." & Format(dtData, "YY")

    ' Se o caminho ja termina com o mes correto, retorna ele
    If InStr(1, pastaMTM, nomeMes, vbTextCompare) > 0 Then
        ObterPastaBaseComMes = pastaMTM
        Exit Function
    End If

    ' Se o caminho termina com o ano, apenda apenas o mes
    If InStr(1, pastaMTM, nomeAno, vbTextCompare) > 0 Then
        ObterPastaBaseComMes = pastaMTM & "\" & nomeMes
        Exit Function
    End If

    ' Caso contrario, assume que pastaMTM eh a raiz e apenda Ano + Mes
    ObterPastaBaseComMes = pastaMTM & "\" & nomeAno & "\" & nomeMes
End Function

'===========================================================================
' Cria hierarquia de pastas recursivamente (equivalente a mkdir -p)
'===========================================================================
Public Sub CriarHierarquiaPastas(sPath As String)
    If sPath = "" Then Exit Sub
    Dim aDirs  As Variant: aDirs  = Split(sPath, "\")
    Dim iStart As Integer: iStart = IIf(Left(sPath, 2) = "\\", 3, 1)
    Dim sCur   As String:  sCur   = Left(sPath, InStr(iStart, sPath, "\"))
    Dim i      As Integer
    For i = iStart To UBound(aDirs)
        sCur = sCur & aDirs(i) & "\"
        If Dir(sCur, vbDirectory) = vbNullString Then
            On Error Resume Next
            MkDir sCur
            On Error GoTo 0
        End If
    Next i
End Sub

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

'===========================================================================
' Lista pastas MTM (diagnostico)
'===========================================================================
Public Function ListarPastasMTM(pastaMTM As String) As String
    Dim fso    As Object
    Dim pasta  As Object
    Dim sub_   As Object
    Dim result As String
    Dim i      As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pastaMTM) Then
        ListarPastasMTM = "PASTA NAO ENCONTRADA: " & pastaMTM
        Exit Function
    End If

    Set pasta = fso.GetFolder(pastaMTM)
    For Each sub_ In pasta.SubFolders
        result = result & sub_.Name & vbCrLf
        i = i + 1
        If i >= 30 Then result = result & "...": Exit For
    Next sub_
    ListarPastasMTM = "Pastas em '" & pastaMTM & "':" & vbCrLf & result
End Function
