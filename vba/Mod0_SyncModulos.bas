Attribute VB_Name = "Mod0_SyncModulos"
'==========================================================================
' ATUALIZADOR DE SWAPS - Modulo 0: Sincronizacao de modulos VBA (v1.0)
'
' OBJETIVO
'   Re-importar automaticamente os .bas atualizados do disco para dentro
'   do Config_AtualizadorSWAPs.xlsm. Resolve o problema de "editei o
'   .bas no disco mas o .xlsm continua com versao antiga".
'
'   Substitui o passo manual:
'     Alt+F11 > Project Explorer > [click direito em cada modulo] >
'     Remover Modulo > Arquivo > Importar Arquivo > selecionar .bas
'   ...por um unico clique na macro SincronizarModulosVBA.
'
'   Este modulo NAO faz parte do fluxo de atualizacao das calculadoras;
'   eh uma ferramenta de manutencao do proprio projeto VBA.
'
' USO TIPICO
'   1. Edita-se qualquer Mod*.bas no disco (Claude, outro dev, etc.)
'   2. Abre o Config_AtualizadorSWAPs.xlsm
'   3. Roda SincronizarModulosVBA (Alt+F8 > selecionar > Executar,
'      ou clica num botao Forms ligado a essa macro)
'   4. Confere o relatorio na MsgBox final
'   5. Alt+F11 > Depurar > Compilar VBAProject para validar
'
' PRE-REQUISITO (configurar 1 vez por usuario, sem admin)
'   Excel > Arquivo > Opcoes > Centro de Confiabilidade >
'     Configuracoes do Centro de Confiabilidade > Configuracoes de Macro >
'     [marcar] "Confiar no acesso ao modelo de objeto do projeto VBA"
'
' INSTALACAO INICIAL DESTE MODULO
'   Este modulo precisa estar dentro do .xlsm para funcionar. Importe-o
'   UMA UNICA VEZ manualmente pelo VBE:
'     Alt+F11 > Arquivo > Importar Arquivo > Mod0_SyncModulos.bas
'   A partir dai, ele mesmo cuida das atualizacoes seguintes dos outros
'   modulos (Mod1..Mod6). Para atualizar este modulo no futuro, repita
'   a importacao manual (ele nao se auto-atualiza para evitar crash).
'
' PARAMETROS OPCIONAIS (aba Parametros do Config, coluna A = chave)
'   PastaCodigo  -> path da pasta com os .bas. Default: pasta abaixo
'                   com {USER} substituido pelo usuario do Windows.
'==========================================================================
Option Explicit

' Pasta padrao dos .bas (SharePoint Bloomberg, sincronizado via OneDrive).
' Usado como ultimo fallback se nenhum candidato automatico for encontrado.
' Pode ser sobrescrito pelo parametro "PastaCodigo" na aba Parametros do Config.
' O placeholder {USER} eh substituido pelo nome do usuario do Windows em runtime.
Private Const PASTA_CODIGO_DEFAULT As String = _
    "C:\Users\{USER}\Empresa - Bloomberg\Documentos Compartilhados\Sakura\Código"

' Lista de modulos a sincronizar, em ordem de import.
' Mod0_SyncModulos (este) NAO esta na lista de proposito: ele nao pode
' se remover enquanto esta rodando (causa crash do VBE). Para atualizar
' este modulo, faca import manual.
Private Const MODULOS_SYNC As String = _
    "Mod1_Orquestrador;Mod2_DiasUteis;Mod3_PastasMTM;Mod4_INETX;" & _
    "Mod5_Calculadoras;Mod6_Bloomberg;Mod7_IncluirCalculadora;" & _
    "Mod8_DownloadINETX"

'==========================================================================
' MACRO PRINCIPAL - execute esta
'==========================================================================
Public Sub SincronizarModulosVBA()
    On Error GoTo ErrHandler

    ' --- 1. Verifica pre-requisito (trust VBA project model) ------------
    If Not TrustVBAHabilitado() Then
        MsgBox _
            "Acesso ao modelo de objeto do projeto VBA NAO esta habilitado." & _
            vbCrLf & vbCrLf & _
            "Para habilitar (sem precisar de admin):" & vbCrLf & _
            "  Arquivo > Opcoes > Centro de Confiabilidade >" & vbCrLf & _
            "  Configuracoes do Centro de Confiabilidade >" & vbCrLf & _
            "  Configuracoes de Macro > marcar a opcao" & vbCrLf & _
            "  'Confiar no acesso ao modelo de objeto do projeto VBA'" & _
            vbCrLf & vbCrLf & _
            "Feche e reabra o Excel apos habilitar.", _
            vbExclamation, "Pre-requisito faltando"
        Exit Sub
    End If

    ' --- 2. Resolve a pasta dos .bas ------------------------------------
    Dim pastaCodigo As String
    pastaCodigo = ResolverPastaCodigo()
    If Right(pastaCodigo, 1) <> "\" Then pastaCodigo = pastaCodigo & "\"

    If Dir(pastaCodigo, vbDirectory) = "" Then
        MsgBox _
            "Pasta de codigo nao encontrada:" & vbCrLf & _
            "  " & pastaCodigo & vbCrLf & vbCrLf & _
            "Possiveis causas:" & vbCrLf & _
            "  - O SharePoint Bloomberg nao esta sincronizado nesta maquina." & vbCrLf & _
            "    Abra o OneDrive e sincronize a pasta:" & vbCrLf & _
            "    Bloomberg > Documentos Compartilhados > Sakura > Codigo" & vbCrLf & vbCrLf & _
            "  - O caminho de sync e diferente do esperado." & vbCrLf & _
            "    Nesse caso, adicione na aba Parametros a chave" & vbCrLf & _
            "    'PastaCodigo' (coluna A) com o caminho completo (coluna B).", _
            vbCritical, "Pasta invalida"
        Exit Sub
    End If

    ' --- 3. Loop pelos modulos: remove + importa cada um ----------------
    Dim modulos() As String: modulos = Split(MODULOS_SYNC, ";")
    Dim i         As Long
    Dim totalOK   As Long:   totalOK   = 0
    Dim totalErro As Long:   totalErro = 0
    Dim totalSkip As Long:   totalSkip = 0
    Dim relatorio As String: relatorio = ""

    For i = LBound(modulos) To UBound(modulos)
        Dim nomeMod As String: nomeMod = Trim(modulos(i))
        Dim arqBas  As String: arqBas  = pastaCodigo & nomeMod & ".bas"

        If Dir(arqBas) = "" Then
            totalSkip = totalSkip + 1
            relatorio = relatorio & "  [PULADO] " & nomeMod & _
                        "  (arquivo nao existe no disco)" & vbCrLf
        Else
            Dim erroSub As String: erroSub = SubstituirModulo(nomeMod, arqBas)
            If Len(erroSub) = 0 Then
                totalOK = totalOK + 1
                relatorio = relatorio & "  [OK]     " & nomeMod & vbCrLf
            Else
                totalErro = totalErro + 1
                relatorio = relatorio & "  [ERRO]   " & nomeMod & _
                            "  (" & erroSub & ")" & vbCrLf
            End If
        End If
    Next i

    ' --- 4. Salva o workbook se tudo deu certo --------------------------
    If totalErro = 0 Then
        On Error Resume Next
        ThisWorkbook.Save
        On Error GoTo 0
    End If

    ' --- 5. Resumo na tela ----------------------------------------------
    Dim msg As String
    msg = "Sincronizacao concluida." & vbCrLf & vbCrLf & _
          "  OK:     " & totalOK & vbCrLf & _
          "  ERRO:   " & totalErro & vbCrLf & _
          "  PULADO: " & totalSkip & vbCrLf & vbCrLf & _
          relatorio & vbCrLf & _
          "Pasta: " & pastaCodigo & vbCrLf & vbCrLf

    If totalErro = 0 Then
        msg = msg & _
            "Proximo passo: Alt+F11 > Depurar > Compilar VBAProject" & vbCrLf & _
            "(Alt+D depois L) para validar que tudo compila sem erro."
        MsgBox msg, vbInformation, "SincronizarModulosVBA"
    Else
        msg = msg & _
            "Houve erro em algum modulo. Reveja a mensagem acima e tente" & vbCrLf & _
            "importar manualmente pelo VBE > Arquivo > Importar Arquivo."
        MsgBox msg, vbExclamation, "SincronizarModulosVBA"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Erro inesperado durante a sincronizacao:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "SincronizarModulosVBA"
End Sub

'==========================================================================
' HELPERS PRIVADOS
'==========================================================================

' Remove o modulo atual (se existir) e importa o .bas do disco.
' Retorna "" em caso de sucesso, ou a mensagem de erro em caso de falha.
Private Function SubstituirModulo(nomeModulo As String, _
                                  caminhoBas As String) As String
    On Error GoTo Falha

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    ' Remove o modulo existente (se houver). Se nao existir, segue direto
    ' para o import.
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(nomeModulo)
    On Error GoTo Falha
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If

    ' Importa a nova versao do disco
    vbProj.VBComponents.Import caminhoBas

    SubstituirModulo = ""   ' sucesso
    Exit Function

Falha:
    SubstituirModulo = Err.Description
End Function

' Resolve a pasta dos .bas com tres estrategias em cascata:
'   1. Parametro "PastaCodigo" na aba Parametros (override manual)
'   2. Busca automatica: varre USERPROFILE procurando pasta "*Bloomberg*"
'      e testa subcaminhos conhecidos dentro dela (cobre qualquer variacao
'      de maiusculas ou nome de organizacao que o OneDrive possa usar)
'   3. Fallback: PASTA_CODIGO_DEFAULT com {USER} substituido
'
' Implementado sem depender de LerParametro (que esta em Mod1) para
' funcionar mesmo se o Mod1 estiver temporariamente fora do projeto.
Private Function ResolverPastaCodigo() As String

    ' --- 1. Override via aba Parametros ----------------------------------
    Dim raw As String: raw = LerParamSeguro("PastaCodigo")
    If Len(Trim(raw)) > 0 Then
        ResolverPastaCodigo = ExpandirTokens(raw)
        Exit Function
    End If

    ' --- 2. Busca automatica na pasta sincronizada do SharePoint ---------
    ' O OneDrive cria a pasta como "{Org} - Bloomberg" em USERPROFILE.
    ' O nome exato varia por maquina (maiusculas, abreviacao do tenant),
    ' por isso varremos todas as subpastas procurando "bloomberg".
    Dim perfil As String: perfil = Environ$("USERPROFILE")
    If Right(perfil, 1) = "\" Then perfil = Left(perfil, Len(perfil) - 1)

    ' Subcaminhos possiveis dentro da pasta Bloomberg (com e sem acento)
    Dim subPaths(3) As String
    subPaths(0) = "Documentos Compartilhados\Sakura\C" & Chr(243) & "digo"
    subPaths(1) = "Documentos Compartilhados\Sakura\Codigo"
    subPaths(2) = "Sakura\C" & Chr(243) & "digo"
    subPaths(3) = "Sakura\Codigo"

    On Error Resume Next
    Dim fso    As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim rootFo As Object: Set rootFo = fso.GetFolder(perfil)
    Dim subFo  As Object
    Dim k      As Long

    For Each subFo In rootFo.SubFolders
        If InStr(LCase(subFo.Name), "bloomberg") > 0 Then
            For k = 0 To UBound(subPaths)
                Dim candidato As String
                candidato = subFo.Path & "\" & subPaths(k)
                If Dir(candidato, vbDirectory) <> "" Then
                    ResolverPastaCodigo = candidato
                    Set fso = Nothing: Set rootFo = Nothing
                    On Error GoTo 0
                    Exit Function
                End If
            Next k
        End If
    Next subFo
    Set fso = Nothing: Set rootFo = Nothing
    On Error GoTo 0

    ' --- 3. Fallback: PASTA_CODIGO_DEFAULT com {USER} substituido --------
    Dim user   As String
    Dim objNet As Object
    On Error Resume Next
    Set objNet = CreateObject("WScript.Network")
    user = objNet.UserName
    Set objNet = Nothing
    On Error GoTo 0

    If Len(user) = 0 Then user = Environ$("USERNAME")

    ResolverPastaCodigo = Replace(PASTA_CODIGO_DEFAULT, "{USER}", user)
End Function

' Versao standalone de ExpandirCaminho (nao depende de Mod1, pois Mod0
' precisa funcionar mesmo quando Mod1 ainda nao foi sincronizado).
' Resolve %USERPROFILE%, %USERNAME%, %ONEDRIVE%, %DESKTOP% e prefixo ~\.
Private Function ExpandirTokens(s As String) As String
    Dim r As String: r = s
    If Len(r) = 0 Then ExpandirTokens = "": Exit Function
    Dim perfil As String: perfil = Environ$("USERPROFILE")
    If Len(perfil) > 0 And Right(perfil, 1) = "\" Then perfil = Left(perfil, Len(perfil) - 1)
    Dim onedrive As String: onedrive = Environ$("OneDrive")
    If Len(onedrive) > 0 And Right(onedrive, 1) = "\" Then onedrive = Left(onedrive, Len(onedrive) - 1)
    r = Replace(r, "%USERPROFILE%", perfil)
    r = Replace(r, "%UserProfile%", perfil)
    r = Replace(r, "%userprofile%", perfil)
    r = Replace(r, "%USERNAME%", Environ$("USERNAME"))
    r = Replace(r, "%DESKTOP%", perfil & "\Desktop")
    r = Replace(r, "%ONEDRIVE%", onedrive)
    If Left(r, 2) = "~\" Then r = perfil & "\" & Mid(r, 3)
    ExpandirTokens = r
End Function

' Versao standalone de LerParametro (nao depende de Mod1).
Private Function LerParamSeguro(chave As String) As String
    On Error GoTo SemParam

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Parametros")

    Dim i   As Long
    Dim ult As Long: ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ult
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = LCase(Trim(chave)) Then
            LerParamSeguro = Trim(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i

SemParam:
    LerParamSeguro = ""
End Function

' Detecta se "Confiar no acesso ao modelo de objeto do projeto VBA" esta
' habilitado. Se nao estiver, qualquer acesso a ThisWorkbook.VBProject
' levanta erro 1004 ou similar.
Private Function TrustVBAHabilitado() As Boolean
    On Error Resume Next
    Dim n As Long
    n = ThisWorkbook.VBProject.VBComponents.Count
    TrustVBAHabilitado = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
