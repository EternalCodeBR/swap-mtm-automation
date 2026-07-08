Attribute VB_Name = "Mod4_INETX"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 4: Localizacao do Curvas_Separadas (v4)
'
' ALTERACOES v4:
'   - Formato da pasta MTM corrigido: DD.MM.YY (ano 2 digitos)
'   - Validacao atualizada para incluir "Daily Compound SOFR"
'     (aba esperada no Curvas_Separadas para contratos que usam essa curva)
'
' PRE-REQUISITO: usuario deve ter rodado BaixaArquivosINETX.xlsm antes.
'===========================================================================
Option Explicit

'===========================================================================
' Localiza Curvas_Separadas_YYYYMMDD.xlsx
' Prioridade: pasta MTM D-1 -> Desktop
'===========================================================================
Public Function LocalizarCurvasSeparadas(dataD1 As Date) As String
    Dim sUserName   As String
    Dim nomeArquivo As String

    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")
    sUserName = objNetwork.UserName
    Set objNetwork = Nothing

    nomeArquivo = "Curvas_Separadas_" & Format(dataD1, "YYYYMMDD") & ".xlsx"

    ' --- 1a tentativa: pasta MTM D-1 ------------------------------------
    Dim pastaMTM As String
    pastaMTM = ObterPastaMTM_D1(dataD1, sUserName)

    If Len(pastaMTM) > 0 Then
        Dim caminhoMTM As String
        caminhoMTM = pastaMTM & nomeArquivo
        If Dir(caminhoMTM) <> "" Then
            GravarLog "OK", "INETX", "Curvas_Separadas localizado em MTM D-1: " & caminhoMTM
            LocalizarCurvasSeparadas = caminhoMTM
            Exit Function
        End If
    End If

    ' --- 2a tentativa: Desktop ------------------------------------------
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim caminhoDesktop As String
    caminhoDesktop = wsh.SpecialFolders("Desktop") & "\" & nomeArquivo
    Set wsh = Nothing

    If Dir(caminhoDesktop) <> "" Then
        GravarLog "AVISO", "INETX", "Curvas_Separadas encontrado no Desktop: " & caminhoDesktop
        LocalizarCurvasSeparadas = caminhoDesktop
        Exit Function
    End If

    GravarLog "ERRO", "INETX", "Curvas_Separadas nao encontrado para " & Format(dataD1, "dd/mm/yyyy")
    LocalizarCurvasSeparadas = ""
End Function

'===========================================================================
' Verifica existencia e orienta usuario caso nao encontre
'===========================================================================
Public Function ValidarCurvasSeparadas(dataD1 As Date) As Boolean
    Dim caminho As String
    caminho = LocalizarCurvasSeparadas(dataD1)

    If Len(caminho) > 0 Then
        ValidarCurvasSeparadas = True
    Else
        Dim sUserName As String
        Dim objNetwork As Object
        Set objNetwork = CreateObject("WScript.Network")
        sUserName = objNetwork.UserName
        Set objNetwork = Nothing

        MsgBox "Curvas_Separadas_" & Format(dataD1, "YYYYMMDD") & ".xlsx nao encontrado." & vbCrLf & vbCrLf & _
               "Locais verificados:" & vbCrLf & _
               "  - MTM & PU (" & Year(dataD1) & ")\MTM & PU - " & AbrevMesPT(Month(dataD1)) & "." & Format(dataD1, "YY") & "\MTM - " & Format(dataD1, "DD.MM.YY") & "\" & vbCrLf & _
               "  - Desktop" & vbCrLf & vbCrLf & _
               "Execute o BaixaArquivosINETX.xlsm primeiro e tente novamente.", _
               vbExclamation, "Curvas_Separadas nao encontrado"
        ValidarCurvasSeparadas = False
    End If
End Function

'===========================================================================
' Valida que o Curvas_Separadas contem as abas necessarias
' Abas obrigatorias: Pre, FWD_USD_BRL, Cupom_SOFR
' Aba opcional (Voxa):         Daily Compound SOFR
'===========================================================================
Public Function ValidarConteudoCurvas(caminhoCurvas As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=caminhoCurvas, UpdateLinks:=0, ReadOnly:=True)
    On Error GoTo 0

    If wb Is Nothing Then
        GravarLog "ERRO", "INETX", "Nao foi possivel abrir: " & caminhoCurvas
        ValidarConteudoCurvas = False
        Exit Function
    End If

    ' Abas obrigatorias
    Dim abasObrigatorias As Variant
    abasObrigatorias = Array("Pre", "FWD_USD_BRL", "Cupom_SOFR")

    Dim abaFaltando As String
    Dim a As Variant
    For Each a In abasObrigatorias
        If Not AbaExiste(wb, CStr(a)) Then
            abaFaltando = abaFaltando & "  - " & CStr(a) & vbCrLf
        End If
    Next a

    ' Aba opcional - apenas loga se estiver faltando (nao bloqueia)
    If Not AbaExiste(wb, "Daily Compound SOFR") Then
        GravarLog "AVISO", "INETX", "Aba 'Daily Compound SOFR' ausente no Curvas_Separadas (necessaria para Voxa)"
    End If

    wb.Close SaveChanges:=False

    If Len(abaFaltando) > 0 Then
        GravarLog "ERRO", "INETX", "Abas obrigatorias faltando: " & Replace(abaFaltando, vbCrLf, " | ")
        MsgBox "O Curvas_Separadas esta incompleto." & vbCrLf & _
               "Abas nao encontradas:" & vbCrLf & abaFaltando & vbCrLf & _
               "Regenere via BaixaArquivosINETX.xlsm.", vbExclamation, "Curvas Incompletas"
        ValidarConteudoCurvas = False
    Else
        GravarLog "OK", "INETX", "Curvas_Separadas validado (abas obrigatorias OK)"
        ValidarConteudoCurvas = True
    End If
End Function

'===========================================================================
' Helper: retorna caminho da pasta de curvas do dia correspondente
' Usa a logica centralizada de Mod3 para respeitar a hierarquia Ano/Mes.
'===========================================================================
Private Function ObterPastaMTM_D1(dtData As Date, sUserName As String) As String
    Dim pastaRaiz As String
    pastaRaiz = LerParametro("PastaMTM_PU")
    
    ' Fallback dinamico caso o parametro esteja vazio
    If Len(pastaRaiz) = 0 Then
        On Error Resume Next
        pastaRaiz = Application.Run("ObterCaminhoBaseUsuario") & "Empresa\Bloomberg - Documentos\Sakura\MTM & PU"
        On Error GoTo 0
    End If
    
    If Len(pastaRaiz) = 0 Then pastaRaiz = "C:\Users\" & sUserName & "\Empresa\Bloomberg - Documentos\Sakura\MTM & PU\"

    Dim baseMes As String
    baseMes = ObterPastaBaseComMes(pastaRaiz, dtData)
    
    Dim nomeDia As String
    nomeDia = "MTM - " & Format(dtData, "DD.MM.YY")

    Dim caminho As String
    caminho = baseMes & "\" & nomeDia & "\"
    
    If Dir(caminho, vbDirectory) <> "" Then
        ObterPastaMTM_D1 = caminho
    Else
        ObterPastaMTM_D1 = ""
    End If
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
