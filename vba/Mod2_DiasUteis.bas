Attribute VB_Name = "Mod2_DiasUteis"
'===========================================================================
' ATUALIZADOR DE SWAPS - Modulo 2: Dias Uteis
'
' Calcula dias uteis usando a aba "Feriados_Conf" do workbook Config.
' A aba Feriados_Conf deve ter:
'   - Linha 1: cabecalho (qualquer texto)
'   - Coluna A: datas dos feriados (formato data)
'
' Para popular a aba, copie a coluna de datas da aba "Feriados" de qualquer
' uma das 14 calculadoras de SWAP.
'===========================================================================
Option Explicit

'===========================================================================
' Carrega array com todos os feriados da aba Feriados_Conf
'===========================================================================
Public Function CarregarFeriados() As Date()
    Dim ws As Worksheet
    Dim arr() As Date
    Dim i As Long, j As Long, ult As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_FERIADOS)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Sem aba de feriados -> retorna array de 1 elemento vazio
        ' (o calculo continuara sem feriados; apenas fins de semana serao ignorados)
        GravarLog "AVISO", "DiasUteis", "Aba '" & SHEET_FERIADOS & _
                  "' nao encontrada. Usando apenas fins de semana."
        ReDim arr(0)
        CarregarFeriados = arr
        Exit Function
    End If

    ult = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ult < 2 Then
        ReDim arr(0)
        CarregarFeriados = arr
        Exit Function
    End If

    ReDim arr(1 To ult - 1)
    j = 1
    For i = 2 To ult
        Dim v As Variant
        v = ws.Cells(i, 1).Value
        If IsDate(v) Then
            arr(j) = CDate(v)
            j = j + 1
        End If
    Next i
    ReDim Preserve arr(1 To j - 1)

    GravarLog "INFO", "DiasUteis", j - 1 & " feriados carregados"
    CarregarFeriados = arr
End Function

'===========================================================================
' Retorna a data N dias uteis ANTES de dataRef
' Parametros:
'   dataRef  - data de referencia (normalmente Date = hoje)
'   n        - quantos dias uteis voltar (1=D-1, 2=D-2, ...)
'   feriados - array de datas de feriado (de CarregarFeriados)
'===========================================================================
Public Function DiasUteisVoltar(dataRef As Date, n As Integer, _
                                 feriados() As Date) As Date
    Dim d As Date
    Dim count As Integer
    d = dataRef
    count = 0

    Do While count < n
        d = d - 1
        If EhDiaUtil(d, feriados) Then
            count = count + 1
        End If
    Loop

    DiasUteisVoltar = d
End Function

'===========================================================================
' Retorna a data N dias uteis APOS dataRef
'===========================================================================
Public Function DiasUteisAvancar(dataRef As Date, n As Integer, _
                                  feriados() As Date) As Date
    Dim d As Date
    Dim count As Integer
    d = dataRef
    count = 0

    Do While count < n
        d = d + 1
        If EhDiaUtil(d, feriados) Then
            count = count + 1
        End If
    Loop

    DiasUteisAvancar = d
End Function

'===========================================================================
' Verifica se uma data e dia util
' Retorna False se for sabado, domingo ou feriado do array
'===========================================================================
Public Function EhDiaUtil(d As Date, feriados() As Date) As Boolean
    ' Fim de semana (1=domingo, 7=sabado na convencao VBA com vbSunday)
    Dim dw As Integer
    dw = Weekday(d, vbSunday)
    If dw = 1 Or dw = 7 Then   ' domingo=1, sabado=7
        EhDiaUtil = False
        Exit Function
    End If

    ' Verifica feriado
    If UBound(feriados) >= LBound(feriados) Then
        Dim i As Long
        For i = LBound(feriados) To UBound(feriados)
            If feriados(i) = d Then
                EhDiaUtil = False
                Exit Function
            End If
        Next i
    End If

    EhDiaUtil = True
End Function

'===========================================================================
' Conta quantos dias uteis ha entre duas datas (inclusivo em startDate)
'===========================================================================
Public Function ContarDiasUteis(startDate As Date, endDate As Date, _
                                  feriados() As Date) As Long
    Dim d As Date
    Dim count As Long
    count = 0
    For d = startDate To endDate
        If EhDiaUtil(d, feriados) Then count = count + 1
    Next d
    ContarDiasUteis = count
End Function
