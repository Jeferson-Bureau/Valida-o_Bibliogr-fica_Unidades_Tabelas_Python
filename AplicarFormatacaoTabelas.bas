Attribute VB_Name = "ModFormatacaoTabelas"
Option Explicit

' ============================================================
'  Aplica o padrão de tabela científica de "três linhas"
'  (Acta Scientiarum) a TODAS as tabelas do documento ativo:
'  - Largura = 100% da área de texto
'  - Sem bordas verticais e sem linhas internas entre as
'    linhas de dados; apenas 3 linhas horizontais (topo,
'    sob o cabeçalho, base), simples, 0,5 pt, preta
'  - Fonte PT Serif 8 pt, centralizado horizontal/vertical
'  - Sem sombreamento de célula
'  - Margem interna lateral de 0,05 pol.
'  - Altura de linha automática
'
'  COMO USAR:
'    1) Abra o documento no Word.
'    2) Alt+F11 -> Inserir -> Módulo -> cole todo este código.
'    3) Feche o editor e rode a macro "AplicarFormatacaoTabelas"
'       (Exibir > Macros, ou F5 com o cursor dentro da sub).
'    4) Salve o documento.
' ============================================================

' ---- Parâmetros de formatação (edite aqui se precisar) ----
Const NOME_FONTE As String = "PT Serif"
Const TAM_FONTE As Single = 8              ' pontos (corpo e cabeçalho)
Const MARGEM_LATERAL_POL As Single = 0.05  ' polegadas (~70 twips)
Const COR_BORDA As Long = wdColorAutomatic
Const ESPESSURA_BORDA As Long = wdLineWidth050pt ' 0,5 pt

' ------------------------------------------------------------
' Macro principal: percorre todas as tabelas do documento
' ------------------------------------------------------------
Sub AplicarFormatacaoTabelas()
    Dim tbl As Table
    Dim contador As Integer
    Dim totalErros As Integer
    contador = 0
    totalErros = 0

    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "Nenhuma tabela encontrada neste documento.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each tbl In ActiveDocument.Tables
        If FormatarTabelaEstiloActa(tbl) Then
            contador = contador + 1
        Else
            totalErros = totalErros + 1
        End If
    Next tbl

    Application.ScreenUpdating = True

    If totalErros = 0 Then
        MsgBox contador & " tabela(s) formatada(s) com sucesso.", vbInformation
    Else
        MsgBox contador & " tabela(s) formatada(s); " & totalErros & _
               " tabela(s) apresentaram erro e foram ignoradas.", vbExclamation
    End If
End Sub

' ------------------------------------------------------------
' Aplica a uma única tabela toda a especificação de layout.
' Retorna True se concluiu sem erro fatal, False caso contrário.
' ------------------------------------------------------------
Function FormatarTabelaEstiloActa(tbl As Table) As Boolean
    Dim lin As Row
    Dim cel As Cell
    Dim ultimaLinha As Long

    On Error GoTo Falhou

    ' --- 1) Largura da tabela = 100% da área de texto ---
    On Error Resume Next
    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100
    On Error GoTo 0

    ' --- 2) Margens internas de célula / espaçamento ---
    tbl.TopPadding = 0
    tbl.BottomPadding = 0
    tbl.LeftPadding = InchesToPoints(MARGEM_LATERAL_POL)
    tbl.RightPadding = InchesToPoints(MARGEM_LATERAL_POL)
    tbl.Spacing = 0

    ' --- 3) Remove TODAS as bordas (internas e externas) ---
    tbl.Borders.InsideLineStyle = wdLineStyleNone
    tbl.Borders.OutsideLineStyle = wdLineStyleNone

    ' --- 4) Recria o padrão de "3 linhas": topo, sob o
    '        cabeçalho e base da tabela ---
    ultimaLinha = tbl.Rows.Count

    With tbl.Rows(1).Borders(wdBorderTop)
        .LineStyle = wdLineStyleSingle
        .LineWidth = ESPESSURA_BORDA
        .Color = COR_BORDA
    End With
    With tbl.Rows(1).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = ESPESSURA_BORDA
        .Color = COR_BORDA
    End With
    With tbl.Rows(ultimaLinha).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = ESPESSURA_BORDA
        .Color = COR_BORDA
    End With

    ' --- 5) Altura de linha automática (sem altura fixa) ---
    For Each lin In tbl.Rows
        On Error Resume Next
        lin.HeightRule = wdRowHeightAuto
        On Error GoTo 0
    Next lin

    ' --- 6) Fonte, alinhamento e sombreamento, célula a célula ---
    For Each cel In tbl.Range.Cells
        On Error Resume Next
        With cel.Range
            .Font.Name = NOME_FONTE
            .Font.NameFarEast = NOME_FONTE
            .Font.Size = TAM_FONTE
            .Font.Color = wdColorAutomatic
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.LeftIndent = 0
        End With
        cel.VerticalAlignment = wdCellAlignVerticalCenter
        cel.Shading.Texture = wdTextureNone
        cel.Shading.BackgroundPatternColor = wdColorAutomatic
        cel.Shading.ForegroundPatternColor = wdColorAutomatic
        On Error GoTo 0
    Next cel

    FormatarTabelaEstiloActa = True
    Exit Function

Falhou:
    FormatarTabelaEstiloActa = False
End Function
