Attribute VB_Name = "Módulo1"
Option Explicit
Dim firstSheetName As String
Dim DownloadedSheet As String
Dim GeneratedTable As String

Sub Main()

    Workbooks(1).Activate
    If IsEmpty(Sheets(1).Range("A1").Value) = True Then
        GeneratedTable = ActiveWorkbook.Name
        DownloadedSheet = Workbooks(2).Name
    Else
        DownloadedSheet = ActiveWorkbook.Name
        GeneratedTable = Workbooks(2).Name
    End If
    
    Workbooks(DownloadedSheet).Activate
    firstSheetName = Sheets(1).Name 'Gets first sheet name because it keeps changing for every download
    Workbooks(GeneratedTable).Activate
    Workbooks(GeneratedTable).Sheets.Add.Name = "Tabela"
    Sheets("Tabela").Range("A1").Value = "Hora"
    Sheets("Tabela").Range("B1").Value = "Tipo"
    Sheets("Tabela").Range("C1").Value = "Vara"
    Sheets("Tabela").Range("D1").Value = "Número/Pasta/Jurisdicionados"
    Sheets("Tabela").Range("E1").Value = "Observações" 'new table is created
    
    Workbooks(DownloadedSheet).Sheets(firstSheetName).Activate
    Range("A2:M2").Select 'I don't know what any of this does, I just stole it from a macro
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets(firstSheetName).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(firstSheetName).Sort.SortFields.Add2 Key:=Range( _
        "A2:A104"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(firstSheetName).Sort
        .SetRange Range("A2:M104")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With 'sorts original table by alphabetical order, so hearings come first since they have numbers on them
    AudFinder 'finds hearings by looking for the words AUDIÊNCIA or PERÍCIA
End Sub
    
Sub AudFinder()

    Dim substrings() As String
    Dim rowAddress As Integer 'rows on the original sheet
    Dim rowCounter As Integer 'rows on the generated sheet
    Dim i As Integer
    
    rowCounter = 1
    
    Workbooks(DownloadedSheet).Sheets(firstSheetName).Range("A2").Select
    While IsNumeric(Left(ActiveCell, 1)) = True 'checks if the leftmost character is numeric
        substrings = Split(ActiveCell, " ") 'splits the entire string everywhere there is a space, creating an array
        rowAddress = ActiveCell.Row 'gets current row address
        For i = 0 To UBound(substrings)
            If StrComp(substrings(i), "AUDIÊNCIA", 1) = 0 Or StrComp(substrings(i), "PERÍCIA", 1) = 0 Then 'i looks wether it's a hearing or not
                rowCounter = rowCounter + 1
                Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 1).Value = substrings(0) 'pastes onto new table the first substring (which we hope is the substrings)
                
                TipoAud i, substrings, rowCounter 'finds hearing types
                VaraNPJ rowCounter, rowAddress 'finds process number, file number and involved parties, as well as vara name from process number using the database

            End If
        Next i
        rowAddress = rowAddress + 1
        Workbooks(DownloadedSheet).Sheets(firstSheetName).Activate
        Workbooks(DownloadedSheet).Sheets(firstSheetName).Cells(rowAddress, 1).Select 'goes back to original sheet and goes to next line
    Wend
    Workbooks(GeneratedTable).Sheets("Tabela").Columns("A:E").EntireColumn.AutoFit 'fits to length
    
End Sub
Sub TipoAud(i, substrings, rowCounter)

Dim j As Integer
Dim k As Integer

        For j = i To UBound(substrings) 'j looks for the type of hearing
                If StrComp(substrings(i), "AUDIÊNCIA", 1) = 0 Then
                    If StrComp(substrings(j), "CONCILIAÇÃO", 1) = 0 Then 'if it's conciliação, it writes that, but it might be conciliação e saneamento
                        Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "CONCILIAÇÃO"
                        For k = j To UBound(substrings) 'k looks for saneamento
                            If StrComp(substrings(k), "SANEAMENTO", 1) = 0 Then
                                Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "CONCILIAÇÃO E SANEAMENTO"
                                k = UBound(substrings)
                                j = UBound(substrings)
                                i = UBound(substrings)
                            End If
                        Next k
                    ElseIf StrComp(substrings(j), "ENCERRAMENTO", 1) = 0 Then 'now for instruction ending hearings
                        Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "ENCERRAMENTO DA INSTRUÇÃO"
                        j = UBound(substrings)
                        i = UBound(substrings)
                    ElseIf StrComp(substrings(j), "INAUGURAL", 1) = 0 Then 'now for inaugural
                        Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "INAUGURAL"
                        j = UBound(substrings)
                        i = UBound(substrings)
                    ElseIf StrComp(substrings(j), "INICIAL", 1) = 0 Then 'same
                        Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "INICIAL"
                        j = UBound(substrings)
                        i = UBound(substrings)
                    ElseIf StrComp(substrings(j), "INSTRUÇÃO", 1) = 0 Then
                        Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "INSTRUÇÃO" 'you got it now
                        For k = j To UBound(substrings) 'k looks for civil
                            If StrComp(substrings(k), "CÍVEL", 1) = 0 Then
                                Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "INSTRUÇÃO CÍVEL"
                                k = UBound(substrings)
                                j = UBound(substrings)
                                i = UBound(substrings)
                            End If
                        Next k
                    End If
                ElseIf StrComp(substrings(i), "PERÍCIA", 1) = 0 Then 'same but for perícia
                Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 2).Value = "PERÍCIA"
                i = UBound(substrings)
                End If
        Next j
End Sub
Sub VaraNPJ(rowCounter, rowAddress) 'Vara and Número/Pasta/Jurisdicionados
    
    Dim VaraNum As Integer
    Dim ProcessoPasta As String
    Dim Envolvido As String
    Dim Cliente As String
    Dim VaraPasta() As String
    
    ProcessoPasta = Workbooks(DownloadedSheet).Sheets(firstSheetName).Cells(rowAddress, 3).Value 'gets the process number/file number cell from the downloaded sheet
    Cliente = Workbooks(DownloadedSheet).Sheets(firstSheetName).Cells(rowAddress, 10).Value 'gets client name
    Envolvido = Workbooks(DownloadedSheet).Sheets(firstSheetName).Cells(rowAddress, 11).Value 'gets other party name
    Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 4).Value = ProcessoPasta + " / " + "CLIENTE: " + Cliente + " / ENVOLVIDO: " + Envolvido 'big concatenation for formatting purposes
    VaraPasta = Split(ProcessoPasta, " ")
    VaraNum = CInt(Right(VaraPasta(0), 4))
    Workbooks(GeneratedTable).Sheets("Base de dados").Activate
    Workbooks(GeneratedTable).Sheets("Base de dados").Range("A2").Select 'checks the database
    While IsEmpty(ActiveCell.Value) = False 'while the list goes on
        If ActiveCell.Value = VaraNum Then 'check if number corresponds
            Workbooks(GeneratedTable).Sheets("Tabela").Cells(rowCounter, 3).Value = ActiveCell.Offset(0, 1) 'get corresponding vara name
            GoTo Line128 'if so, just end the loop
        End If
        ActiveCell.Offset(1, 0).Activate
    Wend
Line128:
End Sub

