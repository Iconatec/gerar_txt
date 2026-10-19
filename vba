
   
Sub GerarTXT()

    Dim caminhoExcel As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim linha As Long, coluna As Long
    Dim ultimaLinha As Long, ultimaColuna As Long
    Dim textoLinha As String
    Dim arquivoTXT As Integer
    Dim caminhoArquivoTXT As String
    Dim celulaTexto As String

    ' Abre seletor para escolher o arquivo Excel
    caminhoExcel = Application.GetOpenFilename( _
        FileFilter:="Arquivos do Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="Selecione o arquivo Excel")

    If caminhoExcel = False Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    Set wb = Workbooks.Open(caminhoExcel)
    Set ws = wb.ActiveSheet

    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    caminhoArquivoTXT = Application.GetSaveAsFilename( _
        InitialFileName:="exportacao.txt", _
        FileFilter:="Arquivo de Texto (*.txt), *.txt", _
        Title:="Salvar como")

    If caminhoArquivoTXT = "Falso" Then
        MsgBox "Operação cancelada.", vbInformation
        Exit Sub
    End If

    If Right(caminhoArquivoTXT, 4) <> ".txt" Then
        caminhoArquivoTXT = caminhoArquivoTXT & ".txt"
    End If

    arquivoTXT = FreeFile
    Open caminhoArquivoTXT For Output As #arquivoTXT

    ' Exporta os dados com TAB como delimitador (ignorando linha 1)
    For linha = 2 To ultimaLinha
        textoLinha = ""
        For coluna = 1 To ultimaColuna
            celulaTexto = Trim(ws.Cells(linha, coluna).Text)
            If coluna = ultimaColuna Then
                textoLinha = textoLinha & celulaTexto
            Else
                textoLinha = textoLinha & celulaTexto & vbTab
            End If
        Next coluna
        Print #arquivoTXT, textoLinha
    Next linha

    Close #arquivoTXT
    ' wb.Close SaveChanges:=False ' Mantém o Excel aberto

    MsgBox "Arquivo exportado com sucesso para:" & vbCrLf & caminhoArquivoTXT, vbInformation

End Sub
