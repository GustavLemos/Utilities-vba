Sub AtualizarDados()
    Dim teste As Integer
    Dim copia As Excel.Range
      
    ActiveWorkbook.RefreshAll
    Application.Wait (Now + TimeValue("0:00:20"))
    'AQUI FAZ O SET DA PLANILHA ATUAL REFERENCIADA PARA COPIAR(COPY)
    Set copia = ThisWorkbook.Worksheets("NOME DA ABA").Range("A2:U10000")
    copia.Copy
    'AQUI ABRE O ARQUIVO/ABA/INTERVALO(RANGE) A SER COLADO(PASTE) COM SELEÇÃO A SO ADICIONAR DADOS SEM SUBSTITUIR OS QUE JÁ ESTÃO LÁ
    Workbooks.Open Filename:="C:\Users\GUSTA\DOCUMENTS\PLANILHA.xlsm"
    Sheets("NOME DA ABA").Activate
    Range("A2:U2").Select
    Cells(4, 2).Select
    Selection.End(xlDown).Select
    Cells(Selection.Row, 2).Select
        :=False, Transpose:=False
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
     :=False, Transpose:=False
    'AQUI SALVA OS DADOS
    Workbooks("MOBILIZAÇÃO_CORREDOR SUDESTE_REV02_20.12.2020 (2020).xlsm").Close SaveChanges:=True

End Sub