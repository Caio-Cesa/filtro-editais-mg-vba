Attribute VB_Name = "Filtro"
Option Explicit
Public FiltroMG As Boolean
Private Sub Filtro_MG()
Let FiltroMG = True
'Macro faz tratamento de dados do edital e tranforma em uma lista pesquisando apenas os dados de MG
'Macro desenvolvido e propriedade intelectual direcionadas a Nome Empresarial:CAIO CESAR DE ALBUQUERQUE CNPJ:36.611.073/0001-67
'Este produto é final e não é permitida a recomercialização do mesmo
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim w As Worksheet
Set w = Sheets("Controle")
Range("2:1048576").ClearContents
Range("A1").Select
If Range("A1").Value = "" Then
    MsgBox ("Célula A1 está vazia, favor colocar os dados!"), vbOKOnly, "Aviso de erro!"
    w.Visible = True
    w.Select
    Range("L1").Value = Range("L1").Value + 1
    w.Visible = xlSheetVeryHidden
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End
End If
Range("B1").FormulaLocal = "=SUBSTITUIR(A1;CARACT(10);"" "")"
Range("B1").Copy
Range("C1").PasteSpecial Paste:=xlPasteValues
Range("C1").Select
Application.CutCopyMode = False
Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False, Space:=False, Other:=False
Range("A2").Select
Dim coluna_fim As Integer
Dim linha_fim As Integer
Dim ALLcount As Integer
Dim MGcount As Integer
coluna_fim = Range("A2").End(xlToRight).Column
If coluna_fim = 1 Then
    Range("A2").Copy
    Range("A3").PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Range("2:2").ClearContents
    Range("B1:I1").ClearContents
    ALLcount = Application.WorksheetFunction.CountA(Range("A3:A1048576"))
    Range("B3").FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR("":"";A3;1)+2;PROCURAR(""("";A3;1)-PROCURAR("":"";A3;1)-3);EXT.TEXTO(A3;2;PROCURAR(""("";A3;1)-2))"
    Range("C3").FormulaLocal = "=EXT.TEXTO(A3;PROCURAR(""NB:"";A3;1)+4;PROCURAR("","";A3;PROCURAR(""NB:"";A3;1))-PROCURAR(""NB:"";A3;1)-4)"
    Range("D3").FormulaLocal = "=EXT.TEXTO(A3;PROCURAR(""CPF:"";A3;1)+5;11)"
    Range("E3").FormulaLocal = "=EXT.TEXTO(D3;9;1)"
    Range("F3").FormulaLocal = "=SEERRO(SEERRO(EXT.TEXTO(A3;PROCURAR(""Protocolo:"";A3;1)+11;PROCURAR("","";A3;PROCURAR(""Protocolo:"";A3;1))-PROCURAR(""Protocolo:"";A3;1)-11);EXT.TEXTO(A3;PROCURAR(""Protocolo:"";A3;1)+11;PROCURAR("")"";A3;1)-PROCURAR(""Protocolo:"";A3;1)-11));""Sem PROT"")"
    Range("G3").FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR(""Representante Legal:"";A3;1)+21;PROCURAR("","";A3;PROCURAR(""Representante Legal:"";A3;1))-PROCURAR(""Representante Legal:"";A3;1)-21);"""")"
    Range("H3").FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR(""CPF "";A3;1)+4;11);"""")"
    GoTo continua
End If
Range("A2", Selection.End(xlToRight)).Select
Selection.Copy
Range("A3").PasteSpecial Paste:=xlPasteValues, Transpose:=True
Range("2:2").ClearContents
Range("B1:I1").ClearContents
ALLcount = Application.WorksheetFunction.CountA(Range("A3:A1048576")) - 2
linha_fim = Range("A3").End(xlDown).Row
Range("B3:B" & linha_fim).FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR("":"";A3;1)+2;PROCURAR(""("";A3;1)-PROCURAR("":"";A3;1)-3);EXT.TEXTO(A3;2;PROCURAR(""("";A3;1)-2))"
Range("C3:C" & linha_fim).FormulaLocal = "=EXT.TEXTO(A3;PROCURAR(""NB:"";A3;1)+4;PROCURAR("","";A3;PROCURAR(""NB:"";A3;1))-PROCURAR(""NB:"";A3;1)-4)"
Range("D3:D" & linha_fim).FormulaLocal = "=EXT.TEXTO(A3;PROCURAR(""CPF:"";A3;1)+5;11)"
Range("E3:E" & linha_fim).FormulaLocal = "=EXT.TEXTO(D3;9;1)"
Range("F3:F" & linha_fim).FormulaLocal = "=SEERRO(SEERRO(EXT.TEXTO(A3;PROCURAR(""Protocolo:"";A3;1)+11;PROCURAR("","";A3;PROCURAR(""Protocolo:"";A3;1))-PROCURAR(""Protocolo:"";A3;1)-11);EXT.TEXTO(A3;PROCURAR(""Protocolo:"";A3;1)+11;PROCURAR("")"";A3;1)-PROCURAR(""Protocolo:"";A3;1)-11));""Sem PROT"")"
Range("G3:G" & linha_fim).FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR(""Representante Legal:"";A3;1)+21;PROCURAR("","";A3;PROCURAR(""Representante Legal:"";A3;1))-PROCURAR(""Representante Legal:"";A3;1)-21);"""")"
Range("H3:H" & linha_fim).FormulaLocal = "=SEERRO(EXT.TEXTO(A3;PROCURAR(""CPF "";A3;1)+4;11);"""")"
continua:
Range("B2").FormulaLocal = "NOME"
Range("C2").FormulaLocal = "NB"
Range("D2").FormulaLocal = "CPF"
Range("E2").FormulaLocal = "CÓDIGO UF"
Range("F2").FormulaLocal = "PROTOCOLO"
Range("G2").FormulaLocal = "NOME REPRESENTANTE"
Range("H2").FormulaLocal = "CPF REPRESENTANTE"
Range("C2").Select
Selection.AutoFilter
ActiveSheet.Range("$A$2:$H$" & linha_fim + 1).AutoFilter Field:=5, Criteria1:="6"
linha_fim = Range("A3").End(xlDown).Row
Range("B2:H" & linha_fim).Copy
Range("K2").PasteSpecial Paste:=xlPasteValues
Range("A2:H10000").ClearContents
Range("A2:H10000").ClearContents
Selection.Cut Range("A2")
Range("A2").Select
MGcount = Application.WorksheetFunction.CountA(Range("A3:A1048576"))
MsgBox ("Dos " & ALLcount & " CPF'S apenas " & MGcount & " são de MG!"), vbOKOnly, "Aviso de relatório"
w.Visible = True
w.Select
Range("I1").Value = Range("I1").Value + 1
Range("I2").Value = Range("I2").Value + ALLcount
Range("I3").Value = Range("I3").Value + MGcount
w.Visible = xlSheetVeryHidden
Let FiltroMG = False
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
