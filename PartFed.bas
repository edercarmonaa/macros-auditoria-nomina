Attribute VB_Name = "PartFedModule"
Option Explicit

Public Sub PartFed()
Attribute PartFed.VB_ProcData.VB_Invoke_Func = "P\n14"
    Dim f As Long
    Dim bankNames As Variant, bankAccounts(0 To 4) As String
    Dim bankIndex As Long
    Dim existingBanks As Worksheet
    Dim Folder As String
    Dim Archivos As String
    Dim i As Long
    Dim target As Workbook
    Dim logSheet As Worksheet
    Dim oldAlerts As Boolean, oldUpdating As Boolean, oldEvents As Boolean
    Dim oldSecurity As Long
    Dim completed As Boolean
    Dim opened As Workbook
    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldSecurity = Application.AutomationSecurity
    On Error GoTo Failed
    Set logSheet = ThisWorkbook.Worksheets("Hoja1")
    bankNames = Array("SERFIN", "BANCOMER", "BANAMEX", "BANORTE", "HSBC")
    For bankIndex = 0 To 4
        bankAccounts(bankIndex) = Environ$("PARTFED_CUENTA_" & bankNames(bankIndex))
        If Len(bankAccounts(bankIndex)) < 10 Or Len(bankAccounts(bankIndex)) > 18 _
            Or bankAccounts(bankIndex) Like "*[!0-9]*" Then
            MsgBox "Configure las cinco variables PARTFED_CUENTA con digitos (10 a 18).", vbExclamation
            GoTo CleanUp
        End If
    Next bankIndex
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Seleccionar Carpeta"
    .ButtonName = "Aceptar"
        If .Show = -1 Then
             Application.DisplayAlerts = False
             Application.EnableEvents = False
             Application.AutomationSecurity = msoAutomationSecurityForceDisable
            i = 1
            Folder = .SelectedItems(1)
            Archivos = Dir(Folder & "\*.xlsx")
            Do While Archivos <> ""
            Application.ScreenUpdating = False
             If Left$(Archivos, 2) <> "~$" Then
             For Each opened In Application.Workbooks
                 If StrComp(opened.FullName, Folder & "\" & Archivos, vbTextCompare) = 0 Then
                     Err.Raise vbObjectError + 1
                 End If
             Next opened
             Set target = Workbooks.Open(Filename:=Folder & "\" & Archivos, _
                 UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
             If target.ReadOnly Then Err.Raise vbObjectError + 2
             Set existingBanks = Nothing
             On Error Resume Next
             Set existingBanks = target.Worksheets("Bancos")
             On Error GoTo Failed
             If Not existingBanks Is Nothing Then Err.Raise vbObjectError + 3
             target.Worksheets(1).Activate
             Columns("A:A").Select
             Selection.Delete Shift:=xlToLeft
             Columns("F:F").Select
            Selection.Cut
            Columns("B:B").Select
            Selection.Insert Shift:=xlToRight
            Columns("D:D").Select
            Selection.Cut
            Columns("C:C").Select
            Selection.Insert Shift:=xlToRight
            Columns("E:E").Select
            Selection.Cut
            Columns("D:D").Select
            Selection.Insert Shift:=xlToRight
            Columns("H:H").Select
            Selection.Cut
            Columns("E:E").Select
            Selection.Insert Shift:=xlToRight
            Columns("F:F").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("F1").Select
            ActiveCell.FormulaR1C1 = "PERIORICIDAD_PAGO"
            Columns("N:N").Select
            Selection.Cut
            Columns("G:G").Select
            Selection.Insert Shift:=xlToRight
            Columns("P:P").Select
            Selection.Cut
            Columns("H:H").Select
            Selection.Insert Shift:=xlToRight
            Range("H1").Select
            ActiveCell.FormulaR1C1 = "NOM_EMP"
            Columns("P:P").Select
            Selection.Cut
            Columns("I:I").Select
            Selection.Insert Shift:=xlToRight
            ActiveCell.FormulaR1C1 = "NUM_EMP"
            Columns("J:J").Select
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            ActiveWindow.SmallScroll ToRight:=6
            Columns("U:U").Select
            Selection.Cut
            Columns("K:K").Select
            Selection.Insert Shift:=xlToRight
            ActiveWindow.SmallScroll ToRight:=6
            Columns("AA:AA").Select
            Selection.Cut
            ActiveWindow.SmallScroll ToRight:=-7
            Columns("L:L").Select
            Selection.Insert Shift:=xlToRight
            Columns("M:M").Select
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            ActiveWindow.SmallScroll ToRight:=6
            Columns("Z:Z").Select
            Selection.Cut
            ActiveWindow.SmallScroll ToRight:=-5
            Columns("M:M").Select
            Selection.Insert Shift:=xlToRight
            Columns("U:U").Select
            Selection.Cut
            Columns("N:N").Select
            Selection.Insert Shift:=xlToRight
            Columns("P:R").Select
            Selection.Cut
            Columns("O:O").Select
            Selection.Insert Shift:=xlToRight
            Range("O1").Select
            ActiveCell.FormulaR1C1 = "FECHA_INI"
            Range("P1").Select
            ActiveCell.FormulaR1C1 = "FECHA_FIN"
            Range("Q1").Select
            ActiveCell.FormulaR1C1 = "FECHA_PROC"
            Range("Q2").Select
            ActiveWindow.SmallScroll ToRight:=9
            Columns("AA:AC").Select
            Selection.Cut
            ActiveWindow.SmallScroll ToRight:=-6
            Columns("R:R").Select
            Selection.Insert Shift:=xlToRight
            Range("R1").Select
            ActiveCell.FormulaR1C1 = "PERIODO_INI"
            Range("S1").Select
            ActiveCell.FormulaR1C1 = "PERIODOD_FIN"
            Range("T1").Select
            ActiveCell.FormulaR1C1 = "PERIODO_PROC"
            Columns("U:U").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("U1").Select
            ActiveCell.FormulaR1C1 = "DIAS_PAG"
            Range("V1").Select
            ActiveCell.FormulaR1C1 = "ANTIG"
            Columns("W:W").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("W1").Select
            ActiveCell.FormulaR1C1 = "N_QUINQ"
            Range("W2").Select
            ActiveWindow.SmallScroll ToRight:=6
            Columns("X:X").Select
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Selection.Delete Shift:=xlToLeft
            Columns("Z:Z").Select
            Selection.Cut
            Columns("Y:Y").Select
            Selection.Insert Shift:=xlToRight
            Columns("AA:AB").Select
            Selection.Delete Shift:=xlToLeft
            Range("AD1").Select
            ActiveWindow.SmallScroll ToRight:=279
            Range("KR1").Select
            ActiveCell.FormulaR1C1 = "NUM_POL_EGRE"
            Range("KR2").Select
'BUSCARV
            ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Bancos"
            With target.Worksheets("Bancos")
                .Range("B1:B5").NumberFormat = "@"
                For bankIndex = 0 To 4
                    .Cells(bankIndex + 1, 1).Value2 = bankNames(bankIndex)
                    .Cells(bankIndex + 1, 2).Value2 = bankAccounts(bankIndex)
                Next bankIndex
            End With
            Worksheets(1).Select
            f = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
            If f >= 2 Then
            Range("Z2").Select
            ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[1],Bancos!C1:C2,2,FALSE)"
            If f > 2 Then Selection.AutoFill Destination:=Range("Z2:Z" & f)
            Range("Z2:Z" & f).Calculate
            If Application.WorksheetFunction.CountIf(Range("Z2:Z" & f), "#N/A") > 0 Then _
                Err.Raise vbObjectError + 4
            Range("Z2:Z" & f).Select
            Columns("Z:Z").Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            End If
            Sheets("Bancos").Select
                ActiveWindow.SelectedSheets.Delete

             Application.CutCopyMode = False

             target.Close SaveChanges:=True
             Set target = Nothing
             logSheet.Range("A" & i).Value = Archivos
             i = i + 1
             End If
             Archivos = Dir
             Loop
             completed = True
        Else
            'cancel clicked
        End If
    End With
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Close SaveChanges:=False
    Application.CutCopyMode = False
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldUpdating
    On Error GoTo 0
    If completed Then MsgBox "Proceso completado. Revise los archivos y el registro local.", vbInformation
    Exit Sub
Failed:
    MsgBox "Proceso detenido. El archivo actual no se guardara. Los anteriores pueden estar modificados. Revise copias, formato y configuracion.", vbExclamation
    Resume CleanUp
End Sub
