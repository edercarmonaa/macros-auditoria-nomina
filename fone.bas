Attribute VB_Name = "FoneModule"
Option Explicit

Public Sub FONE()
Attribute FONE.VB_ProcData.VB_Invoke_Func = "F\n14"
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
             target.Worksheets(1).Activate
            Columns("C:C").Select
            Selection.Delete Shift:=xlToLeft
            Columns("H:L").Select
            Selection.Cut
            Columns("D:D").Select
            Selection.Insert Shift:=xlToRight
            Columns("AH:AH").Select
            Selection.Cut
            ActiveWindow.ScrollColumn = 19
            ActiveWindow.ScrollColumn = 17
            ActiveWindow.ScrollColumn = 15
            ActiveWindow.ScrollColumn = 14
            ActiveWindow.ScrollColumn = 10
            ActiveWindow.ScrollColumn = 9
            ActiveWindow.ScrollColumn = 8
            ActiveWindow.ScrollColumn = 7
            ActiveWindow.ScrollColumn = 6
            ActiveWindow.ScrollColumn = 5
            ActiveWindow.ScrollColumn = 3
            ActiveWindow.ScrollColumn = 2
            ActiveWindow.ScrollColumn = 1
            Columns("I:I").Select
            Selection.Insert Shift:=xlToRight
            Range("I1").Select
            ActiveCell.FormulaR1C1 = "IMP_DEP"
            Columns("AB:AB").Select
            Selection.Cut
            ActiveWindow.SmallScroll ToRight:=-10
            Columns("J:J").Select
            Selection.Insert Shift:=xlToRight
            Columns("K:K").Select
            ActiveWindow.SmallScroll ToRight:=13
            Columns("AE:AE").Select
            Selection.Cut
            ActiveWindow.ScrollColumn = 16
            ActiveWindow.ScrollColumn = 14
            ActiveWindow.ScrollColumn = 13
            ActiveWindow.ScrollColumn = 12
            ActiveWindow.ScrollColumn = 11
            ActiveWindow.ScrollColumn = 10
            ActiveWindow.ScrollColumn = 9
            ActiveWindow.ScrollColumn = 8
            ActiveWindow.ScrollColumn = 6
            ActiveWindow.ScrollColumn = 5
            ActiveWindow.ScrollColumn = 4
            ActiveWindow.ScrollColumn = 2
            ActiveWindow.ScrollColumn = 1
            Columns("L:L").Select
            Selection.Insert Shift:=xlToRight
            ActiveWindow.SmallScroll ToRight:=16
            Columns("AF:AF").Select
            Selection.Cut
            ActiveWindow.SmallScroll ToRight:=-13
            Columns("M:M").Select
            Selection.Insert Shift:=xlToRight
            Columns("Q:Q").Select
            Selection.Cut
            Columns("N:N").Select
            Selection.Insert Shift:=xlToRight
            Columns("O:LE").Select
            Selection.Delete Shift:=xlToLeft
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
