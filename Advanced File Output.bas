Attribute VB_Name = "Module1"
Option Explicit

' --- Constants ---
Const swDocDRAWING       As Long = 3
Const swSaveVerCurrent   As Long = 0
Const swSaveOpSilent     As Long = 1

' DXF export prefs (explicit company-wide defaults)
Const DXF_FORMAT         As Long = swDxfFormat_R2013
Const DXF_MULTISHEET     As Long = swDxfMultiSheet

' --- Helper: read custom prop via Get5 ---
Private Function TryGetProp5( _
    mgr As SldWorks.CustomPropertyManager, _
    propName As String, _
    ByRef valOut As String _
) As Boolean
    Dim rawVal    As String
    Dim resVal    As String
    Dim wasRes    As Boolean
    Dim ret       As Long

    ret = mgr.Get5(propName, False, rawVal, resVal, wasRes)
    If ret <> 0 Then
        valOut = IIf(Len(resVal) > 0, resVal, rawVal)
        TryGetProp5 = True
    Else
        TryGetProp5 = False
    End If
End Function

' --- Helper: check file exists on disk ---
Private Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

' --- Helper: sanitize filename for Windows ---
Private Function SanitizeFileName(ByVal nameIn As String) As String
    Dim s As String
    Dim ch As Variant
    s = nameIn

    ' replace invalid characters
    For Each ch In Array("\\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "_")
    Next ch

    ' normalize whitespace/newlines
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")

    ' trim trailing dots/spaces (invalid in Windows filenames)
    Do While Len(s) > 0 And (Right$(s, 1) = " " Or Right$(s, 1) = ".")
        s = Left$(s, Len(s) - 1)
    Loop

    If Len(s) = 0 Then s = "Untitled"
    SanitizeFileName = s
End Function

' --- Helper: perform SaveAs4 + verify ---
Private Function SaveAndCheck( _
    doc As SldWorks.ModelDoc2, _
    path As String, _
    fmtName As String _
) As Boolean
    Dim errs   As Long, warns As Long
    doc.SaveAs4 path, swSaveVerCurrent, swSaveOpSilent, errs, warns
    If errs <> 0 Or Not FileExists(path) Then
        MsgBox fmtName & " export failed." & vbCrLf & _
               "Errors: " & errs & "  Warnings: " & warns, vbCritical
        SaveAndCheck = False
    Else
        SaveAndCheck = True
    End If
End Function

' --- Helper: DXF export with explicit preferences (restored after) ---
Private Function SaveDxfWithPrefs( _
    swApp As SldWorks.SldWorks, _
    doc As SldWorks.ModelDoc2, _
    path As String _
) As Boolean
    Dim errs   As Long, warns As Long
    Dim prevVer As Long, prevMulti As Long

    On Error Resume Next
    prevVer = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfVersion)
    prevMulti = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swDxfMultiSheetOption)

    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfVersion, DXF_FORMAT
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, DXF_MULTISHEET
    On Error GoTo 0

    doc.SaveAs4 path, swSaveVerCurrent, swSaveOpSilent, errs, warns

    On Error Resume Next
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfVersion, prevVer
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, prevMulti
    On Error GoTo 0

    If errs <> 0 Or Not FileExists(path) Then
        MsgBox "DXF export failed." & vbCrLf & _
               "Errors: " & errs & "  Warnings: " & warns, vbCritical
        SaveDxfWithPrefs = False
    Else
        SaveDxfWithPrefs = True
    End If
End Function

' === Main routine ===
Public Sub ExportDrawingAndReferencedModel()
    Dim swApp   As SldWorks.SldWorks
    Dim swDrv   As SldWorks.ModelDoc2
    Dim swDrw   As SldWorks.DrawingDoc
    Dim swView  As SldWorks.View
    Dim swRef   As SldWorks.ModelDoc2
    Dim mgr     As SldWorks.CustomPropertyManager

    Dim partNo  As String
    Dim desc    As String
    Dim rev     As String
    Dim baseNm  As String
    Dim folder  As String
    Dim dxfP    As String, pdfP As String, stepP As String
    Dim resp    As VbMsgBoxResult

    ' -- init & checks --
    Set swApp = Application.SldWorks
    Set swDrv = swApp.ActiveDoc
    If swDrv Is Nothing Or swDrv.GetType <> swDocDRAWING Then
        MsgBox "Open a drawing before running this macro.", vbCritical
        Exit Sub
    End If
    Set swDrw = swDrv

    If Len(swDrv.GetPathName) = 0 Then
        MsgBox "Please save the drawing before running this macro.", vbCritical
        Exit Sub
    End If

    ' -- find first referenced model --
    Set swView = swDrw.GetFirstView.GetNextView
    Do While Not swView Is Nothing
        If Not swView.ReferencedDocument Is Nothing Then
            Set swRef = swView.ReferencedDocument
            Exit Do
        End If
        Set swView = swView.GetNextView
    Loop
    If swRef Is Nothing Then
        MsgBox "No referenced part/assembly found in the drawing.", vbCritical
        Exit Sub
    End If

    ' -- pull Part Number & Description from the model --
    Set mgr = swRef.Extension.CustomPropertyManager("")
    If Not TryGetProp5(mgr, "Part Number", partNo) _
       Or Not TryGetProp5(mgr, "Description", desc) Then
        MsgBox "Could not read Part Number or Description from the model.", vbCritical
        Exit Sub
    End If

    ' -- pull optional Revision from the drawing --
    Set mgr = swDrv.Extension.CustomPropertyManager("")
    If Not TryGetProp5(mgr, "Revision", rev) Then rev = ""

    ' -- build the base filename --
    baseNm = partNo & ", " & desc
    If Len(rev) > 0 Then
        baseNm = baseNm & ", Rev " & rev
    End If
    baseNm = SanitizeFileName(baseNm)

    ' -- prepare full paths --
    folder = Left$(swDrv.GetPathName, InStrRev(swDrv.GetPathName, "\") - 1)
    dxfP = folder & "\" & baseNm & ".dxf"
    pdfP = folder & "\" & baseNm & ".pdf"
    stepP = folder & "\" & baseNm & ".step"

    ' -- overwrite check (all-or-nothing) --
    If FileExists(dxfP) Or FileExists(pdfP) Or FileExists(stepP) Then
        resp = MsgBox("One or more export files already exist." & vbCrLf & _
                      "Overwrite all?", vbQuestion + vbYesNo, "Overwrite files?")
        If resp <> vbYes Then Exit Sub
    End If

    ' -- exports --
    If Not SaveDxfWithPrefs(swApp, swDrv, dxfP) Then Exit Sub
    If Not SaveAndCheck(swDrv, pdfP, "PDF") Then Exit Sub
    If Not SaveAndCheck(swRef, stepP, "STEP") Then Exit Sub

    MsgBox "All exports succeeded:" & vbCrLf & _
           dxfP & vbCrLf & pdfP & vbCrLf & stepP, vbInformation
End Sub