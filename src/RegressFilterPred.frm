VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressFilterPred 
   Caption         =   "Select Predictors"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   OleObjectBlob   =   "RegressFilterPred.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressFilterPred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        RegressFilterPred.frm
' # Purpose:     UserForm for selecting predictors to include/exclude from
' #               the active multiple regression analysis.
' #               Part of the "Multiple Regression Explorer" Excel VBA Add-In
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     24 Feb 2014
' # Copyright:   (c) Brian Skinn 2017
' # License:     The MIT License; see "LICENSE.txt" for full license terms.
' #
' #       http://www.github.com/bskinn/excel-mregress
' #
' # ------------------------------------------------------------------------------

Option Compare Text
Option Explicit

Private reg As ClsRegression, includeRg As Range, excludeRg As Range
Private initOK As Boolean, exitCancel As Boolean
Const emptyExcludeVal As String = "0"

Public Function isInitialized() As Boolean
    isInitialized = initOK
End Function

Public Function exitByCancel() As Boolean
    exitByCancel = exitCancel
End Function

Public Function getNewFilteredPreds() As String
    
    Dim workStr As String, cel As Range
    
    ' Should just be an iteration over excludeRg, kicking the numbers into a string?
    If excludeRg.Rows.Count = 1 And excludeRg.Cells(1, 1) = "0" Then
        ' No filters
        getNewFilteredPreds = "()"
    Else
        ' Must build the string
        workStr = "("
        For Each cel In excludeRg.Columns(1).Cells
            workStr = workStr & CStr(cel.value) & ","
        Next cel
        workStr = Left(workStr, Len(workStr) - 1) & ")"
        getNewFilteredPreds = workStr
    End If
    
End Function

Public Function populateForm(inReg As ClsRegression) As Boolean
    populateForm = False
    
    Set reg = inReg
    
    ' Populate the name field
    TBxName.value = reg.Name
    
    ' Set up the listboxes
    popLBoxes
    
    ' Attach the ranges to the controls
    attachLBoxRanges
    
    ' Should be it?
    initOK = True
    
End Function

Private Sub popLBoxes()
    ' Iterators, dim and init
    Dim iter As Long, incIter As Long, excIter As Long
    incIter = 1
    excIter = 1
    
    ' Safety check
    If includeRg Is Nothing Or excludeRg Is Nothing Then initLBoxes
    
    ' Resize ranges to right # cols & rows and set number format to string
    '  Must trap the exclude range for the possiblity of zero excluded predictors
    Set includeRg = includeRg.Resize(reg.numPredictors(True), 2)
            'If reg.numPredictors(False) - reg.numPredictors(True) = 0 Then
    If reg.countFilters(crfPredictor) = 0 Then
        ' No filtered predictors
        Set excludeRg = excludeRg.Resize(1, 2)
        excludeRg.Cells(1, 1) = emptyExcludeVal
        excludeRg.Cells(1, 2) = ""
    Else
        ' Some filtered predictors
        Set excludeRg = excludeRg.Resize(reg.numPredictors(False) - reg.numPredictors(True), 2)
    End If
    includeRg.NumberFormat = "@"
    excludeRg.NumberFormat = "@"
    
    ' Populate Ranges from the regression
    For iter = 1 To reg.numPredictors(False)
        Select Case reg.isFiltered(iter, crfPredictor)
        Case vbUseDefault
            ' Should never hit; hard kill if it does
            Call MsgBox("Invalid predictor index! Exiting...", vbOKOnly + vbCritical, _
                        "Critical Error!!")
            Unload RegressMain
            Unload Me
            Exit Sub
        Case vbFalse
            ' Add to 'include' and increment counter
            includeRg.Cells(incIter, 1) = iter
            includeRg.Cells(incIter, 2) = reg.predictorName(iter, False)
            incIter = incIter + 1
        Case vbTrue
            ' Add to 'exclude'
            excludeRg.Cells(excIter, 1) = iter
            excludeRg.Cells(excIter, 2) = reg.predictorName(iter, False)
            excIter = excIter + 1
        End Select
    Next iter
    
    ' Should be good to go?
    
End Sub

Private Sub attachLBoxRanges()
    ' Attach to controls
    LBxInclude.RowSource = includeRg.Address(External:=True)
    LBxExclude.RowSource = excludeRg.Address(External:=True)
End Sub

Private Sub initLBoxes()
    ' Anchor the binding ranges
    Set includeRg = reg.scratchCell
    Set excludeRg = includeRg.Offset(reg.numPredictors(False) + 1, 0)
End Sub

Private Sub BtnCancel_Click()
    ' Should just be able to hide the form?  _Terminate() will clean up the Ranges;
    '  nothing else should have changed.
    Me.Hide
End Sub

Private Sub BtnDone_Click()
    ' Set a non-cancel return and hide the form.
    ' Responsibility of the calling form to call getNewFilteredPreds() to retrieve the new
    '  set of predictor filters
    exitCancel = False
    Me.Hide
End Sub

Private Sub BtnExcludeOne_Click()
    ' This doesn't worry about retaining sorting while making these selections
    '  reg.filterString shouldn't care...
    
    Dim iter As Long
    
    ' Do nothing if nothing selected on Include side
    If LBxInclude.ListIndex = -1 Then Exit Sub
    
    ' Pitch error and do nothing if this would empty the include range
    If includeRg.Rows.Count < 2 Then
        Call MsgBox("Regression must include at least one predictor!", vbOKOnly + vbCritical, _
                "Error")
        Exit Sub
    End If

    ' Else, transfer the excluded predictor to the other side
    ' Check for empty exclude range
    
    If excludeRg.Cells(1, 1) = emptyExcludeVal Then
        ' Just transfer values
        excludeRg.Cells(1, 1) = includeRg.Cells(LBxInclude.ListIndex + 1, 1)
        excludeRg.Cells(1, 2) = includeRg.Cells(LBxInclude.ListIndex + 1, 2)
    Else
        ' Expand excludeRg and transfer values
        Set excludeRg = excludeRg.Resize(excludeRg.Rows.Count + 1, 2)
        excludeRg.Cells(excludeRg.Rows.Count, 1) = includeRg.Cells(LBxInclude.ListIndex + 1, 1)
        excludeRg.Cells(excludeRg.Rows.Count, 2) = includeRg.Cells(LBxInclude.ListIndex + 1, 2)
    End If
    
    ' Either way, contract includeRg safely
    For iter = LBxInclude.ListIndex + 1 To includeRg.Rows.Count - 1
        includeRg.Cells(iter, 1) = includeRg.Cells(iter + 1, 1)
        includeRg.Cells(iter, 2) = includeRg.Cells(iter + 1, 2)
    Next iter
    'includeRg.Rows(includeRg.Rows.Count).Formula = ""
    includeRg.Rows(includeRg.Rows.Count).ClearContents
    Set includeRg = includeRg.Resize(includeRg.Rows.Count - 1, 2)
    
    ' Do need to reassign ranges
    attachLBoxRanges
    
End Sub

Private Sub BtnIncludeAll_Click()
    ' Can I just call the event function for BtnIncludeOne repeatedly?...?
    LBxExclude.ListIndex = 0
    Do Until LBxExclude.value = emptyExcludeVal Or excludeRg.Cells(1, 1).value = emptyExcludeVal
        BtnIncludeOne_Click
        LBxExclude.ListIndex = 0
    Loop
    
    ' Don't need to reassign ranges?
    'attachLBoxRanges
End Sub

Private Sub BtnIncludeOne_Click()
    ' This doesn't worry about retaining sorting while making these selections
    '  reg.filterString shouldn't care...
    
    Dim iter As Long
    
    ' Do nothing if nothing selected on Exclude side, or if empty entry selected
    If LBxExclude.ListIndex = -1 Or _
            (LBxExclude.ListIndex = 0 And LBxExclude.value = emptyExcludeVal) Then Exit Sub
    
    ' Transfer the included predictor to the other side. Empty exclude range is fine.
    ' Should not need to check for empty include range!
    ' Expand includeRg and transfer values
    Set includeRg = includeRg.Resize(includeRg.Rows.Count + 1, 2)
    includeRg.Cells(includeRg.Rows.Count, 1) = excludeRg.Cells(LBxExclude.ListIndex + 1, 1)
    includeRg.Cells(includeRg.Rows.Count, 2) = excludeRg.Cells(LBxExclude.ListIndex + 1, 2)
    
    ' Contract excludeRg safely, checking for removal of last entry
    If excludeRg.Rows.Count = 1 Then
        ' Just clear the values, don't resize
        excludeRg.Cells(1, 1) = emptyExcludeVal
        excludeRg.Cells(1, 2) = ""
    Else
        ' Do contract the exclude Range
        For iter = LBxExclude.ListIndex + 1 To excludeRg.Rows.Count - 1
            excludeRg.Cells(iter, 1) = excludeRg.Cells(iter + 1, 1)
            excludeRg.Cells(iter, 2) = excludeRg.Cells(iter + 1, 2)
        Next iter
        excludeRg.Rows(excludeRg.Rows.Count).ClearContents
        'excludeRg.Cells(excludeRg.Rows.Count, 1).Formula = emptyExcludeVal
        'excludeRg.Cells(excludeRg.Rows.Count, 2).Formula = ""
        Set excludeRg = excludeRg.Resize(excludeRg.Rows.Count - 1, 2)
    End If
    
    ' Do need to reassign ranges
    attachLBoxRanges
    
End Sub

Private Sub UserForm_Activate()
    If Not isInitialized Then
        Call MsgBox("Predictor filter form activated without proper initalization!", _
                vbOKOnly + vbCritical, "Error")
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    initOK = False
    exitCancel = True
End Sub

Private Sub UserForm_Terminate()
    ' Clear the scratch ranges regardless of what happened, if bound
    If Not includeRg Is Nothing Then
        includeRg.Clear
        includeRg.NumberFormat = "General"
    End If
    If Not excludeRg Is Nothing Then
        excludeRg.Clear
        excludeRg.NumberFormat = "General"
    End If
    ' if bound, write changes to the Regression so that it's not left in a funny state
    reg.writeChanges
End Sub
