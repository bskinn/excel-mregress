VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressSetup 
   Caption         =   "Regression Setup"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "RegressSetup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        RegressSetup.frm
' # Purpose:     UserForm enabling definition of the predictor and response variable
' #               source data for a multiple regression analysis.
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

Private reg As ClsRegression
Private isInitialized As Boolean
Private editExistingRegBook As Boolean
Private Const defaultInclConst As Boolean = True
Private closeCancel As Boolean

Public Function getReg() As ClsRegression
    Set getReg = reg
End Function

Public Function closedByCancel() As Boolean
    closedByCancel = closeCancel
End Function

Public Function initNewReg(regName As String, regRespName As String, regDesc As String, regPath As String) As Boolean
    ' Boolean success/failure return
    initNewReg = False
    
    ' Initialize form with blank contents, for definition of new ClsRegression
    Set reg = New ClsRegression
    If Not reg.setName(regName, regPath) Then
        Call MsgBox("Name assignment to new Regression failed!", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    reg.responseName = regRespName

    If Not reg.setDescription(regDesc) Then
        Call MsgBox("Description assignment to new Regression failed!", _
                vbOKOnly + vbCritical, "error")
        Exit Function
    End If
    
    ' Populate the name display box
    TBxName.value = reg.Name
    
    RefEdXData.value = ""
    RefEdYData.value = ""
    RefEdNameData.value = ""
    ChBxInclConst.value = defaultInclConst
    
    isInitialized = True
    editExistingRegBook = False
    
    ' Successful return
    initNewReg = True
    
End Function

Public Function initLoadReg(inReg As ClsRegression) As Boolean
    ' Populate form with the indicated input regression
    ' For now, assume it's properly loaded/configured/etc. and that all range references
    '  are valid, etc.
    initLoadReg = False
    
    Set reg = inReg
    CBxSrcBook.value = inReg.SourceBook.Name
    RefEdXData.value = inReg.SourceAddress(crsXData)
    RefEdYData.value = inReg.SourceAddress(crsYData)
    RefEdNameData.value = inReg.SourceAddress(crsNameData)
    ChBxInclConst.value = inReg.includeConstant
    
    ' Check for temp blip to accommodate editing w/same name and location
    If InStr(inReg.Name, RegressAux.regNameTempBlip) > 0 Then
        TBxName.value = Left(inReg.Name, Len(inReg.Name) - Len(RegressAux.regNameTempBlip))
    Else ' No temp blip found
        TBxName.value = inReg.Name
    End If
    
    isInitialized = True
    editExistingRegBook = True
    
    initLoadReg = True
    
End Function

Private Sub BtnCancel_Click()
    RegressSetup.Hide
End Sub

Private Sub BtnDoRegress_Click()
    Dim wb As Workbook, xRg As Range, yRg As Range, nameRg As Range
    
'    Dim regressSucceeded As Boolean
'    regressSucceeded = False
    
    If Not isInitialized Then
        Call MsgBox("Form not properly initialized", vbOKOnly + vbCritical, "Error")
        Exit Sub
    End If
    
    ' Attach to the selected data source workbook
    Set wb = Workbooks(CBxSrcBook.value)
    
    With wb.Worksheets(1)
        ' Confirm that the RefEd values map to actual ranges within the selected
        '  data source workbook
        ' Don't halt overall execution for these, just drop from Sub and return to form
        If Not TypeOf .Evaluate(RefEdXData.value) Is Range Then
            Call MsgBox("Invalid predictor data range", vbOKOnly + vbCritical, _
                    "Error")
            Exit Sub
        End If
        If Not TypeOf .Evaluate(RefEdYData.value) Is Range Then
            Call MsgBox("Invalid response data range", vbOKOnly + vbCritical, _
                    "Error")
            Exit Sub
        End If
        If Not TypeOf .Evaluate(RefEdNameData.value) Is Range Then
            Call MsgBox("Invalid predictor name range", vbOKOnly + vbCritical, _
                    "Error")
            Exit Sub
        End If
        
        ' Bind ranges for passing into the Regression creation routine
        Set xRg = .Evaluate(RefEdXData.value)
        Set yRg = .Evaluate(RefEdYData.value)
        Set nameRg = .Evaluate(RefEdNameData.value)
        
        ' Complain of any data columns that are constant across the entire dataset
        
        
        
        If editExistingRegBook Then
            ' Confirm to overwrite existing regression
            If vbYes = MsgBox("Overwrite existing regression data?", vbExclamation + vbYesNo, _
                    "Confirm overwrite") Then
'                regressSucceeded = modifyRegression(xRg, yRg, nameRg, ChBxInclConst.Value)
                ' Presume sufficient notification included w/in function
                If Not modifyRegression(xRg, yRg, nameRg, ChBxInclConst.value) Then Exit Sub
            Else
                Exit Sub ' Return control to form
            End If
        Else
            ' Just proceed to create the regression
'            regressSucceeded = createRegression(xRg, yRg, nameRg, ChBxInclConst.Value)
            ' Presume sufficient notification included w/in function
            If Not createRegression(xRg, yRg, nameRg, ChBxInclConst.value) Then Exit Sub
            editExistingRegBook = True
        End If
    End With
    
'    ' Final details if regression successful; assume if not successful, will have exited prior
''    If regressSucceeded Then
'        ' Refresh the open books list, activate the source book, and repopulate
'        '  the source entries
'        popSrcBooks
'        CBxSrcBook.Value = reg.SourceBook.Name
'        RefEdXData.Value = reg.SourceAddress(crsxData)
'        RefEdYData.Value = reg.SourceAddress(crsYData)
'        RefEdNameData.Value = reg.SourceAddress(crsNameData)
'    'End If
    
    ' Set successful regression generation flags and hide the form
    closeCancel = False
    RegressSetup.Hide
    
End Sub

Private Sub BtnOpenWB_Click()
    Dim fd As FileDialog
    Dim fStr As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Open"
        .InitialView = msoFileDialogViewList
        If RegressMain.lastPath = "" Then
            .InitialFileName = "%homepath%\Documents"
        Else
            .InitialFileName = RegressMain.lastPath
        End If
        .Title = "Select workbook to open"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm", 1
        .Filters.Add "All files", "*.*", 2
        If .Show = -1 Then
            ' FileDialog already performs file-exist checking
            Call Workbooks.Open(.SelectedItems(1))
        Else
            ' Do nothing; user presumably canceled
        End If
    End With
    
    ' Repopulate the workbook selector NEEDS ROBUSTIFYING?
    popSrcBooks
    
End Sub

Private Sub BtnCloseWB_Click()
    ' Probably needs robustifying/error-checking
    ' May not be able to perform close here; might need to live on the main form, since
    '  this form won't naturally know what input/output files are needed for all open
    '  ClsRegressions
    ' ##REVIEW##
    Dim checkStr As String
    
    If vbYes = MsgBox("Really close?", vbYesNo + vbExclamation, "Confirm") Then
        checkStr = RegressMain.bookLinked(CBxSrcBook.value)
        If checkStr <> "" Then
            Call MsgBox("Workbook """ & CBxSrcBook.value & """ is linked to open Regression """ & _
                    checkStr & """ and cannot be closed.", vbOKOnly + vbCritical, "Error")
            Exit Sub
        End If
        
        Workbooks(CBxSrcBook.value).Close
        popSrcBooks
        
    End If
End Sub

Private Sub CBxSrcBook_Change()
    ' Activate the selected workbook
    Workbooks.Item(CBxSrcBook.value).Activate
    
    ' Clear the selected ranges, because they pretty much definitely aren't
    '  valid after the source book has changed.
    ' The exception is if this is an opened book for editing, in which case restore the
    '  range values and the constant yes/no value when the source book is selected
    RefEdXData.value = ""
    RefEdYData.value = ""
    RefEdNameData.value = ""
    ChBxInclConst.value = defaultInclConst
    
    If editExistingRegBook Then
        If reg.SourceBookPath = Application.Workbooks(CBxSrcBook.value).fullName Then
            RefEdXData.value = reg.SourceAddress(crsXData)
            RefEdYData.value = reg.SourceAddress(crsYData)
            RefEdNameData.value = reg.SourceAddress(crsNameData)
            ChBxInclConst.value = reg.includeConstant
        End If
    End If
    
    ' Soft way of making sure DoRegress isn't processed if no workbooks are actually open
    BtnDoRegress.Enabled = True
    
End Sub

Private Sub UserForm_Activate()
    If Not isInitialized Then
        Call MsgBox("Regression Setup form not properly initialized!" & vbLf & _
                vbLf & "Exiting...", vbOKOnly + vbCritical, "Critical Error")
        Unload RegressMain
        Unload RegressSetup
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()
    isInitialized = False
    closeCancel = True
    If Not popSrcBooks Then
        Call MsgBox("List of workbooks failed to initialize.", vbOKOnly + vbCritical, _
                "Error")
        RegressSetup.Hide
    End If
End Sub

Private Function popSrcBooks() As Boolean
    Dim bk As Workbook
    popSrcBooks = False
    
    With CBxSrcBook
        ' Clear the list of workbooks, if any
        Do Until .ListCount = 0
            Call .RemoveItem(.ListCount - 1)
        Loop
        
        ' Check for whether any books open
        If Application.Workbooks.Count > 0 Then
            ' Repopulate with the open workbooks
            For Each bk In Application.Workbooks
                Call .AddItem(bk.Name)
            Next bk
            
            ' Select the currently active workbook
            .value = ActiveWorkbook.Name
            
            ' Enable 'Close' and 'Generate Reg' buttons
            BtnCloseWB.Enabled = True
            BtnDoRegress.Enabled = True
        Else
            ' No books open - disable 'Close' and 'Generate Reg' buttons
            BtnCloseWB.Enabled = False
            BtnDoRegress.Enabled = False
        End If
    End With
    
    popSrcBooks = True
    
End Function

Private Function createRegression(srcX_Vals As Range, srcY_Vals As Range, _
        srcName_Vals As Range, Optional includeConstant As Boolean = True) As Boolean
    
    createRegression = False
    
    ' ID ranges; create new book; prep with input data?
    
    'Set createRegression = Nothing
    
    ' Proof the intended data; any problem with the data already noted within the proofing function,
    '  so just quietly return control to form
    If Not proofRegSrcData(srcX_Vals, srcY_Vals, srcName_Vals) Then Exit Function
    
    ' Create the actual thing, attaching source and generating the regression book
    If Not reg.attachSource(srcX_Vals.Worksheet.Parent, srcX_Vals, srcY_Vals, srcName_Vals) Then
        Call MsgBox("Error binding source data.", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    If Not reg.createNewRegression(includeConstant) Then
        Call MsgBox("Error generating regression workbook.", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    createRegression = True
    
End Function

Private Function modifyRegression(srcX_Vals As Range, srcY_Vals As Range, _
        srcName_Vals As Range, Optional includeConstant As Boolean = True) As Boolean
    
    ' This routine is for when the source of a regression is being (potentially) redefined
    modifyRegression = False
    
    ' Proof the indicated source ranges; if invalid, doesn't matter whether they match the
    '  currently linked sources.  Really, if they *are* identical, then the regression was
    '  improperly constructed to begin with!
    If Not proofRegSrcData(srcX_Vals, srcY_Vals, srcName_Vals) Then Exit Function
    
    ' Create the regression, attaching new source but re-generating the regression book in-place
    If Not reg.attachSource(srcX_Vals.Worksheet.Parent, srcX_Vals, srcY_Vals, srcName_Vals) Then
        Call MsgBox("Error binding source data.", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    If Not reg.modifyRegression(False, includeConstant) Then
        Call MsgBox("Error generating regression workbook.", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    modifyRegression = True
    
End Function

Private Function proofRegSrcData(xD As Range, yD As Range, nameD As Range) As Boolean
    
    ' Initialize fail return
    proofRegSrcData = False
    
    ' Kick out if containing workbook is not identical for all three sources (should never fail!)
    If Not (xD.Worksheet.Parent Is yD.Worksheet.Parent And _
            xD.Worksheet.Parent Is nameD.Worksheet.Parent) Then
        Call MsgBox("X/Y/Name data sourcing from separate workbooks is not currently supported.", _
                vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    ' Kick out if X_ or Y_ or Name_Vals has more than one area
    If xD.Areas.Count > 1 Or yD.Areas.Count > 1 Or nameD.Areas.Count > 1 Then
        Call MsgBox("Predictor and response data ranges must be a single, rectangular selection.", _
                vbOKOnly + vbCritical, "Improper data selection(s)")
        Exit Function
    End If
    
    ' Kick out if row dims mismatch
    If xD.Rows.Count <> yD.Rows.Count Then
        Call MsgBox("Row counts of predictor and response data do not match.", _
                    vbOKOnly + vbCritical, "Row count mismatch")
        Exit Function
    End If
    
    ' Kick out if regression is overdetermined
    If reg.includeConstant Then
        If xD.Rows.Count <= xD.Columns.Count Then
            Call MsgBox("Regression is overdetermined." & vbLf & vbLf & _
                        "Increase number of data points or decrease number of predictors.", _
                        vbOKOnly + vbCritical, "Overdetermined regression")
            Exit Function
        End If
    Else
        If xD.Rows.Count < xD.Columns.Count Then
            Call MsgBox("Regression is overdetermined." & vbLf & vbLf & _
                        "Increase number of data points or decrease number of predictors.", _
                        vbOKOnly + vbCritical, "Overdetermined regression")
            Exit Function
        End If
    End If
    
    ' Kick out if Y_Vals has more than one column
    If yD.Columns.Count > 1 Then
        Call MsgBox("Response data must contain only a single column.", _
                    vbOKOnly + vbCritical, "Improper response data")
        Exit Function
    End If
    
    ' Kick out if Name_Vals is not a vector, size-matched to columns of X
    If Not ((nameD.Rows.Count = 1 And nameD.Columns.Count = xD.Columns.Count) Or _
            (nameD.Columns.Count = 1 And nameD.Rows.Count = xD.Columns.Count)) Then
        Call MsgBox("Size of predictor name vector does not match number of predictor columns.", _
                vbOKOnly + vbCritical, "Predictor name mismatch")
        Exit Function
    End If
    
    ' Kick out if there is only one data point; obsoleted, from above check for
    '  overspecified regressions
    If xD.Rows.Count < 2 Then
        Call MsgBox("Creation of Regressions with a single data point is not allowed.", _
                vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    proofRegSrcData = True

End Function

Private Sub UserForm_Terminate()
    ' Unlink any attached regression and de-initialize
    Set reg = Nothing
    isInitialized = False
End Sub
