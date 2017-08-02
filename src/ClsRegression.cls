VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsRegression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        ClsRegression.cls
' # Purpose:     Custom class defining a multiple regression analysis. Part of
' #               the "Multiple Regression Explorer" Excel VBA Add-In
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

' Enum for range type
Public Enum CRSrcType
    crsXData
    crsYData
    crsNameData
End Enum

' Enum for config setting -- value doubles as the row number
Public Enum CRConfig
    crcSourceBook = 1
    crcSourceXRange
    crcSourceYRange
    crcSourceNameRange
    crcIncludeConstant
    crcDescription
    crcFiltPts
    crcFiltPreds
    crcRespName
    crcLastChartX
    crcLastChartY
    crcLastChartXPred
    crcLastChartYPred
    crcLastChartXNorm
    crcLastChartYNorm
    crcLastChartDoOutliers
    crcLastChartOutlierAlpha
    crcLastChartSize
End Enum

' Enum for filter type
Public Enum CRFilterType
    crfDataPoint
    crfPredictor
End Enum

Public Enum degsOfFreedom
    DOFTotal
    DOFRegression
    DOFResidual
End Enum

' Enum for regression statistic type
Public Enum StatType
    ST_SSTotal = 1
    ST_SSRegression
    ST_SSResidual
    ST_RSq
    ST_FStat
    ST_FStatPVal
    ST_AIC
    ST_CorrAIC
End Enum

' Enum for charting variables
Public Enum CRChartVar
    crcvResponse
    crcvPredictor
    crcvFittedResponse
    crcvDirectResidual
    crcvStudentizedResidual
    crcvTStatResidual
    crcvCookDistance
    crcvPointSequence
End Enum

' Enum for charting size
Public Enum CRChartSize
    crcsFullPage = 1
    crcsFigure
    crcsThumbnail
End Enum

' Enum for copying internal ranges
Public Enum CRCopyRange
    crcrStats = 1
End Enum

Const firstStatTypeIndex As Long = 1
Const lastStatTypeIndex As Long = 7

Const firstChartSizeIndex As Long = 1
Const lastChartSizeIndex As Long = 3

' References to input data for the regression
Private srcBk As Workbook
'Private SrcBkPath As String

Private SrcXRng As Range
'Private SrcXRef As String

Private SrcYRng As Range
'Private SrcYRef As String

Private SrcNameRng As Range
'Private SrcNameRef As String

Private SrcInclConst As Boolean

Private regName As String
Private regDesc As String
Private regRespName As String
Private lastChartX As CRChartVar
Private lastChartY As CRChartVar
Private lastChartXPred As Long
Private lastChartYPred As Long
Private lastChartXNorm As Boolean
Private lastChartYNorm As Boolean
Private lastChartDoOutliers As Boolean
Private lastChartOutlierAlpha As Double
Private lastChartSize As CRChartSize

' References to in-use regression info will be detected/generated on the fly,
'  not ever loaded directly from disk. So, just Ranges are needed(?)
' Workbook and Worksheets
' !!!HERE!!! if new worksheet/chartsheet added
Private rBook As Workbook
Private rConfigSheet As Worksheet
Private rDataSheet As Worksheet
Private rFiltDataSheet As Worksheet
Private rVSheet As Worksheet
Private rResSheet As Worksheet
Private rStatSheet As Worksheet
Private rDynChart As Chart

' !!!HERE!!! when new result or other fields added
' Ranges
Private rXData As Range, rYData As Range, rNameData As Range
Private rXFiltData As Range, rYFiltData As Range, rNameFiltData As Range
Private rVMat As Range
Private rResYRg As Range, rResBetaRg As Range
Private rResBetaSERg As Range, rResTBetaRg As Range
Private rResYHatRg As Range, rResEHatRg As Range
Private rResVIIRg As Range, rResRRg As Range
Private rResTRg As Range, rResDRg As Range
Private rResPredNameRg As Range
Private rResSYYRg As Range, rResSYYNameRg As Range
Private rResRSSRg As Range, rResRSSNameRg As Range
Private rResRegSSRg As Range, rResRegSSNameRg As Range
Private rResR2Rg As Range, rResR2NameRg As Range
Private rResFStatRg As Range, rResFStatNameRg As Range
Private rResDegTotRg As Range, rResDegTotNameRg As Range
Private rResDegRegRg As Range, rResDegRegNameRg As Range
Private rResDegResRg As Range, rResDegResNameRg As Range
Private rResAICRg As Range, rResAICNameRg As Range
Private rResCorrAICRg As Range, rResCorrAICNameRg As Range
Private rStatRg As Range

' Container for *all* objects, for simple unsetting, etc. May not be useful; ignoring for now
'Private ObjColl As New Collection

' Filter arrays
Private filtPts As Variant
Private filtPreds As Variant

' Flag
Private sourceAttached As Boolean
Private regCreated As Boolean

' Internal strings
Const invalidIndexText As String = "#INVALID INDEX#"



Public Function loadRegression(rBookToLoad As Workbook, _
        Optional overwriteOldData As Boolean = False) As Boolean
    ' Populates the entirety of the Reg, to include opening the source book if requred
    ' Check to ensure the last modification timestamp is within +/- 5 sec; if not, notify(?)
    '  and perform additional validation checks.  Also do additional validation checks if
    '  the filename doesn't contain the special extension flag.
    ' ##NEED TO DEAL WITH if this function is called when the Reg has already been initialized
    ' Currently just **always overwrites**
    
    Dim workStr As String, workWb As Workbook, bookOpen As Boolean
    Dim iter As Long
    
    Dim doRigorousChecks As Boolean
    doRigorousChecks = False    ' Presume ok to start
    
    ' Pre-validation checking
    ' Timestamps
    ' Check for custom save time stamp
    If Not CustDocPropExists(RegressAux.timeStampDocPropName, rBookToLoad) Then
        doRigorousChecks = True
    Else
        ' Check +/- of timestamp
        If 86400 * Abs(CDate(rBookToLoad.BuiltinDocumentProperties(RegressAux.lastSaveDocPropName)) - _
                CDate(rBookToLoad.CustomDocumentProperties(RegressAux.timeStampDocPropName))) _
                > RegressAux.stampTimeWindowSecs Then
            doRigorousChecks = True
        End If
    End If
    
    ' If extension is not the special one, do rigorous checks
    If InStr(rBookToLoad.Name, RegressAux.rBookExtension) < 1 Then doRigorousChecks = True
    
    ' Do the rigorous checking
    If doRigorousChecks Then
        ' ##NEED TO ADD some rigorous checking, or if it proves too dangerous, halt the load ##
        ' Incl ensuring relevant sheets exist and content is right
        '  Config sheet should only have cells in the one column, with count matching number of
        '   known config settings
        '  Data sheets (main and filtered) should have one column, and then one block of data;
        '   non-overlapping; row counts should match; name range should have same #cols as X data
        '  V-Matrix sheet should just be one big block of data, with the label at top.  Should be
        '   SQUARE, though
        '  Results sheet should be a series of columns of data, save for the beta values & S.E.'s, and
        '   the single results like R^2 &c.  .CurrentRegion from the expected locations should all just
        '   have ONE column; row counts on all of them (but the betas, etc.) should be EQUAL
        '
        '  At least alert that rigorous structure checking was flagged
        Call MsgBox("Rigorous checking of Regression """ & rBookToLoad.Name & """ was triggered," & _
                "but has not yet been implemented.", vbOKOnly + vbInformation, "Alert")
    End If
    
    ' Proceed assuming okay, both content and structure
    ' Set the rBook
    Set rBook = rBookToLoad
    
    ' Set the Reg name
    ' ## ASSUMES the special extension -- will hack off part of the name of a regular
    '  *.xlsx book.
    regName = Left(rBookToLoad.Name, Len(rBookToLoad.Name) - Len(RegressAux.rBookExtension))
    
    ' Attach the rBook sheets; assumes exist..!
    Set rConfigSheet = rBook.Sheets(RegressAux.configShtName)
    Set rDataSheet = rBook.Sheets(RegressAux.dataShtName)
    Set rFiltDataSheet = rBook.Sheets(RegressAux.dataFiltShtName)
    Set rVSheet = rBook.Sheets(RegressAux.vMatShtName)
    Set rResSheet = rBook.Sheets(RegressAux.resShtName)
    Set rStatSheet = rBook.Sheets(RegressAux.statShtName)
    
    '' Retrieve configuration stuff from the config sheet
    ' Set the Reg description and 'include constant' Boolean
    regDesc = rConfigSheet.Cells(crcDescription, RegressAux.configValCol).Formula
    SrcInclConst = CBool(rConfigSheet.Cells(crcIncludeConstant, RegressAux.configValCol).Formula)
    
    ' Bind the source ranges from config
    Me.bindSourceFromConfig
    
    ' And, set to bind the results ranges too.
    ' Data sheet
    Set rXData = rDataSheet.Cells(RegressAux.firstDataRow, RegressAux.dataXCol).CurrentRegion
    Set rYData = rDataSheet.Cells(RegressAux.firstDataRow, RegressAux.dataYCol).CurrentRegion
    Set rNameData = rDataSheet.Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).CurrentRegion
    ' Filtered data sheet
    Set rXFiltData = rFiltDataSheet.Cells(RegressAux.firstDataRow, RegressAux.dataXCol).CurrentRegion
    Set rYFiltData = rFiltDataSheet.Cells(RegressAux.firstDataRow, RegressAux.dataYCol).CurrentRegion
    Set rNameFiltData = rFiltDataSheet.Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).CurrentRegion
    ' V-Matrix sheet
    Set rVMat = rVSheet.Cells(RegressAux.firstDataRow, RegressAux.vMatCol).CurrentRegion
    ' !!!HERE!!! when new result fields added
    ' Results sheet
    Set rResYRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultYCol).CurrentRegion
    Set rResBetaRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultBetaCol).CurrentRegion
    Set rResBetaSERg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultBetaSECol).CurrentRegion
    Set rResTBetaRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultTBetaCol).CurrentRegion
    Set rResYHatRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultYHCol).CurrentRegion
    Set rResEHatRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultEHCol).CurrentRegion
    Set rResVIIRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultVIICol).CurrentRegion
    Set rResRRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultRCol).CurrentRegion
    Set rResTRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultTCol).CurrentRegion
    Set rResDRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultDCol).CurrentRegion
    Set rResPredNameRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultPredNameCol).CurrentRegion
    
    With rResBetaRg
        Set rResSYYRg = .Offset(.Rows.Count + RegressAux.resultSYYRowOffset - 1, 0).Resize(1, 1)
        Set rResRSSRg = .Offset(.Rows.Count + RegressAux.resultRSSRowOffset - 1, 0).Resize(1, 1)
        Set rResRegSSRg = .Offset(.Rows.Count + RegressAux.resultRegSSRowOffset - 1, 0).Resize(1, 1)
        Set rResR2Rg = .Offset(.Rows.Count + RegressAux.resultR2RowOffset - 1, 0).Resize(1, 1)
        Set rResFStatRg = .Offset(.Rows.Count + RegressAux.resultFStatRowOffset - 1, 0).Resize(1, 1)
        Set rResDegTotRg = .Offset(.Rows.Count + RegressAux.resultDegTotRowOffset - 1, 0).Resize(1, 1)
        Set rResDegRegRg = .Offset(.Rows.Count + RegressAux.resultDegRegRowOffset - 1, 0).Resize(1, 1)
        Set rResDegResRg = .Offset(.Rows.Count + RegressAux.resultDegResRowOffset - 1, 0).Resize(1, 1)
        Set rResAICRg = .Offset(.Rows.Count + RegressAux.resultAICRowOffset - 1, 0).Resize(1, 1)
        Set rResCorrAICRg = .Offset(.Rows.Count + RegressAux.resultCorrAICRowOffset - 1, 0).Resize(1, 1)
    End With
    With rResPredNameRg
        Set rResSYYNameRg = .Offset(.Rows.Count + RegressAux.resultSYYRowOffset - 1, 0).Resize(1, 1)
        Set rResRSSNameRg = .Offset(.Rows.Count + RegressAux.resultRSSRowOffset - 1, 0).Resize(1, 1)
        Set rResRegSSNameRg = .Offset(.Rows.Count + RegressAux.resultRegSSRowOffset - 1, 0).Resize(1, 1)
        Set rResR2NameRg = .Offset(.Rows.Count + RegressAux.resultR2RowOffset - 1, 0).Resize(1, 1)
        Set rResFStatNameRg = .Offset(.Rows.Count + RegressAux.resultFStatRowOffset - 1, 0).Resize(1, 1)
        Set rResDegTotNameRg = .Offset(.Rows.Count + RegressAux.resultDegTotRowOffset - 1, 0).Resize(1, 1)
        Set rResDegRegNameRg = .Offset(.Rows.Count + RegressAux.resultDegRegRowOffset - 1, 0).Resize(1, 1)
        Set rResDegResNameRg = .Offset(.Rows.Count + RegressAux.resultDegResRowOffset - 1, 0).Resize(1, 1)
        Set rResAICNameRg = .Offset(.Rows.Count + RegressAux.resultAICRowOffset - 1, 0).Resize(1, 1)
        Set rResCorrAICNameRg = .Offset(.Rows.Count + RegressAux.resultCorrAICRowOffset - 1, 0).Resize(1, 1)
    End With
    
    ' Retrieve the data point and predictor filters
    '  Can only be done after certain of the Ranges are bound
    Me.filterString(crfDataPoint) = rConfigSheet.Cells(crcFiltPts, RegressAux.configValCol).Formula
    Me.filterString(crfPredictor) = rConfigSheet.Cells(crcFiltPreds, RegressAux.configValCol).Formula
    
    ' Store the response variable name
    Me.responseName = rConfigSheet.Cells(crcRespName, RegressAux.configValCol).Formula
    
    ' Bind the stats range on the stats sheet and generate the stats
    Set rStatRg = rStatSheet.Cells(RegressAux.statStartRow, RegressAux.statStartCol)
    Me.writeStats
    
    regCreated = True  ' Flag that the results stuff is bound
    
    ' Look for the DynChart and attach if present
    For iter = 1 To rBook.Charts.Count
        If rBook.Charts(iter).Name = RegressAux.chartShtName Then
            Set rDynChart = rBook.Charts(iter)
        End If
    Next iter
    
    ' Retrieve the last-charted settings
    lastChartX = CLng(rConfigSheet.Cells(crcLastChartX, RegressAux.configValCol).Formula)
    lastChartY = CLng(rConfigSheet.Cells(crcLastChartY, RegressAux.configValCol).Formula)
    lastChartXPred = CLng(rConfigSheet.Cells(crcLastChartXPred, RegressAux.configValCol).Formula)
    lastChartYPred = CLng(rConfigSheet.Cells(crcLastChartYPred, RegressAux.configValCol).Formula)
    lastChartXNorm = CBool(rConfigSheet.Cells(crcLastChartXNorm, RegressAux.configValCol).Formula)
    lastChartYNorm = CBool(rConfigSheet.Cells(crcLastChartYNorm, RegressAux.configValCol).Formula)
    lastChartDoOutliers = CBool(rConfigSheet.Cells(crcLastChartDoOutliers, RegressAux.configValCol).Formula)
    lastChartOutlierAlpha = CDbl(rConfigSheet.Cells(crcLastChartOutlierAlpha, RegressAux.configValCol).Formula)
    lastChartSize = CLng(rConfigSheet.Cells(crcLastChartSize, RegressAux.configValCol).Formula)
    
    ' DO NOT write config since regression was not regenerated -- all data should be unchanged
    '  from prior write
    
    ' Successful load
    loadRegression = True
    
End Function

Public Sub setDefaultLastChartedVars()
    lastChartX = crcvFittedResponse
    lastChartY = crcvStudentizedResidual
    lastChartXPred = 0
    lastChartYPred = 0
    lastChartXNorm = False
    lastChartYNorm = False
    lastChartDoOutliers = True
    lastChartOutlierAlpha = 0.05
    lastChartSize = crcsFigure
End Sub

Public Sub bindSourceFromConfig()
    ' Bind/rebind source ranges from config. Copes with object referencing failures when
    '  sheets are modified underneath previously-consistent Range bindings
    
    Dim workStr As String, bookOpen As Boolean, workWb As Workbook
    
    ' Extract source book name
    workStr = rConfigSheet.Cells(crcSourceBook, RegressAux.configValCol).Formula
    Do While InStr(workStr, "\") ' Could replace this with a FileSystemObject call...
        workStr = Right(workStr, Len(workStr) - InStr(workStr, "\"))
    Loop
    
    ' Check if source book already open
    bookOpen = False
    For Each workWb In Workbooks
        If workWb.Name = workStr Then bookOpen = True
    Next workWb
    
    ' If not open, open it and bind; else, just bind
    If Not bookOpen Then
        Set srcBk = Workbooks.Open(rConfigSheet.Cells(crcSourceBook, RegressAux.configValCol).Formula)
    Else
        Set srcBk = Workbooks(workStr)
    End If
    workStr = ""
    Set workWb = Nothing
    
    ' Should be all set to bind the source ranges
    Set SrcXRng = srcBk.Worksheets(1).Evaluate( _
            rConfigSheet.Cells(crcSourceXRange, RegressAux.configValCol).Formula)
    Set SrcYRng = srcBk.Worksheets(1).Evaluate( _
            rConfigSheet.Cells(crcSourceYRange, RegressAux.configValCol).Formula)
    Set SrcNameRng = srcBk.Worksheets(1).Evaluate( _
            rConfigSheet.Cells(crcSourceNameRange, RegressAux.configValCol).Formula)
    sourceAttached = True  ' Source is now attached
    
End Sub

Public Function closeRegression(Optional saveBeforeClose As Boolean = False) As Boolean
    ' No particular information in the Boolean return, just whether function made it to the end
    closeRegression = False
    
    ' Save changes if indicated
    If saveBeforeClose Then Me.writeChanges
    
    ' Unlink source objects
    UnlinkSource
    sourceAttached = False
    
    ' Unbind all of the results sheet objects
    UnloadResObjects
    regCreated = False
    
    ' Close and unlink the rBook
    Call rBook.Close(SaveChanges:=False)
    Set rBook = Nothing
    
    ' Return success
    closeRegression = True
    
End Function

Public Function deleteRegressionFile(Optional userConfirm As Boolean = True) As Boolean
    ' Closes and deletes the associated regression file, and resets the Regression name to NULL
    '  MUST either Delete or Re-Define the Reg before returning control to user!!
    ' Anticipated to be used mainly as a pre-step before deletion of the Regression
    '  itself, but possibly usable in clearing a Regression object for re-creation, perhaps
    '  in situations where multiple temporary regression objects/files are being created?
    ' Specifically of note, the source links are NOT obliterated
    '
    ' How to leave the Regression in a consistent state, though...? Just set the name to null string?
    
    ' Return: False = File deletion was unsuccessful, either due to error(?) or user cancel
    '         True  = File deletion successful.
    
    Dim fullName As String, fs As FileSystemObject
    
    deleteRegressionFile = False
    
    ' Query to confirm delete, if indicated
    If userConfirm Then
        If Not vbYes = MsgBox("Really delete data file for Regression """ & regName & """?", _
                vbYesNo + vbQuestion, "Confirm deletion") Then Exit Function
    End If
    
    ' Store name for use after deleting the book
    fullName = rBook.fullName
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    UnloadResObjects
    
    Application.DisplayAlerts = False
    Call rBook.Close(SaveChanges:=False)
    Set rBook = Nothing
    Call fs.DeleteFile(fullName, True)
    Application.DisplayAlerts = True
    
    regName = ""
    regCreated = False
    
    deleteRegressionFile = True
    
End Function

Public Function setName(newName As String, Optional savePath As String = "", _
            Optional deleteOldFile As VbTriState = vbUseDefault) As Boolean
    ' newName (String) is the new name to be applied, which will be the index key of the Regression
    '    and the name of the containing xlsx file
    ' savePath (String) is the path for saving; it can be left unspecified for an already-created
    '    Regression to use the same path as the currently-saved rBook
    ' deleteOldFile (VbTriState) indicates whether to delete the old rBook after renaming.  If the book
    '    is not deleted, it can (eventually) be re-opened under the old name as an unmodified copy/version
    '    vbTrue       = Do delete the old book
    '    vbFalse      = Do not delete the old book
    '    vbUseDefault = Query the user whether to delete
    
    Dim fs As FileSystemObject, fd As FileDialog
    Dim oldBookFullName As String, oldName As String
    Dim priorActiveBook As Workbook
    
    ' Initialize return value
    '  False = Name and workbook were not changed
    '  True  = Name and workbook were changed successfully
    setName = False
    
    ' Name must always match the filename; file must always be saved-to-something.  So,
    '  integrating the book-save operations here makes sense
    
    ' Reg name construction has presumably already been proofed, but no harm in adding an extra layer
    '  of trapping.  Provides info as to where the badness happened.
    If Not RegressAux.validRegName(newName) Then
        Call MsgBox("Invalid regression name """ & newName & """passed to [ClsRegression].setName()", _
                vbOKOnly + vbCritical, "Error")
        Exit Function
    Else
        ' Name ok; proceed, but don't actually change it until the end in the event that the file processing
        '  should fail
        ' Now check for whether the name is actually different; silently drop if same
        If newName = regName Then Exit Function
    End If
    
    ' Possible states?  Will the below work for OPENING an existing Regression? Probably -- will require
    '  a DIFFERENT METHOD. This only is for creating a new Reg via the name, or for renaming/copying.
    ' rBook:  Nothing or something.  Nothing is taken to be the proxy indicator for a newly-created Regression
    '   If Nothing, new book needs to be created; savePath needs to be specified, else error
    '       value of deleteOldFile is irrelevant, but throw error (warning?) if non-default is specified
    '   If something, then must parse savePath and deleteOldFile
    '     savePath:
    '       If savePath = "", then retrieve path from prior rBook
    '       If savePath non-null but invalid, throw error and cancel out
    '       If savePath non-null and valid, proceed
    '     deleteOldFile:
    '       If vbTrue, close and delete the old file after assigning the new one
    '       If vbFalse, close and DO NOT delete the old file after assigning the new one
    '       If vbUseDefault, query user whether to delete or not, allowing Cancel
    '
    
    ' Bind FileSystem object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Error checking the inputs
    If rBook Is Nothing Then
        ' savePath must be a folder that exists
        If Not fs.FolderExists(savePath) Then
            Call MsgBox("Save location passed to [ClsRegression].setName does not exist:" & _
                    vbLf & vbLf & """" & savePath & """", vbOKOnly + vbCritical, "Error")
            Exit Function
        Else
            ' Check if vbTrue set; warn if so.
            If deleteOldFile = vbTrue Then
                Call MsgBox("'Delete old file' flag was set to 'vbTrue' in call to [ClsRegression].setName " & _
                        "on a Regression with no existing results Workbook.", vbExclamation + vbOKOnly, _
                        "Warning")
                    ' Do not exit
            End If
            ' Regardless of the setting, coerce the delete flag to vbFalse
            deleteOldFile = vbFalse
        End If
    Else
        ' Check savePath
        If savePath = "" Then
            ' Desire to retrieve path from existing book.  Book is open (and assumed validly saved),
            '  so rBook.Path should be valid...
            savePath = rBook.path
        Else
            ' savePath must exist; throw error if it doesn't, else continue through
            If Not fs.FolderExists(savePath) Then
                Call MsgBox("Save location passed to [ClsRegression].setName does not exist:" & _
                    vbLf & vbLf & """" & savePath & """", vbOKOnly + vbCritical, "Error")
                Exit Function
            End If
        End If
        ' rBook exists; need to create new book; delete of old will be done (or not) later
    End If  '  rBook Is Nothing
    
    ' Acting on the inputs
    ' Set the Regression name
    oldName = regName
    regName = newName
    
    ' Create or re-save the reg book
    If rBook Is Nothing Then ' must create new
        Set priorActiveBook = ActiveWorkbook
        Set rBook = Workbooks.Add
        priorActiveBook.Activate
        Set priorActiveBook = Nothing
    Else ' Save old fullname for later use
        oldBookFullName = rBook.fullName
    End If
    ' Perform the Save
    Call rBook.SaveAs(savePath & "\" & regName & RegressAux.rBookExtension)
    
    ' Successful name change
    setName = True
    'regCreated = True  ' Think this should not be set here? Not sure where/if it's used prior to
                        '  checking whether to write config after attaching source
    
    ' If deleting the old book, do that here.
    ' If selection was vbUseDefault, must query user whether to delete
    If deleteOldFile = vbUseDefault Then
        ' Query user whether to delete
        Select Case MsgBox("Rename complete." & vbLf & vbLf & "Delete file for old Regression """ & _
                    oldName & """?", vbQuestion + vbYesNo, "Delete old file?")
        Case vbYes
            deleteOldFile = vbTrue
        Case vbNo
            deleteOldFile = vbFalse
        End Select
    End If
    
    If deleteOldFile = vbTrue Then
        ' Was either passed directly or user indicated to delete
        ' Re-confirm that the old book exists, just in case, but presumably this will be error-proof.
        If Not fs.FileExists(oldBookFullName) Then
            Call MsgBox("Old Regression file not found -- unable to delete.", vbExclamation + vbOKOnly, "Warning")
            ' Don't kick out of Function, but don't attempt delete
        Else
            Call fs.DeleteFile(oldBookFullName, True) ' Error may throw here; shouldn't, though.
        End If
    End If
    
End Function

Public Function setDescription(newDesc As String) As Boolean
    ' Initialize no-good return
    setDescription = False
    
    ' FRAGILE(?) - no checking yet
    regDesc = newDesc
    
    setDescription = True
    
End Function

Public Function attachSource(book As Workbook, xData As Range, yData As Range, nameData As Range) As Boolean
    ' Only attach fully opened sources, with actual Ranges
'    If Not Me.isInitialized Then
'        Call Err.Raise(1801 + vbObjectError, "ClsRegression.attachSource", _
'            "Object not initialized.")
'    End If
    
    attachSource = False
    
    ' Attach objects
    Set srcBk = book
    Set SrcXRng = xData
    Set SrcYRng = yData
    Set SrcNameRng = nameData
    'SrcInclConst = includeConstant
    
'    ' Store paths and addresses (may be superfluous?)
'    SrcBkPath = SrcBk.fullName
'    SrcXRef = SrcXRng.Address(External:=True)
'    SrcYRef = SrcYRng.Address(External:=True)
'    SrcNameRef = SrcNameRng.Address(External:=True)
    
    ' Write config here, if reg fully created; otherwise, it's in initial stages of creation
    '  and the necessary object(s) won't have been bound yet
    If Me.isRegCreated Then
        writeConfig
    End If
    
    ' Now initialized
    sourceAttached = True
    attachSource = True
    
End Function

Public Function modifyRegression(Optional retainSrc As Boolean = False, _
            Optional includeConstant As Boolean = True) As Boolean
    
    ' Should hopefully be able to weave neatly into the same functions as used with
    '  createNewRegression...
    ' Main difference is just that it works within the (presumed) existing workbook -
    '  everything about the regression is still re-generated.
    
    ' FRAGILE: Assumes source links AND destination regression book have been properly set
    '   and are open.
    '
    
    ' Blank the 'last charted' values here
    Me.setDefaultLastChartedVars
    
    
    
    ' Initialize fail return value and de-set the regCreated flag
    modifyRegression = False
    regCreated = False
    
    ' Apply the includeConstant setting
    SrcInclConst = includeConstant
    
    ' If not retaining old source, pull new source; if sourcebook is .rgn.xlsx then
    '  reattach the source range to avoid fragility
    If Not retainSrc Or InStr(Me.SourceBookPath, RegressAux.rBookExtension) > 0 Then
        Me.bindSourceFromConfig
        ImportSrcData
    End If
        
    ' Recreated filtered data
    CreateFiltData
    
    ' Make the new V-Matrix and the results matrices
    MakeVMatrix
    MakeResMatrices
    
    ' Clear the stats sheet
    rStatSheet.Cells.Clear
    
    ' Write stats, config settings to sheets; and file to disk
    Me.writeStats
    Me.writeConfig
    Me.writeChanges
    
    ' Return success
    regCreated = True
    modifyRegression = True
    
End Function

Public Function createNewRegression(Optional includeConstant As Boolean = True, _
                Optional filtPtsStr As String = "()", _
                Optional filtPredsStr As String = "()") As Boolean
    ' Will want to break some of these functions into separate functions to allow for
    '  reuse in subsequent re-creations of the regression setup.
    
    ' POTENTIALLY FRAGILE: Assumes source links have been properly set
    
    ' Fail initial return value
    createNewRegression = False
    
    ' Apply the includeConstant setting
    SrcInclConst = includeConstant
    
    ' Create new workbook for holding regression info (TRANSFERRED into renaming function)
    'Set rBook = Workbooks.Add
    
    ' Clear all but one worksheet
    Application.DisplayAlerts = False
    Do Until rBook.Sheets.Count = 1
        rBook.Sheets(rBook.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Set as the config sheet; wait to populate until end of function call
    Set rConfigSheet = rBook.Worksheets(1)
    rConfigSheet.Name = RegressAux.configShtName
    rConfigSheet.Cells.NumberFormat = "@"   ' Set to text to avoid weirdness
    
    ' Add data sheet
    With rBook.Worksheets
        Set rDataSheet = .Add(After:=.Item(.Count))
    End With
    
    ' Rename sheet
    rDataSheet.Name = RegressAux.dataShtName
    
    ' Apply labels
    rDataSheet.Name = RegressAux.dataShtName
    rDataSheet.Cells(RegressAux.lblRow, RegressAux.dataYCol) = RegressAux.dataYName
    rDataSheet.Cells(RegressAux.lblRow, RegressAux.dataXCol) = RegressAux.dataXName
    
    ' Perform source data import
    ImportSrcData
    
    ' Apply any supplied filters
    Me.filterString(crfDataPoint) = filtPtsStr
    Me.filterString(crfPredictor) = filtPredsStr
    
    ' Add worksheet for filtered data sheet; CHANGED from just a straight copy of the
    '  source data sheet; should afford more consistent range handling
'    Call rDataSheet.Copy(After:=rDataSheet)
'    With rDataSheet.Parent.Worksheets
'        Set rFiltDataSheet = .Item(.Count)
'    End With
    With rBook.Worksheets
        Set rFiltDataSheet = .Add(After:=.Item(.Count))
    End With
    
    ' Rename sheet
    rFiltDataSheet.Name = RegressAux.dataFiltShtName
    
    ' Apply labels (identical to data sheet, except for sheet name)
    rFiltDataSheet.Name = RegressAux.dataFiltShtName
    rFiltDataSheet.Cells(RegressAux.lblRow, RegressAux.dataYCol) = RegressAux.dataYName
    rFiltDataSheet.Cells(RegressAux.lblRow, RegressAux.dataXCol) = RegressAux.dataXName
    
    ' Perform data copy and any filtering
    CreateFiltData
        
    ' Add worksheet for V-matrix
    With rBook.Worksheets
        Set rVSheet = .Add(After:=.Item(.Count))
    End With
    
    ' Configure the V-matrix sheet
    ' Apply labels
    rVSheet.Name = RegressAux.vMatShtName
    rVSheet.Cells(RegressAux.lblRow, RegressAux.vMatCol) = RegressAux.vMatName
    ' Make matrix
    MakeVMatrix
    
    ' Add worksheet for reporting results of fit
    With rBook.Worksheets
        Set rResSheet = .Add(After:=.Item(.Count))
    End With
    
    ' Configure the results sheet
    ' Add the labels
    ' !!!HERE!!! when new result fields added
    rResSheet.Name = RegressAux.resShtName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultYCol) = RegressAux.resYName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultBetaCol) = RegressAux.resBetaName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultBetaSECol) = RegressAux.resBetaSEName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultTBetaCol) = RegressAux.resTBetaName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultYHCol) = RegressAux.resYHName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultEHCol) = RegressAux.resEHName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultVIICol) = RegressAux.resVIIName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultRCol) = RegressAux.resRName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultTCol) = RegressAux.resTName
    rResSheet.Cells(RegressAux.lblRow, RegressAux.resultDCol) = RegressAux.resDName
    
    ' Make the results ranges
    MakeResMatrices
    
    ' Make and configure the stats sheet, then populate
    With rBook.Worksheets
        Set rStatSheet = .Add(After:=.Item(.Count))
    End With
    rStatSheet.Name = RegressAux.statShtName
    Set rStatRg = rStatSheet.Cells(RegressAux.statStartRow, statStartCol)
    Me.writeStats
    
    ' Write the config and save the file
    writeConfig
    Me.writeChanges
    
    ' Flag that the regression has been fully created
    regCreated = True
    
    ' Set a success return value
    createNewRegression = True
    
End Function

Public Function writeStats() As Boolean
    ' Should be trivial; should always succeed; will provide return value anyways
    writeStats = False
    
    Const RowOffsetANOVA As Long = 6
    
    ' Is (probably?) not going to be accessed programmatically later on, so can just
    '  hard-code the layout.  May regret it, but all of this is going to be drawn from
    '  info already in the Results sheet, so there should be no loss.
    ' Plan is to substantially reproduce the display in the main form, though adjusted
    '  to match the more typical ANOVA layout
    
    Dim iter As Long, offsetVal As Long
    
    ' Clear the stats sheet
    rStatSheet.Cells.Clear
    
    ' Go to! Start with response variable
    With rStatRg
        .Formula = "Response Variable:"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(0, 1).Formula = Me.responseName
    
    ' Data points info
    With rStatRg.Offset(2, 0)
        .Formula = "# Data Points (active/total)"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(2, 1).Formula = "'" & Me.numPoints(True) & _
            "/" & Me.numPoints(False) & " " & Me.filterString(crfDataPoint)
    
    ' Inactive predictors info
    With rStatRg.Offset(3, 0)
        .Formula = "Inactive Predictors:"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(3, 1).Formula = Me.filteredPredictors
    
    '' Summary of coefficients &c.
    ' Headers
    With rStatRg.Offset(5, 0)
        .Formula = "Predictor"
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With rStatRg.Offset(5, 1)
        .Formula = "Value"
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With rStatRg.Offset(5, 2)
        .Formula = "s.e."
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With rStatRg.Offset(5, 3)
        .Formula = "t-stat"
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With rStatRg.Offset(5, 4)
        .Formula = "p-value"
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ' Constant, if used
    If Me.includeConstant Then
        rStatRg.Offset(6, 0).Formula = "Constant"
        rStatRg.Offset(6, 1).Formula = Me.constantBeta
        rStatRg.Offset(6, 2).Formula = Me.constantBetaSE
        rStatRg.Offset(6, 3).Formula = Me.constantBeta / Me.constantBetaSE
        rStatRg.Offset(6, 4).Formula = RegressAux.calcPStat( _
                                        Me.constantBeta, _
                                        Me.constantBetaSE, _
                                        Me.DOF(DOFResidual))
'        Application.WorksheetFunction _
'                .T_Dist_2T(Abs(Me.constantBeta / Me.constantBetaSE), Me.DOF(DOFResidual))
    End If
    
    ' Predictors
    For iter = 1 To Me.numPredictors(True)
        offsetVal = 6 + iter + ConstOffset - 1
        With rStatRg.Offset(offsetVal, 0)
            .NumberFormat = "@"
            .Formula = Me.predictorName(iter, True)
        End With
        rStatRg.Offset(offsetVal, 1).Formula = Me.predictorBeta(iter)
        rStatRg.Offset(offsetVal, 2).Formula = Me.predictorBetaSE(iter)
        rStatRg.Offset(offsetVal, 3).Formula = Me.predictorBeta(iter) / _
                Me.predictorBetaSE(iter)
        rStatRg.Offset(offsetVal, 4).Formula = RegressAux.calcPStat( _
                                        Me.predictorBeta(iter), _
                                        Me.predictorBetaSE(iter), _
                                        Me.DOF(DOFResidual))
'        Application.WorksheetFunction _
'                .T_Dist_2T(Abs(Me.predictorBeta(iter) / Me.predictorBetaSE(iter)), _
'                Me.DOF(DOFResidual))
    Next iter
    
    ' R^2 value
    With rStatRg.Offset(1, 7)
        .Formula = "R^2"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(1, 8) = rResR2Rg.value
    
    ' Akaike stat, incl. corrected
    With rStatRg.Offset(2, 7)
        .Formula = "Akaike Information Criterion:"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(2, 8).Formula = rResAICRg.value
    With rStatRg.Offset(3, 7)
        .Formula = "Bias-Corrected AIC:"
        .HorizontalAlignment = xlHAlignRight
    End With
    rStatRg.Offset(3, 8).Formula = rResCorrAICRg.value
    
    ' ANOVA table
    rStatRg.Offset(RowOffsetANOVA, 8).Formula = "ANOVA"
    ' Sources
    With rStatRg.Offset(RowOffsetANOVA + 1, 7)
        .Formula = "Source"
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 7)
        .Formula = "Regression"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 3, 7)
        .Formula = "Residual"
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 4, 7)
        .Formula = "Total"
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' Degrees of freedom
    With rStatRg.Offset(RowOffsetANOVA + 1, 8)
        .Formula = "d.f."
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 8)
        .Formula = Me.DOF(DOFRegression)
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 3, 8)
        .Formula = Me.DOF(DOFResidual)
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 4, 8)
        .Formula = Me.DOF(DOFTotal)
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' SumSq Errors
    With rStatRg.Offset(RowOffsetANOVA + 1, 9)
        .Formula = "Sum Sq Err"
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 9)
        .Formula = Me.regStat(ST_SSRegression)
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 3, 9)
        .Formula = Me.regStat(ST_SSResidual)
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 4, 9)
        .Formula = Me.regStat(ST_SSTotal)
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' Mean Sq Err
    With rStatRg.Offset(RowOffsetANOVA + 1, 10)
        .Formula = "Mean Sq Err"
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 10)
        .Formula = .Offset(0, -1).value / .Offset(0, -2).value
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 3, 10)
        .Formula = .Offset(0, -1).value / .Offset(0, -2).value
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 4, 10)
        .Formula = .Offset(0, -1).value / .Offset(0, -2).value
        .HorizontalAlignment = xlHAlignCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' F-Stat
    With rStatRg.Offset(RowOffsetANOVA + 1, 11)
        .Formula = "F-stat"
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 11)
        .Formula = .Offset(0, -1).value / .Offset(1, -1).value
        .HorizontalAlignment = xlHAlignCenter
    End With
    rStatRg.Offset(RowOffsetANOVA + 4, 11).Borders(xlEdgeBottom).Weight = xlThin
    ' p-value
    With rStatRg.Offset(RowOffsetANOVA + 1, 12)
        .Formula = "p-value"
        .Borders(xlEdgeBottom).Weight = xlThin
        .HorizontalAlignment = xlHAlignCenter
    End With
    With rStatRg.Offset(RowOffsetANOVA + 2, 12)
        .Formula = Application.WorksheetFunction.F_Dist_RT(Abs(.Offset(0, -1).value), _
                    .Offset(0, -4).value, .Offset(1, -4).value)
        .HorizontalAlignment = xlHAlignCenter
    End With
    rStatRg.Offset(RowOffsetANOVA + 4, 12).Borders(xlEdgeBottom).Weight = xlThin
    
    rStatRg.Offset(RowOffsetANOVA + 1, 0).CurrentRegion.Columns.AutoFit
    rStatRg.Offset(RowOffsetANOVA + 1, 7).CurrentRegion.Columns.AutoFit
    
    Me.writeChanges
    
    writeStats = True
End Function

Public Function writeChanges() As Boolean
    ' Wrapper for anything that has to happen when changes are written to disk
    
    ' Initialize fail return
    writeChanges = False
    
    ' TRAP FOR non-initialized things!
    
    If Not CustDocPropExists(RegressAux.timeStampDocPropName) Then
        Call rBook.CustomDocumentProperties.Add(Name:=RegressAux.timeStampDocPropName, _
                LinkToContent:=False, Type:=4, value:=Format(Now, RegressAux.timeStampFormat)) ' 4 means String type
    Else
        rBook.CustomDocumentProperties(RegressAux.timeStampDocPropName).value = _
                Format(Now, RegressAux.timeStampFormat)
    End If
    rBook.Save
    
    writeChanges = True
    
End Function

Public Function makeChart(abscissa As CRChartVar, ordinate As CRChartVar, _
        normalizeX As Boolean, normalizeY As Boolean, _
        abscissaIndex As Long, ordinateIndex As Long, _
        doOutliers As Boolean, outlierAlpha As Double, chartSize As CRChartSize) As Boolean
    
    Dim iter As Long, testBool As Boolean
    Dim wkSrs As Series, wkPt As Point, wkTl As Trendline, wkTlDl As DataLabel
    Dim tThresh As Double, normXVal As Double, normYVal As Double
    
    ' If properly proofed, this should pretty much never fail
    makeChart = False
    
    ' If either of abscissa or ordinate are predictors, index must be valid
    If abscissa = crcvPredictor And (abscissaIndex < 1 Or abscissaIndex > Me.numPredictors(True)) Then
        Call MsgBox("Valid predictor index must be specified!", vbOKOnly + vbExclamation, "Error")
        Exit Function
    End If
    If ordinate = crcvPredictor And (ordinateIndex < 1 Or ordinateIndex > Me.numPredictors(True)) Then
        Call MsgBox("Valid predictor index must be specified!", vbOKOnly + vbExclamation, "Error")
        Exit Function
    End If
    
    ' If doing outliers, alpha must fall between zero and one
    If doOutliers And Not (outlierAlpha > 0 And outlierAlpha < 1) Then
        Call MsgBox("Outlier alpha value must be between zero and one!", vbOKOnly + vbCritical, "Error")
        Exit Function
    End If
    
    ' If doesn't exist, create new chart sheet and link; else, destroy and re-create
    With rBook
        ' Possibly fragile, if perhaps rDynChart is assigned to something weird
        If Not rDynChart Is Nothing Then
            Application.DisplayAlerts = False
            rDynChart.Delete
            Application.DisplayAlerts = True
        End If
        
        Set rDynChart = .Charts.Add(After:=.Sheets(.Sheets.Count))
        rDynChart.Name = RegressAux.chartShtName
    End With
    
    ' Clear the chart - possibly want to reset other things...?
    For Each wkSrs In rDynChart.SeriesCollection
        wkSrs.Delete
    Next wkSrs
    
    ' Set the chart type to scatter
    rDynChart.ChartType = xlXYScatter
    
    ' Make a new series
    Set wkSrs = rDynChart.SeriesCollection.NewSeries
    
    ' Change to scatter
    ' Add data to it
    Select Case ordinate
    Case crcvCookDistance
        wkSrs.Values = MakeSeriesArray(rResDRg, normalizeY)
        normYVal = GetMaxAmpl(rResDRg)
    Case crcvDirectResidual
        wkSrs.Values = MakeSeriesArray(rResEHatRg, normalizeY)
        normYVal = GetMaxAmpl(rResEHatRg)
    Case crcvFittedResponse
        wkSrs.Values = MakeSeriesArray(rResYHatRg, normalizeY)
        normYVal = GetMaxAmpl(rResYHatRg)
    Case crcvPredictor
        ' Must account for constant or not
        wkSrs.Values = MakeSeriesArray(rXFiltData.Columns(ordinateIndex + ConstOffset), normalizeY)
        normYVal = GetMaxAmpl(rXFiltData.Columns(ordinateIndex + ConstOffset))
    Case crcvResponse
        wkSrs.Values = MakeSeriesArray(rResYRg, normalizeY)
        normYVal = GetMaxAmpl(rResYRg)
    Case crcvStudentizedResidual
        wkSrs.Values = MakeSeriesArray(rResRRg, normalizeY)
        normYVal = GetMaxAmpl(rResRRg)
    Case crcvTStatResidual
        wkSrs.Values = MakeSeriesArray(rResTRg, normalizeY)
        normYVal = GetMaxAmpl(rResTRg)
    Case crcvPointSequence
        wkSrs.Values = MakeNumSeqArray(1, rXFiltData.Rows.Count)
        normYVal = rXFiltData.Rows.Count
    End Select
    
    Select Case abscissa
    Case crcvCookDistance
        wkSrs.XValues = MakeSeriesArray(rResDRg, normalizeX)
        normXVal = GetMaxAmpl(rResDRg)
    Case crcvDirectResidual
        wkSrs.XValues = MakeSeriesArray(rResEHatRg, normalizeX)
        normXVal = GetMaxAmpl(rResEHatRg)
    Case crcvFittedResponse
        wkSrs.XValues = MakeSeriesArray(rResYHatRg, normalizeX)
        normXVal = GetMaxAmpl(rResYHatRg)
    Case crcvPredictor
        wkSrs.XValues = MakeSeriesArray(rXFiltData.Columns(abscissaIndex + ConstOffset), normalizeX)
        normXVal = GetMaxAmpl(rXFiltData.Columns(abscissaIndex + ConstOffset))
    Case crcvResponse
        wkSrs.XValues = MakeSeriesArray(rResYRg, normalizeX)
        normXVal = GetMaxAmpl(rResYRg)
    Case crcvStudentizedResidual
        wkSrs.XValues = MakeSeriesArray(rResRRg, normalizeX)
        normXVal = GetMaxAmpl(rResRRg)
    Case crcvTStatResidual
        wkSrs.XValues = MakeSeriesArray(rResTRg, normalizeX)
        normXVal = GetMaxAmpl(rResTRg)
    Case crcvPointSequence
        wkSrs.XValues = MakeNumSeqArray(1, rXFiltData.Rows.Count)
        normXVal = rXFiltData.Rows.Count
    End Select
    
    '''' Apply chart-size-independent formatting
    ' Direct to series
    With wkSrs
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerBackgroundColorIndex = xlColorIndexNone
        .MarkerForegroundColor = RGB(0, 0, 0)
        
        ' Remove existing trendlines
        Do While .Trendlines.Count > 0
            .Trendlines(1).Delete
        Loop
        
        ' Add a new one; include the datalabel, to possibly be deleted later
        Set wkTl = .Trendlines.Add(Type:=xlLinear, DisplayEquation:=True, DisplayRSquared:=True)
    End With
    
    ' Chart-size-independent formatting on other parts of the chart
    With rDynChart
        ' No ChartArea border
        .ChartArea.Format.Line.Visible = msoFalse
        
        .PlotArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        
        With .Axes(xlValue, xlPrimary)
            ' Delete all gridlines, will re-add later
            If .HasMinorGridlines Then .MinorGridlines.Delete
            If .HasMajorGridlines Then .MajorGridlines.Delete
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True
            .ScaleType = xlScaleLinear
            .MajorTickMark = xlTickMarkCross
        End With
        
        With .Axes(xlCategory, xlPrimary)
            If .HasMajorGridlines Then .MajorGridlines.Delete
            If .HasMinorGridlines Then .MinorGridlines.Delete
            .MaximumScaleIsAuto = True
            .MinimumScaleIsAuto = True
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True
            .ScaleType = xlScaleLinear
            .MajorTickMark = xlTickMarkCross
        End With
        
        ' Delete legend if present
        If .HasLegend Then .Legend.Delete
    End With
    
    ' If outlier highlighting specified, apply it
    ' This uses the Bonferroni inequality as described in Weisberg, p116 (essentially, the
    '  desired alpha is divided by the number of points _n_ being tested for outliers, which
    '  corrects for there effectively being _n_ separate tests performed.
    If doOutliers Then
        tThresh = Application.WorksheetFunction.T_Inv_2T(outlierAlpha / Me.numPoints(True), _
                    Me.DOF(DOFResidual))
        For iter = 1 To rXFiltData.Rows.Count
            If Abs(rResTRg.Cells(iter, 1).value) > tThresh Then
                ' Color the data point red
                wkSrs.Points(iter).MarkerForegroundColor = RGB(255, 0, 0)
            End If
        Next iter
    End If
    
    ' Apply the chart-size-specific formatting
    Call ApplyChartSizing(abscissa, ordinate, normalizeX, normalizeY, _
            normYVal, normXVal, abscissaIndex, ordinateIndex, chartSize)
    
    ' Only some chart sizes follow up on the trendline datalabel; don't do anything here
    
    ' Show the chart
    rBook.Activate
    rDynChart.Activate
    
    ' Update internal tracking values
    lastChartX = abscissa
    lastChartXPred = abscissaIndex
    lastChartY = ordinate
    lastChartYPred = ordinateIndex
    lastChartXNorm = normalizeX
    lastChartYNorm = normalizeY
    lastChartDoOutliers = doOutliers
    lastChartOutlierAlpha = outlierAlpha
    lastChartSize = chartSize
    
    ' Write config changes and save the book
    writeConfig
    writeChanges
    
    makeChart = True
    
End Function

Public Sub copyDataRange(rangeType As CRCopyRange)
    Select Case rangeType
    Case crcrStats
        rStatRg.Worksheet.UsedRange.Copy
    Case Else
        MsgBox "Not implemented!"
    End Select
End Sub

Private Sub ApplyChartSizing(abscissa As CRChartVar, ordinate As CRChartVar, _
        normalizeX As Boolean, normalizeY As Boolean, _
        normYVal As Double, normXVal As Double, _
        abscissaIndex As Long, ordinateIndex As Long, _
        chartSize As CRChartSize)
        
    ' Probably not any very elegant ways to do this...
    
    Dim wkSrs As Series, wkTl As Trendline, wkDl As DataLabel
    
    Dim bigFontSize As Long, smallFontSize As Long
    Dim markerLineWeight As Double, lineLineWeight As Double, plotBorderWeight As XlBorderWeight
    Dim markerPtSize As Double, dataLabelLineWeight As Double, axisLineWeight As Double
    Dim numSigFigs As Long
    Dim showTrendlineDataLabel As Boolean, showNormFactors As Boolean
    Dim doYGridlines As Boolean, showAxisTitles As Boolean, showAxes As Boolean
    Dim gridlineDashStyle As MsoLineDashStyle, gridlineWeight As Double
    
    Set wkSrs = rDynChart.SeriesCollection(1)
    
    Select Case chartSize
    ' Full Page Figure
    Case crcsFullPage
        bigFontSize = 18
        smallFontSize = 16
        markerLineWeight = 1.5
        markerPtSize = 8#
        lineLineWeight = 2.5
        plotBorderWeight = xlHairline
        axisLineWeight = 1#
        dataLabelLineWeight = 2.5
        numSigFigs = 5
        showTrendlineDataLabel = True
        showNormFactors = True
        doYGridlines = True
        gridlineDashStyle = msoLineLongDash
        gridlineWeight = 1.5
        showAxisTitles = True
        showAxes = True
    Case crcsFigure
        bigFontSize = 28
        smallFontSize = 24
        markerLineWeight = 3
        markerPtSize = 10#
        lineLineWeight = 4#
        plotBorderWeight = xlThin
        axisLineWeight = 2#
        dataLabelLineWeight = 3#
        numSigFigs = 3
        showTrendlineDataLabel = True
        showNormFactors = True
        doYGridlines = False
        gridlineDashStyle = msoLineDash ' Specified just in case
        gridlineWeight = 2#
        showAxisTitles = True
        showAxes = True
    Case crcsThumbnail
        bigFontSize = 48
        smallFontSize = 36
        markerLineWeight = 4#
        markerPtSize = 13#
        lineLineWeight = 7#
        plotBorderWeight = xlThick
        axisLineWeight = 4#
        dataLabelLineWeight = 3#  ' Irrelevant, not showing label
        numSigFigs = 2
        showTrendlineDataLabel = False
        showNormFactors = False
        doYGridlines = False
        gridlineDashStyle = msoLineDash ' Specified just in case
        gridlineWeight = 4#
        showAxisTitles = False
        showAxes = True
    End Select
    
    With wkSrs
        ' Visual formatting
        .MarkerSize = markerPtSize
        .Format.Line.Weight = markerLineWeight
          ' Will 'disappear' the connecting line later
        
        ' Should only be one trendline to bind; bind it and the datalabel
        Set wkTl = .Trendlines(1)
        Set wkDl = wkTl.DataLabel
    End With
    
    With rDynChart
        .PlotArea.Border.Weight = plotBorderWeight
        
        If showAxes Then
            With .Axes(xlValue, xlPrimary)
                .Format.Line.Weight = axisLineWeight
                
                ' Add gridlines if indicated
                If doYGridlines Then
                    .HasMajorGridlines = True
                    With .MajorGridlines.Format.Line
                        .ForeColor.RGB = RGB(191, 191, 191)
                        .DashStyle = gridlineDashStyle
                        .Weight = gridlineWeight
                    End With
                End If
                
                .TickLabels.Font.Size = smallFontSize
                .TickLabelPosition = xlTickLabelPositionLow
                
                If showAxisTitles Then
                    .HasTitle = True
                    With .AxisTitle
                        .Font.Size = bigFontSize
                        .Font.Bold = True
                        Select Case ordinate
                        Case crcvCookDistance
                            .Caption = "Cook's Distance Metric"
                        Case crcvDirectResidual
                            .Caption = "Residual"
                        Case crcvFittedResponse
                            .Caption = "Fitted " & Me.responseName
                        Case crcvPredictor
                            .Caption = Me.predictorName(ordinateIndex, True)
                        Case crcvResponse
                            .Caption = Me.responseName
                        Case crcvStudentizedResidual
                            .Caption = "Studentized Residual"
                        Case crcvTStatResidual
                            .Caption = "Residual t-Statistic"
                        Case crcvPointSequence
                            .Caption = "Sequence #"
                        End Select
                        If normalizeY Then
                            .Caption = .Caption & " [normalized"
                            If showNormFactors Then
                                .Caption = .Caption & " (" & RegressAux.roundSigs(normYVal, numSigFigs) & ")"
                            End If
                            .Caption = .Caption & "]"
                        End If
                    End With
                Else
                    .HasTitle = False
                End If
            End With
            
            With .Axes(xlCategory, xlPrimary)
                .Format.Line.Weight = axisLineWeight
                
                ' No gridlines
                
                .TickLabels.Font.Size = smallFontSize
                .TickLabelPosition = xlTickLabelPositionLow
                
                If showAxisTitles Then
                    .HasTitle = True
                    With .AxisTitle
                        .Font.Size = bigFontSize
                        .Font.Bold = True
                        Select Case abscissa
                        Case crcvCookDistance
                            .Caption = "Cook's Distance Metric"
                        Case crcvDirectResidual
                            .Caption = "Residual"
                        Case crcvFittedResponse
                            .Caption = "Fitted " & Me.responseName
                        Case crcvPredictor
                            .Caption = Me.predictorName(abscissaIndex, True)
                        Case crcvResponse
                            .Caption = Me.responseName
                        Case crcvStudentizedResidual
                            .Caption = "Studentized Residual"
                        Case crcvTStatResidual
                            .Caption = "Residual t-Statistic"
                        Case crcvPointSequence
                            .Caption = "Sequence #"
                        End Select
                        If normalizeX Then
                            .Caption = .Caption & " [normalized"
                            If showNormFactors Then
                                .Caption = .Caption & " (" & RegressAux.roundSigs(normXVal, numSigFigs) & ")"
                            End If
                            .Caption = .Caption & "]"
                        End If
                    End With
                Else
                    .HasTitle = False
                End If
            End With
        Else
            .Axes(xlCategory, xlPrimary).Delete
            .Axes(xlValue, xlPrimary).Delete
        End If
    End With
    
    ' Set lineweight and coerce to markers only
    With wkSrs
        .Format.Line.Weight = markerLineWeight
        .ChartType = xlXYScatter
    End With
    
    ' Set trendline lineweight and color
    With wkTl.Format.Line
        .Weight = lineLineWeight
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    ' If keeping it, thicken the datalabel line, make it black, make it white-fill; else delete
    If showTrendlineDataLabel Then
        With wkDl
            With .Format.Line
                .ForeColor.RGB = RGB(0, 0, 0)
                .Weight = dataLabelLineWeight
            End With
            .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .NumberFormat = "0." & String(numSigFigs - 1, "0") ' This may need some tweaking...
            .Font.Size = smallFontSize
            
            ' The positioning changes here must be the LAST THING DONE!
            .Left = rDynChart.PlotArea.Left + rDynChart.Axes(xlValue, xlPrimary).Width + 20
            .Top = rDynChart.PlotArea.Top + 20
        End With
    Else
        wkDl.Delete
    End If
    
End Sub


Public Sub setLastChartedVars(lastX As CRChartVar, lastY As CRChartVar, lastXPred As Long, _
            lastYPred As Long, lastXNorm As Boolean, lastYNorm As Boolean, _
            lastDoOutliers As Boolean, lastOutlierAlpha As Double, lastChSize As CRChartSize)
    
    lastChartX = lastX
    lastChartY = lastY
    lastChartXPred = lastXPred
    lastChartYPred = lastYPred
    lastChartXNorm = lastXNorm
    lastChartYNorm = lastYNorm
    lastChartDoOutliers = lastDoOutliers
    lastChartOutlierAlpha = lastOutlierAlpha
    lastChartSize = lastChSize
    
End Sub

Public Function addFilter(idNum As Long, filterType As CRFilterType) As Boolean
    ' True means Filter was set and regression should be regenerated
    ' False means problem occurred; no changes effected and regression should not be
    '   regenerated
    
    Dim typeStr As String
    
    addFilter = False
    
    ' For later use in errorbox
    Select Case filterType
    Case crfDataPoint
        typeStr = "data point"
    Case crfPredictor
        typeStr = "predictor"
    End Select
    
    ' Check whether the id of the given type is filtered
    Select Case Me.isFiltered(idNum, filterType)
    Case vbUseDefault
        ' ID is invalid, error and exit
        Call MsgBox("Invalid " & typeStr & " filter ID", vbCritical + vbOKOnly, "Error")
        Exit Function
    Case vbFalse
        ' ID is not filtered; filter, if not prohibited by over-determination or the
        '  need for at least one predictor. Return success/fail from the subfunction
        '  if filter-add attempted
        Select Case filterType
        Case crfDataPoint
            If Not (Me.numPoints - Me.countFilters(crfDataPoint)) > _
                    (Me.numPredictors - Me.countFilters(crfPredictor)) Then
                ' Adding a filter would cause overspecification! Can't add another
                Call MsgBox("Addition of another data point filter would cause regression " & _
                        "over-specification!  Canceling...", vbOKOnly + vbCritical, "Error")
                Exit Function
            Else
                ' Go ahead and try adding the filter
                addFilter = AddPtFilt(idNum)
            End If
        Case crfPredictor
            If Not Me.countFilters(crfPredictor) < Me.numPredictors - 1 Then
                ' Too many filters! Can't add another.
                Call MsgBox("Regression must have at least one predictor!", _
                        vbOKOnly + vbCritical, "Error")
                Exit Function
            Else
                ' Go ahead and try adding the filter
                addFilter = AddPredFilt(idNum)
            End If
        End Select
    Case vbTrue
        ' ID already filtered.
        Call MsgBox("The " & typeStr & " #" & idNum & " is already filtered!", _
                vbExclamation + vbOKOnly, "Alert")
        Exit Function
    End Select
    
End Function

Public Function countFilters(filterType As CRFilterType) As Long
    ' Returns the number of filters currently applied to the given type
    ' -1 is failure return
    
    countFilters = -1
    
    Select Case filterType
    Case crfDataPoint
        countFilters = CountPtFilts
    Case crfPredictor
        countFilters = CountPredFilts
    End Select
    
End Function

Public Function delFilter(idNum As Long, filterType As CRFilterType) As Boolean
    ' True return means filter was deleted and regression should be regenerated
    ' False means filter NOT deleted for some reason, and regression should not
    '  be regenerated.
    
    delFilter = False
    
    Dim typeStr As String
    
    ' For later use in errorbox
    Select Case filterType
    Case crfDataPoint
        typeStr = "data point"
    Case crfPredictor
        typeStr = "predictor"
    End Select
    
    ' Check whether the indicated filter is valid/already applied
    Select Case Me.isFiltered(idNum, filterType)
    Case vbUseDefault
        ' ID is invalid, error and exit
        Call MsgBox("Invalid " & typeStr & " filter ID", vbCritical + vbOKOnly, "Error")
        Exit Function
    Case vbTrue
        ' Filter for the ID is present; remove filter, if not prohibited by
        ' over-determination Return success/fail from the subfunction if
        '  filter-delete is attempted
        Select Case filterType
        Case crfDataPoint
            ' Removing a data point filter that exists should always be
            '  acceptable
            delFilter = DelPtFilt(idNum)
        Case crfPredictor
            ' Have to make sure that removing a predictor filter doesn't
            '  lead to an over-determined regression
            If Not (Me.numPredictors - Me.countFilters(crfPredictor)) < _
                    (Me.numPoints - Me.countFilters(crfDataPoint)) Then
                ' Removing a predictor filter would cause over-determination!
                '  Can't remove.
                Call MsgBox("Removal of a predictor filter would cause regression " & _
                        "over-specification!  Canceling...", vbOKOnly + vbCritical, "Error")
                Exit Function
            Else
                ' Go ahead and try removing the filter
                delFilter = DelPredFilt(idNum)
            End If
        End Select
    Case vbFalse
        ' Filter ID not present!
        Call MsgBox("Filter for " & typeStr & " #" & idNum & " is not present!", _
                vbExclamation + vbOKOnly, "Alert")
        Exit Function
    End Select

End Function

Public Function getFilterSeq(seqNum As Long, filterType As CRFilterType) As Long
    ' -1 Return value means invalid sequence number provided
    ' Used for iterating over the filtered values
    
    getFilterSeq = -1
    
    If seqNum < 1 Or seqNum > Me.countFilters(filterType) Then Exit Function
    
    Select Case filterType
    Case crfDataPoint
        getFilterSeq = filtPts(LBound(filtPts) + seqNum - 1)
    Case crfPredictor
        getFilterSeq = filtPreds(LBound(filtPreds) + seqNum - 1)
    End Select
    
End Function

Public Function clearFilters(filterType As CRFilterType) As Boolean
    ' False - clear operation failed somehow -- SHOULD NEVER OCCUR
    ' True - Filters are now cleared
    
    clearFilters = False
    
    Select Case filterType
    Case crfDataPoint
        filtPts = Array(0)
    Case crfPredictor
        filtPreds = Array(0)
    End Select
    
    clearFilters = True
    
End Function

Property Get filteredPredictors() As String
    ' Assemble semicolon-separated list of named, filtered predictors
    Dim cel As Range, workStr As String, iter As Long
    
    ' Initialize
    workStr = ""
    
    ' Special return if nothing filtered
    If Me.countFilters(crfPredictor) < 1 Then
        filteredPredictors = "(none)"
        Exit Property
    End If
    
    ' Cycle through predictor filters
    For iter = 1 To Me.countFilters(crfPredictor)
        workStr = workStr & Trim(Me.predictorName(Me.getFilterSeq(iter, crfPredictor), False)) & "; "
    Next iter
    
    ' Strip final blip if present (should always be present at this point)
    If Right(workStr, 2) = "; " Then workStr = Left(workStr, Len(workStr) - 2)
    
    filteredPredictors = workStr
    
End Property

Property Get filterString(filterType As CRFilterType) As String
    ' Assemble indicated filter array into comma-separated list within parentheses
    Dim ary As Variant, iter As Long
    
    ' Check for no filter case; if filters there, assign working array
    Select Case filterType
    Case crfDataPoint
        If Me.countFilters(crfDataPoint) = 0 Then
            filterString = "()"
            Exit Property
        Else
            ary = filtPts
        End If
    Case crfPredictor
        If Me.countFilters(crfPredictor) = 0 Then
            filterString = "()"
            Exit Property
        Else
            ary = filtPreds
        End If
    End Select
    
    ' Some filters present: assemble string
    filterString = "("
    
    ' Iterate over array
    For iter = LBound(ary) To UBound(ary)
        filterString = filterString & ary(iter) & ","
    Next iter
    
    ' Close the array w/paren, stripping extra comma
    filterString = Left(filterString, Len(filterString) - 1) & ")"
    
End Property

Property Let filterString(filterType As CRFilterType, fString As String)
    ' This OVERWRITES ALL FILTERS for the given CRFilterType.
    ' Should be able to effectively RegEx the proofing of the string...?
    Dim rx As New RegExp, mColl As MatchCollection, mch As Match
    Dim oldFilters As String
    Dim workStr As String
    
    ' Allow a null string input to double as 'no filters'
    If fString = "" Then
        workStr = "()"
    Else
        workStr = fString
    End If
    
    With rx
        .IgnoreCase = True ' Should be irrelevant
        .Global = True  ' Should be irrelevant
        .MultiLine = True  ' Should be irrelevant
        .Pattern = "^\([0-9,]*\)$"  ' Proofing pattern for general structure
        If Not .Test(workStr) Then ' Doesn't match the expected format
            Call MsgBox("Invalid format of filter array string." & vbLf & vbLf & _
                    typeStr(filterType, True) & " filters not updated.", _
                    vbOKOnly + vbCritical, "Error")
            Exit Property
        End If
        
        ' Reset to look for just sequences of digits, and Execute
        .Pattern = "[0-9]+"
        Set mColl = .Execute(workStr)
    End With
    
    ' Save the old filters, in case
    oldFilters = Me.filterString(filterType)
    
    ' Clear the filters
    If Not Me.clearFilters(filterType) Then
        Call MsgBox("Clearing of " & typeStr(filterType) & " filters failed." & vbLf & _
                vbLf & "Aborting filter import.", vbOKOnly + vbCritical, "Error")
        Exit Property
    End If
    
    ' Run through the matches, warning about but otherwise ignoring bad IDs
    For Each mch In mColl
        ' Warn if the filter to be added is already presentor invalid; else, add
        Select Case Me.isFiltered(CLng(mch.value), filterType)
        Case vbUseDefault
            Call MsgBox("Invalid " & typeStr(filterType) & " filter ID #" & _
                    Format(CLng(mch.value), "0") & " present in input." & vbLf & vbLf & _
                    "Ignoring...", vbOKOnly + vbInformation, "Alert")
        Case vbTrue
            Call MsgBox("Duplicate " & typeStr(filterType) & " filter ID #" & _
                    Format(CLng(mch.value), "0") & " present in input." & vbLf & vbLf & _
                    "Ignoring duplicate...", vbOKOnly + vbInformation, _
                    "Alert")
        Case vbFalse
            ' Add the filter; toss error and restore old filters if fails.
            If Not Me.addFilter(CLng(mch.value), filterType) Then
                Me.filterString(filterType) = oldFilters ' Presuming this will work, else recursive
                Call MsgBox("Error attempting additino of " & typeStr(filterType) & _
                        " filter ID #" & Format(CLng(mch.value), "0") & "." & vbLf & vbLf & _
                        "Leaving filters unchanged...", vbOKOnly + vbCritical, "Error")
                Exit Property
            End If
        End Select
    Next mch
    
End Property

Property Get regStat(rStatType As StatType) As Double
    Select Case rStatType
    Case ST_SSTotal
        regStat = rResSYYRg.value
    Case ST_SSResidual
        regStat = rResRSSRg.value
    Case ST_SSRegression
        regStat = rResRegSSRg.value
    Case ST_RSq
        regStat = rResR2Rg.value
    Case ST_FStat
        regStat = rResFStatRg.value
    Case ST_FStatPVal
        regStat = Application.WorksheetFunction.F_Dist_RT(Abs(rResFStatRg.value), _
                Me.DOF(DOFRegression), Me.DOF(DOFResidual))
    Case ST_AIC
        regStat = rResAICRg.value
    Case ST_CorrAIC
        regStat = rResCorrAICRg.value
    End Select
End Property

Property Get regStatName(rStatType As StatType) As String
    Select Case rStatType
    Case ST_SSTotal
        regStatName = rResSYYNameRg.value
    Case ST_SSResidual
        regStatName = rResRSSNameRg.value
    Case ST_SSRegression
        regStatName = rResRegSSNameRg.value
    Case ST_RSq
        regStatName = rResR2NameRg.value
    Case ST_FStat
        regStatName = rResFStatNameRg.value
    Case ST_FStatPVal
        regStatName = rResFStatNameRg.value & " (p-value)"
    Case ST_AIC
        regStatName = rResAICNameRg.value
    Case ST_CorrAIC
        regStatName = rResCorrAICNameRg.value
    End Select
End Property

Private Function typeStr(filterType As CRFilterType, Optional initCaps As Boolean = False) As String
    If initCaps Then
        Select Case filterType
        Case crfDataPoint
            typeStr = "Data point"
        Case crfPredictor
            typeStr = "Predictor"
        End Select
    Else
        Select Case filterType
        Case crfDataPoint
            typeStr = "data point"
        Case crfPredictor
            typeStr = "predictor"
        End Select
    End If
End Function

Private Function CustDocPropExists(Key As String, Optional wbk As Workbook = Nothing) As Boolean
    Dim dp As DocumentProperty
    
    CustDocPropExists = False
    
    ' If no book passed, assume internal book
    If wbk Is Nothing Then Set wbk = rBook
    
    With wbk.CustomDocumentProperties
        If .Count > 0 Then
            For Each dp In wbk.CustomDocumentProperties
                If dp.Name = Key Then CustDocPropExists = True
            Next dp
        End If
    End With
    
End Function

Private Sub WriteSourceConfig()
    With rConfigSheet
        .Cells(crcSourceBook, RegressAux.configValCol).Formula = Me.SourceBookPath
        .Cells(crcSourceXRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsXData)
        .Cells(crcSourceYRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsYData)
        .Cells(crcSourceNameRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsNameData)
    End With
End Sub

Public Sub writeConfig()
    With rConfigSheet
        .Cells(1, RegressAux.configValCol).EntireColumn.ColumnWidth = 50
        .Cells(crcSourceBook, RegressAux.configValCol).Formula = Me.SourceBookPath
        .Cells(crcSourceXRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsXData)
        .Cells(crcSourceYRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsYData)
        .Cells(crcSourceNameRange, RegressAux.configValCol).Formula = "'" & Me.SourceAddress(crsNameData)
        .Cells(crcIncludeConstant, RegressAux.configValCol).Formula = UCase(CStr(Me.includeConstant))
        .Cells(crcDescription, RegressAux.configValCol).Formula = Me.Description
        .Cells(crcFiltPts, RegressAux.configValCol).Formula = Me.filterString(crfDataPoint)
        .Cells(crcFiltPreds, RegressAux.configValCol).Formula = Me.filterString(crfPredictor)
        .Cells(crcRespName, RegressAux.configValCol).Formula = Me.responseName
        .Cells(crcLastChartX, RegressAux.configValCol).Formula = CLng(Me.getLastChartedX)
        .Cells(crcLastChartY, RegressAux.configValCol).Formula = CLng(Me.getLastChartedY)
        .Cells(crcLastChartXPred, RegressAux.configValCol).Formula = Me.getLastChartedXPred
        .Cells(crcLastChartYPred, RegressAux.configValCol).Formula = Me.getLastChartedYPred
        .Cells(crcLastChartXNorm, RegressAux.configValCol).Formula = UCase(CStr(Me.getLastChartedXNorm))
        .Cells(crcLastChartYNorm, RegressAux.configValCol).Formula = UCase(CStr(Me.getLastChartedYNorm))
        .Cells(crcLastChartDoOutliers, RegressAux.configValCol).Formula = UCase(CStr(Me.getLastChartedDoOutliers))
        .Cells(crcLastChartOutlierAlpha, RegressAux.configValCol).Formula = CStr(Me.getLastChartedOutlierAlpha)
        .Cells(crcLastChartSize, RegressAux.configValCol).Formula = CStr(Me.getLastChartedSize)
    End With
    'rBook.CustomDocumentProperties(RegressAux.timeStampDocPropName) = CStr(Now)
End Sub

Private Sub ImportSrcData()
    ' Copy source data from its location and bind the relevant Ranges
    
    ' Scrub contents of old ranges if already defined
    If Not rXData Is Nothing Then rXData.Clear
    If Not rYData Is Nothing Then rYData.Clear
    If Not rNameData Is Nothing Then rNameData.Clear
    
    With rDataSheet
        ' Copy in predictor data, accounting for constant column
        Call SrcXRng.Copy
        If SrcInclConst Then
            Call .Cells(RegressAux.firstDataRow, RegressAux.dataXCol + 1).PasteSpecial(xlPasteValues)
        Else
            Call .Cells(RegressAux.firstDataRow, RegressAux.dataXCol).PasteSpecial(xlPasteValues)
        End If
        
        ' Set number format of name row to Text
        .Cells(RegressAux.dataPredNameRow, 1).EntireRow.NumberFormat = "@"
        
        ' Copy in predictor name data, accounting for constant column and possible need to transpose
        Call SrcNameRng.Copy
        ' SrcNameRng needs already to have been proofed to be a vector of the right size
        If SrcNameRng.Rows.Count > 1 Then  ' Need to transpose
            If SrcInclConst Then ' Shift by one and add the constant name
                Call .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol + 1) _
                                .PasteSpecial(Paste:=xlPasteValues, Transpose:=True)
                .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol) = RegressAux.constName
            Else ' Don't shift or add constant name
                Call .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol) _
                                .PasteSpecial(Paste:=xlPasteValues, Transpose:=True)
            End If  ' includeConstant
        Else    ' Do not transpose
            If SrcInclConst Then  ' Shift by one and add the constant name
                Call .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol + 1) _
                                .PasteSpecial(Paste:=xlPasteValues, Transpose:=False)
                .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol) = RegressAux.constName
            Else ' Don't shift or add constant name
                Call .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol) _
                                .PasteSpecial(Paste:=xlPasteValues, Transpose:=False)
            End If  ' SrcInclConst
        End If  ' SrcNameRng.Rows.Count > 1
        
        ' Copy in response data
        Call SrcYRng.Copy
        Call .Cells(RegressAux.firstDataRow, RegressAux.dataYCol).PasteSpecial(xlPasteValues)
        
        ' Attach the local Ranges for the source data, accounting for constant
        '  column or not
        If SrcInclConst Then
            Set rXData = .Cells(RegressAux.firstDataRow, RegressAux.dataXCol + 1).CurrentRegion
        Else
            Set rXData = .Cells(RegressAux.firstDataRow, RegressAux.dataXCol).CurrentRegion
        End If
        Set rYData = .Cells(RegressAux.firstDataRow, RegressAux.dataYCol).CurrentRegion
        Set rNameData = .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).CurrentRegion
        
        If SrcInclConst Then ' add in the column of ones
            rXData.Offset(0, -1).Resize(rXData.Rows.Count, 1).Formula = "1"
            Set rXData = rXData.CurrentRegion
        End If
    End With
    
End Sub

Private Sub CreateFiltData()
    ' Lays out the filtered data in the filtered-data sheet
    ' At present, fragile to a state where the filtering sheet is unbound, or improperly
    '  configured; or if the output book is uncreated, unopened, or unbound.
    
    Dim filtIter As Long, workRg As Range
    
    ' Scrub contents of filtered ranges, if they exist
    If Not rXFiltData Is Nothing Then rXFiltData.Clear
    If Not rYFiltData Is Nothing Then rYFiltData.Clear
    If Not rNameFiltData Is Nothing Then rNameFiltData.Clear
        
    With rFiltDataSheet
        ' Direct values-only copy of data from the data sheet.
        rXData.Copy
        Call .Cells(RegressAux.firstDataRow, RegressAux.dataXCol).PasteSpecial(xlPasteValues)
        rYData.Copy
        Call .Cells(RegressAux.firstDataRow, RegressAux.dataYCol).PasteSpecial(xlPasteValues)
        rNameData.Copy
        Call .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).PasteSpecial(xlPasteValuesAndNumberFormats)
        
        ' Link ranges
        Set rXFiltData = .Cells(RegressAux.firstDataRow, RegressAux.dataXCol).CurrentRegion
        Set rYFiltData = .Cells(RegressAux.firstDataRow, RegressAux.dataYCol).CurrentRegion
        Set rNameFiltData = .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).CurrentRegion
        
        ' Filtering of the data
        ' First, filter the predictors (arbitrary choice of handling order)
        If Me.countFilters(crfPredictor) > 0 Then
            ' Step BACKWARD, to ensure proper matching of predictor to filter ID
            For filtIter = Me.countFilters(crfPredictor) To 1 Step -1
                ' Start with the name range of the predictor to filter
                Set workRg = rNameFiltData.Cells(1, ConstOffset + Me.getFilterSeq(filtIter, crfPredictor))
                
                ' Combine with the corresponding X data column
                Set workRg = Union(workRg, rXFiltData.Cells(1, _
                            ConstOffset + Me.getFilterSeq(filtIter, crfPredictor)).EntireColumn)
                
                ' Delete the predictor values and label
                Call workRg.Delete(xlShiftToLeft)
            Next filtIter
        End If
        
        ' Then, filter the data points
        If Me.countFilters(crfDataPoint) > 0 Then
            ' Again, step BACKWARD, to delete the rows from the bottom up, to ensure
            '  proper matching of data point to filter ID
            For filtIter = Me.countFilters(crfDataPoint) To 1 Step -1
                ' Start with Y range...
                Set workRg = rYFiltData.Cells(Me.getFilterSeq(filtIter, crfDataPoint), 1)
                
                ' Combine with corresponding X data row
                Set workRg = Union(workRg, rXFiltData.Cells(Me.getFilterSeq(filtIter, crfDataPoint), _
                            1).EntireRow)
                
                ' Delete the row of data
                Call workRg.Delete(xlShiftUp)
            Next filtIter
        End If
        
        ' Re-link ranges
        Set rXFiltData = .Cells(RegressAux.firstDataRow, RegressAux.dataXCol).CurrentRegion
        Set rYFiltData = .Cells(RegressAux.firstDataRow, RegressAux.dataYCol).CurrentRegion
        Set rNameFiltData = .Cells(RegressAux.dataPredNameRow, RegressAux.dataXCol).CurrentRegion
    End With
    
End Sub

Private Sub MakeVMatrix()
    ' Convert to Boolean-ed confirmed-success Function???
    Dim addrStr As String, tempCell As Range
    
    ' FRAGILE to output book unbound, unopened, etc.
    
    ' If rVMat has been set, clear it
    If Not rVMat Is Nothing Then
        rVMat.Clear
    End If
    
    ' Attach the properly sized array of cells to the range
    Set rVMat = rVSheet.Cells(RegressAux.firstDataRow, RegressAux.vMatCol) _
            .Resize(rXFiltData.Rows.Count, rXFiltData.Rows.Count)
    
    ' Store address string
    addrStr = rXFiltData.Address(External:=True)
    
    ' Set the temp cell for the formula limit workaround
    Set tempCell = rVMat.Offset(-1, 0).Resize(1, 1)
    
    ' Apply the array formula for the V-matrix, referencing the filtered data
    tempCell.Formula = "=MMULT(MMULT(" & addrStr & ",MINVERSE(MMULT(TRANSPOSE(" & _
                addrStr & ")," & addrStr & "))),TRANSPOSE(" & addrStr & "))"
    rVMat.FormulaArray = tempCell.Formula
    tempCell.Clear
    
    'rVMat.FormulaArray = rVMat.Formula
    'rVMat.Worksheet.Calculate
    
End Sub

Private Sub MakeResMatrices()
    ' ASSUMES THAT makeVMatrix has already been successfully called. This may need to be
    '  trapped for.
    
    ' FRAGILE to output book unbound, etc.
    
    Dim addrStr As String, rg As Range
    
    ' If the various results ranges have been defined, clear their contents
    '
    ' Could just clear the whole worksheet, but will try this first. Avoids the need
    '  to re-apply labels, if nothing else.  Also will permit user comments.
    ' Could shift the data down a few rows to allow a 'Notes' range at top of
    '  sheet?
    ' !!!HERE!!! when new result fields added
    '  (NEED to clear beta-pendant value ranges; can leave name ranges alone(?))
    If Not rResYRg Is Nothing Then rResYRg.Clear
    If Not rResBetaRg Is Nothing Then rResBetaRg.Clear
    If Not rResBetaSERg Is Nothing Then rResBetaSERg.Clear
    If Not rResTBetaRg Is Nothing Then rResTBetaRg.Clear
    If Not rResYHatRg Is Nothing Then rResYHatRg.Clear
    If Not rResEHatRg Is Nothing Then rResEHatRg.Clear
    If Not rResVIIRg Is Nothing Then rResVIIRg.Clear
    If Not rResRRg Is Nothing Then rResRRg.Clear
    If Not rResTRg Is Nothing Then rResTRg.Clear
    If Not rResDRg Is Nothing Then rResDRg.Clear
    If Not rResPredNameRg Is Nothing Then rResPredNameRg.EntireColumn.Clear ' Wipe the whole thing
    If Not rResSYYRg Is Nothing Then rResSYYRg.Clear
    If Not rResRSSRg Is Nothing Then rResRSSRg.Clear
    If Not rResRegSSRg Is Nothing Then rResRegSSRg.Clear
    If Not rResR2Rg Is Nothing Then rResR2Rg.Clear
    If Not rResFStatRg Is Nothing Then rResFStatRg.Clear
    If Not rResDegTotRg Is Nothing Then rResDegTotRg.Clear
    If Not rResDegRegRg Is Nothing Then rResDegRegRg.Clear
    If Not rResDegResRg Is Nothing Then rResDegResRg.Clear
    If Not rResAICRg Is Nothing Then rResAICRg.Clear
    If Not rResCorrAICRg Is Nothing Then rResCorrAICRg.Clear
    
    ' Attach the properly sized cell arrays to the ranges
    ' !!!HERE!!! when new result fields added
    Set rResYRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultYCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResBetaRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultBetaCol).Resize(rXFiltData.Columns.Count, 1)
    Set rResBetaSERg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultBetaSECol) _
                        .Resize(rXFiltData.Columns.Count, 1)
    Set rResTBetaRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultTBetaCol) _
                        .Resize(rXFiltData.Columns.Count, 1)
    Set rResYHatRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultYHCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResEHatRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultEHCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResVIIRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultVIICol).Resize(rYFiltData.Rows.Count, 1)
    Set rResRRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultRCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResTRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultTCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResDRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultDCol).Resize(rYFiltData.Rows.Count, 1)
    Set rResPredNameRg = rResSheet.Cells(RegressAux.firstDataRow, RegressAux.resultPredNameCol) _
                            .Resize(rXFiltData.Columns.Count, 1)
    With rResBetaRg
        Set rResSYYRg = .Offset(.Rows.Count + RegressAux.resultSYYRowOffset - 1, 0).Resize(1, 1)
        Set rResRSSRg = .Offset(.Rows.Count + RegressAux.resultRSSRowOffset - 1, 0).Resize(1, 1)
        Set rResRegSSRg = .Offset(.Rows.Count + RegressAux.resultRegSSRowOffset - 1, 0).Resize(1, 1)
        Set rResR2Rg = .Offset(.Rows.Count + RegressAux.resultR2RowOffset - 1, 0).Resize(1, 1)
        Set rResFStatRg = .Offset(.Rows.Count + RegressAux.resultFStatRowOffset - 1, 0).Resize(1, 1)
        Set rResDegTotRg = .Offset(.Rows.Count + RegressAux.resultDegTotRowOffset - 1, 0).Resize(1, 1)
        Set rResDegRegRg = .Offset(.Rows.Count + RegressAux.resultDegRegRowOffset - 1, 0).Resize(1, 1)
        Set rResDegResRg = .Offset(.Rows.Count + RegressAux.resultDegResRowOffset - 1, 0).Resize(1, 1)
        Set rResAICRg = .Offset(.Rows.Count + RegressAux.resultAICRowOffset - 1, 0).Resize(1, 1)
        Set rResCorrAICRg = .Offset(.Rows.Count + RegressAux.resultCorrAICRowOffset - 1, 0).Resize(1, 1)
    End With
    With rResPredNameRg
        Set rResSYYNameRg = .Offset(.Rows.Count + RegressAux.resultSYYRowOffset - 1, 0).Resize(1, 1)
        Set rResRSSNameRg = .Offset(.Rows.Count + RegressAux.resultRSSRowOffset - 1, 0).Resize(1, 1)
        Set rResRegSSNameRg = .Offset(.Rows.Count + RegressAux.resultRegSSRowOffset - 1, 0).Resize(1, 1)
        Set rResR2NameRg = .Offset(.Rows.Count + RegressAux.resultR2RowOffset - 1, 0).Resize(1, 1)
        Set rResFStatNameRg = .Offset(.Rows.Count + RegressAux.resultFStatRowOffset - 1, 0).Resize(1, 1)
        Set rResDegTotNameRg = .Offset(.Rows.Count + RegressAux.resultDegTotRowOffset - 1, 0).Resize(1, 1)
        Set rResDegRegNameRg = .Offset(.Rows.Count + RegressAux.resultDegRegRowOffset - 1, 0).Resize(1, 1)
        Set rResDegResNameRg = .Offset(.Rows.Count + RegressAux.resultDegResRowOffset - 1, 0).Resize(1, 1)
        Set rResAICNameRg = .Offset(.Rows.Count + RegressAux.resultAICRowOffset - 1, 0).Resize(1, 1)
        Set rResCorrAICNameRg = .Offset(.Rows.Count + RegressAux.resultCorrAICRowOffset - 1, 0).Resize(1, 1)
    End With
    
    ' Additionally set the number format for the predictor name range; ended up being UNHELPFUL
    '  formulas should retain the text number format..?
    'DO NOT USE: rResPredNameRg.NumberFormat = "@"
    
    ' Store the address string for the filtered X data
    addrStr = rXFiltData.Address(External:=True)
    
    ' Apply the various formulas, drawing from the filtered X and Y data. Lots of temp ranges to
    '  avoid the stinkin' 256-character limit
    ' !!!HERE!!! when new result fields added
    Set rg = rResBetaSERg.Offset(-1, 0).Resize(1, 1)
    
    rResYRg.Formula = "=INDEX(" & rYFiltData.Address(External:=True) & "," & _
            "1+ROW()-ROW(" & rResYRg.Cells(1, 1).Address(ColumnAbsolute:=False) & _
            "),1)"
            
    rg.Formula = "=MMULT(MMULT(MINVERSE(MMULT(TRANSPOSE(" & addrStr & _
            ")," & addrStr & ")),TRANSPOSE(" & addrStr & "))," & _
            rResYRg.Address & ")"
    rResBetaRg.FormulaArray = rg.Formula
    
    
    rResYHatRg.FormulaArray = "=MMULT(" & rVMat.Address(External:=True) & "," & _
            rResYRg.Address(ColumnAbsolute:=False) & ")"
            
    rResEHatRg.FormulaArray = "=" & rResYRg.Address & "-" & rResYHatRg.Address
    
    rResVIIRg.Formula = "=INDEX(" & rVMat.Address(External:=True) & _
            ",1+ROW()-ROW(" & rResVIIRg.Cells(1, 1).Address(ColumnAbsolute:=False) & _
            "),1+ROW()-ROW(" & rResVIIRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "))"
    
'    rResRRg.Formula = "=INDEX(" & rResEHatRg.Address & ",1+ROW()-ROW(" & _
'            rResRRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)/(SQRT(SUMSQ(" & _
'            rResEHatRg.Address & ")/(ROWS(" & addrStr & ")-COLUMNS(" & addrStr & _
'            ")))*SQRT(1-INDEX(" & rResVIIRg.Address & _
'            ",1+ROW()-ROW(" & rResRRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "))))"
    rResRRg.Formula = "=INDEX(" & rResEHatRg.Address & ",1+ROW()-ROW(" & _
            rResRRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)/(SQRT(" & _
            rResRSSRg.Address & "/" & rResDegResRg.Address & ")*SQRT(1-" & _
            "MIN(1e-12,INDEX(" & _
            rResVIIRg.Address & ",1+ROW()-ROW(" & _
            rResRRg.Cells(1, 1).Address(ColumnAbsolute:=False) & ")))))"
    
    rResTRg.Formula = "=INDEX(" & rResRRg.Address & ",1+ROW()-ROW(" & _
            rResTRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)*SQRT((ROWS(" & _
            addrStr & ")-COLUMNS(" & addrStr & ")-1)/(ROWS(" & addrStr & ")-COLUMNS(" & _
            addrStr & ")-INDEX(" & rResRRg.Address & ",1+ROW()-ROW(" & _
            rResTRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)^2))"
    
    rResDRg.Formula = "=(1/COLUMNS(" & addrStr & "))*INDEX(" & rResRRg.Address & _
            ",1+ROW()-ROW(" & rResDRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)^2*(INDEX(" & _
            rResVIIRg.Address & ",1+ROW()-ROW(" & rResDRg.Cells(1, 1).Address(ColumnAbsolute:=False) & _
            "),1)/(1-INDEX(" & rResVIIRg.Address & ",1+ROW()-ROW(" & _
            rResDRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "),1)))"
    
    rResPredNameRg.Formula = "=INDEX(" & rNameFiltData.Address(External:=True) & ",1,1+ROW()-ROW(" & _
            rResPredNameRg.Cells(1, 1).Address(ColumnAbsolute:=False) & "))"
    
    If Me.includeConstant Then
        ' TSS is relative to mean y
        rResSYYRg.FormulaArray = "=SUMSQ(" & rResYRg.Address & "-AVERAGE(" & rResYRg.Address & "))"
    Else
        ' TSS is relative to zero
        rResSYYRg.Formula = "=SUMSQ(" & rResYRg.Address & ")"
    End If
    rResSYYNameRg.Formula = RegressAux.resSYYName
    
    rResRSSRg.Formula = "=SUMSQ(" & rResEHatRg.Address & ")"
    rResRSSNameRg.Formula = RegressAux.resRSSName
    
    rResRegSSRg.Formula = "=" & rResSYYRg.Address & "-" & rResRSSRg.Address
    rResRegSSNameRg.Formula = RegressAux.resRegSSName
    
'    rResR2Rg.FormulaArray = "=1-(SUMSQ(" & rResEHatRg.Address & "))/SUM((" & _
'            rResYRg.Address & "-AVERAGE(" & rResYRg.Address & "))^2)"
    rResR2Rg.Formula = "=" & rResRegSSRg.Address & "/" & rResSYYRg.Address
    rResR2NameRg.Formula = RegressAux.resR2Name
    
    rResFStatRg.Formula = "=(" & rResRegSSRg.Address & "/" & rResDegRegRg.Address & ")/(" & _
            rResRSSRg.Address & "/" & rResDegResRg.Address & ")"
    rResFStatNameRg.Formula = RegressAux.resFStatName
    
    rResAICRg.Formula = "=ROWS(" & addrStr & ")*LOG10(" & rResRSSRg.Address & "/ROWS(" & _
            addrStr & "))+2*COLUMNS(" & addrStr & ")"
    rResAICNameRg.Formula = RegressAux.resAICName
    
    rResCorrAICRg.Formula = "=" & rResAICRg.Address & "+(2*ROWS(" & rResBetaRg.Address & ")*(ROWS(" & _
            rResBetaRg.Address & ")+1))/(ROWS(" & addrStr & ")-ROWS(" & rResBetaRg.Address & ")-1)"
    rResCorrAICNameRg.Formula = RegressAux.resCorrAICName
    
    If Me.includeConstant Then
        rResDegTotRg.Formula = "=ROWS(" & rResYRg.Address & ")-1"
    Else
        rResDegTotRg.Formula = "=ROWS(" & rResYRg.Address & ")"
    End If
    rResDegTotNameRg.Formula = RegressAux.resDegTotName
    
    rResDegResRg.Formula = "=ROWS(" & rResYRg.Address & ")-ROWS(" & rResBetaRg.Address & ")"
    rResDegResNameRg.Formula = RegressAux.resDegResName
    
    rResDegRegRg.Formula = "=" & rResDegTotRg.Address & "-" & rResDegResRg.Address
    rResDegRegNameRg.Formula = RegressAux.resDegRegName
    
'    If Me.includeConstant Then
'        rResDegRegRg.Formula = "=ROWS(" & rResBetaRg.Address & ")-1"
'    Else
'        rResDegRegRg.Formula = "=ROWS(" & rResBetaRg.Address & ")"
'    End If
'    rResDegResRg.Formula = "=" & rResDegTotRg.Address & "-" & rResDegRegRg.Address
    
    
'    rg.Formula = "=SQRT(SUMSQ(" & rResEHatRg.Address & ")/(ROWS(" & addrStr & _
'            ")-COLUMNS(" & addrStr & "))*INDEX(MINVERSE(MMULT(TRANSPOSE(" & addrStr & ")," & _
'            addrStr & ")),1+ROW()-ROW(" & rResBetaSERg.Cells(1, 1).Address(ColumnAbsolute:=False) & _
'            "),1+ROW()-ROW(" & rResBetaSERg.Cells(1, 1).Address(ColumnAbsolute:=False) & ")))"
    rg.Formula = "=SQRT(" & rResRSSRg.Address & "/" & rResDegResRg.Address & _
            "*INDEX(MINVERSE(MMULT(TRANSPOSE(" & addrStr & ")," & _
            addrStr & ")),1+ROW()-ROW(" & rResBetaSERg.Cells(1, 1).Address(ColumnAbsolute:=False) & _
            "),1+ROW()-ROW(" & rResBetaSERg.Cells(1, 1).Address(ColumnAbsolute:=False) & ")))"
    rResBetaSERg.FormulaArray = rg.Formula
    
    rResTBetaRg.FormulaArray = "=" & rResBetaRg.Address & "/" & rResBetaSERg.Address
    
    ' Done with rg; clear it
    rg.Clear
    
End Sub

Private Function MakeNumSeqArray(startVal As Long, endVal As Long) As Variant
    Dim iter As Long, outArray As Variant
    
    If Not endVal > startVal Then
        MakeNumSeqArray = Array(1)
        Exit Function
    End If
    
    ReDim outArray(startVal To endVal)
    For iter = startVal To endVal
        outArray(iter) = iter
    Next iter
    
    MakeNumSeqArray = outArray
    
End Function

Private Function GetMaxAmpl(valRg As Range) As Double
    With Application.WorksheetFunction
        GetMaxAmpl = .Max(.Max(valRg), -.Min(valRg))
    End With
End Function

Private Function MakeSeriesArray(inRg As Range, normalized As Boolean) As Variant
    ' Converts the values of a range (NEEDS TO BE A VECTOR) into a comma-separated list,
    '  surrounded with braces, with prepended equals sign.  This allows the numerical
    '  data to be applied directly to Series data variables
    ' NORMALIZING if indicated
    
    Dim iter As Long, outArray As Variant, maxAmpl As Double
       
    ' inRg needs to be a vector.  If not, complain and crash out.
    '  Will probably want to robustify.
    If inRg.Rows.Count > 1 And inRg.Columns.Count > 1 Then
        MakeSeriesArray = Array(0)
        Exit Function
    End If
    
    maxAmpl = GetMaxAmpl(inRg)
    
    ' Pile in the values
    ReDim outArray(1 To inRg.Cells.Count)
    For iter = 1 To inRg.Cells.Count
        If inRg.Rows.Count > 1 Then
            If normalized Then
                outArray(iter) = inRg.Cells(iter, 1) / maxAmpl
            Else
                outArray(iter) = inRg.Cells(iter, 1)
            End If
        Else
            If normalized Then
                outArray(iter) = inRg.Cells(1, iter) / maxAmpl
            Else
                outArray(iter) = inRg.Cells(1, iter)
            End If
        End If
    Next iter
    
    ' Assign the return
    MakeSeriesArray = outArray
    
End Function

Private Sub UnloadResObjects()
    ' ## Must update if new objects are added!
    ' DOES NOT UNLOAD THE rBook OBJECT!!
    ' !!!HERE!!! when new result fields added
    Set rConfigSheet = Nothing
    Set rDataSheet = Nothing
    Set rFiltDataSheet = Nothing
    Set rVSheet = Nothing
    Set rResSheet = Nothing
    Set rStatSheet = Nothing
    Set rDynChart = Nothing
    Set rXData = Nothing
    Set rYData = Nothing
    Set rNameData = Nothing
    Set rXFiltData = Nothing
    Set rYFiltData = Nothing
    Set rNameFiltData = Nothing
    Set rVMat = Nothing
    Set rResYRg = Nothing
    Set rResBetaRg = Nothing
    Set rResBetaSERg = Nothing
    Set rResYHatRg = Nothing
    Set rResEHatRg = Nothing
    Set rResVIIRg = Nothing
    Set rResRRg = Nothing
    Set rResTRg = Nothing
    Set rResDRg = Nothing
    Set rResPredNameRg = Nothing
    Set rResSYYRg = Nothing
    Set rResSYYNameRg = Nothing
    Set rResRSSRg = Nothing
    Set rResRSSNameRg = Nothing
    Set rResRegSSRg = Nothing
    Set rResRegSSNameRg = Nothing
    Set rResR2Rg = Nothing
    Set rResR2NameRg = Nothing
    Set rResFStatRg = Nothing
    Set rResFStatNameRg = Nothing
    Set rResDegTotRg = Nothing
    Set rResDegTotNameRg = Nothing
    Set rResDegRegRg = Nothing
    Set rResDegRegNameRg = Nothing
    Set rResDegResRg = Nothing
    Set rResDegResNameRg = Nothing
    Set rResAICRg = Nothing
    Set rResAICNameRg = Nothing
    Set rResCorrAICRg = Nothing
    Set rResCorrAICNameRg = Nothing
    Set rStatRg = Nothing
End Sub

Private Sub UnlinkSource()
    Set srcBk = Nothing
    Set SrcXRng = Nothing
    Set SrcYRng = Nothing
    Set SrcNameRng = Nothing
End Sub


' ####################################################################
' ##############      PROPERTY SET/LET/GET        ####################
' ####################################################################


Property Get isSourceAttached() As Boolean
    isSourceAttached = sourceAttached
End Property

Property Get isRegCreated() As Boolean
    isRegCreated = regCreated
End Property

Property Get includeConstant() As Boolean
    includeConstant = SrcInclConst
End Property

Property Get SourceRange(rangeType As CRSrcType) As Range
    ' Must update if CRSrcType enum changes
    Select Case rangeType
    Case crsXData
        Set SourceRange = SrcXRng
    Case crsYData
        Set SourceRange = SrcYRng
    Case crsNameData
        Set SourceRange = SrcNameRng
    End Select
End Property

Property Get SourceAddress(rangeType As CRSrcType) As String
    ' Must update if CRSrcType enum changes
    Select Case rangeType
    Case crsXData
        SourceAddress = SrcXRng.Address(External:=True)
    Case crsYData
        SourceAddress = SrcYRng.Address(External:=True)
    Case crsNameData
        SourceAddress = SrcNameRng.Address(External:=True)
    End Select
End Property

Property Get SourceBook() As Workbook
    Set SourceBook = srcBk
End Property

Property Get SourceBookPath() As String
    SourceBookPath = srcBk.fullName
End Property

Property Get StatsAddress() As String
    StatsAddress = rStatRg.Address(True, True)
End Property

Property Get Name() As String
    Name = regName
End Property

Property Get Description() As String
    Description = regDesc
End Property

Property Get RegFileName() As String
    RegFileName = rBook.Name
End Property

Property Get RegFileFullName() As String
    RegFileFullName = rBook.fullName
End Property

Property Get RegFilePath() As String
    RegFilePath = rBook.path
End Property

Property Get isFiltered(idNum As Long, filterType As CRFilterType) As VbTriState
    ' Return values:
    '   vbUseDefault -- invalid id number
    '   vbTrue -- id number of the given type IS filtered
    '   vbFalse -- id number of the given type is NOT filtered
    Select Case filterType
    Case crfDataPoint
        isFiltered = IsPtFiltered(idNum)
    Case crfPredictor
        isFiltered = IsPredFiltered(idNum)
    End Select
End Property

Property Get numPoints(Optional justActivePoints As Boolean = False) As Long
    If justActivePoints Then
        numPoints = rXFiltData.Rows.Count
    Else
        numPoints = rXData.Rows.Count
    End If
End Property

Property Get numPredictors(Optional justActivePreds As Boolean = False) As Long
    ' Does NOT EVER include the constant, if included, as a predictor
    If justActivePreds Then
        If Me.includeConstant Then
            numPredictors = rXFiltData.Columns.Count - 1
        Else
            numPredictors = rXFiltData.Columns.Count
        End If
    Else
        If Me.includeConstant Then
            numPredictors = rXData.Columns.Count - 1
        Else
            numPredictors = rXData.Columns.Count
        End If
    End If
End Property

Property Get responseName() As String
    responseName = regRespName
End Property

Property Let responseName(newName As String)
    regRespName = newName
End Property

Property Get constantName() As String
    If Me.includeConstant Then
        constantName = RegressAux.constName
    Else
        constantName = "(constant excluded)"
    End If
End Property

Property Get predictorName(index As Long, Optional fromActivePreds As Boolean = True) As String
    
    If fromActivePreds Then
        If index < 1 Or index > Me.numPredictors(True) Then
            predictorName = invalidIndexText
        Else
            ' Value here; name cells are populated by formula - COLUMN orientation
            predictorName = rResPredNameRg.Cells(ConstOffset + index, 1).value
        End If
    Else
        If index < 1 Or index > Me.numPredictors(False) Then
            predictorName = invalidIndexText
        Else
            ' Value here; pull from source sheet - in ROW orientation
            predictorName = rNameData.Cells(1, ConstOffset + index).value
        End If
    End If
End Property

Property Get predictorBeta(index As Long) As Double
    If index < 1 Or index > Me.numPredictors Then
        predictorBeta = invalidIndexText
    Else
        predictorBeta = rResBetaRg.Cells(ConstOffset + index, 1).value
    End If
End Property

Property Get constantBeta() As Double
    ' Not unreasonable to just return zero if constant excluded!
    If Me.includeConstant Then
        constantBeta = rResBetaRg.Cells(1, 1)
    Else
        constantBeta = 0#
    End If
End Property

Property Get predictorBetaSE(index As Long) As Double
    If index < 1 Or index > Me.numPredictors Then
        predictorBetaSE = invalidIndexText
    Else
        predictorBetaSE = rResBetaSERg.Cells(ConstOffset + index, 1)
    End If
End Property

Property Get predictorIndex(Name As String, fromActivePreds As Boolean) As Long
    Dim iter As Long, stopFlag As Boolean
    
    iter = 1
    Do Until Me.predictorName(iter, fromActivePreds) = Name Or iter > Me.numPredictors
        iter = iter + 1
    Loop
    
    predictorIndex = iter
    
End Property

Property Get constantBetaSE() As Double
    ' Zero return if constant excluded; nonphysical, but enh.
    If Me.includeConstant Then
        constantBetaSE = rResBetaSERg.Cells(1, 1)
    Else
        constantBetaSE = 0#
    End If
End Property

Property Get DOF(dofType As degsOfFreedom) As Long
    Select Case dofType
    Case DOFTotal
        DOF = rResDegTotRg.value
    Case DOFRegression
        DOF = rResDegRegRg.value
    Case DOFResidual
        DOF = rResDegResRg.value
    End Select
End Property

Property Get scratchCell() As Range
    Set scratchCell = rConfigSheet.Cells(1, RegressAux.configScratchCol)
End Property

Property Get getLastChartedX() As CRChartVar
    getLastChartedX = lastChartX
End Property

Property Get getLastChartedY() As CRChartVar
    getLastChartedY = lastChartY
End Property

Property Get getLastChartedXPred() As Long
    getLastChartedXPred = lastChartXPred
End Property

Property Get getLastChartedYPred() As Long
    getLastChartedYPred = lastChartYPred
End Property

Property Get getLastChartedXNorm() As Boolean
    getLastChartedXNorm = lastChartXNorm
End Property

Property Get getLastChartedYNorm() As Boolean
    getLastChartedYNorm = lastChartYNorm
End Property

Property Get getLastChartedDoOutliers() As Boolean
    getLastChartedDoOutliers = lastChartDoOutliers
End Property

Property Get getLastChartedOutlierAlpha() As Double
    getLastChartedOutlierAlpha = lastChartOutlierAlpha
End Property

Property Get getLastChartedSize() As CRChartSize
    getLastChartedSize = lastChartSize
End Property

Property Get statTypeIndex(first As Boolean) As Long
    If first Then
        statTypeIndex = firstStatTypeIndex
    Else
        statTypeIndex = lastStatTypeIndex
    End If
End Property

Property Get chartSizeIndex(first As Boolean) As Long
    If first Then
        chartSizeIndex = firstChartSizeIndex
    Else
        chartSizeIndex = lastChartSizeIndex
    End If
End Property

Property Get chartSizeName(chartSize As CRChartSize) As String
    Select Case chartSize
    Case crcsFullPage
        chartSizeName = "Full Page"
    Case crcsFigure
        chartSizeName = "Inline Figure"
    Case crcsThumbnail
        chartSizeName = "Thumbnail"
    End Select
End Property


' ####################################################################
' ##############      PRIVATE FUNCTIONS/SUBS      ####################
' ####################################################################


Private Function ConstOffset() As Long
    If Me.includeConstant Then
        ConstOffset = 1
    Else
        ConstOffset = 0
    End If
End Function

Private Function IsPtFiltered(idNum As Long) As VbTriState
    Dim iter As Long
    
    IsPtFiltered = vbUseDefault ' Means invalid id value
    
    ' Sanity check
    If idNum < 1 Or idNum > rXData.Rows.Count Then Exit Function
    
    ' Check for empty filtering array
    If IsFiltPtsEmpty Then
        IsPtFiltered = vbFalse
        Exit Function
    End If
    
    ' Else, seek through filtering array for the indicated id
    For iter = LBound(filtPts) To UBound(filtPts)
        If filtPts(iter) = idNum Then
            IsPtFiltered = vbTrue
            Exit Function
        End If
    Next iter
    
    ' idNum not found
    IsPtFiltered = vbFalse
    
End Function

Private Function IsPredFiltered(idNum As Long) As VbTriState
    Dim iter As Long
    
    IsPredFiltered = vbUseDefault ' Means invalid id value
    
    ' Sanity check (have to check whether constant is included)
    If idNum < 1 Or idNum > Me.numPredictors Then Exit Function
    
    ' Check for empty filtering array
    If IsFiltPredsEmpty Then
        IsPredFiltered = vbFalse
        Exit Function
    End If
    
    ' Else, seek through filtering array for the indicated id
    For iter = LBound(filtPreds) To UBound(filtPreds)
        If filtPreds(iter) = idNum Then
            IsPredFiltered = vbTrue
            Exit Function
        End If
    Next iter
    
    ' idNum not found
    IsPredFiltered = vbFalse
    
End Function

Private Function AddPtFilt(idNum As Long) As Boolean
    ' False = failure
    ' True = success
    ' Check to ensure that not too many Pt filters are added is handled
    '  within the addFilter wrapper
    
    Dim iter As Long
    
    AddPtFilt = False
    
    ' If no filters
    If LBound(filtPts) = UBound(filtPts) And filtPts(LBound(filtPts)) = 0 Then
        filtPts(LBound(filtPts)) = idNum
    Else ' There are filters
        ReDim Preserve filtPts(LBound(filtPts) To UBound(filtPts) + 1)
        ' Scan from the back end, inserting the value in the appropriate spot.
        '  Presumes proper ordering of values; ##could add a proofing function
        For iter = UBound(filtPts) - 1 To LBound(filtPts) Step -1
            If idNum < filtPts(iter) Then
                filtPts(iter + 1) = filtPts(iter)
            Else ' found the insertion spot
                filtPts(iter + 1) = idNum
                Exit For
            End If
        Next iter
        
        ' Must check that value to add wasn't needing to go at the start of the list
        If iter = LBound(filtPts) - 1 Then filtPts(LBound(filtPts)) = idNum
            
    End If
    
    AddPtFilt = True
    
End Function

Private Function AddPredFilt(idNum As Long) As Boolean
    ' False = failure
    ' True = success
    
    Dim iter As Long
    
    AddPredFilt = False
    
    ' If no filters
    If LBound(filtPreds) = UBound(filtPreds) And filtPreds(LBound(filtPreds)) = 0 Then
        filtPreds(LBound(filtPreds)) = idNum
    Else ' There are filters
        ReDim Preserve filtPreds(LBound(filtPreds) To UBound(filtPreds) + 1)
        ' Scan from the back end, inserting the value in the appropriate spot.
        '  Presumes proper ordering of values; ##could add a proofing function
        For iter = UBound(filtPreds) - 1 To LBound(filtPreds) Step -1
            If idNum < filtPreds(iter) Then
                filtPreds(iter + 1) = filtPreds(iter)
            Else ' found the insertion spot
                filtPreds(iter + 1) = idNum
                Exit For
            End If
        Next iter
        
        ' Must check that value to add wasn't needing to go at the start of the list
        If iter = LBound(filtPreds) - 1 Then filtPreds(LBound(filtPreds)) = idNum
        
    End If
    
    AddPredFilt = True
    
End Function

Private Function CountPtFilts() As Long
    ' -1 Return means failure; should essentially never occur
    CountPtFilts = -1
    
    If IsFiltPtsEmpty Then
        CountPtFilts = 0
    Else
        CountPtFilts = UBound(filtPts) - LBound(filtPts) + 1
    End If
    
End Function

Private Function CountPredFilts() As Long
    ' -1 Return means failure; should essentially never occur
    CountPredFilts = -1
    
    If IsFiltPredsEmpty Then
        CountPredFilts = 0
    Else
        CountPredFilts = UBound(filtPreds) - LBound(filtPreds) + 1
    End If
    
End Function

Private Function DelPtFilt(idNum As Long) As Boolean
    ' False = failure; filter not found
    ' True = success
    '
    ' Assumes competent wrapper function, ensuring that the indicated filter
    '  is present and appropriate to delete
    
    Dim iter As Long
    
    DelPtFilt = False
    
    ' Do do sanity check for empty filter list or length-one filter list
    ' Empty filter list is handled by not having a code block in the below;
    '  the False return value just falls through.
    If UBound(filtPts) = LBound(filtPts) Then
        If filtPts(LBound(filtPts)) <> 0 Then ' Length-one filter list
            filtPts(LBound(filtPts)) = 0
            DelPtFilt = True
        End If
    Else
        ' Scan backward from end of array until id is located, then scan back up, swapping
        '  down filter IDs
        ' Downward scan
        For iter = UBound(filtPts) To LBound(filtPts) Step -1
            If idNum = filtPts(iter) Then Exit For
        Next iter
        
        ' Removing filter entry other than the last in the list
        If iter < UBound(filtPts) Then
            ' Copy down each filter ID one-by-one above the ID to be removed
            For iter = iter To UBound(filtPts) - 1
                filtPts(iter) = filtPts(iter + 1)
            Next iter
        End If
        
        ' Last element of the filter array now gets ReDim-ed away, regardless of whether
        '  it was originally in the middle or at the end of the list
        ReDim Preserve filtPts(LBound(filtPts) To UBound(filtPts) - 1)
        
        ' Success return
        DelPtFilt = True
    End If
    
End Function

Private Function DelPredFilt(idNum As Long) As Boolean
    ' False = failure; filter not found
    ' True = success
    '
    ' Assumes competent wrapper function, ensuring that the indicated filter
    '  is present and appropriate to delete
    
    Dim iter As Long
    
    DelPredFilt = False
    
    ' Do do sanity check for empty filter list or length-one filter list
    ' Empty filter list is handled by not having a code block in the below;
    '  the False return value just falls through.
    If UBound(filtPreds) = LBound(filtPreds) Then
        If filtPreds(LBound(filtPreds)) <> 0 Then ' Length-one filter list
            filtPreds(LBound(filtPreds)) = 0
            DelPredFilt = True
        End If
    Else
        ' Scan backward from end of array until id is located, then scan back up, swapping
        '  down filter IDs
        ' Downward scan
        For iter = UBound(filtPreds) To LBound(filtPreds) Step -1
            If idNum = filtPreds(iter) Then Exit For
        Next iter
        
        ' Removing filter entry other than the last in the list
        If iter < UBound(filtPreds) Then
            ' Copy down each filter ID one-by-one above the ID to be removed
            For iter = iter To UBound(filtPreds) - 1
                filtPreds(iter) = filtPreds(iter + 1)
            Next iter
        End If
        
        ' Last element of the filter array now gets ReDim-ed away, regardless of whether
        '  it was originally in the middle or at the end of the list
        ReDim Preserve filtPreds(LBound(filtPreds) To UBound(filtPreds) - 1)
        
        ' Success return
        DelPredFilt = True
    End If
    
End Function

Private Function IsFiltPtsEmpty() As Boolean
    If LBound(filtPts) = UBound(filtPts) And filtPts(LBound(filtPts)) = 0 Then
        IsFiltPtsEmpty = True
    End If
End Function

Private Function IsFiltPredsEmpty() As Boolean
    If LBound(filtPreds) = UBound(filtPreds) And filtPreds(LBound(filtPreds)) = 0 Then
        IsFiltPredsEmpty = True
    End If
End Function


' ####################################################################
' ##############          CLASS EVENTS            ####################
' ####################################################################

Private Sub Class_Initialize()
    sourceAttached = False
    regCreated = False
    'populateObjColl
    
    ' Initialize the filtering arrays
    filtPts = Array(0)
    filtPreds = Array(0)
    
    ' Initialize default charting behavior
    Me.setDefaultLastChartedVars
    
End Sub

Private Sub Class_Terminate()
    ' Probably don't need anything here that's not a custom object?
    Set srcBk = Nothing
    Set SrcXRng = Nothing
    Set SrcYRng = Nothing
    Set SrcNameRng = Nothing
End Sub

'Private Sub populateObjColl()
'    With ObjColl
'        ' Source objects
'        .Add SrcBk
'        .Add SrcXRng
'        .Add SrcYRng
'        .Add SrcNameRng
'
'        ' Results book and sheet objects
'        .Add rBook
'        .Add rConfigSheet
'        .Add rDataSheet
'        .Add rFiltDataSheet
'        .Add rVSheet
'        .Add rResSheet
'
'        ' Results book Range objects
'        .Add rXData
'        .Add rYData
'        .Add rNameData
'        .Add rXFiltData
'        .Add rYFiltData
'        .Add rNameFiltData
'        .Add rVMat
'        .Add rResYRg
'        .Add rResBetaRg
'        .Add rResBetaSERg
'        .Add rResYHatRg
'        .Add rResEHatRg
'        .Add rResVIIRg
'        .Add rResRRg
'        .Add rResTRg
'        .Add rResDRg
'        .Add rResPredNameRg
'        .Add rResR2Rg
'        .Add rResR2NameRg
'    End With
'End Sub
