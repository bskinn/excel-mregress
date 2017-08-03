Attribute VB_Name = "RegressAux"
' # ------------------------------------------------------------------------------
' # Name:        RegressAux.bas
' # Purpose:     Helper functions for "Multiple Regression Explorer" Excel VBA Add-In
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

' !!!HERE!!! when new result fields added
' Default locations for data in generated workbooks
Public Const lblRow As Long = 1, firstDataRow As Long = 5
Public Const dataPredNameRow As Long = 3
Public Const dataXCol As Long = 3, dataYCol As Long = 1
Public Const vMatCol As Long = 1
Public Const resultYCol As Long = 1, resultBetaCol As Long = 5
Public Const resultPredNameCol As Long = 3
Public Const resultTBetaCol As Long = 9
Public Const resultBetaSECol As Long = 7
Public Const resultYHCol As Long = 11, resultEHCol As Long = 13
Public Const resultVIICol As Long = 15, resultRCol As Long = 17
Public Const resultTCol As Long = 19, resultDCol As Long = 21
Public Const resultSYYRowOffset As Long = 2
Public Const resultRSSRowOffset As Long = 4
Public Const resultRegSSRowOffset As Long = 3
Public Const resultR2RowOffset As Long = 5 ' Row gap is this value minus one
Public Const resultFStatRowOffset As Long = 6
Public Const resultDegTotRowOffset As Long = 8
Public Const resultDegRegRowOffset As Long = 9
Public Const resultDegResRowOffset As Long = 10
Public Const resultAICRowOffset As Long = 11
Public Const resultCorrAICRowOffset As Long = 12
Public Const modelCritColOffset As Long = 0
Public Const modelR2ColOffset As Long = 1
Public Const modelParamColOffset As Long = 2
Public Const modelPValColOffset As Long = 3
Public Const modelWUNameColOffset As Long = 0
Public Const modelWUCritDataColOffset As Long = 1
Public Const modelWUCritMinColOffset As Long = 2
Public Const modelWUCritDiffColOffset As Long = 3
Public Const modelWUR2DataColOffset As Long = 4
Public Const modelWUR2RefColOffset As Long = 5
Public Const modelWUR2DiffColOffset As Long = 6
Public Const modelWUR2BasebarColOffset As Long = 7


'Public Const configIDCol As Long = 1
Public Const configValCol As Long = 2
Public Const configScratchCol As Long = 10
Public Const statStartRow As Long = 4
Public Const statStartCol As Long = 3

' !!!HERE!!! when new result fields added
'  (not needed for new single-cell fields tagged off of the betas)
' Default labels for the various data ranges and worksheets
Public Const dataShtName As String = "Data"
Public Const dataFiltShtName As String = "Filtered Data"
Public Const vMatShtName As String = "V-Matrix"
Public Const resShtName As String = "Results"
Public Const configShtName As String = "Config"
Public Const statShtName As String = "Stats"
Public Const chartShtName As String = "DynChart"
Public Const constName As String = "!Const!"
Public Const dataYName As String = "Responses", dataXName As String = "Predictors"
Public Const vMatName As String = "V-Matrix"
Public Const resYName As String = "y", resBetaName As String = "beta" ' &H3B2
Public Const resBetaSEName As String = "s.e. (beta)"
Public Const resTBetaName As String = "t(beta)"
Public Const resYHName As String = "y^" ' &H177
Public Const resEHName As String = "ê"
Public Const resVIIName As String = "diag(V)", resRName As String = "r"
Public Const resTName As String = "t", resDName As String = "D"
Public Const resSYYName As String = "Total SS"
Public Const resRSSName As String = "Resid SS"
Public Const resRegSSName As String = "Regr SS"
Public Const resR2Name As String = "R^2"
Public Const resFStatName As String = "F-Stat"
Public Const resDegTotName As String = "d.f. Tot"
Public Const resDegRegName As String = "d.f. Regr"
Public Const resDegResName As String = "d.f. Resid"
Public Const resAICName As String = "Akaike Crit"
Public Const resCorrAICName As String = "Corr. AIC"
Public Const modelDataRgCell As String = "$B$4"
Public Const modelWkupRgCell As String = "$H$4"
Public Const modelStatsRgCell As String = "$R$4"
Public Const modelBlueColor As Long = 12419407 'RGB(79, 129, 189)
Public Const modelGreenColor As Long = 5880731 ' RGB(155, 187, 89)
Public Const modelGridlineColor As Long = 12566463 ' RGB(191, 191, 191)
Public Const modelAxisColor As Long = 8355711 ' RGB(127, 127, 127)
Public Const modelWhiteColor As Long = 16777215 ' RGB(255,255,255)
Public Const modelLineWeight As Double = 2#
Public Const modelBarLineWeight As Double = 1.75
Public Const modelBarPattern As Long = 15 ' MsoPatternType = msoPatternDarkDownwardDiagonal
Public Const modelR2BarGapWidth As Double = 265
Public Const modelCritBarGapWidth As Double = 100

' Timestamp custom document property name; misc other stuff
Public Const timeStampDocPropName As String = "RegressTimeStamp"
Public Const lastSaveDocPropName As String = "Last Save Time"
Public Const timeStampFormat As String = "yyyy-mm-dd hh:mm:ss"
Public Const stampTimeWindowSecs As Double = 3#
Public Const rBookExtension As String = ".rgn.xlsx"
Public Const eBookExtension As String = ".xlsx"
Public Const regNameTempBlip As String = "_TMP_"
Public Const maxNameLength As Long = 60
Public Const defaultRoundSigFigs As Long = 3  ' Default value for the roundSigs() function
Public Const statAlpha As Double = 0.05 ' ???ALREADY AUTOMATED???  Doesn't seem to be in use


Public Sub bootRegress()
Attribute bootRegress.VB_Description = "Load Multiple Regression Tool"
Attribute bootRegress.VB_ProcData.VB_Invoke_Func = "R\n14"
    Load RegressMain
    RegressMain.Show
End Sub

Public Function validRegName(testName As String) As Boolean
    Dim RegEx As New RegExp, workStr As String
    
    validRegName = False
    
    workStr = testName
    
    ' Strip temp string to account for possible blip
    If InStr(workStr, RegressAux.regNameTempBlip) > 0 Then
        workStr = Left(workStr, Len(workStr) - Len(RegressAux.regNameTempBlip))
    End If
    
    If Len(workStr) > RegressAux.maxNameLength Then Exit Function
    
    With RegEx
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        '.Pattern = "^[a-zA-Z][0-9a-zA-Z_ ^-]*$"
        'validRegName = .Test(workStr)
        .Pattern = "[\\/:*?""<>|]"
        validRegName = Not .Test(workStr)
    End With
    
End Function

Public Function roundSigs(value As Double, Optional numSFs As Long = defaultRoundSigFigs) As Double
    ' Works with mantissa in the 0-1 interval
    Dim exponent As Long, mantissa As Double, negVal As Long
    
    If value = 0 Then
        roundSigs = 0#
        Exit Function
    End If
    
    If value < 0 Then
        negVal = -1
    Else
        negVal = 1
    End If
    
    With Application.WorksheetFunction
        exponent = .Ceiling(.Log10(Abs(value)), 1)
        mantissa = negVal * .Power(10, .Log10(Abs(value)) - exponent)
        roundSigs = .Round(mantissa, numSFs) * .Power(10, exponent)
    End With
End Function

Public Sub checkXData()
    Dim inRg As Range, col As Range, cel As Range, col2 As Range, testVal As Variant
    Dim row As Range, row2 As Range
    Dim reportStr As String, needToReport As Boolean
    Dim dataRg As Range, nameRg As Range
    Dim wsf As WorksheetFunction
    
    ' Initialize worksheet functions and reporting str
    reportStr = ""
    Set wsf = Application.WorksheetFunction
    
    ' Check if anything passed
    'If inRg Is Nothing Then
    Set inRg = Selection
    
    ' Split into header row and data block
    With inRg
        Set nameRg = .Rows(1)
        Set dataRg = .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count)
    End With
    
    ' Check for all-constant columns
    For Each col In dataRg.Columns
        ' Initialize
        testVal = Empty
        needToReport = True
        
        ' Check if all cells are equal
        For Each cel In col.Cells
            If IsEmpty(testVal) Then
                testVal = cel.value
            Else
                If cel.value <> testVal Then needToReport = False
            End If
        Next cel
        
        If needToReport Then
            reportStr = reportStr & "Values for '" & _
                    Intersect(nameRg, col.EntireColumn).value & _
                    "' are all equal." & vbCrLf
        End If
    Next col
    
    ' Check for linearly dependent columns
    For Each col In dataRg.Columns
        For Each col2 In dataRg.Columns
            ' Only care about distinct columns; only need upper-triangular
            If Intersect(col, col2) Is Nothing And col.Column < col2.Column Then
                With wsf
                    ' Store the dot product
                    testVal = .SumProduct(col, col2)
                    
                    ' Divide by the magnitudes
                    testVal = testVal / Sqr(.SumProduct(col, col)) / _
                                                Sqr(.SumProduct(col2, col2))
                    
                    ' If it's within a tolerance of unity, log
                    If Abs(testVal - 1) < 0.000001 Then
                        reportStr = reportStr & "Columns '" & _
                                Intersect(nameRg, col.EntireColumn).value & _
                                "' and '" & _
                                Intersect(nameRg, col2.EntireColumn).value & _
                                "' are linearly dependent." & vbLf
                    End If
                End With
            End If
        Next col2
    Next col
    
    ' Check for linearly dependent rows
    For Each row In dataRg.Rows
        For Each row2 In dataRg.Rows
            ' Only care about distinct columns; only need upper-triangular
            If Intersect(row, row2) Is Nothing And row.row < row2.row Then
                With wsf
                    ' Store the dot product
                    testVal = .SumProduct(row, row2)
                    
                    ' Divide by the magnitudes
                    testVal = testVal / Sqr(.SumProduct(row, row)) / _
                                                Sqr(.SumProduct(row2, row2))
                    
                    ' If it's within a tolerance of unity, log
                    If Abs(testVal - 1) < 0.000001 Then
                        reportStr = reportStr & "Worksheet rows '" & _
                                row.row & _
                                "' and '" & _
                                row2.row & _
                                "' are linearly dependent." & vbLf
                    End If
                End With
            End If
        Next row2
    Next row
    
    ' Report
    If reportStr <> "" Then
        MsgBox reportStr
    End If

End Sub

Public Function requestFilename(promptStr As String, defaultStr As String) As String
    Dim workStr As String

    ' Initialize to nothing
    workStr = ""
    
    ' Loop to a valid name
    Do
        If workStr <> "" Then
            MsgBox "Invalid file name.", vbOKOnly + vbExclamation, "Invalid name"
        End If
        workStr = InputBox(promptStr, "Enter file name", defaultStr)
    Loop Until RegressAux.validRegName(workStr) Or workStr = ""
    
    ' Return
    requestFilename = workStr
    
End Function

Public Function calcPStat(value As Double, valSE As Double, DOF As Long) As Double
    Dim wsf As WorksheetFunction
    
    Set wsf = Application.WorksheetFunction

    calcPStat = wsf.T_Dist_2T(Abs(value / valSE), DOF)
End Function

Public Function addFilterToString(filtStr As String, idx As Long) As String
    Dim rx As New RegExp, mchs As MatchCollection
    Dim workStr As String, iter As Long
    
    ' CALLING FUNCTION MUST DETERMINE IF ADDING A FILTER IS FEASIBLE!
    
    ' Ignore non-positive indices
    If idx < 1 Then
        addFilterToString = filtStr
        Exit Function
    End If
    
    ' Define Regex
    With rx
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "[(,]([0-9]+)"
        Set mchs = .Execute(filtStr)
    End With
    
    ' Loop through the elements, injecting the new index where appropriate.
    '  Will just not add the index if it's already present!
    workStr = "("
    For iter = 0 To mchs.Count - 1
        If iter > 0 Then
            If CLng(mchs(iter - 1).SubMatches(0)) < idx And _
                        CLng(mchs(iter).SubMatches(0)) > idx Then
                workStr = workStr & idx & "," & mchs(iter).SubMatches(0) & ","
            Else
                workStr = workStr & mchs(iter).SubMatches(0) & ","
            End If
        Else
            If CLng(mchs(iter).SubMatches(0)) > idx Then
                workStr = workStr & idx & "," & mchs(iter).SubMatches(0) & ","
            Else
                workStr = workStr & mchs(iter).SubMatches(0) & ","
            End If
        End If
        
        ' Must check for if new index is larger than all values
        If iter = mchs.Count - 1 And CLng(mchs(iter).SubMatches(0)) < idx Then
            workStr = workStr & idx & ","
        End If
        
    Next iter
    
    ' If the final length is one, then it was an empty list. Populate
    If Len(workStr) = 1 Then
        workStr = workStr & idx & ")"
    Else
        
        ' Replace the last comma with a paren
        workStr = Left(workStr, Len(workStr) - 1) & ")"
    End If
        
    ' Return the string
    addFilterToString = workStr
    
End Function

Public Function delFilterFromString(filtStr As String, idx As Long) As String
    Dim rx As New RegExp, mchs As MatchCollection
    Dim workStr As String, iter As Long
    
    ' CALLING FUNCTION MUST DETERMINE IF REMOVING A FILTER IS FEASIBLE!
    
    ' Ignore non-positive indices
    If idx < 1 Then
        delFilterFromString = filtStr
        Exit Function
    End If
    
    ' Define Regex
    With rx
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "[(,]([0-9]+)"
        Set mchs = .Execute(filtStr)
    End With
    
    ' Loop, skipping any index matching the target
    workStr = "("
    For iter = 0 To mchs.Count - 1
        If idx <> CLng(mchs(iter).SubMatches(0)) Then
            workStr = workStr & mchs(iter).SubMatches(0) & ","
        End If
    Next iter
    
    ' Swap last character for a paren
    ' Trap for if everything removed
    If Len(workStr) = 1 Then
        workStr = workStr & ")"
    Else
        workStr = Left(workStr, Len(workStr) - 1) & ")"
    End If
    
    ' Return the string
    delFilterFromString = workStr
    
End Function

