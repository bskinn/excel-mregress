VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressChart 
   Caption         =   "Select Plot Variables"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   OleObjectBlob   =   "RegressChart.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        RegressChart.frm
' # Purpose:     UserForm for generating charts with summary data from the active
' #               multiple regression analysis.
' #               Part of the "Multiple Regression Explorer" Excel VBA Add-In
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     25 Feb 2014
' # Copyright:   (c) Brian Skinn 2017
' # License:     The MIT License; see "LICENSE.txt" for full license terms.
' #
' #       http://www.github.com/bskinn/excel-mregress
' #
' # ------------------------------------------------------------------------------

Option Compare Text
Option Explicit

Private chartReg As ClsRegression

Private Sub BtnChart_Click()
    ' Should just be a call to the plotting function.  Chart automatically activated by reg.makeChart
    ' Don't really care whether it succeeded or not; any errors will have been critical, and
    '  reported during .makeChart execution
    Dim xType As CRChartVar, yType As CRChartVar, xPred As Long, yPred As Long
    Dim xNorm As Boolean, yNorm As Boolean
    Dim doOutliers As Boolean, outlierAlphaVal As Double
    
    xPred = 0
    yPred = 0
    
    If OBtXPtSeq.value Then xType = crcvPointSequence
    If OBtXCook.value Then xType = crcvCookDistance
    If OBtXResTStat.value Then xType = crcvTStatResidual
    If OBtXStRes.value Then xType = crcvStudentizedResidual
    If OBtXResid.value Then xType = crcvDirectResidual
    If OBtXFitResp.value Then xType = crcvFittedResponse
    If OBtXResp.value Then xType = crcvResponse
    If OBtXPred.value Then
        xType = crcvPredictor
        xPred = CBxXPredIndex.ListIndex + 1
    End If
    
    If OBtYPtSeq.value Then yType = crcvPointSequence
    If OBtYCook.value Then yType = crcvCookDistance
    If OBtYResTStat.value Then yType = crcvTStatResidual
    If OBtYStRes.value Then yType = crcvStudentizedResidual
    If OBtYResid.value Then yType = crcvDirectResidual
    If OBtYFitResp.value Then yType = crcvFittedResponse
    If OBtYResp.value Then yType = crcvResponse
    If OBtYPred.value Then
        yType = crcvPredictor
        yPred = CBxYPredIndex.ListIndex + 1
    End If
    
    xNorm = ChBxNormX.value
    yNorm = ChBxNormY.value
    
    doOutliers = ChBxOutliers.value
    outlierAlphaVal = CDbl(TBxOutlierAlpha.value) ' Proofing done in charting function
    
    If Not chartReg.makeChart(xType, yType, xNorm, yNorm, xPred, yPred, doOutliers, outlierAlphaVal, _
            CBxChartSize.ListIndex + 1) Then
        Exit Sub ' does nothing right now
    End If
    
    ' Enable the 'copy image' button
    BtnCopyImg.Enabled = True
    
End Sub

Private Sub BtnCopyImg_Click()
    ' Going to robustify this by disabling the button until a chart has been generated
    '  in the current instance of the form
    
    Dim cht As Chart
    
    Set cht = Workbooks(chartReg.RegFileName).Sheets(RegressAux.chartShtName)
    
    Call cht.CopyPicture(xlPrinter, xlPicture, xlPrinter)

'    ' Always check if the right thing is active
'    ' Is it a chart?
'    If TypeOf ActiveSheet Is Chart Then
'        ' Store the object for enhanced auto-code-stuff
'        Set actCht = ActiveSheet
'
'        ' ...
'        If actCht.Name = RegressAux.chartShtName Then
'            If actCht.Parent.fullName = chartReg.RegFileFullName Then
'                ' Probably okay
'                Call actCht.CopyPicture(xlScreen, xlPicture, xlPrinter)
'            End If
'        End If
'    End If
End Sub

Private Sub ChBxOutliers_Change()
    TBxOutlierAlpha.Enabled = ChBxOutliers.value
End Sub

Private Sub OBtXPred_Change()
    CBxXPredIndex.Enabled = OBtXPred.value
End Sub

Private Sub oBtYPred_change()
    CBxYPredIndex.Enabled = OBtYPred.value
End Sub

Private Sub TBxOutlierAlpha_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsNumeric(TBxOutlierAlpha.value) Then
        TBxOutlierAlpha.ForeColor = RGB(255, 0, 0)
        Cancel = True
    Else
        If CDbl(TBxOutlierAlpha.value) <= 0 Or CDbl(TBxOutlierAlpha.value) >= 1 Then
            TBxOutlierAlpha.ForeColor = RGB(255, 0, 0)
            Cancel = True
        Else
            TBxOutlierAlpha.ForeColor = &H80000008
            Cancel = False
        End If
    End If
End Sub

Private Sub UserForm_Activate()
    Dim iter As Long
    
    ' Check to ensure reg assigned
    If chartReg Is Nothing Then
        Call MsgBox("No Regression assigned for charting!", vbOKOnly + vbCritical, "Error")
        Me.Hide
        Exit Sub
    End If
    
    ' Populate the form
    ' First, the drop-boxes with the predictors
    For iter = 1 To chartReg.numPredictors(True)
        Call CBxXPredIndex.AddItem(chartReg.predictorName(iter, True))
        Call CBxYPredIndex.AddItem(chartReg.predictorName(iter, True))
    Next iter
    
    CBxXPredIndex.ListIndex = 0
    CBxYPredIndex.ListIndex = 0
    
    ' Now, select the prior option button settings; enable/disable of the combobox
    '  to choose the predictor should be handled automatically by the relevant Change events
    Select Case chartReg.getLastChartedX
    Case crcvCookDistance
        OBtXCook.value = True
    Case crcvDirectResidual
        OBtXResid.value = True
    Case crcvFittedResponse
        OBtXFitResp.value = True
    Case crcvPredictor
        OBtXPred.value = True
        CBxXPredIndex.ListIndex = chartReg.getLastChartedXPred - 1
    Case crcvResponse
        OBtXResp.value = True
    Case crcvStudentizedResidual
        OBtXStRes.value = True
    Case crcvTStatResidual
        OBtXResTStat.value = True
    Case crcvPointSequence
        OBtXPtSeq.value = True
    End Select
    ChBxNormX.value = chartReg.getLastChartedXNorm
    
    Select Case chartReg.getLastChartedY
    Case crcvCookDistance
        OBtYCook.value = True
    Case crcvDirectResidual
        OBtYResid.value = True
    Case crcvFittedResponse
        OBtYFitResp.value = True
    Case crcvPredictor
        OBtYPred.value = True
        CBxYPredIndex.ListIndex = chartReg.getLastChartedYPred - 1
    Case crcvResponse
        OBtYResp.value = True
    Case crcvStudentizedResidual
        OBtYStRes.value = True
    Case crcvTStatResidual
        OBtYResTStat.value = True
    Case crcvPointSequence
        OBtYPtSeq.value = True
    End Select
    ChBxNormY.value = chartReg.getLastChartedYNorm
    
    TBxRegName.value = chartReg.Name
    
    ChBxOutliers.value = chartReg.getLastChartedDoOutliers
    TBxOutlierAlpha.value = CStr(chartReg.getLastChartedOutlierAlpha)
    TBxOutlierAlpha.Enabled = ChBxOutliers.value
    
    With chartReg
        For iter = .chartSizeIndex(True) To .chartSizeIndex(False)
            ' This ties the index of the combobox to the Long-equivalent
            '  value of the underlying Enum.  When chart generation is called for,
            '  the index of the drop-down can then be used directly in chartReg.makeChart
            Call CBxChartSize.AddItem(.chartSizeName(iter), iter - 1)
        Next iter
        CBxChartSize.ListIndex = .getLastChartedSize - 1
    End With
    
    ' Should be it.
    
End Sub

Public Sub setChartReg(inReg As ClsRegression)
    Set chartReg = inReg
End Sub

Private Sub BtnClose_Click()
    ' Save the rBook
    chartReg.writeConfig
    
    ' Nothing to pass back to the main form; just exit
    Unload RegressChart
End Sub

