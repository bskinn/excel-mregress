VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressMain 
   Caption         =   "Main Menu"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   OleObjectBlob   =   "RegressMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        RegressMain.frm
' # Purpose:     Core UserForm for interaction with multiple regression analyses
' #               (ClsRegression/CollRegressions).
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

Private RegColl As New CollRegressions
Private checkingLoadedRegs As Boolean
Private quitBtnPressed As Boolean
Private priorRegSelected As Long
Private firstLoad As Boolean
Public lastPath As String

Public Function readyToQuit() As Boolean
    readyToQuit = quitBtnPressed
End Function

Public Sub RefreshSourceLinks()
    Dim reg As ClsRegression, srcBk As Workbook, tmpWksht As Worksheet
    Dim iter As Long
    
    ' Assume the source books are open for now, and just the Ranges need refreshing in case
    '  they have been disrupted by other operations
    For iter = 1 To RegColl.Count
        RegColl.Item(iter).bindSourceFromConfig
'        Set reg = RegColl.Item(iter)
'        Set srcBk = reg.SourceBook
'        Set tmpWksht = srcBk.Worksheets(1)
'
'        If Not reg.attachSource(srcBk, tmpWksht.Evaluate(reg.SourceAddress(crsXData)), _
'                tmpWksht.Evaluate(reg.SourceAddress(crsYData)), _
'                tmpWksht.Evaluate(reg.SourceAddress(crsNameData))) Then
'            Call MsgBox("Refresh of source objects for Regression """ & reg.Name & _
'                    """ was unsuccessful." & vbLf & vbLf & "Please exit and restart the application.", _
'                    vbOKOnly + vbExclamation, "Alert")
'        End If
    Next iter
    
End Sub

Public Function bookLinked(bookName As String) As String
    ' Checking to see if a given workbook is linked to an open Regression,
    '  either as the source book or as a destination book.
    
    Dim iter As Long
    
    ' Initalize to safe assumption, book is linked
    bookLinked = "<FAILSAFE -- ASSUMING BOOK IS LINKED>"
    
    ' Scan active regressions, dropping from function if found to be linked
    With RegColl
        If .Count > 0 Then  ' No worries if no regressions defined
            For iter = 1 To .Count
                If .Item(iter).SourceBook.Name = bookName Then
                    bookLinked = .Item(iter).Name
                    Exit Function
                End If
                'if .Item(iter).dataBook.Name = bookname then exit function  ' Not implemented yet!
            Next iter
        End If
    End With
    
    ' Made it here; presumably indicated book is not linked
    bookLinked = ""
    
End Function

Public Function rColl() As CollRegressions
    ' TEMPORARY FUNCTION FOR DEBUGGING ONLY!!
    Set rColl = RegColl
End Function

Private Sub BtnChart_Click()
    'Dummy button for now that just calls the charter
    ' If nothing selected in the listbox or no regs loaded, just silently do nothing
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    
    'Call RegColl.ItemKey(LBxOpenRegs.Value).makeChart(crcvFittedResponse, crcvDirectResidual, True, True)
    
    ' Disable buttons
    Call SetControlsStatus(False)
    
    ' Load, initialize and show the charting form
    Load RegressChart
    Call RegressChart.setChartReg(RegColl.ItemKey(LBxOpenRegs.value))
    
    priorRegSelected = LBxOpenRegs.ListIndex + 1
    
    RegressMain.Hide
    RegressChart.Show
    
    ' Should be nothing to capture back, and no need to repopulate anything?
    RegressMain.Show
    
End Sub

Private Sub BtnCopy_Click()
    ' Probably best to pop up the name form to allow custom indicating the
    '  name, save path, etc.
    Dim oldReg As ClsRegression, newReg As New ClsRegression
    Dim regNewName As String, regNewDesc As String, regNewPath As String
    Dim regNewRespName As String
    
    ' If nothing selected in the listbox or no regs loaded, just silently do nothing
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    
    ' Store prior selected Reg
    priorRegSelected = LBxOpenRegs.ListIndex + 1
    
    ' Disable buttons
    Call SetControlsStatus(False)
    
    ' Link reg
    Set oldReg = RegColl.ItemKey(LBxOpenRegs.value)
    
    ' Query for name and description
    Do
        ' Load, populate, and show name query form
        Load RegressNameEntry
        With RegressNameEntry
            Call .setNameDesc(oldReg.Name, oldReg.responseName, oldReg.Description, _
                        False, oldReg.RegFilePath)
            .Show
            
            ' Check whether canceled; drop from Sub if so
            If .closedByCancel Then
                Unload RegressNameEntry
                Exit Sub
            End If
            
            ' Store name, description, and target path
            regNewName = .enteredName
            regNewRespName = .enteredRespName
            regNewDesc = .enteredDescription
            regNewPath = .pickedFolderPath
            
        End With
        
        ' Either way, unload the name form
        Unload RegressNameEntry
        
        ' Check name -- one or both of name and path must differ, and name must not duplicate
        '  an already-open Reg ##LOOSE HERE, assuming compare string?
        If RegColl.RegExists(regNewName) Then
            Call MsgBox("The new Regression must not duplicate the name of any other open Regression.", _
                    vbOKOnly + vbCritical, "Error")
        End If
    Loop Until Not RegColl.RegExists(regNewName)
    
    ' Create the new regression
    If Not newReg.setName(regNewName, regNewPath, vbUseDefault) Then Exit Sub
    
    ' Assign response variable name
    newReg.responseName = regNewRespName
    
    ' Try to assign description
    If Not newReg.setDescription(regNewDesc) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Try to attach source.  If fails, must delete the new book created just now
    If Not newReg.attachSource(oldReg.SourceBook, oldReg.SourceRange(crsXData), _
            oldReg.SourceRange(crsYData), oldReg.SourceRange(crsNameData)) Then
        ' If deleting the new book fails, kill the app
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Create the regression content in the new book. If fails, must delete the new book
    If Not newReg.createNewRegression(oldReg.includeConstant) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Transfer the last-charted information
    With oldReg
        Call newReg.setLastChartedVars(.getLastChartedX, .getLastChartedY, .getLastChartedXPred, _
                .getLastChartedYPred, .getLastChartedXNorm, .getLastChartedYNorm, _
                .getLastChartedDoOutliers, .getLastChartedOutlierAlpha, .getLastChartedSize)
    End With
    
    ' Transfer the filters
    newReg.filterString(crfDataPoint) = oldReg.filterString(crfDataPoint)
    newReg.filterString(crfPredictor) = oldReg.filterString(crfPredictor)
    
    ' Regenerate the Reg content now that filters have been transferred
    If Not newReg.modifyRegression(False, newReg.includeConstant) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Add the new copied Reg to the collection; error-quit if fails
    If Not RegColl.Add(newReg) Then GoTo ErrorQuit
    
    ' Re-save the new book, just in case
    newReg.writeChanges
    
    ' Refresh the list of Regressions and the Reg info box
    popRegsList
    popRegInfo
    
    ' This operation should not change existing Regressions.
    '   No need to RefreshSourceLinks?
    
    Exit Sub
    
ErrorQuit:
    Call MsgBox("Error during Regression copy operation." & vbLf & vbLf & _
                "Exiting...", vbOKOnly + vbCritical, "Error")
    Unload RegressMain
    
End Sub

Private Sub BtnClose_Click()
    ' DO NOT query for delete of Reg file.  DO CHECK for if no other open Regs reference the source
    '  book of the Reg to be closed, and inquire if desire to close that book.
    
    Dim srcBook As Workbook, closeReg As ClsRegression, index As Long
    Dim otherSrcLinkFound As Boolean
    
    ' If nothing selected in the listbox or no regs loaded, just silently do nothing
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    
    ' Store the Reg for closing
    Set closeReg = RegColl.ItemKey(LBxOpenRegs.value)
    
    ' ### CHECK IF THIS REG IS SOURCE FOR ANOTHER
    
    ' Store sourcebook and check all other open Regs
    Set srcBook = closeReg.SourceBook
    otherSrcLinkFound = False
    For index = 1 To RegColl.Count
        If srcBook.fullName = RegColl.Item(index).SourceBook.fullName And _
                    Not closeReg.Name = RegColl.Item(index).Name Then
            otherSrcLinkFound = True
        End If
    Next index
    
    ' Query what to do if source is not linked to any other Regs
    If Not otherSrcLinkFound Then
        Select Case MsgBox("No other open Regressions are linked to the following source data book:" & _
                    vbLf & vbLf & srcBook.fullName & vbLf & vbLf & _
                    "Close this source book?" & vbLf & vbLf & _
                    "(Cancel aborts closing the Regression)", vbYesNoCancel + vbQuestion, "Close source book?")
        Case vbYes
            ' Check for if it's been changed; if so, ask about saving changes
            If Not srcBook.Saved Then
                Select Case MsgBox("Source data book has been changed.  Save before closing?" & _
                        vbLf & vbLf & "(Cancel aborts closing the Regression)", _
                        vbYesNoCancel + vbQuestion, "Save changes?")
                Case vbYes  ' do save
                    Call srcBook.Close(SaveChanges:=True)
                Case vbNo   ' don't save
                    Call srcBook.Close(SaveChanges:=False)
                Case vbCancel ' Cancel the Reg close
                    Exit Sub
                End Select
            Else
                ' Book already saved; just close
                Call srcBook.Close(SaveChanges:=True)
            End If
        Case vbNo
            ' Do not close the source book, just continue
        Case vbCancel
            ' Cancel the Reg close
            Exit Sub
        End Select
    End If
    
    ' Source dealt with; now remove the Reg from RegColl, save/close the data book,
    '  and repopulate the listbox
    Call RegColl.Remove(closeReg.Name)
    Call closeReg.closeRegression(saveBeforeClose:=True)
    Set closeReg = Nothing
    popRegsList
    popRegInfo
    RefreshSourceLinks
    
End Sub

Private Sub BtnDefineNew_Click()
    Dim regName As String, regDesc As String, regPath As String, regRespName As String
    
    ' Query for name and description
    Do
        ' Load, populate, and show name query form
        Load RegressNameEntry
        With RegressNameEntry
            Call .setNameDesc("[Enter Name]", "[Response]", "", True)
            .Show
            
            ' Check whether canceled; drop from Sub if so
            If .closedByCancel Then
                Unload RegressNameEntry
                Exit Sub
            End If
            
            ' Store name, response variable name, description, and target path
            regName = .enteredName
            regRespName = .enteredRespName
            regDesc = .enteredDescription
            regPath = .pickedFolderPath
        End With
        
        ' Either way, unload the name form
        Unload RegressNameEntry
        
        ' Check name
        If RegColl.RegExists(regName) Then
            Call MsgBox("A Regression with the name """ & regName & """ is already loaded!", _
                    vbOKOnly + vbCritical, "Error")
        End If
    Loop Until Not RegColl.RegExists(regName)
    
    ' Load regression definition form, initialize, and create new regression
    Load MultipleRegression.RegressSetup
    If Not RegressSetup.initNewReg(regName, regRespName, regDesc, regPath) Then Exit Sub
    
    Me.Hide
    RegressSetup.Show  ' Modal, so will wait until execution completes
    
    ' If Regression creation successful, try adding new regression to the collection
    With RegressSetup
        If Not .closedByCancel Then
            ' Write changes to disk
            If Not .getReg.writeChanges Then ' Something very weird happened
                Call MsgBox("Write of changes for Regression """ & .getReg.Name & """ failed!" & _
                        vbLf & vbLf & "Exiting...", vbOKOnly + vbCritical, "Error")
                Unload RegressSetup
                Unload RegressMain
            End If
            
            If RegColl.Add(.getReg) Then
                ' Refresh the list of Regressions
                popRegsList
                LBxOpenRegs.value = .getReg.Name
                popRegInfo
            Else
                ' Report error and delete the file on disk since it's a new Regression
                Call MsgBox("Addition of Regression with name """ & .getReg.Name & _
                        """ failed.", vbOKOnly + vbCritical, "Error")
                If Not .getReg.deleteRegressionFile(False) Then
                    ' Should never happen, with False passed
                End If
            End If
        Else
            ' User cancelled in the 'RegressSetup' phase, so the created workbook
            '  needs to be deleted
            If Not .getReg.deleteRegressionFile(False) Then
                ' Should never reach here, with False passed. Still, handle.
                Call MsgBox("Deletion of temporary regression file failed!" & _
                        vbLf & vbLf & "Exiting...", vbOKOnly + vbCritical, _
                        "Critical Error")
                Unload RegressSetup
                Unload RegressMain
                Exit Sub
            End If
            ' Regression object should die when RegressSetup is Unloaded
        End If

    End With
    
    ' Unload the regression setup form regardless of success or failure
    '  and re-show the main form
    Unload RegressSetup
    Me.Show
    
End Sub

Private Sub popRegsList()
    Dim iter As Long
    
    ' Clear the list
    LBxOpenRegs.Clear
    
    ' Cycle through the collection and populate with names (could have hidden columns
    '  with other things, if needed...)
    With RegColl
        If .Count > 0 Then
            For iter = 1 To .Count
                Call LBxOpenRegs.AddItem(.Item(iter).Name)
            Next iter
        End If
    End With
    
End Sub

Private Sub BtnDelete_Click()
    ' Deletes regression file from disk and eliminates the regression from the collection
    Dim delName As String
    
    ' Only do if a reg is selected
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    
    delName = LBxOpenRegs.value
    
    ' ##DO A CHECK FOR THIS REGRESSION BEING THE SOURCE FOR ANOTHER!
    
    ' Confirm delete
    If Not vbYes = MsgBox("Really delete Regression """ & delName & """?", vbQuestion + vbYesNo, _
            "Confirm Delete") Then Exit Sub
    
    ' Do the delete
    If Not RegColl.ItemKey(delName).deleteRegressionFile(False) Then
        Call MsgBox("Error while deleting Regression file!" & vbLf & vbLf & "Exiting...", _
                vbOKOnly + vbCritical, "Critical Error")
        Unload RegressMain
        Exit Sub
    End If
    
    ' Remove the entry from the listing and refresh
    Call RegColl.CullEmpty
    popRegsList
    popRegInfo
    
End Sub

Private Sub BtnEdit_Click()
    ' Call to name/description/folder form to see if changes desired there; then
    '  pass on to setup form to possibly change source.
    '
    ' If name or save folder changed, must keep track of old rBook and delete after
    '  changes applied -- or, recover if user cancels.
    '
    ' ## For now, regenerate Regression after ALL edits. Simple rename
    '     not implemented
    
    Dim oldReg As ClsRegression, newReg As New ClsRegression
    Dim regOldName As String, regOldDesc As String, regOldPath As String, regOldRespName As String
    Dim regNewName As String, regNewDesc As String, regNewPath As String, regNewRespName As String
    
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    ' Link reg
    Set oldReg = RegColl.ItemKey(LBxOpenRegs.value)
    
    ' ### ADD CHECK AND ALERT if this reg is the source to another open Reg
    
    ' Retain old name, description, locationsettings
    regOldName = oldReg.Name
    regOldRespName = oldReg.responseName
    regOldDesc = oldReg.Description
    regOldPath = oldReg.RegFilePath
    
    ' Query for name and description
    Do
        ' Load, populate, and show name query form
        Load RegressNameEntry
        With RegressNameEntry
            Call .setNameDesc(oldReg.Name, oldReg.responseName, oldReg.Description, _
                        False, oldReg.RegFilePath)
            .Show
            
            ' Check whether canceled; drop from Sub if so
            If .closedByCancel Then
                Unload RegressNameEntry
                Exit Sub
            End If
            
            ' Store name, description, and target path
            regNewName = .enteredName
            regNewRespName = .enteredRespName
            regNewDesc = .enteredDescription
            regNewPath = .pickedFolderPath
        End With
        
        ' Either way, unload the name form
        Unload RegressNameEntry
        
        ' Check name -- can be the same as the old name, but not that of any other Reg
        If RegColl.RegExists(regNewName) And regNewName <> regOldName Then
            Call MsgBox("A Regression with the name """ & regNewName & """ is already loaded!", _
                    vbOKOnly + vbCritical, "Error")
        End If
    Loop Until Not (RegColl.RegExists(regNewName) And regNewName <> regOldName)
    
    ' Load regression definition form, initialize, and create new regression with values
    '  populated from the prior, and pass the new reg into the form
    ' THIS WILL HAVE A PROBLEM IF NAME AND PATH ARE IDENTICAL.
    ' Trap for this, here and in Setup form
    If regNewName = regOldName And regNewPath = regOldPath Then
        regNewName = regNewName & RegressAux.regNameTempBlip
    End If
    
    ' Create the new regression
    If Not newReg.setName(regNewName, regNewPath, vbUseDefault) Then Exit Sub
    
    ' Try to assign description
    If Not newReg.setDescription(regNewDesc) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Assign response name
    newReg.responseName = regNewRespName
    
    ' Try to attach source.  If fails, must delete the new book created just now
    If Not newReg.attachSource(oldReg.SourceBook, oldReg.SourceRange(crsXData), _
            oldReg.SourceRange(crsYData), oldReg.SourceRange(crsNameData)) Then
        ' If deleting the new book fails, kill the app
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' Transfer the last-charted information
    With oldReg
        Call newReg.setLastChartedVars(.getLastChartedX, .getLastChartedY, .getLastChartedXPred, _
                .getLastChartedYPred, .getLastChartedXNorm, .getLastChartedYNorm, _
                .getLastChartedDoOutliers, .getLastChartedOutlierAlpha, .getLastChartedSize)
    End With
    
    ' Create the regression content in the new book. If fails, must delete the new book
    If Not newReg.createNewRegression(oldReg.includeConstant, _
                oldReg.filterString(crfDataPoint), oldReg.filterString(crfPredictor)) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    ' NO, DO NOT TRANSFER FILTERS! Can result in inconsistent state and/or unexpected
    '  behavior if number of predictors is changed, especially if reduced.
'    ' Transfer the filters (must fall after content is created in the new book)
'    newReg.filterString(crfDataPoint) = oldReg.filterString(crfDataPoint)
'    newReg.filterString(crfPredictor) = oldReg.filterString(crfPredictor)
    
    ' Try to configure the setup form.  If fails, must delete the new book
    If Not RegressSetup.initLoadReg(newReg) Then
        If Not newReg.deleteRegressionFile(False) Then GoTo ErrorQuit
        Exit Sub
    End If
    
    Me.Hide
    RegressSetup.Show  ' Modal, so will wait until execution completes
    
    ' If Regression creation successful, try adding new regression to the collection
    '  newReg and RegressSetup.getReg now refer to the same object, so dereferencing newReg
    '  here to avoid confusion
    Set newReg = Nothing
    With RegressSetup
        If Not .closedByCancel Then ' Committed to applying the edit at this point; any failures
                                    '  will result in an error-quit
            ' Pull the old Reg from the collection first
            If Not RegColl.Remove(oldReg.Name) Then GoTo ErrorQuit
'                    Call MsgBox("Error during Regression edit operation." & vbLf & vbLf & _
'                            "Exiting...", vbOKOnly + vbCritical, "Error")
'                    Unload RegressSetup
'                    Unload RegressMain
'                End If

            ' Need to cull the file from the old one before obliterating
            If Not oldReg.deleteRegressionFile(False) Then GoTo ErrorQuit
'                    Call MsgBox("Error during Regression edit operation." & vbLf & vbLf & _
'                            "Exiting...", vbOKOnly + vbCritical, "Error")
'                    Unload RegressSetup
'                    Unload RegressMain
'                End If
            
            ' Dereference the old Reg
            Set oldReg = Nothing

            ' Write changes to disk for the new Reg
            If Not .getReg.writeChanges Then ' Something very weird happened
                Call MsgBox("Write of changes for Regression """ & .getReg.Name & """ failed!" & _
                        vbLf & vbLf & "Exiting...", vbOKOnly + vbCritical, "Error")
                Unload RegressSetup
                Unload RegressMain
            End If
            
            ' If temp blip is in Name, rename it back to w/o blip
            With RegressSetup
                If InStr(.getReg.Name, RegressAux.regNameTempBlip) > 0 Then
                    ' Must rename; committed to edit at this point, so
                    If Not .getReg.setName(Left(.getReg.Name, Len(.getReg.Name) - Len(RegressAux.regNameTempBlip)), _
                            "", vbTrue) Then GoTo ErrorQuit
'                            call msgbox("Error during
                End If
            End With
            
            ' Add the new Reg into the collection
            If Not RegColl.Add(.getReg) Then GoTo ErrorQuit
'                    Call MsgBox("Error during Regression edit operation." & vbLf & vbLf & _
'                            "Exiting...", vbOKOnly + vbCritical, "Error")
'                    Unload RegressSetup
'                    Unload RegressMain
'                End If
                
            ' Dereference the old Reg
            Set oldReg = Nothing
            
            ' Re-save the new book, just in case
            RegressSetup.getReg.writeChanges
            
            ' Refresh the list of Regressions and the reg info, and refresh the source links
            popRegsList
            popRegInfo
            RefreshSourceLinks
'                Else
'                    ' Report error and delete the file on disk since it's a new Regression
'                    Call MsgBox("Addition of Regression with name """ & .getReg.Name & _
'                            """ failed.", vbOKOnly + vbCritical, "Error")
'                    If Not .getReg.deleteRegressionFile(False) Then
'                        ' Should never happen, with False passed
'
'                End If
        Else
            ' User cancelled in the 'RegressSetup' phase, so the created workbook
            '  needs to be deleted
            If Not .getReg.deleteRegressionFile(False) Then
                ' Should never reach here, with False passed. Still, handle.
                Call MsgBox("Deletion of temporary regression file failed!" & _
                        vbLf & vbLf & "Exiting...", vbOKOnly + vbCritical, _
                        "Critical Error")
                Unload RegressSetup
                Unload RegressMain
                Exit Sub
            End If
            ' Regression object should die when RegressSetup is Unloaded
        End If

    End With
    
    ' Unload the regression setup form regardless of success or failure
    '  and re-show the main form
    Unload RegressSetup
    Me.Show
    
    Exit Sub ' To avoid trespassing into ErrorQuit block
    
ErrorQuit:
    Call MsgBox("Error during Regression edit operation." & vbLf & vbLf & _
                "Exiting...", vbOKOnly + vbCritical, "Error")
    Unload RegressSetup
    Unload RegressMain

End Sub

Private Sub BtnFilterPreds_Click()
    ' Only do anything if something selected
    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
    
    ' Store selected Reg
    priorRegSelected = LBxOpenRegs.ListIndex + 1
    
    ' Disable controls
    Call SetControlsStatus(False)
    
    Load RegressFilterPred
    
    Call RegressFilterPred.populateForm(RegColl.ItemKey(LBxOpenRegs.value))
    
    If Not RegressFilterPred.isInitialized Then
        Call MsgBox("Unable to populate predictor filtering form!", vbOKOnly + vbCritical, _
                    "Error")
        Call SetControlsStatus(True) ' Re-enable before exit
        Exit Sub
    End If
    
    ' Re-enable before showing filter form
    Call SetControlsStatus(True)
    RegressFilterPred.Show
    
    ' Handle post-stuff -- if canceled, just unload the form
    If RegressFilterPred.exitByCancel Then
        Unload RegressFilterPred
    Else
        ' Disable controls
        Call SetControlsStatus(False)
        
        With RegColl.ItemKey(LBxOpenRegs.value)
            ' Pull the new filters and push into the regression
            .filterString(crfPredictor) = RegressFilterPred.getNewFilteredPreds
            
            ' Unload the filter form
            Unload RegressFilterPred
            
            ' Update to reflect the new filters; inconsistent state if failed; exit
            If Not .modifyRegression(True, .includeConstant) Then
                Call MsgBox("Regeneration of Regression with new filter set failed!  Exiting...", _
                        vbOKOnly + vbCritical, "Critical Error")
                Unload RegressMain
                Exit Sub
            End If
        End With
        
        ' Repopulate the info box
        popRegInfo
        
        ' Always refresh source links here
        RefreshSourceLinks
        
        ' ### NOTIFY USER if this Reg is the source for another Reg
        
        ' Re-enable controls
        Call SetControlsStatus(True)
    End If
    
End Sub

Private Sub BtnModSelAIC_Click()
    runModelSelection ST_AIC, "AIC"
End Sub

Private Sub BtnModSelAICc_Click()
    runModelSelection ST_CorrAIC, "AICc"
End Sub

Private Sub BtnOpen_Click()
    ' Query for file
    ' Check to ensure not the same target as an rBook of an already open Reg
    ' Check to ensure name is not identical to an already open Reg,
    '  with a rBook saved in a different location.
    ' Look to see if rBook is open but Reg is not created; if so, query
    '  whether desire to pull rBook in or create new.
    '  SHOULD NOT HAPPEN NOW, if _Activate() of main form is tied into checking
    '  to ensure that no .rgn.xlsx files are open that don't match open Regs,
    '  and the code is working as intended...
    ' Create a reg, pass it the book; reg should be able to recreate itself
    '  using the RegressAux constants and the information in the Config sheet.
    '
    ' Proofing will be required to ensure that the passed book is actually
    '  a Regression results book, and that the info hasn't been mangled since last saved
    
    Dim openBookFullName As String, fd As FileDialog, fs As FileSystemObject
    Dim index As Long
    Dim wb As Workbook
    Dim reg As ClsRegression
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    With fd
        ' Configure
        .AllowMultiSelect = False
        .ButtonName = "Open"
        .Filters.Clear
        Call .Filters.Add("Regression Excel files", "*" & RegressAux.rBookExtension)
        Call .Filters.Add("Excel files", "*.xlsx; *.xlsm; *.xls")
        If lastPath = "" Then
            .InitialFileName = "%homepath%\Documents"
        Else
            .InitialFileName = lastPath
        End If
        .Title = "Choose Regression for Import"
        
        ' Execute query for file
        If Not .Show = -1 Then Exit Sub ' b/c assume user cancelled
        openBookFullName = .SelectedItems(1)
        lastPath = fs.GetParentFolderName(.SelectedItems(1))
    End With
    
    ' Check for rBook or name collision with any open Regs
    With RegColl
        If .Count > 0 Then
            For index = 1 To .Count
                ' rBook collision; just a notification, not an error
                If openBookFullName = .Item(index).RegFileFullName Then
                    Call MsgBox("Selected Regression is already open!", vbExclamation + vbOKOnly, "Warning")
                    Exit Sub
                End If
                ' Name collision
                If fs.GetFile(openBookFullName).Name = .Item(index).Name & RegressAux.rBookExtension Then
                    Call MsgBox("A Regression with the name """ & .Item(index).Name & """ is already open!" _
                            & vbLf & vbLf & "Rename and retry.", _
                            vbOKOnly + vbCritical, "Error")
                    Exit Sub
                End If
            Next index
        End If
    End With
    
'    ' Check for rBook having been opened after app execution but somehow having been
'    '  missed by the trap in RegressMain_Activate(). Needs to fall AFTER the check of
'    '  already-opened Regs
'    If Application.Workbooks.Count > 0 Then
'        For Each wb In Application.Workbooks
'            If wb.fullName = openBookFullName Then
'                ''''select case msgbox("
'                ' ##Ask if want to draw this Reg into the thingy##
'                '   Is basically just 'opening,' but with the xlsx already having been opened.
'                ' ##CONSIDER## adding auto-detect feature during form load so that it automatically
'                '  pulls in any already-open rBook books.  Would make sense to
'                '  NOT necessarily force-close all rBooks on quitting from app.
'            End If
'        Next wb
'    End If
    
    ' Create new Regression and load the book into it
    Set reg = New ClsRegression
    Set wb = Workbooks.Open(openBookFullName)
    If Not reg.loadRegression(wb) Then
        Call MsgBox("Load of indicated Regression file failed.", vbOKOnly + vbCritical, "Error")
        Exit Sub
    End If
    
    ' Add the reg to the collection; presume any failure in .Add is already reported there. Nothing
    '  really to do otherwise if the add failed..?
    Call RegColl.Add(reg)
    
    ' Repopulate the regs list
    popRegsList
    popRegInfo
    ' SHould be no need to RefreshSourceLinks
    
End Sub

Private Sub BtnQuit_Click()
    ' Probably robustify with 'confirm' dialog tied into whether
    '  things are saved or not.
    ' Also, definitely close any referenced rBooks, or at least query whether to.
    ' ### Definitely need to add things here ###
    
    Unload RegressMain
End Sub

Private Sub BtnShowDesc_Click()
    If LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0 Then
        With RegColl.ItemKey(LBxOpenRegs.value)
            Call MsgBox(.Description, vbOKOnly, .Name)
        End With
    End If
End Sub

Private Function getUnlinkedRegBook() As Workbook
    ' Scan all open workbooks; ignore those not ending with right extension
    '  If any is open, confirm whether linked to an open Reg
    Dim wb As Workbook, index As Long, RegFound As Boolean
    
    ' Init return to Nothing
    Set getUnlinkedRegBook = Nothing
    
    ' No need if no workbooks open
    If Workbooks.Count > 0 Then
        For Each wb In Workbooks
            ' Reset the flag
            RegFound = False
            ' Check each book for the rBook extension
            If Right(wb.fullName, Len(RegressAux.rBookExtension)) = RegressAux.rBookExtension Then
                With RegColl
                    ' If no Regs open, book definitely needs flagging for attachment
                    If .Count > 0 Then  ' Have to check the linked Regs
                        For index = 1 To .Count
                            ' If one of the opened Regs matches the book, don't need to link it
                            If .Item(index).RegFileFullName = wb.fullName Then RegFound = True
                        Next index
                        ' If no opened Regs match the book, return as unlinked
                        If Not RegFound Then
                            Set getUnlinkedRegBook = wb
                            Exit Function
                        End If
                    Else  ' No Regs open, definitely return as unlinked
                        Set getUnlinkedRegBook = wb
                        Exit Function
                    End If
                End With
            End If
        Next wb
    End If
End Function

Private Sub popRegInfo()
    Dim textStr As String, reg As ClsRegression
    Dim indStr As String, wf As WorksheetFunction
    Dim iter As Long
    Dim tempBeta As Double, tempSE As Double
    Const lineChrs As Long = 55
    
    ' Couple of spaces to lead each line
    indStr = " "
    
    ' Bind worksheet function object
    Set wf = Application.WorksheetFunction
    
    ' Something selected - do the population
    ' ##Later, clean up the box appearance by coercing to fixed width.  Will require
    '  determining the max widths needed for all fields, then padding the added
    '  strings accordingly.  fprint would be nice...
    If LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0 Then
        Set reg = rColl.ItemKey(LBxOpenRegs.value)
        ' Response variable name
        textStr = indStr & "Response Variable: " & reg.responseName & vbLf
        
        ' Number of predictors and data points
        textStr = textStr & indStr & "Original Dataset: " & reg.numPredictors & _
                    " Predictors, " & reg.numPoints & " Data Points" & vbLf & vbLf
        
        ' Include/exclude constant
        If reg.includeConstant Then
            textStr = textStr & indStr & "Constant INCLUDED" & vbLf
        Else
            textStr = textStr & indStr & "Constant EXCLUDED" & vbLf
        End If
        
        ' Filtering status
        textStr = textStr & indStr & "Filtered Predictors: " & reg.filteredPredictors & vbLf
        textStr = textStr & indStr & "Filtered Data Points: " & reg.filterString(crfDataPoint) & vbLf & vbLf
        
        ' Fit parameters
        'textStr = textStr & indStr & "Predictor: Value  ( s.e. | t-stat || p-value )" & vbLf
        textStr = textStr & indStr & betaString("Predictor", "Value", "s.e.", "t-stat", "p-value") & vbLf
        textStr = textStr & indStr & wf.Rept("-", lineChrs) & vbLf
        If reg.includeConstant Then
            tempBeta = reg.constantBeta
            tempSE = reg.constantBetaSE
            textStr = textStr & indStr & betaString("Constant", roundSigs(tempBeta), _
                    roundSigs(tempSE), roundSigs(tempBeta / tempSE), _
                    roundSigs(wf.T_Dist_2T(Abs(tempBeta / tempSE), reg.DOF(DOFResidual)), 2)) & vbLf
'            textStr = textStr & indStr & "Constant : " & roundSigs(tempBeta) & "  ( " & _
'                    roundSigs(tempSE) & " | " & roundSigs(tempBeta / tempSE) & " || " & _
'                    roundSigs(wf.T_Dist_2T(Abs(tempBeta / tempSE), reg.DOF(DOFResidual)), 2) & _
'                    " )" & vbLf
        End If
        For iter = 1 To reg.numPredictors(True) ' Has to be at least one; want the number left after filtering
            tempBeta = reg.predictorBeta(iter)
            tempSE = reg.predictorBetaSE(iter)
            textStr = textStr & indStr & betaString(reg.predictorName(iter, True), _
                        roundSigs(tempBeta), roundSigs(tempSE), roundSigs(tempBeta / tempSE), _
                        roundSigs(wf.T_Dist_2T(Abs(tempBeta / tempSE), reg.DOF(DOFResidual)), 2)) & vbLf
'            textStr = textStr & indStr & reg.predictorName(iter, True) & ": " & _
'                    roundSigs(tempBeta) & "  ( " & roundSigs(tempSE) & " | " & _
'                    roundSigs(tempBeta / tempSE) & " || " & _
'                    roundSigs(wf.T_Dist_2T(Abs(tempBeta / tempSE), reg.DOF(DOFResidual)), 2) & _
'                    " )" & vbLf
        Next iter
        textStr = textStr & vbLf
        
        ' Fit statistics
        textStr = textStr & indStr & "Fit Statistics" & vbLf
        textStr = textStr & indStr & wf.Rept("=", lineChrs) & vbLf
        ' Sum-squares
        textStr = textStr & indStr & colSpreadThree( _
                reg.regStatName(ST_SSTotal) & ": " & _
                    RegressAux.roundSigs(reg.regStat(ST_SSTotal), 4), _
                reg.regStatName(ST_SSResidual) & ": " & _
                    RegressAux.roundSigs(reg.regStat(ST_SSResidual), 4), _
                reg.regStatName(ST_SSRegression) & ": " & _
                    RegressAux.roundSigs(reg.regStat(ST_SSRegression), 4)) & _
                vbLf
        ' Degrees of freedom
        textStr = textStr & indStr & colSpreadThree( _
                            "d.f.  " & reg.DOF(DOFTotal), _
                            reg.DOF(DOFResidual), _
                            reg.DOF(DOFRegression), , , "RRR") & vbLf
        ' R^2
        textStr = textStr & indStr & reg.regStatName(ST_RSq) & ": " & _
                RegressAux.roundSigs(reg.regStat(ST_RSq), 3) & vbLf
        'F-stat and p-value
        textStr = textStr & indStr & reg.regStatName(ST_FStatPVal) & ": " & _
                RegressAux.roundSigs(reg.regStat(ST_FStat), 3) & " (" & _
                RegressAux.roundSigs(reg.regStat(ST_FStatPVal), 2) & ")" & vbLf
        ' Akaike information criterion, incl. corrected version
        textStr = textStr & indStr & reg.regStatName(ST_AIC) & ": " & _
                RegressAux.roundSigs(reg.regStat(ST_AIC)) & vbLf
        textStr = textStr & indStr & reg.regStatName(ST_CorrAIC) & ": " & _
                RegressAux.roundSigs(reg.regStat(ST_CorrAIC)) & vbLf
        
'        textStr = textStr & indStr
'        For iter = reg.statTypeIndex(True) To reg.statTypeIndex(False)
'            textStr = textStr & Trim(reg.regStatName(iter)) & ": " & _
'                    RegressAux.roundSigs(reg.regStat(iter), 4)
'            If iter Mod 2 = 0 Then
'                textStr = textStr & vbLf & indStr ' Two per line
'            Else
'                textStr = textStr & "; "
'            End If
'        Next iter
'        ' If last char is not vbLf, add one
'        If Not Right(textStr, 1) = vbLf Then textStr = textStr & vbLf
        
        
        TBxRegInfo.Text = textStr
    Else ' Nothing selected; clear the thing
        TBxRegInfo.Text = ""
    End If
End Sub

Private Function colSpreadThree(in1 As String, in2 As String, in3 As String, _
        Optional colWidth As Long = 16, Optional separator As String = " |", _
        Optional alignCode As String = "LLL") As String
    
    Dim workStr As String
    
    alignCode = UCase(alignCode)
    
    If Len(alignCode) <> 3 Then alignCode = "LLL"
    
    If Len(in1) <= colWidth Then
        If Mid(alignCode, 1, 1) = "R" Then
            workStr = String(colWidth - Len(in1), " ") & in1 & separator
        Else ' default L
            workStr = in1 & String(colWidth - Len(in1), " ") & separator
        End If
    Else
        workStr = Left(in1, colWidth)
    End If
    
    If Len(in2) <= colWidth Then
        If Mid(alignCode, 2, 1) = "R" Then
            workStr = workStr & String(colWidth - Len(in2), " ") & in2 & separator
        Else ' default L
            workStr = workStr & in2 & String(colWidth - Len(in2), " ") & separator
        End If
    Else
        workStr = workStr & Left(in2, colWidth)
    End If
    
    If Len(in3) <= colWidth Then
        If Mid(alignCode, 3, 1) = "R" Then
            workStr = workStr & String(colWidth - Len(in3), " ") & in3 & separator
        Else ' default L
            workStr = workStr & in3 & String(colWidth - Len(in3), " ") & separator
        End If
    Else
        workStr = workStr & Left(in3, colWidth)
    End If
    
    colSpreadThree = workStr
    
End Function

Private Function betaString(nameS As String, ByVal betaV As Variant, ByVal betaSEV As Variant, _
            ByVal tStatV As Variant, ByVal pV As Variant) As String
    Dim workStr As String
    Const nameSpan As Long = 10
    Const valSpan As Long = 8
    Const separator As String = " |"
    
    ' Convert if numeric; any Variant but an array should be coercible; array = BLOWUP is fine.
    If IsNumeric(betaV) Then betaV = CStr(betaV)
    If IsNumeric(betaSEV) Then betaSEV = CStr(betaSEV)
    If IsNumeric(tStatV) Then tStatV = CStr(tStatV)
    If IsNumeric(pV) Then pV = CStr(pV)
    
    If Len(nameS) <= nameSpan Then
        workStr = nameS & String(nameSpan - Len(nameS), " ") & separator
    Else
        workStr = Left(nameS, nameSpan) & separator
    End If
    
    If Len(betaV) <= valSpan Then
        workStr = workStr & betaV & String(valSpan - Len(betaV), " ") & separator
    Else
        workStr = workStr & Left(betaV, valSpan) & separator
    End If
    
    If Len(betaSEV) <= valSpan Then
        workStr = workStr & betaSEV & String(valSpan - Len(betaSEV), " ") & separator
    Else
        workStr = workStr & Left(betaSEV, valSpan) & separator
    End If
    
    If Len(tStatV) <= valSpan Then
        workStr = workStr & tStatV & String(valSpan - Len(tStatV), " ") & separator
    Else
        workStr = workStr & Left(tStatV, valSpan) & separator
    End If
    
    If Len(pV) <= valSpan Then
        workStr = workStr & pV & String(valSpan - Len(pV), " ") & separator
    Else
        workStr = workStr & Left(pV, valSpan) & separator
    End If
    
    betaString = workStr
End Function

'Private Sub BtnWriteStats_Click()
'    If Not (LBxOpenRegs.ListIndex >= 0 And RegColl.Count > 0) Then Exit Sub
'
'    If Not rColl.ItemKey(LBxOpenRegs.Value).writeStats Then
'        Call MsgBox("Write of statistics to """ & RegressAux.statShtName & _
'                """ worksheet failed.", vbOKOnly + vbExclamation, "Alert")
'    End If
'End Sub

Private Sub LBxOpenRegs_Change()
    popRegInfo
End Sub

Private Sub SetControlsStatus(btnsEnabled As Boolean)
    BtnChart.Enabled = btnsEnabled
    BtnClose.Enabled = btnsEnabled
    BtnCopy.Enabled = btnsEnabled
    BtnDefineNew.Enabled = btnsEnabled
    BtnDelete.Enabled = btnsEnabled
    BtnEdit.Enabled = btnsEnabled
    BtnFilterPreds.Enabled = btnsEnabled
    BtnOpen.Enabled = btnsEnabled
    BtnQuit.Enabled = btnsEnabled
    BtnShowDesc.Enabled = btnsEnabled
    BtnModSelAICc.Enabled = btnsEnabled
    BtnModSelAIC.Enabled = btnsEnabled
    LBxOpenRegs.Enabled = btnsEnabled
    LblWorking.Visible = Not btnsEnabled
End Sub

Private Sub UserForm_Activate()
    ' Probably some stuff to do here...
    ' Every time the form returns activated, sweep in any open rBooks not attached to Regs
    Dim rBk As Workbook, reg As ClsRegression
    Dim index As Long, wb As Workbook, bookFound As Boolean
    
    ' Disable buttons
    'Call SetControlsStatus(False)
    
    ' Refresh the source links on other than first load; change first-load flag if is first load
    If Not firstLoad Then
        RefreshSourceLinks
    Else
        firstLoad = False
    End If
    
    If Not checkingLoadedRegs Then
        ' Set flag to stop infinite-recursive checking
        checkingLoadedRegs = True
        
        ' Scan open workbooks
        If Workbooks.Count > 0 Then
            Set rBk = getUnlinkedRegBook
            Do Until rBk Is Nothing
                ' Create new regression and load rBk into it
                Set reg = New ClsRegression
                If Not reg.loadRegression(rBk) Then
                    ' Load failed, warn and continue
                    Call MsgBox("Load of open Regression book """ & _
                            Left(rBk.Name, Len(rBk.Name) - Len(RegressAux.rBookExtension)) & _
                            """ was unsuccessful.", vbOKOnly + vbExclamation, "Warning")
                Else
                    ' Load successful; add reg to RegColl
                    Call RegColl.Add(reg)
                End If
                ' Re-set the book to the next unlinked Reg book, if any
                Set rBk = getUnlinkedRegBook
            Loop
        End If
        
        popRegsList
        
        ' Clear flag
        checkingLoadedRegs = False
    End If
    
    ' Populate the info box
    popRegInfo
    
    ' If prior reg specified, select it and clear the retention variable
    If priorRegSelected > 0 Then
        LBxOpenRegs.ListIndex = priorRegSelected - 1
        priorRegSelected = 0
    End If
    
    ' Re-enable buttons
    Call SetControlsStatus(True)
    
    ' Put focus on quit
    BtnQuit.SetFocus
    
'    ' Check all Regs - if any source or rBook Workbooks are not open, open them
'    ' INCOMPLETE - MAY NOT NEED
'    If RegColl.Count > 0 Then
'        For index = 1 To RegColl.Count
'            Set reg = RegColl.Item(index)
'            bookFound = False
'            ' Checking for sourceBooks
'            For Each wb In Workbooks
'                If wb.Name = reg.SourceBook.Name Then bookFound = True
'            Next wb
'            If Not bookFound Then Workbooks.Open (reg.SourceBookPath)
'        Next index
'    End If
    
End Sub

Private Sub UserForm_Initialize()
    ' Initialize the reg-checking status flag
    checkingLoadedRegs = False
    quitBtnPressed = False
    priorRegSelected = 0
    firstLoad = True
    lastPath = ""
End Sub

Private Sub runModelSelection(critType As StatType, critName As String)
    Dim wkBk As Workbook, wkSht As Worksheet
    Dim fs As FileSystemObject, reg As ClsRegression, wsf As WorksheetFunction
    Dim dataRg As Range, workupRg As Range, sortRg As Range, statsRg As Range
    Dim workStr As String, rowOffset As Long, lastCrit As Double
    Dim iter As Long, workVal As Double, maxPVal As Double, maxPValIdx As Long
    Dim errNum As Long
    Dim workPredName As String
    Dim keepLooping As Boolean, cannotSave As Boolean
    Dim resp As VbMsgBoxResult
    Dim wkChOb As ChartObject, wkCh As Chart, wkSrs As Series
    Dim critRange As Double, critMin As Double, critMax As Double, critOrder As Long
    
    Dim basePredCount As Long, redPredCount As Long, removedPredCount As Long
    
    ' If no reg selected, just drop
    If LBxOpenRegs.ListIndex < 0 Then Exit Sub
    
    ' Disable controls
    Call SetControlsStatus(False)
    
    ' Bind objects
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set reg = rColl.ItemKey(LBxOpenRegs.value)
    Set wsf = Application.WorksheetFunction
    
    ' Create destination for analysis
    Set wkBk = Workbooks.Add
    
    ' Kill all but one sheet
    Do While wkBk.Sheets.Count > 1
        wkBk.Sheets(1).Delete
    Loop
    
    ' Bind the remaining and rename
    Set wkSht = wkBk.Sheets(1)
    wkSht.Name = "Model Selection"
    
    ' Bind the data range ref cell
    Set dataRg = wkSht.Range(RegressAux.modelDataRgCell)
    
    ' Add headers
    dataRg.Offset(-1, modelCritColOffset) = critName
    dataRg.Offset(-1, modelR2ColOffset) = "R^2"
    dataRg.Offset(-1, modelParamColOffset) = "Param"
    dataRg.Offset(-1, modelPValColOffset) = "p-value"
    
    ' Initialize the row offset and reference AIC (dummy value to ensure starting)
    rowOffset = 0
    lastCrit = reg.regStat(critType) + 1
    
    ' Store the base number of predictors
    basePredCount = reg.numPredictors(True)
    
    ' Loop until criterion increases
    Do
        ' Store the criterion and R^2 value. Also store criterion to
        '  helper variable
        dataRg.Offset(rowOffset, modelCritColOffset) = reg.regStat(critType)
        dataRg.Offset(rowOffset, modelR2ColOffset) = reg.regStat(ST_RSq)
        lastCrit = reg.regStat(critType)
        
        ' Identify the predictor with the highest p-value
        maxPVal = 0#
        maxPValIdx = 0
        For iter = 1 To reg.numPredictors(True)
            workVal = RegressAux.calcPStat( _
                                reg.predictorBeta(iter), _
                                reg.predictorBetaSE(iter), _
                                reg.DOF(DOFResidual))
            If maxPVal < workVal Then
                maxPVal = workVal
                maxPValIdx = iter
            End If
        Next iter
        
        ' Store the name and the p-value
        dataRg.Offset(rowOffset + 1, modelPValColOffset) = maxPVal
        workPredName = reg.predictorName(maxPValIdx, True)
        dataRg.Offset(rowOffset + 1, modelParamColOffset) = workPredName
        
        ' Change the predictor filter
        reg.filterString(crfPredictor) = _
                RegressAux.addFilterToString(reg.filterString(crfPredictor), _
                                                reg.predictorIndex(workPredName, False))
        
        ' Update to reflect the new filters
        regenerateRegression reg, True, reg.includeConstant
        
        ' Increment the row index
        rowOffset = rowOffset + 1
    
    Loop Until reg.regStat(critType) > lastCrit Or reg.numPredictors(True) < 2
    
    ' Ok, pulled one too many predictors at this point, or down only to one predictor
    If reg.regStat(critType) < lastCrit Then
        ' Only one predictor remains
        'dataRg.Offset(rowOffset + 3, 0) = "ONE PREDICTOR REMAINS; CHECK FOR SIGNIFICANCE"
        dataRg.Offset(rowOffset, modelCritColOffset) = reg.regStat(critType)
        dataRg.Offset(rowOffset, modelR2ColOffset) = reg.regStat(ST_RSq)
        rowOffset = rowOffset + 1
        dataRg.Offset(rowOffset, modelParamColOffset).value = reg.predictorName(1, True)
        dataRg.Offset(rowOffset, modelPValColOffset).value = RegressAux.calcPStat( _
                                                        reg.predictorBeta(1), _
                                                        reg.predictorBetaSE(1), _
                                                        reg.DOF(DOFResidual))
        redPredCount = reg.numPredictors(True) ' Should be one, always
        removedPredCount = basePredCount - redPredCount
    Else
        ' Delete the just inserted param and p-value, and add a divider line
        With dataRg.Offset(rowOffset, modelParamColOffset)
            .value = ""
            .Borders(xlEdgeTop).Weight = xlThin
        End With
        With dataRg.Offset(rowOffset, modelPValColOffset)
            .value = ""
            .Borders(xlEdgeTop).Weight = xlThin
        End With
        dataRg.Offset(rowOffset, modelCritColOffset).Borders(xlEdgeTop).Weight = xlThin
        dataRg.Offset(rowOffset, modelR2ColOffset).Borders(xlEdgeTop).Weight = xlThin
        
        ' Change the predictor filter and reload; workPredName holds the last-filtered name
        reg.filterString(crfPredictor) = _
                RegressAux.delFilterFromString(reg.filterString(crfPredictor), _
                                                reg.predictorIndex(workPredName, False))
        
        ' Update regression to reflect the new filters; inconsistent state if failed; exit
        regenerateRegression reg, True, reg.includeConstant
        
        ' Store the reduced number of predictors
        redPredCount = reg.numPredictors(True)
        removedPredCount = basePredCount - redPredCount
        
        ' If model already reduced, just inform, re-enable controls, and exit
        If basePredCount = redPredCount Then
            MsgBox "Model is already reduced.", vbOKOnly, "Model already reduced"
            wkBk.Close SaveChanges:=False
            SetControlsStatus True
            Exit Sub
        End If
        
        ' Loop through the remaining predictors, storing name, p-value; and criterion/R^2 when
        '  removed
        For iter = 1 To reg.numPredictors(True)
            ' Store the name of the predictor to be removed; placing into sheet, with p-val
            workPredName = reg.predictorName(iter, True)
            dataRg.Offset(rowOffset, modelParamColOffset).value = workPredName
            dataRg.Offset(rowOffset, modelPValColOffset).value = _
                            RegressAux.calcPStat( _
                                        reg.predictorBeta(iter), _
                                        reg.predictorBetaSE(iter), _
                                        reg.DOF(DOFResidual) _
                                                )
            
            ' Filter out the predictor
            reg.filterString(crfPredictor) = _
                    RegressAux.addFilterToString(reg.filterString(crfPredictor), _
                                                reg.predictorIndex(workPredName, False))
            
            ' Update the regression
            regenerateRegression reg, True, reg.includeConstant
            
            ' Store the criterion and the R^2 value
            dataRg.Offset(rowOffset, modelCritColOffset).value = reg.regStat(critType)
            dataRg.Offset(rowOffset, modelR2ColOffset).value = reg.regStat(ST_RSq)
            
            ' Restore the predictor
            reg.filterString(crfPredictor) = _
                    RegressAux.delFilterFromString(reg.filterString(crfPredictor), _
                                                reg.predictorIndex(workPredName, False))
            
            ' Update the regression
            regenerateRegression reg, True, reg.includeConstant
            
            ' Increment the row counter
            rowOffset = rowOffset + 1
        Next iter
        
    End If
    
    ' Sort the data rows of predictors staying in the model, if there is more than one
    If reg.numPredictors(True) > 1 Then
        Set sortRg = Intersect(dataRg.Offset(removedPredCount + 1, 0) _
                    .Resize(redPredCount, 1).EntireRow, dataRg.CurrentRegion)
        sortRg.Sort Key1:=sortRg.Columns(modelCritColOffset + 1).Cells, Order1:=xlDescending, Header:=xlNo
    End If
    
'    ' Populate the workup block
    ' Link the cell reference
    Set workupRg = wkSht.Range(RegressAux.modelWkupRgCell)

    ' Headers
    workupRg.Offset(-1, modelWUNameColOffset).value = "Description"
    workupRg.Offset(-1, modelWUCritDataColOffset).value = "Adj Crit Value"
    workupRg.Offset(-1, modelWUCritMinColOffset).value = "Min Crit"
    workupRg.Offset(-1, modelWUCritDiffColOffset).value = "Crit " & Chr(150) & " Min"
    workupRg.Offset(-1, modelWUR2DataColOffset).value = "R^2 Value"
    workupRg.Offset(-1, modelWUR2RefColOffset).value = "Ref R^2"
    workupRg.Offset(-1, modelWUR2DiffColOffset).value = "Ref " & Chr(150) & " R^2"
    workupRg.Offset(-1, modelWUR2BasebarColOffset).value = "R^2 Base"

    ' Fill all of the various criterion and R^2 values
    With workupRg
        For iter = 0 To basePredCount
            ' Name
            If iter = 0 Then
                .Offset(iter, modelWUNameColOffset).value = "BASE MODEL"
            ElseIf iter = basePredCount And reg.numPredictors(True) = 1 Then
                ' Do nothing; don't add in the name for a single remaining predictor
            Else
                .Offset(iter, modelWUNameColOffset).value = _
                            Chr(150) & " " & dataRg.Offset(iter, modelParamColOffset).value
            End If
    
            ' Criterion and R^2 values from the various regression forms
            .Offset(iter, modelWUCritDataColOffset).value = _
                            dataRg.Offset(iter, modelCritColOffset).value - _
                            dataRg.Offset(removedPredCount, modelCritColOffset).value
            .Offset(iter, modelWUR2DataColOffset).value = _
                            dataRg.Offset(iter, modelR2ColOffset).value
            
            ' Calculated values for the charting
            If iter < removedPredCount Then
                .Offset(iter, modelWUCritMinColOffset).Formula = "=NA()"
                .Offset(iter, modelWUCritDiffColOffset).Formula = "=NA()"
                .Offset(iter, modelWUR2RefColOffset).Formula = "=NA()"
                .Offset(iter, modelWUR2DiffColOffset).Formula = "=NA()"
                .Offset(iter, modelWUR2BasebarColOffset).Formula = "=NA()"
            ElseIf iter = removedPredCount Then
                .Offset(iter, modelWUCritMinColOffset).Formula = "=" & _
                        .Offset(removedPredCount, modelWUCritDataColOffset).Address(True, True)
                .Offset(iter, modelWUCritDiffColOffset).Formula = "=NA()"
                .Offset(iter, modelWUR2RefColOffset).Formula = "=" & _
                        .Offset(removedPredCount, modelWUR2DataColOffset).Address(True, True)
                .Offset(iter, modelWUR2DiffColOffset).Formula = "=NA()"
                .Offset(iter, modelWUR2BasebarColOffset).Formula = "=NA()"
            Else
                .Offset(iter, modelWUCritMinColOffset).Formula = "=" & _
                        .Offset(removedPredCount, modelWUCritDataColOffset).Address(True, True)
                .Offset(iter, modelWUCritDiffColOffset).Formula = "=" & _
                        .Offset(iter, modelWUCritDataColOffset).Address(True, True) & _
                        "-" & .Offset(iter, modelWUCritMinColOffset).Address(True, True)
                .Offset(iter, modelWUR2RefColOffset).Formula = "=" & _
                        .Offset(removedPredCount, modelWUR2DataColOffset).Address(True, True)
                .Offset(iter, modelWUR2DiffColOffset).Formula = "=" & _
                        .Offset(iter, modelWUR2RefColOffset).Address(True, True) & _
                        "-" & .Offset(iter, modelWUR2DataColOffset).Address(True, True)
                .Offset(iter, modelWUR2BasebarColOffset).Formula = "=" & _
                        .Offset(iter, modelWUR2DataColOffset).Address(True, True)
            End If
        Next iter
    End With
    ' Autofit the sheet contents
    wkSht.UsedRange.EntireColumn.AutoFit
    
    ' Paste in the regression stats
    Set statsRg = wkSht.Range(RegressAux.modelStatsRgCell)
    reg.copyDataRange crcrStats
    statsRg.PasteSpecial xlPasteValues
    statsRg.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    
    ' Calculate some stats on the criterion
    critMin = wsf.Min(dataRg.CurrentRegion.Columns(modelCritColOffset + 1))
    critMax = wsf.Max(dataRg.CurrentRegion.Columns(modelCritColOffset + 1))
    critRange = critMax - critMin
    
    
    ' Define the chart object, & set chart type
    Set wkChOb = wkSht.ChartObjects.Add(350, 250, 400, 300)
    wkChOb.Chart.ChartType = xlLine
    
    ' Criterion data series during down-selection
    With wkChOb.Chart.SeriesCollection.Add( _
                Source:=workupRg.Offset(0, modelWUCritDataColOffset) _
                                        .Resize(removedPredCount + 1, 1), _
                rowcol:=xlColumns, _
                SeriesLabels:=False, _
                CategoryLabels:=False, _
                Replace:=False)
        .XValues = workupRg.Offset(0, modelWUNameColOffset).Resize(basePredCount + 1, 1)
        .MarkerStyle = xlMarkerStyleDiamond
        .MarkerBackgroundColor = RegressAux.modelBlueColor
        .MarkerForegroundColor = RegressAux.modelBlueColor
        .MarkerSize = 5
        .Format.Line.ForeColor.RGB = RegressAux.modelBlueColor
        .Format.Line.DashStyle = msoLineSolid
        .Format.Line.Weight = RegressAux.modelLineWeight
        .Name = workupRg.Offset(-1, modelWUCritDataColOffset)
    End With
    
    ' Pull off the legend
    wkChOb.Chart.HasLegend = False
    
    ' R^2 data series
    With wkChOb.Chart.SeriesCollection.Add( _
                Source:=workupRg.Offset(0, modelWUR2DataColOffset) _
                                        .Resize(removedPredCount + 1, 1), _
                rowcol:=xlColumns, _
                SeriesLabels:=False, _
                CategoryLabels:=False, _
                Replace:=False)
        .AxisGroup = xlSecondary
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerBackgroundColor = RegressAux.modelGreenColor
        .MarkerForegroundColor = RegressAux.modelGreenColor
        .MarkerSize = 5
        .Format.Line.ForeColor.RGB = RegressAux.modelGreenColor
        .Format.Line.DashStyle = msoLineSolid
        .Format.Line.Weight = RegressAux.modelLineWeight
        .Name = workupRg.Offset(-1, modelWUR2DataColOffset)
    End With
    
    ' Add the dashed reference lines across the bar chart area
    With wkChOb.Chart.SeriesCollection.Add( _
                Source:=workupRg.Offset(0, modelWUCritMinColOffset) _
                    .Resize(basePredCount + 1, 1), _
                rowcol:=xlColumns, _
                SeriesLabels:=False, _
                CategoryLabels:=False, _
                Replace:=False)
        .AxisGroup = xlPrimary
        .MarkerStyle = xlMarkerStyleNone
        .Format.Line.ForeColor.RGB = RegressAux.modelBlueColor
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = RegressAux.modelLineWeight
        .Name = workupRg.Offset(-1, modelWUCritMinColOffset)
    End With
    With wkChOb.Chart.SeriesCollection.Add( _
                Source:=workupRg.Offset(0, modelWUR2RefColOffset) _
                    .Resize(basePredCount + 1, 1), _
                rowcol:=xlColumns, _
                SeriesLabels:=False, _
                CategoryLabels:=False, _
                Replace:=False)
        .AxisGroup = xlSecondary
        .MarkerStyle = xlMarkerStyleNone
        .Format.Line.ForeColor.RGB = RegressAux.modelGreenColor
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = RegressAux.modelLineWeight
        .Name = workupRg.Offset(-1, modelWUR2RefColOffset)
    End With
    
    ' Only add bars if more than one predictor remains
    If reg.numPredictors(True) > 1 Then
        ' Add crit removal-from-final-model bars
        With wkChOb.Chart.SeriesCollection.Add( _
                    Source:=workupRg.Offset(0, modelWUCritDiffColOffset) _
                        .Resize(basePredCount + 1, 1), _
                    rowcol:=xlColumns, _
                    SeriesLabels:=False, _
                    CategoryLabels:=False, _
                    Replace:=False)
            .AxisGroup = xlPrimary
            .MarkerStyle = xlMarkerStyleNone
            .ChartType = xlColumnStacked
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RegressAux.modelBlueColor
            .Format.Fill.Visible = msoTrue
            .Format.Fill.Patterned RegressAux.modelBarPattern
            .Format.Fill.ForeColor.RGB = RegressAux.modelBlueColor
            .Format.Fill.BackColor.RGB = RegressAux.modelWhiteColor
            .Parent.GapWidth = RegressAux.modelCritBarGapWidth
            .Name = workupRg.Offset(-1, modelWUCritDiffColOffset)
        End With
        
        ' Add R^2 base bars
        With wkChOb.Chart.SeriesCollection.Add( _
                    Source:=workupRg.Offset(0, modelWUR2BasebarColOffset) _
                        .Resize(basePredCount + 1, 1), _
                    rowcol:=xlColumns, _
                    SeriesLabels:=False, _
                    CategoryLabels:=False, _
                    Replace:=False)
            .AxisGroup = xlSecondary
            .MarkerStyle = xlMarkerStyleNone
            .ChartType = xlColumnStacked
            .Format.Line.Visible = msoFalse
            .Format.Fill.Visible = msoFalse
            .Name = workupRg.Offset(-1, modelWUR2BasebarColOffset)
        End With
        
        ' Add R^2 delta bars, tweak gap width
        With wkChOb.Chart.SeriesCollection.Add( _
                    Source:=workupRg.Offset(0, modelWUR2DiffColOffset) _
                        .Resize(basePredCount + 1, 1), _
                    rowcol:=xlColumns, _
                    SeriesLabels:=False, _
                    CategoryLabels:=False, _
                    Replace:=False)
            .AxisGroup = xlSecondary
            .MarkerStyle = xlMarkerStyleNone
            .ChartType = xlColumnStacked
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RegressAux.modelGreenColor
            .Format.Fill.Visible = msoTrue
            .Format.Fill.Patterned RegressAux.modelBarPattern
            .Format.Fill.ForeColor.RGB = RegressAux.modelGreenColor
            .Format.Fill.BackColor.RGB = RegressAux.modelWhiteColor
            .Parent.GapWidth = RegressAux.modelR2BarGapWidth
            .Name = workupRg.Offset(-1, modelWUR2DiffColOffset)
        End With
    End If

    ' Reformat axes
    With wkChOb.Chart.Axes(xlValue, xlPrimary)
        .MajorUnit = wsf.Ceiling_Precise(critRange / 4, 4)
        .MinimumScale = -1 * .MajorUnit
        .MaximumScale = 4 * .MajorUnit
        .TickLabels.Font.Size = 14
        .TickLabels.Font.Color = RegressAux.modelBlueColor
        .TickLabels.NumberFormat = "[>=]0;[<0]"""""
        .HasTitle = True
        .AxisTitle.Text = critName & " " & Chr(150) & " min(" & critName & ")"
        .AxisTitle.Characters.Font.Color = RegressAux.modelBlueColor
        .AxisTitle.Characters.Font.Size = 16
        .CrossesAt = .MinimumScale
        .HasMajorGridlines = True
        .MajorGridlines.Border.Color = modelGridlineColor
        .MajorGridlines.Border.LineStyle = xlDash
    End With
    With wkChOb.Chart.Axes(xlValue, xlSecondary)
        .MinimumScale = 0
        .MaximumScale = 1
        .MajorUnit = 0.2
        .TickLabels.Font.Size = 14
        .TickLabels.Font.Color = RegressAux.modelGreenColor
        .TickLabels.NumberFormat = "0.0"
        .HasTitle = True
        .AxisTitle.Text = "R2"
        .AxisTitle.Characters.Font.Color = RegressAux.modelGreenColor
        .AxisTitle.Characters.Font.Size = 16
        .AxisTitle.Characters(2).Font.Superscript = True
    End With
    With wkChOb.Chart.Axes(xlCategory, xlPrimary)
        .TickLabels.Font.Size = 14
        .TickLabels.Orientation = xlUpward
        .MajorGridlines.Border.Color = modelGridlineColor
        .MajorGridlines.Border.LineStyle = xlDash
    End With

    ' Some general touchup
    ' Remove border line on chart object and add to the PlotArea
    wkChOb.ShapeRange(1).Line.Visible = msoFalse
    wkChOb.Chart.PlotArea.Format.Line.Visible = msoTrue
    'wkChOb.Chart.PlotArea.Format.Line.Style = msoLineSingle
    wkChOb.Chart.PlotArea.Format.Line.ForeColor.RGB = RegressAux.modelAxisColor ' Middlin' gray
    
    
    ' Query for name and save, if name given
    keepLooping = True
    Do While keepLooping
        ' First-pass initialization
        workStr = ""
        Do
            If workStr <> "" Then
                ' Complain if not the first time through; means a bad filename given
                MsgBox "Invalid filename, please re-enter", vbOKOnly + vbExclamation, "Invalid filename"
            End If
            workStr = RegressAux.requestFilename("Workbook name for model selection output?" & _
                                    vbLf & vbLf & "(Cancel to leave unsaved)", reg.Name & " - Model")
        Loop Until RegressAux.validRegName(workStr) Or workStr <> ""
        
        cannotSave = False
        If workStr <> "" Then
            ' Check for open workbooks of the same name
            For iter = 1 To Workbooks.Count
                If Workbooks(iter).Name = workStr & RegressAux.eBookExtension Then
                    MsgBox "Cannot save: A workbook with that name is currently open.", _
                            vbOKOnly + vbExclamation, "Cannot save"
                    cannotSave = True
                End If
            Next iter
            
            ' Check for an existing file
            If Not cannotSave Then
                If fs.FileExists(reg.SourceBook.path & "\" & workStr & RegressAux.eBookExtension) Then
                    resp = MsgBox("Cannot save: File already exists." & vbLf & vbLf & _
                            "Overwrite?" & vbLf & vbLf & "(Cancel skips saving.)", _
                            vbYesNoCancel + vbExclamation, "Cannot save")
                    Select Case resp
                    Case vbYes
                        ' Close workbook if open
                        For iter = 1 To Workbooks.Count
                            If Workbooks(iter).Name = workStr & RegressAux.eBookExtension Then
                                Workbooks(iter).Close SaveChanges:=False
                                Exit For
                            End If
                        Next iter
                        
                        ' Delete workbook
                        fs.DeleteFile reg.SourceBook.path & "\" & workStr & RegressAux.eBookExtension, True
    
                        ' Proceed with save
                    Case vbNo
                        ' Don't overwrite; go back to name entry
                        cannotSave = True
                        
                    Case vbCancel
                        ' Don't save and don't keep looping
                        cannotSave = True
                        keepLooping = False
                    End Select
                End If
            End If
            
            ' Try to save if conditions are ok
            If Not cannotSave Then
                On Error Resume Next
                wkBk.SaveAs reg.SourceBook.path & "\" & workStr
                errNum = Err.Number
                Err.Clear
                On Error GoTo 0
                Select Case errNum
                Case 0
                    ' All fine; proceed
                    keepLooping = False
                Case Else
                    ' Some other error; let user deal with the open sheet.
                    MsgBox "An error occurred while trying to save the workbook", _
                            vbOKOnly + vbExclamation, "Save error"
                    keepLooping = False
                End Select
            End If
        Else
            ' User cancelled filename entry; exit the loop
            keepLooping = False
        End If
    Loop
    

    ' Re-enable controls
    Call SetControlsStatus(True)

End Sub

Private Sub regenerateRegression(reg As ClsRegression, retainSrc As Boolean, _
            includeConstant As Boolean)
    ' Update to reflect the new filters; inconsistent state if failed; exit
    If Not reg.modifyRegression(retainSrc, includeConstant) Then
        Call MsgBox("Regeneration of Regression failed!  Exiting...", _
                vbOKOnly + vbCritical, "Critical Error")
        Unload RegressMain
        Exit Sub
    End If
    
    ' Repopulate the info box
    popRegInfo
    
    ' Always refresh source links here
    RefreshSourceLinks
    
End Sub


