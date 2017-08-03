VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegressNameEntry 
   Caption         =   "Enter regression name and description"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   OleObjectBlob   =   "RegressNameEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegressNameEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        RegressNameEntry.frm
' # Purpose:     UserForm for defining the name, response variable name,
' #               verbose description, and save location of a multiple
' #               regression analysis file.
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

Private closeCancel As Boolean
Private folderPicked As Boolean
Private saveFolder As Folder
Private Const saveFolderDefault As String = "..."

Public Function closedByCancel() As Boolean
    closedByCancel = closeCancel
End Function

Public Function enteredName() As String
    enteredName = TBxName.value
End Function

Public Function enteredRespName() As String
    enteredRespName = TBxRespName.value
End Function

Public Function enteredDescription() As String
    enteredDescription = TBxDesc.value
End Function

Public Function pickedFolderPath() As String
    pickedFolderPath = saveFolder.path
End Function

Public Sub setNameDesc(inName As String, inRespName As String, inDesc As String, _
                Optional isFillerName As Boolean = False, _
                Optional inFolderPath As String = "")
    
    Dim fs As FileSystemObject
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Fill in default name and description
    TBxName.value = inName ' Any starting name accepted; check for validity at form OK close
    TBxRespName.value = inRespName
    TBxDesc.value = inDesc
    
    ' Check whether input string is a real path.  If it's blank, don't notify
    If Not fs.FolderExists(inFolderPath) Then
        folderPicked = False
        If Not inFolderPath = "" Then
            Call MsgBox("Indicated save folder does not exist. Please re-select.", vbExclamation + vbOKOnly, "Alert")
        End If
    Else
        Set saveFolder = fs.GetFolder(inFolderPath)
        folderPicked = True
        Call popFolderTBx(inFolderPath)
    End If
    
    ' If it's a filler name, focus on the textbox, highlight, set italics, and turn fonts black
    If isFillerName Then
        LblNameNote.ForeColor = blackColor
        With TBxName
            .SetFocus
            .ForeColor = blackColor
            .Font.Italic = True
            .SelStart = 0
            .SelLength = Len(.value)
        End With
    End If
    
    ' Set the OK button
    setOkBtn
    
End Sub

Private Sub popFolderTBx(path As String)
    If Len(path) < RegressAux.maxNameLength Then
        TBxSaveFolder.value = path
    Else
        TBxSaveFolder.value = Left(path, 25) & " ... " & Right(path, 25)
    End If
End Sub

Private Sub setOkBtn()
    ' Unless the name field is valid AND the folder is set AND a non-null response
    '  variable name is set, disable OK button
    If RegressAux.validRegName(TBxName.value) And folderPicked And _
            Len(TBxRespName.value) > 0 Then
        BtnOK.Enabled = True
    Else
        BtnOK.Enabled = False
    End If
End Sub

Private Sub BtnCancel_Click()
    RegressNameEntry.Hide
End Sub

Private Sub BtnOK_Click()
    If Not RegressAux.validRegName(TBxName.value) Then
        ' Should never see this; button should be enabled only if valid name entered
        Call MsgBox("Please enter a valid Regression name.", vbOKOnly + vbCritical, "Error")
        Exit Sub
    End If
    
    closeCancel = False
    RegressNameEntry.Hide
    
End Sub

Private Sub BtnPickFolder_Click()
    Dim fd As FileDialog, fs As FileSystemObject
    
    ' Initialize objects
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    With fd
        ' Configure
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialView = msoFileDialogViewList
        If Not fs.FolderExists(TBxSaveFolder.value) Then
            If RegressMain.lastPath = "" Then
                .InitialFileName = "%homepath%\Documents"
            Else
                .InitialFileName = RegressMain.lastPath
            End If
        Else
            .InitialFileName = TBxSaveFolder.value
        End If
        .Title = "Choose Folder..."
        
        ' Show, check -- exit if cancelled
        If Not .Show = -1 Then Exit Sub  ' Depends on FileDialog existence checking
        
        ' If folder set, attach and configure form
        Set saveFolder = fs.GetFolder(.SelectedItems(1))
        RegressMain.lastPath = .SelectedItems(1)
        Call popFolderTBx(saveFolder.path)
        folderPicked = True
        setOkBtn
    End With
    
End Sub

Private Sub TBxName_Change()
    If Not RegressAux.validRegName(TBxName.value) Then
        TBxName.ForeColor = redColor
        'LblName.ForeColor = red
        LblNameNote.ForeColor = redColor
        'BtnOK.Enabled = False
    Else
        TBxName.ForeColor = blackColor
        'LblName.ForeColor = black
        LblNameNote.ForeColor = blackColor
        'BtnOK.Enabled = True
    End If
    
    ' Set the OK button status
    setOkBtn
    
    ' Font defaults to italics if it's a placeholder text string; want it as normal text
    '  pretty much all other times
    TBxName.Font.Italic = False
End Sub

Private Sub TBxRespName_Change()
    setOkBtn
End Sub

Private Sub UserForm_Initialize()
    closeCancel = True
    folderPicked = False
    TBxSaveFolder.value = saveFolderDefault
End Sub

Private Function redColor() As Long
    redColor = &HFF
End Function

Private Function blackColor() As Long
    blackColor = -2147483640
End Function
