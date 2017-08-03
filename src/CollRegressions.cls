VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollRegressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        CollRegressions.cls
' # Purpose:     Custom wrapper class for a Collection of ClsRegression objects.
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

Private Regs As New Collection


Public Function Add(newReg As ClsRegression) As Boolean
    Dim reg As ClsRegression ' Collection contains anonymous objects
    Dim iter As Long
    
    ' Failure return value
    Add = False
    
    ' Check whether name of desired-to-add regression is valid
    ' Right now, just a check for a Reg with a duplicate name - presuming that
    '  the Reg creation, etc. functionality is successfully preventing bad
    '  regression names from being chosen
    For Each reg In Regs
        If reg.Name = newReg.Name Then
            Call MsgBox("Regression with name """ & newReg.Name & """ already exists.", _
                    vbOKOnly + vbCritical, "Error")
            Exit Function
        End If
    Next reg
    
    ' Name appears ok
    
    ' Find insertion point
    For iter = 1 To Regs.Count
        If newReg.Name < Regs.Item(iter).Name Then  ' insert before the current Item
            With Regs
                Call .Add(Item:=newReg, Key:=newReg.Name, Before:=iter)
            End With
            Add = True
            Exit Function
        End If
    Next iter
    
    ' If reaches here, need to add at end of collection
    With Regs
        Call .Add(Item:=newReg, Key:=newReg.Name) ', After:=.Item(.Count))
    End With
    
    ' Successful return
    Add = True
    
End Function

Public Function CullEmpty() As Boolean
    ' Return value indicates whether empty reg was found and removed from collection
    Dim iter As Long
    
    CullEmpty = False
    
    ' Check for empty-name regs
    For iter = 1 To Regs.Count
        If Regs.Item(iter).Name = "" Then
            CullEmpty = True
            Call Regs.Remove(iter)
            Exit Function
        End If
    Next iter
    
End Function

Public Function Remove(regName As String) As Boolean
    Dim reg As ClsRegression ' Collection contains anonymous objects
    
    ' Failure return value
    Remove = False
      
    ' Check whether exists; if does, remove and exit success
    For Each reg In Regs
        If reg.Name = regName Then
            Remove = True
            Call Regs.Remove(regName)
            Exit Function
        End If
    Next reg
    
    ' Not found; alert and exit (default) fail
    Call MsgBox("Regression with name """ & regName & """ not found to delete", _
            vbOKOnly + vbCritical, "Error")
    
End Function

Public Function Item(index As Long) As ClsRegression
    
    Set Item = Nothing
    
    ' Check that index is ok
    If index < 1 Or index > Regs.Count Then
        Call MsgBox("Index out of bounds", vbCritical + vbOKOnly, "Error")
        Exit Function
    End If
    
    Set Item = Regs.Item(index)
    
End Function

Public Function ItemKey(regName As String) As ClsRegression
    Dim reg As ClsRegression ' Collection contains anonymous objects
    
    Set ItemKey = Nothing ' Nothing return value if not found
    
    For Each reg In Regs
        If regName = reg.Name Then
            Set ItemKey = Regs(regName)
            Exit Function
        End If
    Next reg
    
    ' Throw not-found errorbox
    Call MsgBox("Regression with name """ & regName & """ not found.", _
            vbOKOnly + vbCritical, "Error")
    
End Function

Public Function Count() As Long
    Count = Regs.Count
End Function

Public Function RegExists(regName As String) As Boolean
    Dim reg As ClsRegression
    
    RegExists = False
    
    For Each reg In Regs
        If regName = reg.Name Then RegExists = True
    Next reg
    
End Function
