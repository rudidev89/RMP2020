VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_component"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Exits(ComponentNameAndExtention As String) As Boolean
    Dim Cm As Object
    For Each Cm In ThisWorkbook.VBProject.VBComponents
        If Cm.Name = ComponentNameAndExtention Then
            Exits = True
        End If
    Next Cm
End Function


Public Function Comp(ComponentNameAndExtention As String) As Object
    If Exits(ComponentName) Then
        Set Comp = ThisWorkbook.VBProject.VBComponents(ComponentNameAndExtention)
    Else
        With ThisWorkbook.VBProject
            .Import
        End With
    End If
End Function
