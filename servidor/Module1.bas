Attribute VB_Name = "Module1"
Option Explicit

Public Function ruta() As String
    Dim strruta As String
    strruta = App.Path
    
    If Right(strruta, 1) <> "\" Then
        strruta = strruta & "\"
        ruta = strruta
    End If

End Function
