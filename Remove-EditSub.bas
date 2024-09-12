Attribute VB_Name = "Module1"
Sub RemoveSubstringAndSubsequent()
    Dim ws As Worksheet
    Dim cell As Range
    Dim substring As String
    Dim rng As Range
    Dim pos As Integer
    
    Set ws = ActiveSheet
    
    Set rng = Selection

    substring = InputBox("What string do you want to remove?")
    
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            pos = InStr(cell.Value, substring)
            If pos > 0 Then
                cell.Value = Left(cell.Value, pos - 1)
                '  substring is found within string1, pos = position of sub
            End If
        End If
    Next cell
End Sub

