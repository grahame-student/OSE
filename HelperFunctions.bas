Attribute VB_Name = "HelperFunctions"
Option Explicit
DefObj A-Z

Public Function Ordinal(ByVal DayValue As Integer) As String

    Select Case DayValue
         Case 1, 21, 31
            Ordinal = DayValue & "st"
         Case 2, 22
            Ordinal = DayValue & "nd"
         Case 3, 23
            Ordinal = DayValue & "rd"
         Case Else
            Ordinal = DayValue & "th"
      End Select

End Function

Public Sub ValidateInput(ByRef CheckTextBox As TextBox, ByVal MinVal As Long, ByVal MaxVal As Long)
    
    If Not (IsNumeric(CheckTextBox.Text)) Then
        SendMessage CheckTextBox.hwnd, EM_UNDO, 0, ByVal 0&
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
    ElseIf CheckTextBox.Text < MinVal Or CheckTextBox.Text > MaxVal Then
        SendMessage CheckTextBox.hwnd, EM_UNDO, 0, ByVal 0&
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
    Else
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
        CheckTextBox.Text = Int(CheckTextBox.Text)
    End If

End Sub

