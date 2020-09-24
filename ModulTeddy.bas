Attribute VB_Name = "ModulTeddy"
'Created By: Teddy Siswoyo(TedzHutz)
Public Function GetTextName(TextValue As String, ObjectTarget As Integer, Optional OutputString As String) As String
Dim NowObject As Integer 'The Text Value must be in this format ",[txt],[txt],[txt],"
Dim TextResult As String 'Every Value is seperated by ","
Dim i As Integer
    For i = 1 To Len(TextValue) 'To Get a text, This Function work similar with select case
    If Mid(TextValue, i, 1) = "," Then NowObject = NowObject + 1
        If NowObject = ObjectTarget Then
           If Mid(TextValue, i + 1, 1) <> "," Then TextResult = TextResult + Mid(TextValue, i + 1, 1)
        ElseIf ObjectTarget < NowObject Then
           OutputString = TextResult
           GetTextName = TextResult
           Exit For
        End If
    Next i
End Function
Public Sub DelSpace(Txtstring As String) 'To delete the empty space, work similar with Trim or Ltrim or Rtrim
Dim TempString As String
Dim i As Integer
    For i = 1 To Len(Txtstring)
        If Mid(Txtstring, i, 1) <> " " Then 'If not empty space then update the string
           TempString = TempString + Mid(Txtstring, i, 1)
        End If
    Next i
    Txtstring = TempString
End Sub



