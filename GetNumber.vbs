Sub getUrl()

'
'
Dim hlink As String

Dim rng As Range

Dim str_start As Long

Dim str_end As Long
Sheets("sheet1").Activate


For i = 1 To 200
    str_start = 0
    str_end = 0
    Set rng = Worksheets("sheet1").Range("B" & i)
    If rng.Hyperlinks.Count Then
        hlink = rng(1).Hyperlinks(1).Address
        If InStr(hlink, "view('") <> 0 Then
            str_start = InStr(hlink, "view('") + 6
            str_end = InStr(hlink, "')") - 2
        End If

    If str_start - str_end <> 0 Then
    Worksheets("sheet1").Range("J" & i).NumberFormat = "@"
       Worksheets("sheet1").Range("J" & i).Value = CStr(Mid(hlink, str_start, 18))
       
    End If
'''
'   Worksheets("sheet1").Range("H" & i).Value = rng.Hyperlinks(1).Address
    End If
    Next
End Sub

