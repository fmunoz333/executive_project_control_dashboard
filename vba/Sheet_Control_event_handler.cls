VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim c As Range, bad As Boolean
    Set c = Me.Range("C6")
    If Intersect(Target, c) Is Nothing Then Exit Sub

    On Error GoTo EH
    Application.EnableEvents = False
    bad = Not IsDate(c.Value)
    If Not bad Then bad = (c.Value < Me.Range("C13").Value Or c.Value > Me.Range("C15").Value)
    If bad Then
        MsgBox "The status date must be a valid date between " & _
               Format(Me.Range("C13").Value, "dd-mmm-yyyy") & " and " & _
               Format(Me.Range("C15").Value, "dd-mmm-yyyy") & "." & vbCrLf & _
               "It has been reset to the latest actual date in the data.", _
               vbExclamation, "Terranova Dashboard"
        c.Value = Me.Range("C14").Value
    End If
EH:
    Application.EnableEvents = True
End Sub


