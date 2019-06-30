VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_data2()
Dim ticket_summary As String
Dim total_volume As Double
Dim summary_table As Integer
total_volume = 0
summary_table = 2
For i = 2 To 797711
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticket_summary = Cells(i, 1).Value
   
    ' print these statement in the table
         Range("I" & summary_table).Value = ticket_summary
    'sum of total value
         total_volume = total_volume + Cells(i, 7).Value
     ' print it in the summary
         Range("J" & summary_table).Value = total_volume
    ' add the table
    summary_table = summary_table + 1
    'reset the total
    total_volume = 0
    Else
        total_volume = total_volume + Cells(i, 7).Value
    End If
Next i

   
    
End Sub

