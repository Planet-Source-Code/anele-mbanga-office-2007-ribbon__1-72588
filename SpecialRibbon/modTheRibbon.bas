Attribute VB_Name = "modTheRibbon"
Option Explicit
' these variables are used to store menu details
' do not delete
Public menuCnt As Long
Public menuKeys() As String
Public Sub Main()
    On Error Resume Next
    Load Form1
    Form1.Show
    Form1.Command1_Click
    Err.Clear
End Sub
