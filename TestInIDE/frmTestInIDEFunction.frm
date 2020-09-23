VERSION 5.00
Begin VB.Form frmTestInIDEFunction 
   Caption         =   "Test InIde function"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test InIDE functions"
      Height          =   495
      Left            =   1755
      TabIndex        =   0
      Top             =   1350
      Width           =   2025
   End
End
Attribute VB_Name = "frmTestInIDEFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub cmdTest_Click()
   Dim lngStart As Long
    Dim i As Long
    Dim blnInIDE As Boolean
    '// Testing InIDE Function in this article
    '// http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=49142
    lngStart = GetTickCount
    For i = 0 To 100000
        blnInIDE = InIDE1
    Next
    MsgBox "InIde Method used in this article " & vbNewLine & _
            "Time taken - " & GetTickCount - lngStart & " ms"
    
    DoEvents
    
    '//Testing InIDE Function using on error resume next
    lngStart = GetTickCount
    For i = 0 To 100000
        blnInIDE = InIDE2
    Next
    MsgBox "InIde Method using on error resume next" & vbNewLine & _
            "Time taken - " & GetTickCount - lngStart & " ms"
    
    DoEvents
    
    '//Testing InIDE Function using App.logmode
    lngStart = GetTickCount
    For i = 0 To 100000
        blnInIDE = InIDE3
    Next
    MsgBox "InIde Method using App.LogMode" & vbNewLine & _
            "Time taken - " & GetTickCount - lngStart & " ms"
End Sub

Private Function InIDE1() As Boolean
    On Error GoTo Xit
    Debug.Print 1 / 0
    Exit Function
Xit:
    InIDE1 = True
End Function

Private Function InIDE2() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    InIDE2 = Err > 0
End Function

Private Function InIDE3() As Boolean
    InIDE3 = (App.LogMode = 0)
End Function
