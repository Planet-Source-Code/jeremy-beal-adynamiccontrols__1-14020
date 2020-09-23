VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//Using WithEvents allows all events to occur within the textbox and
'//are passed thru this class module.//

Public WithEvents weTextBox As TextBox
Attribute weTextBox.VB_VarHelpID = -1

'//Copy the Control's Event(s) and Paste here. Rename the Control to the
'//Withevents variable. Such as Text1_change must be:  weTextBox_Change//
'//You can actually display or change the Properties of the TextBox such as weTextBox.Name//

Private Sub weTextBox_Change()
  MsgBox "Change Made on TextBox: " & weTextBox.Name
End Sub

Private Sub weTextBox_Click()
'//display Left and Top Properties of TextBox using this Event//

 MsgBox "Clicked on TextBox: " & weTextBox.Name & _
            " Left: " & weTextBox.Left & " Right: " & weTextBox.Top
End Sub

Private Sub weTextBox_DblClick()
'//Add Text to TextBox//
  weTextBox.Text = weTextBox.Text & "W"
  MsgBox "Double Clicked on TextBox: " & weTextBox.Name
End Sub

Private Sub weTextBox_GotFocus()
  MsgBox "Got Focus on TextBox: " & weTextBox.Name
End Sub

Private Sub weTextBox_KeyPress(KeyAscii As Integer)
  MsgBox "KeyPress: " & KeyAscii & " on TextBox: " & weTextBox.Name
End Sub

