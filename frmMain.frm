VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dynamic Controls"
   ClientHeight    =   3030
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "TextBox Control Array w/ events"
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton cmdRemoveControl 
      Caption         =   "Remove Control"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Non-Intrinsic Control w/ events"
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "TextBox VBControl Object w/ events"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Textbox VBControl Ojbect w/o Events"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton cmdChoice 
      Caption         =   "Textbox w/o Events"
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------READ HERE FIRST-------------------
'//This Example Shows How to Dynamically add Controls in Various Ways.
'//Please Read under cmdChoice_Click event to learn about the various assignment.
'//Debug the program and step-thru to see what is really happening.
'//Written by: Jeremy Beal  1/1/2001  - My First Source Code

'//Please RATE this at www.pscode.com if you like it. The more rates the more code i write.//
'--------------------------------------------

Public clsTextBox As New clsTextEvents '//This looks at the Class object clsTextEvents
Attribute clsTextBox.VB_VarHelpID = -1
Public ctrlTextBox As TextBox
Public WithEvents ctrlDynamic As VBControlExtender
Attribute ctrlDynamic.VB_VarHelpID = -1
Public WithEvents objTextBox As TextBox
Attribute objTextBox.VB_VarHelpID = -1

Private Sub cmdChoice_Click(Index As Integer)
'//Example of Which Command Button Selected From CommandButton Array//

'//Be Sure to have RemoveControl Button selected and remove added control
If cmdRemoveControl.Visible = True Then
  MsgBox "Please Click Remove Control Button!"
  Exit Sub
End If
  
'//Checks which Command Button index Selected//

Select Case Index
  
  Case 0
  '//Create TextBox On Form as is without events associated to TextBox//
  
    frmMain.Controls.Add "VB.TextBox", "Text2"
    '//Center TextBox//
    '//Lets put it to the corner for now
    frmMain.Controls("Text2").Left = 0 '(frmMain.ScaleWidth / 2)
    frmMain.Controls("Text2").Top = 0 '(frmMain.ScaleHeight / 2) - (frmMain.Controls("Text2").Height / 2)
    '//Display TextBox//
    frmMain.Controls("Text2").Visible = True
    
  Case 1
    '//Create TextBox on Form and Assign to ctrlTextBox without events associated to Textbox//
    '//Simplier to Modify Control without using frmMain.Controls("")//
    
    Set ctrlTextBox = frmMain.Controls.Add("VB.TextBox", "Text2")
    '//Center Textbox//
    '//Lets put it to the corner for now
    ctrlTextBox.Left = 0 '(frmMain.ScaleWidth / 2)
    ctrlTextBox.Top = 0 '(frmMain.ScaleHeight / 2) - (ctrlTextBox.Height / 2)
    '//Display TextBox//
    ctrlTextBox.Visible = True
    
    Set ctrlTextBox = Nothing '//We are done setting New TextBox
    
    '//If wanting to Keep ctrlTextBox active thru-out program, you can.
    '//Only One ctrlTextBox variable is assigned to each TextBox control created
    '//If you have 50 new TextBox then ctrTextBox1, ctrlTextBox2.... each
    '//variable is assigned to each new TextBox control
  
  Case 2
    '//Create TextBox on Form and Assign to objTextBox with events associated to it.
    '//One WithEvent variable per Control... ie: objTextBox1, objTextBox2.... for each control//
    '//However you must have each WithEvent Variable its own Event.. ie: objTextBox1_Click, ObjTextBox2_Click...
    '//This needs to call all events that is being used.//
    '//Naming the variable is your choice.. not limited to objTextBox...//
    
    Set objTextBox = frmMain.Controls.Add("VB.TextBox", "Text2")
    '//Center Textbox//
    '//Lets put it to the corner for now
    objTextBox.Left = 0 '(frmMain.ScaleWidth / 2)
    objTextBox.Top = 0 '(frmMain.ScaleHeight / 2) - (objTextBox.Height / 2)
    '//Display TextBox//
    objTextBox.Visible = True
    
    '//Set objTextBox = nothing when unloading Form//
    
  Case 3
    '//This Example cannot use intrinsic controls ie: TextBox, commandButton...etc....
    '//Create non-intrinsic controls on Form and Assign to CtrlDynamic with events associated to it.
    '//One CtrlDynamic variable per control... ie: ctrlDynamic1, ctrlDynamic2.... for each control.//
    '//However you must have each ctrlDynamic Variable its own event..
    '//ie: ctrlDynamic1_ObjectEvent(info as EventInfo), ctrlDynamic2_ObjectEvent(Info as EventInfo)....
    '//This only needs one event for many events to call.//
    '//Naming the variable is your choice.. not limited to ctrlDynamic...//
    Licenses.Add "MSComctlLib.TreeCtrl" '//Add if Component is not added to Project/component
    Set ctrlDynamic = frmMain.Controls.Add("MSComctlLib.TreeCtrl", "TreeView1")
    '//Center Textbox//
    '//Lets put it to the corner for now
    ctrlDynamic.Left = 0 '(frmMain.ScaleWidth / 2)
    ctrlDynamic.Top = 0 '(frmMain.ScaleHeight / 2) - (ctrlDynamic.Height / 2)
    '//Display TextBox//
    ctrlDynamic.Visible = True
    '//Set ctrDynamic = nothing when unloading form//
    
  Case 4
    '//Create Textbox on Form and assign to clsTextbox with events associated to it in
    '//class module. One clsTextbox variable per control... ie: clsTextBox1, clsTextBox2... for each control.//
    '//All the clsTextBox variable only need ONE Class Module and events within it. Do not need
    '//to create events for each variable...(consider as Subclassing)//
    '//Naming the variable is your choice.. not limited to clsTextBox...//
    
    Set ctrlTextBox = frmMain.Controls.Add("VB.TextBox", "Text2")
    '//If there is a TextBox already on form you can assign that Control to
    '//clsTextBox: Set clsTextBox = txtText1  and use the events for that control.//
    '//Center Textbox//
    '//Lets put it to the corner for now
    ctrlTextBox.Left = 0 '(frmMain.ScaleWidth / 2)
    ctrlTextBox.Top = 0 '(frmMain.ScaleHeight / 2) - (ctrlTextBox.Height / 2)
    '//Display TextBox//
    ctrlTextBox.Visible = True
    
    Set clsTextBox.weTextBox = ctrlTextBox '//Important if wanting to use Class Events
    '//OR set clstextbox.weTextBox = frmMain.controls("Text2")
    
  Case 5
    '//Create TextBox on form from TextBox Array//
    '//TextBox Control Arrays use one TextBox events using Index in the Parameter
    
    Load Text1(1)
    '//Center TextBox//
    '//Lets put it to the corner for now
    Text1(1).Left = 0 '(frmMain.ScaleWidth / 2)
    Text1(1).Top = 0 '(frmMain.ScaleHeight / 2) - (Text1(1).Height / 2)
    '//Display Textbox//
    Text1(1).Visible = True
    
    '//Or
    'frmMain.Controls.Add "Text1", 1   '// 1 indicates index number of Text1 control
    '' you can assign ctrlTextBox = frmMain.controls.... to simplify coding
    'frmMain.Controls("Text1", 1).Left = (frmMain.ScaleWidth / 2) - (frmMain.Controls("Text1", 1).Width / 2)
    'frmMain.Controls("Text1", 1).Top = (frmMain.ScaleWidth / 2) - (frmMain.Controls("Text1", 1).Height / 2)
    'frmMain.Controls("Text1", 1).Visible = True
    
End Select

'//Show CmdRemoveControl to allow removal of TextBox Control//
  cmdRemoveControl.Tag = Index
  cmdRemoveControl.Visible = True
  
End Sub

Private Sub cmdRemoveControl_Click()
'//Remove New Controls//

Select Case cmdRemoveControl.Tag
 
  Case 0
    frmMain.Controls.Remove "Text2"
  Case 1
    frmMain.Controls.Remove "Text2"
  Case 2
    frmMain.Controls.Remove "Text2"
    Set objTextBox = Nothing
  Case 3
    frmMain.Controls.Remove "TreeView1"
    Licenses.Remove "MsComctlLib.TreeCtrl"
    Set ctrlDynamic = Nothing
  Case 4
    frmMain.Controls.Remove "Text2"
    Set clsTextBox = Nothing
  Case 5
    Unload Text1(1)
    '//Can use  frmMain.Controls.Remove "Text2",1
End Select
  
  '//Reset cmdRemoveControl//
  cmdRemoveControl.Tag = ""
  cmdRemoveControl.Visible = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '//Make sure controls are removed before exiting
  If cmdRemoveControl.Visible = True Then
    Call cmdRemoveControl_Click
  End If
  
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub objTextBox_Click()
  MsgBox "I am Clicked! Using WithEvents.... YEAH!"
End Sub

Private Sub objTextBox_DblClick()
  MsgBox "I am Double Clicked! Using WithEvents.... YEAH!"
End Sub

Private Sub ctrlDynamic_ObjectEvent(Info As EventInfo)
  '//Test for the Click Event of the Control Clicked//
  If Info.Name = "Click" Then
    MsgBox "You Clicked on: " & ctrlDynamic.Name
  Else
    '//Include as Many Events as you need. Remove Comment from next line to see what it does.//
    'MsgBox "Event activated: " & Info.Name
  End If
End Sub

Private Sub Text1_Click(Index As Integer)
  MsgBox "I am Clicked! TextBox: " & Text1(Index).Name
End Sub

Private Sub Text1_DblClick(Index As Integer)
  MsgBox "I am Double Clicked! TextBox: " & Text1(Index).Name
End Sub
