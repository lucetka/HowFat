VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Band description"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      DisabledPicture =   "Form2.frx":0000
      DownPicture     =   "Form2.frx":1C62
      Height          =   375
      Left            =   1440
      Picture         =   "Form2.frx":38C4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "Band 1"
      Top             =   600
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "1000"
      Top             =   2160
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Standard"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Experimental"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Selection - properties:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim toggleButton As Boolean

Private Sub Command1_Click()


standard = Option2.Value
'convert all mass values to microG
Select Case Combo1.Text
    Case "ng"
        mass = CDbl(Text1.Text) / 1000               '*** ADD NON-NUMBER ERROR PROTECTION***
    Case "micro g"
        mass = CDbl(Text1.Text)
    Case "mg"
        mass = CDbl(Text1.Text) * 1000
End Select
If standard = False Then mass = 0

If toggleButton = True Then
maskX1 = myX1
maskX2 = myX2
maskY1 = myY1
maskY2 = myY2
End If


processBand cislobandu, (myX1), (myX2), (myY1), (myY2), intenzita, standard, mass

End Sub

Private Sub Command2_Click()
cislobandu = cislobandu - 1
'Unload Form2
Form2.Hide
 Form1.Enabled = True
 Form1.ZOrder
End Sub

Private Sub Command3_Click()
If toggleButton = True Then
Command3.Picture = Command3.DisabledPicture
toggleButton = False
Exit Sub
End If


If toggleButton = False Then
Command3.Picture = Command3.DownPicture
toggleButton = True
End If
End Sub

Private Sub Form_Load()
toggleButton = False

Form1.Enabled = False
Combo1.AddItem "ng"
Combo1.AddItem "micro g"
Combo1.AddItem "mg"
'Combo1.AddItem "a.u."
Combo1.Text = "ng"
Option1 = True
Combo1.Enabled = False
Text1.Enabled = False
'MsgBox cislobandu
Text2.Text = "Band" & cislobandu
End Sub



Private Sub Option1_Click()
If Option1 Then
Text1.Enabled = False
  Combo1.Enabled = False
 End If
End Sub

Private Sub Option2_Click()
If Option2 Then
  Text1.Enabled = True
  Combo1.Enabled = True
 End If
End Sub

