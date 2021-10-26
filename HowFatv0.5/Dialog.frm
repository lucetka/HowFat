VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gel/picture type"
   ClientHeight    =   2070
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6030
   Begin VB.OptionButton Option2 
      Caption         =   "White bands on dark bakcground, GRAYSCALE PICTURES ONLY"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Dark (black or color) bands on light background"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Ethidium bromide DNA/RNA gels, zinc stained protein gels "
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "(Coomassie staining, silver staining, western blots...)"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Option1.Value = True
End Sub

Private Sub OKButton_Click()
If Option1.Value = True Then geltype = "DarkBands"
If Option2.Value = True Then geltype = "LightBands"
Unload Me
Form1.Enabled = True
Form1.ZOrder
End Sub

Private Sub Option1_Click()
If Option2 Then Option1 = True
End Sub
Private Sub Option2_Click()
If Option1 Then Option2 = True
End Sub
