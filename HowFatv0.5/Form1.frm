VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "HowFat v 0.5"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   482
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6015
      Left            =   8040
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   6000
      Width           =   7815
   End
   Begin VB.PictureBox Picture1 
      Height          =   6375
      Left            =   240
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox Picture2 
         Height          =   6015
         Left            =   120
         ScaleHeight     =   5955
         ScaleWidth      =   7875
         TabIndex        =   3
         Top             =   120
         Width           =   7935
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu mnuOpen 
         Caption         =   "Open picture"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit HowFat"
      End
   End
   Begin VB.Menu mnuMask 
      Caption         =   "Use saved mask"
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run!"
      Begin VB.Menu mnuExportRaw 
         Caption         =   "Export raw band information"
      End
      Begin VB.Menu mnuExportPro 
         Caption         =   "Export processed"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private myfile As String
Public forceexit As Boolean

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private focusrec As RECT
Private X1, Y1
Private X2, Y2



Private Type POINTAPI   'typ pro sledovani mysi
  x As Long
  y As Long
End Type








Private Sub sledujmys()
Dim mousePT As POINTAPI
Dim x As Long, y As Long
Dim stopIt As Boolean
  
  stopIt = False

Do
    If stopIt = True Then Exit Do
        Call GetCursorPos(mousePT)
        x = mousePT.x
        y = mousePT.y
       Picture2.ToolTipText = "x = " & x & ", y = " & y
        DoEvents
Loop
        
        
End Sub




Private Sub Form_Load()

Picture2.ScaleMode = vbPixels
Picture3.ScaleMode = vbPixels

cislobandu = 0

'kod pro ovladani scrollovani
' Set ScaleMode to pixels.
   Form1.ScaleMode = vbPixels
   Picture1.ScaleMode = vbPixels

   ' Autosize is set to True so that the boundaries of
   ' Picture2 are expanded to the size of the actual
   ' bitmap.
   Picture2.AutoSize = True

   ' Set the BorderStyle of each picture box to None.
   Picture1.BorderStyle = 0
   Picture2.BorderStyle = 0

End Sub



Private Sub mnuAbout_Click()
'MsgBox "stupid piece of shitty windows software for YOU from Lucie K."
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuCal_Click()
'idea je vypocitat regresi primku
'je dulezite definovat referencni hladinu bud jako default absolutne bilou,
'nebo uzivatelem definovanou poci set background, nebo dopocitanou jako
'prusecik s osou y
'ale pro zatim bude stacit, kdyz exportuju textak ve formatu
'     ng     I
' a dodelam si to rucne v excelu

End Sub

Private Sub mnuCloneMask_Click()

End Sub

Private Sub mnuExportPro_Click()
MsgBox "Corrected info"
Dim i
Open "corrected.txt" For Output As #1
i = 1
'Do While myBAND(i).name <> ""
For i = 1 To cislobandu
Write #1, "------------------------------------------------------"
    Write #1, myBAND(i).name
    
    If myBAND(i).standard Then
       If myBAND(i).mass = 0 Then
        Write #1, "background" & " " & myBAND(i).mass
       Else
       Write #1, "standard" & " " & myBAND(i).mass
       End If
     Else
     Write #1, "experimental"
     End If
    'Write #1, myBAND(i).standard
    Write #1, myBAND(i).x_1 & " " & myBAND(i).x_2 & " " & myBAND(i).y_1 & " " & myBAND(i).y_2
    Write #1, myBAND(i).integralIntenzity / ((myBAND(i).x_1 - myBAND(i).x_2 - 2) * (myBAND(i).y_1 - myBAND(i).y_2 - 2) * bcgIperPX)
    'i = i + 1
Next i
Close #1
End Sub

Private Sub mnuExportRaw_Click()
Dim typ As String 'experimental, standard, background?


MsgBox "Band info will be in a file of the same name as the picture with .txt extension in the same directory as the picture."
Dim i
Open myfile & ".txt" For Output As #1
'Open "bands.txt" For Output As #1
If geltype = "LightBands" Then
   Print #1, "Bands are considered to be white on black background."
   End If
If geltype = "DarkBands" Then
   Print #1, "Bands are considered to be colored/dark on light background."
   End If
Print #1, "X1, Y1 - coordinates of the upper left corner of the selected area"
Print #1, "X2, Y2 - coordinates of the lower right corner of the selected area"
Print #1, "For opening in Excel say start import at row 5, delimited - semicolon"
Print #1,
Print #1, ";", "Band name", ";", "Type", ";", "Mass(microg)", ";", "X1", ";", "X2", ";", "Y1", ";", "Y2", ";", "Integral intensity"

i = 1
For i = 1 To cislobandu

If myBAND(i).standard Then
       If myBAND(i).mass = 0 Then
        typ = "background"
       Else
       typ = "standard"
       End If
  Else
  typ = "experimental"
End If
    
Print #1, ";", myBAND(i).name, ";", typ, ";", myBAND(i).mass, ";", myBAND(i).x_1, ";", myBAND(i).x_2, ";", myBAND(i).y_1, ";", myBAND(i).y_2, ";", myBAND(i).integralIntenzity

Next i
Close #1

End Sub

Private Sub mnuMask_Click()
'MsgBox maskX1 & maskX2
movingmask = True


Picture2.MousePointer = ccNoDrop


DrawFocusRect Picture2.hdc, focusrec
'sledujmys

'move_and_anchorMask maskX1, maskX2, maskY1, maskY2



End Sub

Private Sub picture2_mousedown(button As Integer, shift As Integer, x As Single, y As Single)
Dim inside_x, inside_y As Boolean
inside_x = False
inside_y = False
'make sure left mouse button is pressed
If (button And vbLeftButton) = 0 Then Exit Sub
'If movingmask Then
'  If x < maskX2 And x > maskX1 Then inside_x = True
'  If y < maskY2 And y > maskY1 Then inside_y = True
'  If inside_x And inside_y Then Picture2.MousePointer = ccSizeAll
'End If

'set upper left corner
  
X1 = x
Y1 = y
End Sub

Private Sub picture2_mousemove(button As Integer, shift As Integer, x As Single, y As Single)
Dim inside_x, inside_y As Boolean
inside_x = False
inside_y = False
Picture2.MousePointer = ccCross
If (button And vbLeftButton) = 0 Then
   ' If movingmask = True Then
   'Picture2.MousePointer = ccNoDrop
   'movemask
   'End If
Exit Sub
End If

 ' If movingmask Then
 '   If x < maskX2 And x > maskX1 Then inside_x = True
 '   If y < maskY2 And y > maskY1 Then inside_y = True
 '   If inside_x And inside_y Then Picture2.ToolTipText = ("x=" & x & ", y=" & y)
 '   If inside_x And inside_y Then Picture2.MousePointer = ccSizeAll
'Exit Sub
'  End If

'make sure left mouse button is pressed
If (button And vbLeftButton) = 0 Then Exit Sub
'delete rectangle in focus, if exists
 If (X2 <> 0) Or (Y2 <> 0) Then
  DrawFocusRect Picture2.hdc, focusrec
 End If
'update coordinates
X2 = x
Y2 = y
'update rectangle
focusrec.left = X1
focusrec.right = X2
focusrec.top = Y1
focusrec.bottom = Y2
'orient rectangle
If Y2 < Y1 Then swap focusrec.top, focusrec.bottom
If X2 < X1 Then swap focusrec.left, focusrec.right
DrawFocusRect Picture2.hdc, focusrec
Refresh
End Sub


Private Sub picture2_mouseup(button As Integer, shift As Integer, x As Single, y As Single)
Dim i, j As Single
Dim pxR, pxG, pxB As Integer
Dim pxGrayScale As Long

Dim intenzitanapixel As Long

intenzita = 0
Dim Ret%
'make sure left mouse button is pressed
If (button And vbLeftButton) = 0 Then Exit Sub
'delete rectangle in focus, if exists
 If focusrec.right Or focusrec.bottom Then
   DrawFocusRect Picture2.hdc, focusrec
End If
'draw the rectangle
'MsgBox X1, Y1
Picture2.Line (X1, Y1)-(X2, Y2), QBColor(12), B
'MsgBox X1, Y1
Picture2.ScaleMode = vbPixels


For j = (Y1 + 1) To (Y2 - 1)
    ReDim Preserve Ipx(j)
    Ipx(j) = intenzita / (X2 - X1 - 2)
intenzita = 0
'pro kazdou radu pixelu v lajne spocitej prumer
  For i = (X1 + 1) To (X2 - 1)
  longcolor = Picture2.Point(i, j)
  pxR = GetRGB(longcolor, 1)
  pxG = GetRGB(longcolor, 2)
  pxB = GetRGB(longcolor, 3)
  pxGrayScale = Round((pxR + pxG + pxB) / 3)
  
  If geltype = "LightBands" Then intenzita = intenzita + 0.00001 * longcolor
  If geltype = "DarkBands" Then intenzita = intenzita + (255 - pxGrayScale)
  

   
  Next i
 
Next j
'cislobandu = cislobandu + 1

myX1 = X1
myX2 = X2
myY1 = Y1
myY2 = Y2

writeProfile
'Load Form2
'Form2.Show

X1 = 0
Y1 = 0
X2 = 0
Y2 = 0

End Sub


Private Sub writeProfile()


MsgBox "Writing to file of the same name as the picture with .txt extension in the same directory as the picture."
Dim i
Open myfile & ".txt" For Output As #1
'Open "bands.txt" For Output As #1
If geltype = "LightBands" Then
   Print #1, "Bands are considered to be white on black background."
   End If
If geltype = "DarkBands" Then
   Print #1, "Bands are considered to be colored/dark on light background."
   End If


For i = (myY1 + 1) To (myY2 - 1)

  
Print #1, Ipx(i)

Next i
Close #1

End Sub




'-----------------------------------------------------------------
'PURPOSE: Returns red/green/blue color from RGB color value.
'
'ACCEPTS: RGB color value as Long, and component number as integer
'         that represents the component color to return (1=red,
'         2=green, 3=blue).
'
'RETURNS: The intensity of the color component (0 - 255) as an
'         integer or -1 indicating that an argument was invalid.
'-----------------------------------------------------------------

Function GetRGB(RGBval As Long, Num As Integer) As Integer
   ' Check if Num, RGBval are valid.
   If Num > 0 And Num < 4 And RGBval > -1 And RGBval < 16777216 Then
     GetRGB = RGBval \ 256 ^ (Num - 1) And 255
   Else
     ' Return True (-1) if Num or RGBval are invalid.
     GetRGB = True
   End If
End Function


Private Sub swap(a, b)
 Dim t
 t = a
 a = b
 b = t
End Sub
Private Sub mnuOpen_Click()
Dim fileJPEG, fileBMP, fileTIFF, allfiles As String

allfiles = "All files (*.*)|*.*"

fileJPEG = "JPG/JPEG (*.JPG)|*.JPG"
fileBMP = "bitmap BMP (*.BMP)|*.BMP"
'fileTIFF = "Tagged image format (*.TIF)|*.TIF"
CommonDialog1.Filter = allfiles & "|" & fileJPEG & "|" & fileBMP
beforeopen:
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub



Picture2.Picture = LoadPicture(CommonDialog1.FileName)

' Initialize location of both pictures.
   Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, _
   ScaleHeight - HScroll1.Height
   Picture2.Move 0, 0

   ' Position the horizontal scroll bar.
   HScroll1.top = Picture1.Height
   HScroll1.left = 0
   HScroll1.Width = Picture1.Width

   ' Position the vertical scroll bar.
   VScroll1.top = 0
   VScroll1.left = Picture1.Width
   VScroll1.Height = Picture1.Height

   ' Set the Max property for the scroll bars.
   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Height

   ' Determine if the child picture will fill up the
   ' screen.
   ' If so, there is no need to use scroll bars.
   VScroll1.Visible = (Picture1.Height < _
   Picture2.Height)
   HScroll1.Visible = (Picture1.Width < _
   Picture2.Width)


Load Dialog
Dialog.Show
Form1.Enabled = False

myfile = CommonDialog1.FileName
cislobandu = 0


End Sub



'The horizontal and vertical scroll bars' Change event is used to move the child picture box up and down or left and right within the parent picture box. Add the following code to the Change event of both scroll bar controls:

Private Sub HScroll1_Change()
   Picture2.left = -HScroll1.Value
End Sub





Private Sub VScroll1_Change()
   Picture2.top = -VScroll1.Value
End Sub

'The Left and Top properties of the child picture box are set to the negative value of the horizontal and vertical scroll bars so that as you scroll up or down or right or left, the display moves appropriately.

'At run time, the graphic will be displayed as shown in Figure 7.26.

'Figure 7.26   Scrolling the bitmap at run time



'Resizing the Form at Run Time
'In the example described above, the viewable size of the graphic is limited by the original size of the form. To resize the graphic viewport application when the user adjusts the size of the form at run time, add the following code to the form's Form_Resize event procedure:

Private Sub Form_Resize()
   ' When the form is resized, change the Picture1
   ' dimensions.
   Picture1.Height = Form1.Height
   Picture1.Width = Form1.Width

   ' Reinitialize the picture and scroll bar
   ' positions.
   Picture1.Move 0, 0, ScaleWidth - VScroll1.Width, _
   ScaleHeight - HScroll1.Height
   Picture2.Move 0, 0
   HScroll1.top = Picture1.Height
   HScroll1.left = 0
   HScroll1.Width = Picture1.Width
   VScroll1.top = 0
   VScroll1.left = Picture1.Width
   VScroll1.Height = Picture1.Height
   HScroll1.Max = Picture2.Width - Picture1.Width
   VScroll1.Max = Picture2.Height - Picture1.Width

   ' Check to see if scroll bars are needed.
   VScroll1.Visible = (Picture1.Height < _
   Picture2.Height)
   HScroll1.Visible = (Picture1.Width < _
   Picture2.Width)

End Sub


Private Sub mnuQuit_Click()
End
End Sub

