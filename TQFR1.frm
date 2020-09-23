VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   1575
   ClientTop       =   1470
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   Picture         =   "TQFR1.frx":0000
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   Begin VB.TextBox cmdbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Text            =   "command box"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2520
      Top             =   2160
   End
   Begin MSComDlg.CommonDialog CDbox1 
      Left            =   480
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".map"
      Filter          =   "TQFR map file | *.map"
   End
   Begin VB.Image man 
      Height          =   960
      Left            =   3600
      Picture         =   "TQFR1.frx":63FB
      Top             =   2280
      Width           =   585
   End
   Begin VB.Image cigar_status 
      Height          =   450
      Left            =   3480
      Picture         =   "TQFR1.frx":6C9F
      Top             =   1920
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Shape real_health 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   630
      Left            =   120
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image DPdrop 
      Height          =   330
      Left            =   3600
      Picture         =   "TQFR1.frx":705D
      Top             =   1920
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image object 
      Height          =   840
      Index           =   20
      Left            =   -2520
      Picture         =   "TQFR1.frx":7409
      Top             =   -1200
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   3840
      Picture         =   "TQFR1.frx":7A4C
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape health 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      Top             =   6000
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' make it so that the playa chooses player type
' and a name for the char.
' ******
' make dos-castle style fighting where you run
' into the monster.
' ******
' You can put different things into different
' control arrays, objects are things that the
' man cannot walk thru, items could be things
' he can pick up and use and you can put in some
' NPC's
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim x As Integer 'no of objects()
Dim MAN1progress As Integer
Dim SUCCESS As Long ' used to bitblt
Dim PicTotal As Long ' total amount of pics and (x) of last placed pic
Dim PICx(1 To 1000) As Integer ' y of pic(x)
Dim PICy(1 To 1000) As Integer ' x of pic(x)
Dim oPICx(1 To 1000) As Integer ' original y of pic(x)
Dim oPICy(1 To 1000) As Integer ' original x of pic(x)
Dim PICPic(1 To 1000) As String 'picture of pic(x)
Dim CurrentPic As String 'current picture to place on pics


Private Sub cmdbox_Change()
If cmdbox.Text = "zwee" And real_health.Height < 50 Then
real_health.Height = real_health.Height + 1
real_health.Top = real_health.Top - 1
cmdbox.Text = ""
End If

End Sub

Private Sub cmdbox_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyX Then
    If cmdbox.Visible = True Then
    cmdbox.Visible = False
    cmdbox.Text = ""
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyX Then
    If cmdbox.Visible = False Then
    cmdbox.Visible = True
    cmdbox.Text = ""
    End If
End If

If KeyCode = vbKeyEscape Then End

End Sub

Sub Movement()
Dim i As Long

If GetAsyncKeyState(vbKeyLeft) <> 0 Then
man.Picture = LoadPicture(App.Path & "\man_left.gif")
For i = 1 To PicTotal
PICx(i) = PICx(i) + 10
Next i
Form1.Refresh ' clear the screen to prevent skid marks
RedrawSprites

ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
man.Picture = LoadPicture(App.Path & "\man_right.gif")
For i = 1 To PicTotal
PICx(i) = PICx(i) - 10
Next i
Form1.Refresh ' clear the screen to prevent skid marks
RedrawSprites

ElseIf GetAsyncKeyState(vbKeyUp) <> 0 Then
man.Picture = LoadPicture(App.Path & "\man_up.gif")
For i = 1 To PicTotal
PICy(i) = PICy(i) + 10
Next i
Form1.Refresh ' clear the screen to prevent skid marks
RedrawSprites

ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
man.Picture = LoadPicture(App.Path & "\man_down.gif")
For i = 1 To PicTotal
PICy(i) = PICy(i) - 10
Next i
Form1.Refresh ' clear the screen to prevent skid marks
RedrawSprites
End If
End Sub

Private Sub Form_Load()
Close
On Error GoTo ErrHandler
'CDbox1.DialogTitle = "Select file to be loaded into map editor"
'CDbox1.ShowOpen
Open "C:\Program Files\Microsoft Visual Studio\VB98\My VB\The Quest for Redemption\terrain\map1.map" For Input As #1
Do Until EOF(1)
Input #1, PicDat1, PicDat2, PicDat3, PicDat4, PicDat5
PicTotal = PicTotal + 1
PICx(PicTotal) = PicDat1
PICy(PicTotal) = PicDat2
oPICx(PicTotal) = PicDat3
oPICy(PicTotal) = PicDat4
PICPic(PicTotal) = PicDat5
Loop
Close
Me.Show
DoEvents
RedrawSprites
Exit Sub
ErrHandler:
'hoo bah dug gah
' bup choo wurrywurry
'homboohoombolunbooloomboh
' a-meegoh a-mung-go
Debug.Print Err.Description
End Sub

Private Sub Timer1_Timer()
Movement
End Sub

Sub RedrawSprites()
Dim i ' for the loop
For i = 1 To PicTotal ' cycle through all the sprites
If PICPic(i) = "rock" Then
Form2.Picture1 = Form2.Picture5 'if the pics pic is rock then redraw its pic as a rock
Form2.Picture2 = Form2.Picture6
ElseIf PICPic(i) = "tree" Then
Form2.Picture1 = Form2.Picture3 'if the pics pic is tree then redraw its pic as a tree
Form2.Picture2 = Form2.Picture4
ElseIf PICPic(i) = "start" Then
Form2.Picture1 = Form2.Picture7 'if the pics pic is start then redraw its pic as a tree
Form2.Picture2 = Form2.Picture8
ElseIf PICPic(i) = "siggi" Then
Form2.Picture1 = Form2.Picture9 'if the pics pic is siggi then redraw its pic as a tree
Form2.Picture2 = Form2.Picture10
ElseIf PICPic(i) = "peedtree" Then
Form2.Picture1 = Form2.Picture11 'set image to the picture of the tree.
Form2.Picture2 = Form2.Picture12 'set mask
ElseIf PICPic(i) = "cigar" Then
Form2.Picture1 = Form2.Picture13 'set image to the picture of the tree.
Form2.Picture2 = Form2.Picture14 'set mask
End If

SUCCESS = BitBlt(Form1.hDC, PICx(i), PICy(i), _
Form2.Picture1.ScaleWidth, Form2.Picture1.ScaleHeight, _
Form2.Picture1.hDC, 0, 0, SRCAND) 'paints the pic

SUCCESS = BitBlt(Form1.hDC, PICx(i), PICy(i), _
Form2.Picture1.ScaleWidth, Form2.Picture1.ScaleHeight, _
Form2.Picture2.hDC, 0, 0, SRCPAINT) 'paints the pic
Next i 'cycle

End Sub

