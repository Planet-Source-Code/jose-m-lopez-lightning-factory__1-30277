VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmScreen"
   ClientHeight    =   6795
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
frmControls.txtParameters(0).Text = 250 'ray x start
frmControls.txtParameters(1).Text = 150 'ray y start
frmControls.txtParameters(2).Text = 0 'lower angle
frmControls.txtParameters(3).Text = 45 'higher angle

frmControls.txtParameters(4).Text = 5 'QtyRaysPerDisplay
frmControls.txtParameters(5).Text = 50 'QtySegments
frmControls.txtParameters(6).Text = 5 'How many fade steps, 0 and negative = no fade and no clear
frmControls.txtParameters(7).Text = 100 'DesiredRadius

frmControls.txtParameters(8).Text = 10 'x rnd
frmControls.txtParameters(9).Text = 10 'y rnd
frmControls.txtParameters(10).Text = 5 'x rnd/2
frmControls.txtParameters(11).Text = 5 'y rnd/2
frmControls.txtParameters(12).Text = 1 'drawwidth

frmControls.txtParameters(13).Text = 255 'R inner color
frmControls.txtParameters(14).Text = 255 'G inner color
frmControls.txtParameters(15).Text = 0 'B inner color
frmControls.txtParameters(16).Text = 255 'R outer color
frmControls.txtParameters(17).Text = 255 'G outer color
frmControls.txtParameters(18).Text = 255 'B outer color

frmControls.txtParameters(19).Text = 1 'quantity displays
frmControls.txtParameters(20).Text = 0 'delay between rays in milliseconds
frmControls.txtParameters(21).Text = 0 'delay between segments in milliseconds
frmControls.txtParameters(22).Text = 0 'delay between displays in milliseconds

frmControls.optAngle(1).Value = True '"Pie Chart" type ray angling on
frmControls.txtTitle.Text = "Design 1" 'Title

frmControls.optHighLightExe(0) = True 'Highlight playlist line by line execution

frmScreen.Show
frmControls.Show
booStop = False
booStopped = True
strFileName = ""
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If booStopped = False Then booStop = True
frmControls.Show
End Sub

Private Sub Form_DblClick()
Unload frmControls
Unload frmScreen
End
End Sub


