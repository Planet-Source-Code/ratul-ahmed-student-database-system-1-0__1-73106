VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash10c.ocx"
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Student Database"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   770
   ShowInTaskbar   =   0   'False
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash mflash 
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      Height          =   4455
      Left            =   3360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1920
      Width           =   7095
      _cx             =   12515
      _cy             =   7858
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin Student_database.GUI_Rollover cmd_settings_but 
      Height          =   885
      Left            =   600
      TabIndex        =   5
      Top             =   5640
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1561
      Enabled         =   0   'False
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":27A2
      ImageHover      =   "frmmain.frx":3432
      ImageDown       =   "frmmain.frx":4126
      ImageDisabled   =   "frmmain.frx":4DB6
      ImageMask       =   "frmmain.frx":5A46
      ImageSelected   =   "frmmain.frx":66D6
      ImageSelectedHover=   "frmmain.frx":7366
   End
   Begin Student_database.GUI_Rollover cmd_search_but 
      Height          =   870
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1535
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":7FF6
      ImageHover      =   "frmmain.frx":8CB2
      ImageDown       =   "frmmain.frx":99D0
      ImageDisabled   =   "frmmain.frx":A68C
      ImageMask       =   "frmmain.frx":B348
      ImageSelected   =   "frmmain.frx":C004
      ImageSelectedHover=   "frmmain.frx":CCC0
   End
   Begin Student_database.GUI_Rollover cmd_restult_but 
      Height          =   825
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1455
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":D97C
      ImageHover      =   "frmmain.frx":EEB0
      ImageDown       =   "frmmain.frx":1046B
      ImageDisabled   =   "frmmain.frx":1199F
      ImageMask       =   "frmmain.frx":12ED3
      ImageSelected   =   "frmmain.frx":14407
      ImageSelectedHover=   "frmmain.frx":1593B
   End
   Begin Student_database.GUI_Rollover cmd_student_but 
      Height          =   840
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1482
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":16E6F
      ImageHover      =   "frmmain.frx":17F63
      ImageDown       =   "frmmain.frx":190B8
      ImageDisabled   =   "frmmain.frx":1A1AC
      ImageMask       =   "frmmain.frx":1B2A0
      ImageSelected   =   "frmmain.frx":1C394
      ImageSelectedHover=   "frmmain.frx":1D488
   End
   Begin Student_database.GUI_Rollover cmd_about 
      Height          =   405
      Left            =   10350
      TabIndex        =   1
      Top             =   510
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":1E57C
      ImageHover      =   "frmmain.frx":1EB65
      ImageDown       =   "frmmain.frx":1F695
      ImageDisabled   =   "frmmain.frx":1FC7E
      ImageMask       =   "frmmain.frx":20267
      ImageSelected   =   "frmmain.frx":20850
      ImageSelectedHover=   "frmmain.frx":20E39
   End
   Begin Student_database.GUI_Rollover cmd_exit 
      Height          =   405
      Left            =   10350
      TabIndex        =   0
      Top             =   60
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   714
      Selectable      =   0   'False
      ImageNormal     =   "frmmain.frx":21422
      ImageHover      =   "frmmain.frx":21948
      ImageDown       =   "frmmain.frx":2236D
      ImageDisabled   =   "frmmain.frx":22893
      ImageMask       =   "frmmain.frx":22DB9
      ImageSelected   =   "frmmain.frx":232DF
      ImageSelectedHover=   "frmmain.frx":23805
   End
   Begin VB.Label lblcap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Student Database System 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00715240&
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GWL_EXSTYLE As Long = -20
Private Const LWA_COLORKEY As Long = &H1
Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Dim FileRes As Integer
Dim Buffer() As Byte
Sub Center(FormName As Form)
Move (Screen.Width - FormName.Width) \ 2, (Screen.Height - FormName.Height) \ 2
End Sub
Sub MoveForm(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
End Sub



Private Sub cmd_about_OnMouseClick()
aerofrm.Visible = False
aerofrm.tcount.Enabled = False
frmmain.Visible = False
frmabut.Visible = True
End Sub

Private Sub cmd_exit_OnMouseClick()
Unload aerofrm
End Sub

Private Sub cmd_restult_but_OnMouseClick()
aerofrm.Visible = False
aerofrm.tcount.Enabled = False
frmmain.Visible = False
frmsres.Visible = True
End Sub

Private Sub cmd_search_but_OnMouseClick()
Dim sStr As String
sStr = InputBox("Enter Student ID :", "ID")
frmstud.txtsearch.Text = sStr
aerofrm.Visible = False
aerofrm.tcount.Enabled = False
frmmain.Visible = False
frmstud.Visible = True

End Sub

Private Sub cmd_student_but_OnMouseClick()

aerofrm.Visible = False
aerofrm.tcount.Enabled = False
frmmain.Visible = False
frmstud.Visible = True

End Sub

Private Sub Form_Load()
    Dim tmpSt As Long
    frmmain.Height = 7680
    frmmain.Width = 11550
    
    Call Center(frmmain)
    tmpSt = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    tmpSt = tmpSt Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, tmpSt
    SetLayeredWindowAttributes Me.hWnd, RGB(255, 0, 0), 0, LWA_COLORKEY
    Load aerofrm
    aerofrm.Show
    
    On Error Resume Next
    mflash.WMode = 1
    mflash.LoadMovie 0, App.Path & "\gui\main_anim.swf"
End Sub



