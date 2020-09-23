VERSION 5.00
Object = "{3A94456F-33AF-4D40-A77B-936366DCE6FB}#1.0#0"; "AeroSuite.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmimgedit 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crop Image"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9990
   Icon            =   "frmimgedit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   666
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox stdpic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   11160
      Picture         =   "frmimgedit.frx":27A2
      ScaleHeight     =   1680
      ScaleWidth      =   1680
      TabIndex        =   21
      Top             =   3960
      Width           =   1680
   End
   Begin AeroSuite.AeroButton cmd_exit 
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   5640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "Exit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicNormal       =   "frmimgedit.frx":7796
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroButton cmd_save 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicNormal       =   "frmimgedit.frx":9B78
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroTextBox txtpic 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5700
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      BackColor       =   16777215
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
   End
   Begin AeroSuite.AeroButton cmdload 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   3
      PicNormal       =   "frmimgedit.frx":BF5A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroGroupBox frmprepic 
      Height          =   5415
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9551
      BorderColor     =   14277081
      BackColor       =   -2147483633
      BackColor2      =   15263976
      HeadColor1      =   -2147483633
      HeadColor2      =   14869218
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin AeroSuite.AeroGroupBox frmcpi 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   6588
         BorderColor     =   14277081
         BackColor       =   -2147483633
         BackColor2      =   15263976
         HeadColor1      =   -2147483633
         HeadColor2      =   14869218
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Preview"
         Begin VB.PictureBox picSt 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   3360
            Left            =   720
            ScaleHeight     =   3360
            ScaleWidth      =   2760
            TabIndex        =   22
            Top             =   240
            Width           =   2760
         End
      End
      Begin VB.Label moy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label mox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label ory 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label orx 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   3120
         TabIndex        =   15
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label cry 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblcx 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00715240&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Y :"
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   12
         Top             =   4920
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse X :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Original Y :"
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   10
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Original X :"
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   9
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crop Y :"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crop X :"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   4200
         Width           =   615
      End
   End
   Begin AeroSuite.AeroGroupBox frmpic 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   9551
      BorderColor     =   14277081
      BackColor       =   -2147483633
      BackColor2      =   15263976
      HeadColor1      =   -2147483633
      HeadColor2      =   14869218
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.PictureBox picOriginal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         ForeColor       =   &H80000008&
         Height          =   5040
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   336
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   336
         TabIndex        =   19
         Top             =   240
         Width           =   5040
         Begin VB.PictureBox picCrop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            DrawWidth       =   2
            ForeColor       =   &H80000008&
            Height          =   3360
            Left            =   1200
            ScaleHeight     =   224
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   184
            TabIndex        =   20
            Top             =   720
            Width           =   2760
         End
      End
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   10080
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmimgedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartX As Single, StartY As Single
Private m_Selecting As Boolean
Private m_X1 As Single
Private m_Y1 As Single
Private m_X2 As Single
Private m_Y2 As Single

Private Sub cmd_exit_Click()
frmstud.Visible = True
Unload Me
End Sub

Private Sub cmd_save_Click()
    SaveJPG picSt.Image, App.Path & "\pics\" & frmstud.picname.Caption & ".jpg", 100
    frmstud.Visible = True
    Unload Me
End Sub

Private Sub cmdload_Click()
    On Error Resume Next
    cmdlg.Filter = "JPG Image (*.jpg) | *.jpg"
    cmdlg.ShowOpen
    txtpic.Text = cmdlg.filename
    If cmdlg.filename = "" Then
        'NOTHING
    Else
        Call Resize_image(cmdlg.filename, picOriginal, stdpic)
        stdpic.Picture = picOriginal.Image
        picOriginal.Cls
        picOriginal.Picture = stdpic.Image
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmstud.Visible = True
    Unload Me
End Sub

Private Sub picCrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartX = X
    StartY = Y
    picCrop.AutoRedraw = False
End Sub

Private Sub picCrop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim wid As Single
    Dim hig As Single
    
    If Button = 1 Then
        picCrop.Left = IIf(X < StartX, picCrop.Left - (StartX - X), picCrop.Left + (X - StartX))
        picCrop.Top = IIf(Y < StartY, picCrop.Top - (StartY - Y), picCrop.Top + (Y - StartY))
        picCrop.PaintPicture picOriginal.Image, 0, 0, picCrop.Width, picCrop.Height, picCrop.Left, picCrop.Top, picCrop.Width, picCrop.Height
        picCrop.Line (1, 1)-(picCrop.Width - 1, picCrop.Height - 1), , B
        lblcx = picCrop.Left
        cry = picCrop.Top
        m_X1 = picCrop.Left * 15
        m_Y1 = picCrop.Top * 15
        

        wid = picCrop.Width * 15
        hig = picCrop.Height * 15
        picSt.Cls
        picSt.AutoSize = True
        picSt.PaintPicture picOriginal.Picture, 0, 0, wid, hig, m_X1, m_Y1, wid, hig
        picSt.Picture = picSt.Image
    End If
End Sub



Private Sub picOriginal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mox = m_X1
    moy = m_Y1
End Sub

Private Sub picOriginal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mox = m_X1
    moy = m_Y1
End Sub

Private Sub picOriginal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mox = m_X1
    moy = m_Y1
End Sub
