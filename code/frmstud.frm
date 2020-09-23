VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3A94456F-33AF-4D40-A77B-936366DCE6FB}#1.0#0"; "AeroSuite.ocx"
Begin VB.Form frmstud 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11250
   Icon            =   "frmstud.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      Height          =   1680
      Left            =   11520
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   44
      Top             =   240
      Width           =   1800
   End
   Begin AeroSuite.AeroButton cmd_del 
      Height          =   495
      Left            =   10080
      TabIndex        =   7
      Top             =   7725
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "Delete"
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
      PicHot          =   "frmstud.frx":27A2
      PicNormal       =   "frmstud.frx":4B84
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroButton cmd_edit 
      Height          =   495
      Left            =   9000
      TabIndex        =   6
      Top             =   7725
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "Edit"
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
      PicHot          =   "frmstud.frx":6F66
      PicNormal       =   "frmstud.frx":9348
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroButton cmd_add 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   7725
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "Add"
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
      PicHot          =   "frmstud.frx":B72A
      PicNormal       =   "frmstud.frx":DB0C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroGroupBox frm_search 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   7560
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      BorderColor     =   14408667
      BackColor       =   15790320
      BackColor2      =   15395562
      HeadColor1      =   15790320
      HeadColor2      =   15000804
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Search"
      Begin AeroSuite.AeroButton cmd_search 
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicHot          =   "frmstud.frx":FEEE
         PicNormal       =   "frmstud.frx":122D0
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin AeroSuite.AeroTextBox txtsearch 
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   300
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         BackColor       =   16777215
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID :"
         ForeColor       =   &H00715240&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   300
         Width           =   375
      End
   End
   Begin AeroSuite.AeroButton cmd_print 
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   7725
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Print"
      Enabled         =   0   'False
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
      PicDown         =   "frmstud.frx":146B2
      PicHot          =   "frmstud.frx":16A94
      PicNormal       =   "frmstud.frx":18E76
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
      State           =   3
   End
   Begin AeroSuite.AeroButton cmd_main 
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   7725
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "Main Menu"
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
      PicDown         =   "frmstud.frx":1B258
      PicHot          =   "frmstud.frx":1D63A
      PicNormal       =   "frmstud.frx":1FA1C
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroGroupBox frmdb 
      Height          =   7335
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12938
      BorderColor     =   14408667
      BackColor       =   15790320
      BackColor2      =   15395562
      HeadColor1      =   15790320
      HeadColor2      =   15000804
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin AeroSuite.AeroProgressBar pb 
         Height          =   270
         Left            =   120
         Top             =   6960
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   476
      End
      Begin MSComctlLib.ListView lststd 
         Height          =   6735
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   11880
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin AeroSuite.AeroGroupBox frinfo 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12938
      BorderColor     =   14408667
      BackColor       =   15790320
      BackColor2      =   15395562
      HeadColor1      =   15790320
      HeadColor2      =   15000804
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin AeroSuite.AeroButton cmd_upload 
         Height          =   375
         Left            =   2160
         TabIndex        =   46
         Top             =   1700
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Upload Photo"
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
         PicHot          =   "frmstud.frx":21DFE
         PicNormal       =   "frmstud.frx":241E0
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin AeroSuite.AeroGroupBox frmstd_i 
         Height          =   2655
         Left            =   120
         TabIndex        =   31
         Top             =   4560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4683
         BorderColor     =   14408667
         BackColor       =   15790320
         BackColor2      =   15395562
         HeadColor1      =   15790320
         HeadColor2      =   15000804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin AeroSuite.AeroTextBox txtlastcirtificate 
            Height          =   255
            Left            =   1560
            TabIndex        =   43
            Top             =   2160
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtlastpayment 
            Height          =   255
            Left            =   1560
            TabIndex        =   42
            Top             =   1800
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtpresents 
            Height          =   255
            Left            =   1560
            TabIndex        =   41
            Top             =   1440
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtgrade 
            Height          =   255
            Left            =   1560
            TabIndex        =   40
            Top             =   1080
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
            Enabled         =   0   'False
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
         Begin AeroSuite.AeroTextBox txtsem 
            Height          =   255
            Left            =   1560
            TabIndex        =   39
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtdep 
            Height          =   255
            Left            =   1560
            TabIndex        =   38
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Last Certificate :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   37
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Last Payment :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   36
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Grade :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presents :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Semester  :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Department  :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
      End
      Begin AeroSuite.AeroGroupBox frm_root 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4260
         BorderColor     =   14408667
         BackColor       =   15790320
         BackColor2      =   15395562
         HeadColor1      =   15790320
         HeadColor2      =   15000804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin AeroSuite.AeroTextBox txtperadd 
            Height          =   255
            Left            =   1560
            TabIndex        =   30
            Top             =   2040
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtpresentadd 
            Height          =   255
            Left            =   1560
            TabIndex        =   29
            Top             =   1680
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtbirth 
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtmname 
            Height          =   255
            Left            =   1560
            TabIndex        =   27
            Top             =   960
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtfname 
            Height          =   255
            Left            =   1560
            TabIndex        =   26
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin AeroSuite.AeroTextBox txtname 
            Height          =   255
            Left            =   1560
            TabIndex        =   25
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16777215
            ForeColor       =   7426624
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
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Parmanent Add. :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   24
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Present Add. :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   22
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mothers Name :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   21
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fathers Name :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin AeroSuite.AeroOptionButton opt_female 
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         Align           =   0
         Caption         =   "Female"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   15790320
         ForeColor       =   7426624
         MousePointer    =   0
         MouseIcon       =   "frmstud.frx":265C2
         Value           =   0   'False
      End
      Begin AeroSuite.AeroOptionButton opt_male 
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         Align           =   0
         Caption         =   "Male"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocusRect   =   -1  'True
         BackColor       =   15790320
         ForeColor       =   7426624
         MousePointer    =   0
         MouseIcon       =   "frmstud.frx":265DE
         Value           =   0   'False
      End
      Begin AeroSuite.AeroTextBox txtroll 
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         ForeColor       =   7426624
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
      Begin AeroSuite.AeroTextBox txtid 
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BackColor       =   16777215
         ForeColor       =   7426624
         Enabled         =   0   'False
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
      Begin AeroSuite.AeroGroupBox frmstdimg 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   3625
         BorderColor     =   14408667
         BackColor       =   15790320
         BackColor2      =   15395562
         HeadColor1      =   15790320
         HeadColor2      =   15000804
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.PictureBox stdpic 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1680
            Left            =   200
            Picture         =   "frmstud.frx":265FA
            ScaleHeight     =   1680
            ScaleWidth      =   1500
            TabIndex        =   45
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label picname 
            Caption         =   "Label1"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Label lblRoll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Roll :"
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
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblID 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID :"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmstud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_add_Click()

    Call Add_data


End Sub

Private Sub cmd_del_Click()
Delete_data
End Sub

Private Sub cmd_edit_Click()
Call Edit_data
End Sub

Private Sub cmd_main_Click()
aerofrm.Visible = True
aerofrm.tcount.Enabled = True
frmmain.Visible = True
frmstud.Visible = False
Call DisconnectDB
Unload frmstud
End Sub

Private Sub cmd_search_Click()
Call search(txtsearch.Text, lststd)
Call lststd_Click
End Sub

Private Sub cmd_upload_Click()
    frmimgedit.Visible = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Height = 8850
Me.Width = 11250
Call lstConfig(lststd)
Call connectDB
Call LoadSTD_data(lststd)
Call lststd_Click


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
aerofrm.Visible = True
aerofrm.tcount.Enabled = True
frmmain.Visible = True
frmstud.Visible = False
Call DisconnectDB
Unload frmstud
End Sub

Private Sub lststd_Click()
Dim picn As String
Dim fex As Integer

    frmstud.txtname.Refresh
    frmstud.txtdep.Refresh
    frmstud.txtsem.Refresh
    frmstud.txtroll.Refresh
    frmstud.txtpresents.Refresh
    frmstud.txtlastpayment.Refresh
    frmstud.txtlastcirtificate.Refresh
    frmstud.txtfname.Refresh
    frmstud.txtmname.Refresh
    frmstud.txtbirth.Refresh
    frmstud.txtpresentadd.Refresh
    frmstud.txtperadd.Refresh


Call change_Objects(lststd)
picn = App.Path & "\pics\" & picname.Caption & ".jpg"
fex = FileExists(picn)
If fex = "-1" Then
    Call Streach_pic(picn, picOriginal, stdpic, 130, 110, 184, 224)
Else
    'NOTHING
End If

End Sub


