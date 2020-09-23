VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3A94456F-33AF-4D40-A77B-936366DCE6FB}#1.0#0"; "AeroSuite.ocx"
Begin VB.Form frmsres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Result"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11280
   Icon            =   "frmsres.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtscore 
      Height          =   375
      Left            =   11520
      TabIndex        =   48
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      Height          =   1680
      Left            =   11520
      ScaleHeight     =   108
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   116
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
   Begin AeroSuite.AeroButton cmd_reset 
      Height          =   495
      Left            =   10080
      TabIndex        =   1
      Top             =   7605
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Caption         =   "Reset"
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
      PicDown         =   "frmsres.frx":27A2
      PicHot          =   "frmsres.frx":4B84
      PicNormal       =   "frmsres.frx":6F66
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroButton cmd_edit 
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   7605
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
      PicHot          =   "frmsres.frx":9348
      PicNormal       =   "frmsres.frx":B72A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroGroupBox frm_search 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   7440
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
         TabIndex        =   4
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
         PicHot          =   "frmsres.frx":DB0C
         PicNormal       =   "frmsres.frx":FEEE
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin AeroSuite.AeroTextBox txtsearch 
         Height          =   255
         Left            =   600
         TabIndex        =   5
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
         Caption         =   "ID :"
         ForeColor       =   &H00715240&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   375
      End
   End
   Begin AeroSuite.AeroButton cmd_print 
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   7605
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
      PicDown         =   "frmsres.frx":122D0
      PicHot          =   "frmsres.frx":146B2
      PicNormal       =   "frmsres.frx":16A94
      PicSize         =   1
      PicSizeH        =   16
      PicSizeW        =   16
      State           =   3
   End
   Begin AeroSuite.AeroButton cmd_main 
      Height          =   495
      Left            =   4080
      TabIndex        =   8
      Top             =   7605
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
      PicDown         =   "frmsres.frx":18E76
      PicHot          =   "frmsres.frx":1B258
      PicNormal       =   "frmsres.frx":1D63A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin AeroSuite.AeroGroupBox frmdb 
      Height          =   7335
      Left            =   4080
      TabIndex        =   9
      Top             =   0
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
         LabelWrap       =   0   'False
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
      TabIndex        =   11
      Top             =   0
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
         TabIndex        =   12
         Top             =   1700
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Upload Photo"
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
         PicHot          =   "frmsres.frx":1FA1C
         PicNormal       =   "frmsres.frx":21DFE
         PicSizeH        =   16
         PicSizeW        =   16
         State           =   3
      End
      Begin AeroSuite.AeroGroupBox frmstd_i 
         Height          =   2055
         Left            =   120
         TabIndex        =   13
         Top             =   5160
         Width           =   3615
         _ExtentX        =   6376
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
         Begin AeroSuite.AeroTextBox txtcgpa 
            Height          =   255
            Left            =   1560
            TabIndex        =   50
            Top             =   1680
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16761024
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
         Begin AeroSuite.AeroTextBox txtgrade 
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            BackColor       =   16761024
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
         Begin AeroSuite.AeroTextBox txtsem 
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   960
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
         Begin AeroSuite.AeroTextBox txtdep 
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   600
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
         Begin AeroSuite.AeroTextBox txtname 
            Height          =   255
            Left            =   1560
            TabIndex        =   38
            Top             =   240
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
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "cGPA :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   39
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Department  :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   600
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
            TabIndex        =   18
            Top             =   960
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
            TabIndex        =   17
            Top             =   1320
            Width           =   1335
         End
      End
      Begin AeroSuite.AeroGroupBox frm_root 
         Height          =   3135
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   5530
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
         Begin AeroSuite.AeroTextBox txtbizorg 
            Height          =   255
            Left            =   1560
            TabIndex        =   47
            Top             =   2760
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
         Begin AeroSuite.AeroTextBox txtbookkeep 
            Height          =   255
            Left            =   1560
            TabIndex        =   46
            Top             =   2400
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
         Begin AeroSuite.AeroTextBox txtenviron 
            Height          =   255
            Left            =   1560
            TabIndex        =   45
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
         Begin AeroSuite.AeroTextBox txtdatacomfund 
            Height          =   255
            Left            =   1560
            TabIndex        =   44
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
         Begin AeroSuite.AeroTextBox txtvisual 
            Height          =   255
            Left            =   1560
            TabIndex        =   21
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
         Begin AeroSuite.AeroTextBox txtmicro 
            Height          =   255
            Left            =   1560
            TabIndex        =   22
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
         Begin AeroSuite.AeroTextBox txtcomarch 
            Height          =   255
            Left            =   1560
            TabIndex        =   23
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
         Begin AeroSuite.AeroTextBox txtdatabaseman 
            Height          =   255
            Left            =   1560
            TabIndex        =   24
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
            Caption         =   "Bisiness Org. :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   43
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Book Keeping :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   42
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Environmental :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   41
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Data Comm. Fund."
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Database Man. :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Computer Arc. :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   27
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Microprocessor :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Visual Program :"
            ForeColor       =   &H00715240&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
         End
      End
      Begin AeroSuite.AeroOptionButton opt_female 
         Height          =   255
         Left            =   2880
         TabIndex        =   29
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
         MouseIcon       =   "frmsres.frx":241E0
         Value           =   0   'False
      End
      Begin AeroSuite.AeroOptionButton opt_male 
         Height          =   255
         Left            =   2160
         TabIndex        =   30
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
         MouseIcon       =   "frmsres.frx":241FC
         Value           =   0   'False
      End
      Begin AeroSuite.AeroTextBox txtroll 
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   840
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
      Begin AeroSuite.AeroTextBox txtid 
         Height          =   255
         Left            =   2760
         TabIndex        =   32
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
         TabIndex        =   33
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
            Picture         =   "frmsres.frx":24218
            ScaleHeight     =   1680
            ScaleWidth      =   1500
            TabIndex        =   34
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label picname 
            Caption         =   "Label1"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
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
         TabIndex        =   37
         Top             =   360
         Width           =   615
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
         TabIndex        =   36
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmsres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fres As Boolean


Private Sub cmd_edit_Click()
Call GenResult
Call Edit_data_Result
Call GenResult
End Sub

Private Sub cmd_main_Click()
aerofrm.Visible = True
aerofrm.tcount.Enabled = True
frmmain.Visible = True
frmsres.Visible = False
Call DisconnectDB
Unload frmsres
End Sub

Private Sub cmd_reset_Click()
frmsres.txtcomarch.Text = ""
frmsres.txtmicro.Text = ""
frmsres.txtdatabaseman.Text = ""
frmsres.txtvisual.Text = ""
frmsres.txtdatacomfund.Text = ""
frmsres.txtenviron.Text = ""
frmsres.txtbookkeep.Text = ""
frmsres.txtbizorg.Text = ""
End Sub

Private Sub cmd_search_Click()
Call search(txtsearch.Text, lststd)
Call lststd_Click
End Sub

Private Sub Form_Load()
fres = False
Me.Height = 8850
Me.Width = 11250
Call lstConfig_result(lststd)
Call connectDB
Call LoadSTD_data_result(lststd)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
aerofrm.Visible = True
aerofrm.tcount.Enabled = True
frmmain.Visible = True
frmsres.Visible = False
Call DisconnectDB
Unload frmsres
End Sub

Private Sub lststd_Click()

Dim picn As String
Dim fex As Integer


frmsres.txtname.BackColor = &HFFFFFF
frmsres.txtdep.BackColor = &HFFFFFF
frmsres.txtsem.BackColor = &HFFFFFF
frmsres.txtgrade.BackColor = &HFFFFFF
frmsres.txtcomarch.BackColor = &HFFFFFF
frmsres.txtmicro.BackColor = &HFFFFFF
frmsres.txtdatabaseman.BackColor = &HFFFFFF
frmsres.txtvisual.BackColor = &HFFFFFF
frmsres.txtdatacomfund.BackColor = &HFFFFFF
frmsres.txtenviron.BackColor = &HFFFFFF
frmsres.txtbookkeep.BackColor = &HFFFFFF
frmsres.txtbizorg.BackColor = &HFFFFFF

frmsres.txtname.Refresh
frmsres.txtdep.Refresh
frmsres.txtsem.Refresh
frmsres.txtgrade.Refresh
frmsres.txtcomarch.Refresh
frmsres.txtmicro.Refresh
frmsres.txtdatabaseman.Refresh
frmsres.txtvisual.Refresh
frmsres.txtdatacomfund.Refresh
frmsres.txtenviron.Refresh
frmsres.txtbookkeep.Refresh
frmsres.txtbizorg.Refresh

Call change_Objects_result(lststd)
picn = App.Path & "\pics\" & picname.Caption & ".jpg"
fex = FileExists(picn)
If fex = "-1" Then
    Call Streach_pic(picn, picOriginal, stdpic, 130, 110, 184, 224)
Else
    'NOTHING
End If
fres = False

Call GenResult
End Sub

Public Function GenResult()

If ((frmsres.txtcomarch.Text = "") Or (frmsres.txtmicro.Text = "") Or (frmsres.txtdatabaseman.Text = "") Or (frmsres.txtvisual.Text = "") Or (frmsres.txtdatacomfund.Text = "") Or (frmsres.txtenviron.Text = "") Or (frmsres.txtbookkeep.Text = "") Or (frmsres.txtbizorg.Text = "")) Then
    Call chkEmpty
    Call genSGPA
Else
    'CODE
    Call genSGPA
End If

End Function

Public Function genSGPA()

Dim db_sub_computer_arc As Double
Dim db_sub_microprocessor As Double
Dim db_sub_database_manage As Double
Dim db_sub_visual_programming As Double
Dim db_sub_data_communication As Double
Dim db_sub_environmanal As Double
Dim db_sub_book_keeping As Double
Dim db_sub_business_org As Double
Dim db_sub_score As Double
Dim db_Grade
Dim db_cGPA As Double

Dim tmpGrd As Double

db_sub_database_manage = sGPA(frmsres.txtdatabaseman, 75)       '   Database management
db_sub_computer_arc = sGPA(frmsres.txtcomarch, 50)              '   Computer Architecther
db_sub_microprocessor = sGPA(frmsres.txtmicro, 50)              '   Microprocessor
db_sub_visual_programming = sGPA(frmsres.txtvisual, 50)         '   Visual programing
db_sub_data_communication = sGPA(frmsres.txtdatacomfund, 50)    '   Data Communication
db_sub_environmanal = sGPA(frmsres.txtenviron, 50)              '   Environmental
db_sub_book_keeping = sGPA(frmsres.txtbookkeep, 50)             '   Book Keeping
db_sub_business_org = sGPA(frmsres.txtbizorg, 50)               '   Business Org

tmpGrd = (db_sub_database_manage + db_sub_computer_arc + db_sub_microprocessor + db_sub_visual_programming + db_sub_data_communication + db_sub_environmanal + db_sub_book_keeping + db_sub_business_org)

db_cGPA = tmpGrd / 25 '(*25 is a sum of total Credit)

If fres = False Then

    'If (db_cGPA <= 3.75) And (db_cGPA >= 3.51) Then b_Grade = "A"
    If (db_cGPA <= 4) And (db_cGPA >= 3.76) Then
        txtgrade.Text = "A+"
    End If
    If (db_cGPA <= 3.75) And (db_cGPA >= 3.51) Then
        txtgrade.Text = "A"
    End If
    If (db_cGPA <= 3.5) And (db_cGPA >= 3.26) Then
        txtgrade.Text = "A-"
    End If
    If (db_cGPA <= 3.25) And (db_cGPA >= 3.01) Then
        txtgrade.Text = "B+"
    End If
    If (db_cGPA = 3) Then
        txtgrade.Text = "B"
    End If
    If (db_cGPA <= 2.99) And (db_cGPA >= 2.76) Then
        txtgrade.Text = "B-"
    End If
    If (db_cGPA <= 2.75) And (db_cGPA >= 2.51) Then
        txtgrade.Text = "C+"
    End If
    If (db_cGPA <= 2.5) And (db_cGPA >= 2.25) Then
        txtgrade.Text = "C"
    End If
    If (db_cGPA <= 2.24) And (db_cGPA >= 2) Then
        txtgrade.Text = "D"
    End If
    If (db_cGPA < 2) Then
        txtgrade.Text = "F"
    End If
Else
    txtgrade.Text = "F"
    txtgrade.BackColor = &HC0C0FF
End If

txtcgpa.Text = db_cGPA
'MsgBox db_cGPA

db_sub_score = val(txtdatabaseman.Text) + val(txtcomarch.Text) + val(txtmicro.Text) + _
                val(txtvisual.Text) + val(txtdatacomfund.Text) + val(txtenviron.Text) + val(txtbookkeep.Text) + val(txtbizorg.Text)

txtscore.Text = db_sub_score
'MsgBox db_sub_score
End Function


Public Function sGPA(subtebox As AeroTextBox, subMark As Double) As Double


Dim sMark As Double
Dim sCredit As Double

sMark = val(subtebox.Text)
'MsgBox subMark

    If sMark < (subMark * 0.4) Then ' If fail
        subtebox.Text = sMark
        subtebox.BackColor = &HC0C0FF
        fres = True
    Else
            fres = False
            If subMark = 75 Then sCredit = 4
            If subMark = 50 Then sCredit = 3
            If subMark = 25 Then sCredit = 2
            If subMark = 12.5 Then sCredit = 1
            
            'If ((sMark >= (subMark * 0.9))) Then sGPA = 4 * sCredit
            'If ((sMark >= (subMark * 0.8)) And (sMark <= (subMark * 0.8) + 9)) Then sGPA = 3.75 * sCredit
            'If ((sMark >= (subMark * 0.7)) And (sMark <= (subMark * 0.7) + 9)) Then sGPA = 3.5 * sCredit
            'If ((sMark >= (subMark * 0.6)) And (sMark <= (subMark * 0.6) + 9)) Then sGPA = 3 * sCredit
            'If ((sMark >= (subMark * 0.5)) And (sMark <= (subMark * 0.5) + 9)) Then sGPA = 2.5 * sCredit
            'If ((sMark >= (subMark * 0.4)) And (sMark <= (subMark * 0.4) + 9)) Then sGPA = 2 * sCredit
            
            If ((sMark >= (subMark * 0.8))) Then sGPA = 4 * sCredit '80 A+
            If ((sMark >= (subMark * 0.8) - 9) And (sMark <= (subMark * 0.8) - 5)) Then sGPA = 3.75 * sCredit '75 A
            If ((sMark >= (subMark * 0.8) - 11) And (sMark <= (subMark * 0.8) - 10)) Then sGPA = 3.5 * sCredit  '70 A-
            If ((sMark >= (subMark * 0.8) - 16) And (sMark <= (subMark * 0.8) - 15)) Then sGPA = 3.25 * sCredit '65 B+
            If ((sMark >= (subMark * 0.8) - 21) And (sMark <= (subMark * 0.8) - 20)) Then sGPA = 3 * sCredit '60 B
            If ((sMark >= (subMark * 0.8) - 26) And (sMark <= (subMark * 0.8) - 25)) Then sGPA = 2.75 * sCredit '55 B-
            If ((sMark >= (subMark * 0.8) - 31) And (sMark <= (subMark * 0.8) - 30)) Then sGPA = 2.5 * sCredit '50 C+
            If ((sMark >= (subMark * 0.8) - 36) And (sMark <= (subMark * 0.8) - 35)) Then sGPA = 2.25 * sCredit '45 C
            If ((sMark >= (subMark * 0.8) - 40) And (sMark <= (subMark * 0.8) - 40)) Then sGPA = 2 * sCredit '40 D
            
            
            'MsgBox sGPA
    End If
    
End Function


Public Function chkEmpty()

If frmsres.txtname.Text = "" Then frmsres.txtname.BackColor = &HC0C0FF
If frmsres.txtdep.Text = "" Then frmsres.txtdep.BackColor = &HC0C0FF
If frmsres.txtsem.Text = "" Then frmsres.txtsem.BackColor = &HC0C0FF
If frmsres.txtgrade.Text = "" Then frmsres.txtgrade.BackColor = &HC0C0FF
If frmsres.txtcomarch.Text = "" Then frmsres.txtcomarch.BackColor = &HC0C0FF
If frmsres.txtmicro.Text = "" Then frmsres.txtmicro.BackColor = &HC0C0FF
If frmsres.txtdatabaseman.Text = "" Then frmsres.txtdatabaseman.BackColor = &HC0C0FF
If frmsres.txtvisual.Text = "" Then frmsres.txtvisual.BackColor = &HC0C0FF
If frmsres.txtdatacomfund.Text = "" Then frmsres.txtdatacomfund.BackColor = &HC0C0FF
If frmsres.txtenviron.Text = "" Then frmsres.txtenviron.BackColor = &HC0C0FF
If frmsres.txtbookkeep.Text = "" Then frmsres.txtbookkeep.BackColor = &HC0C0FF
If frmsres.txtbizorg.Text = "" Then frmsres.txtbizorg.BackColor = &HC0C0FF
End Function

