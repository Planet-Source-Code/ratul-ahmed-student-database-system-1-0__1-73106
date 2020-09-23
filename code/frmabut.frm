VERSION 5.00
Begin VB.Form frmabut 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmabut.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmabut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
aerofrm.Visible = True
aerofrm.tcount.Enabled = True
frmmain.Visible = True
frmabut.Visible = False
Unload frmabut
End Sub

