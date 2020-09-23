VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Serverstatus"
   ClientHeight    =   2430
   ClientLeft      =   4080
   ClientTop       =   3405
   ClientWidth     =   4575
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4575
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   4560
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Newsletter.Enabled = True
Newsletter.Show
End Sub
