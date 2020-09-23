VERSION 5.00
Begin VB.Form frmSMTP 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "SMTP Einstellungen"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmSMTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdSave 
      Caption         =   "Speichern"
      Height          =   375
      Left            =   2970
      TabIndex        =   12
      Top             =   75
      Width           =   1425
   End
   Begin VB.TextBox txtBetreff 
      Height          =   285
      Left            =   1395
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtEMail 
      Height          =   285
      Left            =   1395
      TabIndex        =   9
      Top             =   1500
      Width           =   1455
   End
   Begin VB.TextBox txtSendName 
      Height          =   285
      Left            =   1395
      TabIndex        =   7
      Top             =   1185
      Width           =   1455
   End
   Begin VB.TextBox txtPasswort 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1395
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   705
      Width           =   1455
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1395
      TabIndex        =   4
      Top             =   405
      Width           =   1455
   End
   Begin VB.TextBox txtSMTP 
      Height          =   285
      Left            =   1395
      TabIndex        =   3
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Betreff:"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1875
      Width           =   1110
   End
   Begin VB.Label Label5 
      Caption         =   "Absender E-Mail:"
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "Absendername:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1245
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Login-Name:"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   465
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "SMTP-Server:"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1020
   End
End
Attribute VB_Name = "frmSMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
Open "server.cfg" For Output As 3
Print #3, txtSMTP.Text
Print #3, txtUsername.Text
Print #3, txtPasswort.Text
Print #3, txtSendName.Text
Print #3, txtEMail.Text
Print #3, txtBetreff.Text
Close #3
Me.Hide
Newsletter.Enabled = True
Newsletter.Show
Server = txtSMTP.Text
Username = txtUsername.Text
Passwort = txtPasswort.Text
Absendername = txtSendName.Text
Absenderemail = txtEMail.Text
Betreff = txtBetreff.Text
End Sub

Private Sub Form_Load()
txtSMTP.Text = Server
txtUsername.Text = Username
txtPasswort.Text = Passwort
txtSendName.Text = Absendername
txtEMail.Text = Absenderemail
txtBetreff.Text = Betreff
End Sub

Private Sub Form_Unload(Cancel As Integer)
Newsletter.Enabled = True
Newsletter.Show

End Sub
