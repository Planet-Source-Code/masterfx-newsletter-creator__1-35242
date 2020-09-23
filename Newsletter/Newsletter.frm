VERSION 5.00
Begin VB.Form Newsletter 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Newsletter"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "Newsletter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Einstellungen"
      Height          =   360
      Left            =   6330
      TabIndex        =   19
      Top             =   3945
      Width           =   1335
   End
   Begin VB.HScrollBar scrFeld 
      Height          =   225
      Left            =   6345
      Max             =   1
      Min             =   1
      TabIndex        =   18
      Top             =   1950
      Value           =   1
      Width           =   1155
   End
   Begin VB.VScrollBar scrFelder 
      Height          =   405
      Left            =   7200
      Max             =   1
      Min             =   100
      TabIndex        =   15
      Top             =   1140
      Value           =   1
      Width           =   225
   End
   Begin VB.CommandButton cmdEnde 
      Caption         =   "Beenden"
      Height          =   360
      Left            =   6330
      TabIndex        =   12
      Top             =   4305
      Width           =   1335
   End
   Begin VB.CommandButton cmdVorschau 
      Caption         =   "Vorschau"
      Height          =   360
      Left            =   6330
      TabIndex        =   11
      Top             =   615
      Width           =   1335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Senden"
      Height          =   360
      Left            =   6330
      TabIndex        =   10
      Top             =   255
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nachricht"
      Height          =   4515
      Left            =   1860
      TabIndex        =   5
      Top             =   165
      Width           =   4380
      Begin VB.TextBox txtNachricht 
         Height          =   3345
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   4095
      End
      Begin VB.TextBox txtTitel 
         Height          =   300
         Left            =   615
         TabIndex        =   7
         Top             =   285
         Width           =   3660
      End
      Begin VB.Label Label4 
         Caption         =   "Nachricht:"
         Height          =   210
         Left            =   165
         TabIndex        =   8
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Titel:"
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   330
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Löschen"
      Height          =   345
      Left            =   45
      TabIndex        =   4
      Top             =   4320
      Width           =   1380
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Hinzufügen"
      Height          =   345
      Left            =   45
      TabIndex        =   3
      Top             =   3975
      Width           =   1380
   End
   Begin VB.ListBox lstUser 
      Height          =   3180
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label lblFeld 
      Alignment       =   2  'Zentriert
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7350
      TabIndex        =   17
      Top             =   1695
      Width           =   285
   End
   Begin VB.Label Label6 
      Caption         =   "Aktuelles Feld:"
      Height          =   240
      Left            =   6345
      TabIndex        =   16
      Top             =   1695
      Width           =   1185
   End
   Begin VB.Label lblFelder 
      Alignment       =   2  'Zentriert
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6870
      TabIndex        =   14
      Top             =   1335
      Width           =   300
   End
   Begin VB.Label Label5 
      Caption         =   "Anzahl der Felder:"
      Height          =   465
      Left            =   6390
      TabIndex        =   13
      Top             =   1125
      Width           =   810
   End
   Begin VB.Label lblUser 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   3735
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Newletter-Abonenten:"
      Height          =   270
      Left            =   75
      TabIndex        =   0
      Top             =   165
      Width           =   1590
   End
   Begin VB.Menu menue 
      Caption         =   "Hauptmenue"
      Visible         =   0   'False
      Begin VB.Menu mnuAddLink 
         Caption         =   "Link einfügen"
      End
      Begin VB.Menu mnuAddPic 
         Caption         =   "Bild einfügen"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "Format"
         Begin VB.Menu mnuFett 
            Caption         =   "Fettgedruckt"
         End
         Begin VB.Menu mnuUnder 
            Caption         =   "Unterstreichen"
         End
         Begin VB.Menu mnuFettUnder 
            Caption         =   "Fett und Unterstreichen"
         End
      End
   End
End
Attribute VB_Name = "Newsletter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SendMail As vbSendMail.clsSendMail
Attribute SendMail.VB_VarHelpID = -1
Private Titel() As String, Nachricht() As String

Dim Adressen As String, Adresse1 As String, DaNr As Integer

Private Sub cmdAdd_Click()
Dim NewUser As String
NewUser = InputBox("Bitte die neue E-Mail Adresse eingeben!", "Neue E-Mail Adresse")
If NewUser <> "" Then
   lstUser.AddItem (NewUser)
End If
End Sub

Private Sub cmdDel_Click()
If lstUser.List(lstUser.ListIndex) <> "" Then
   lstUser.RemoveItem (lstUser.ListIndex)
   Else
   MsgBox "Bitte wähle eine E-Mail Adresse aus den Du entfernen willst!", vbCritical, "Fehler"
End If
End Sub

Private Sub cmdEnde_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
Dim MsgNachricht As String, Temp As String
Dim i As Integer, Groesse As Long
Call HTMLKopf
For i = 1 To scrFelder.Value
  Call HTMLNachricht(Titel(i), Nachricht(i))
Next i
Call HTMLEnde
DaNr = FreeFile
Open "Newsletter.htm" For Input As DaNr
Do
  Line Input #DaNr, Temp
  MsgNachricht = MsgNachricht & Temp & vbCr
Loop While EOF(DaNr) = False
Close #DaNr
Call Aufreihen
Me.Enabled = False
frmStatus.Show
frmStatus.List2.Clear
With SendMail
     .SMTPHostValidation = VALIDATE_HOST_DNS
     .EmailAddressValidation = VALIDATE_SYNTAX
     .SMTPHost = Server
     .UseAuthentication = True
     .Username = Username
     .Password = Passwort
     .EncodeType = MIME_ENCODE
     .AsHTML = True
     .Subject = Betreff
     .From = Absenderemail
     .FromDisplayName = Absendername
     .Recipient = Adresse1
     .BccRecipient = Adressen
     .RecipientDisplayName = "Mitglieder der SCO-Allianz"
     .Message = MsgNachricht
     .Send
End With
End Sub

Private Sub cmdVorschau_Click()
Dim i As Integer
Call HTMLKopf
For i = 1 To scrFeld.Value
  Call HTMLNachricht(Titel(i), Nachricht(i))
Next i
Call HTMLEnde
Shell "c:\Programme\Internet Explorer\iexplore.exe " & App.Path & "\Newsletter.htm"
End Sub


Private Sub mnuAddLink_Click()
Dim Link As String, LinkName As String
Link = InputBox("Bitte geben Sie die Adresse ein!", "Link einfügen", "http://")
If Len(Link) > 10 Then
  LinkName = InputBox("Bitte geben Sie den Namen von dem Link ein!", "Name eingeben", Right(Link, Len(Link) - 11))
  Else
  LinkName = InputBox("Bitte geben Sie den Namen von dem Link ein!", "Name eingeben", Link)
End If
txtNachricht.Text = txtNachricht.Text & "<a href=""" & Link & """>" & LinkName & "</a>"
txtNachricht.SelStart = Len(txtNachricht.Text)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub

Private Sub mnuAddPic_Click()
Dim Pic As String, Breite As Integer, Hoehe As Integer
Pic = InputBox("Bitte die Adresse des Bildes eingeben", "Bild einfügen", "http://")
Breite = InputBox("Bitte geben Sie die Breite des Bildes ein!", "Bild einfügen", "300")
Hoehe = InputBox("Bitte geben Sie die Höhe des Bildes ein!", "Bild einfügen", "60")
txtNachricht.Text = txtNachricht.Text & "<img scr=""" & Pic & """ width=""" & Breite & """ height=""" & Hoehe & """>"
txtNachricht.SelStart = Len(txtNachricht.Text)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub

Private Sub mnuFett_Click()
Dim Temp As String
Temp = InputBox("Bitte geben Sie den Text ein der fettgedruckt sein soll!", "Fettgedruckt", "Was fettes")
txtNachricht.Text = txtNachricht.Text & "<b>" & Temp & "</b>"
txtNachricht.SelStart = Len(txtNachricht.Text)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub

Private Sub mnuFettUnder_Click()
Dim Temp As String
Temp = InputBox("Bitte geben Sie den Text ein der fett und unterstrichen sein soll!", "Fett und unterstrichen", "Was fettes und unterstrichenes")
txtNachricht.Text = txtNachricht.Text & "<b><u>" & Temp & "</u></b>"
txtNachricht.SelStart = Len(txtNachricht.Text)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub

Private Sub mnuUnder_Click()
Dim Temp As String
Temp = InputBox("Bitte geben Sie den Text ein der unterstrichen sein soll!", "Unterstrichen", "Was unterstrichenes")
txtNachricht.Text = txtNachricht.Text & "<u>" & Temp & "</u>"
txtNachricht.SelStart = Len(txtNachricht.Text)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub

Private Sub SendMail_Status(Status As String)
  frmStatus.List2.AddItem Status
  frmStatus.List2.ListIndex = frmStatus.List2.ListCount - 1
End Sub

Private Sub SendMail_SendFailed(Explanation As String)
  MsgBox ("Fehler beim Senden: " & vbCrLf & Explanation)
  Me.Enabled = True
End Sub

Private Sub SendMail_SendSuccesful()
  MsgBox "E-Mail erfolgreich versendet!"
  Me.Enabled = True
  Me.Show
  Unload frmStatus
End Sub

Sub Aufreihen()
Dim i As Integer
Adresse1 = lstUser.List(0)
For i = 1 To lstUser.ListCount - 1
    Adressen = Adressen & lstUser.List(i) & ";"
Next i
If lstUser.ListCount > 1 Then
Adressen = Left(Adressen, Len(Adressen) - 1)
End If
End Sub


Private Sub cmdSettings_Click()
frmSMTP.Show
Newsletter.Enabled = False
End Sub

Private Sub Form_Load()
Dim Temp As String
ReDim Preserve Titel(1 To 1)
ReDim Preserve Nachricht(1 To 1)
DaNr = FreeFile
Open "emails.txt" For Input As DaNr
Do
Line Input #DaNr, Temp
lstUser.AddItem (Temp)
Loop While EOF(DaNr) = False
Close #DaNr
Set SendMail = New clsSendMail
Call ServerCFGLaden
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
DaNr = FreeFile
Open "emails.txt" For Output As DaNr
For i = 0 To lstUser.ListCount - 1
Print #DaNr, lstUser.List(i)
Next i
Close #DaNr
Unload frmSMTP
SendMail.Shutdown
Set SendMail = Nothing
End Sub

Private Sub lstUser_Click()
lblUser.Caption = lstUser.List(lstUser.ListIndex)
End Sub

Private Sub scrFeld_Change()
lblFeld.Caption = scrFeld.Value
If Titel(scrFeld.Value) <> "" Then              'Wenn Titel(i) schon beschrieben dann Anzeigen
   txtTitel.Text = Titel(scrFeld.Value)
   txtNachricht.Text = Nachricht(scrFeld.Value)
   Else
     txtTitel.Text = ""
     txtNachricht.Text = ""
End If
End Sub

Private Sub scrFelder_Change()
lblFelder.Caption = scrFelder.Value
scrFeld.Max = scrFelder.Value                   'Anzahl der Nachrichten festlegen
ReDim Preserve Titel(1 To scrFelder.Value)
ReDim Preserve Nachricht(1 To scrFelder.Value)
End Sub

Sub ServerCFGLaden()
Open "server.cfg" For Input As 3
Line Input #3, Server
Line Input #3, Username
Line Input #3, Passwort
Line Input #3, Absendername
Line Input #3, Absenderemail
Line Input #3, Betreff
Close #3
End Sub

'HTML Kopf
Sub HTMLKopf()
DaNr = FreeFile
Open "Newsletter.htm" For Output As DaNr
Print #DaNr, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
Print #DaNr, "<html>"
Print #DaNr, "<head>"
Print #DaNr, "</head>"
Print #DaNr, "<body bgcolor=""#000000"" text=""#FFFFFF"">"
Print #DaNr, "<div align=""center"">"
End Sub

'HTML Nachricht
Sub HTMLNachricht(MsgTitel As String, MsgText As String)
Print #DaNr, "<table width=""400"">"
Print #DaNr, "<tr>"
Print #DaNr, "<td align=""center"" bgcolor=""#103050""><b><u>" & MsgTitel & "</u></b></td>"
Print #DaNr, "</tr>"
Print #DaNr, "<tr>"
Print #DaNr, "<td align=""center"" bgcolor=""#808080"">" & MsgText & "</td>"
Print #DaNr, "</tr>"
Print #DaNr, "</table>"
Print #DaNr, "<br><br>"
End Sub

'HTML Ende
Sub HTMLEnde()
Print #DaNr, "</div>"
Print #DaNr, "</body>"
Print #DaNr, "</html>"
Close #DaNr
End Sub

Private Sub txtNachricht_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   txtNachricht.Text = txtNachricht.Text & "<br>"
   txtNachricht.SelStart = Len(txtNachricht.Text)
End If
End Sub

Private Sub txtNachricht_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    txtNachricht.Enabled = False
    txtNachricht.Enabled = True
    txtNachricht.SetFocus
    PopupMenu menue
End If
End Sub

Private Sub txtTitel_KeyUp(KeyCode As Integer, Shift As Integer)
Titel(scrFeld.Value) = txtTitel.Text
End Sub

Private Sub txtNachricht_KeyUp(KeyCode As Integer, Shift As Integer)
Nachricht(scrFeld.Value) = txtNachricht.Text
End Sub
