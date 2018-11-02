VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Misafir Email G�nderici..."
   ClientHeight    =   2085
   ClientLeft      =   18705
   ClientTop       =   3840
   ClientWidth     =   4155
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4155
   Begin VB.CommandButton Command1 
      Caption         =   "G�nder"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bilgiler"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   960
         TabIndex        =   4
         Text            =   "Email"
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   390
         Left            =   960
         TabIndex        =   1
         Text            =   "Ad Soyad"
         Top             =   230
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Email:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Ad-Soyad:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Durum"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
     bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.HTMLBody = sBody
    'lobj_cdomsg.TextBody = sBody
    lobj_cdomsg.Send
    
     
    Set lobj_cdomsg = Nothing
    SendMail = "Email ba�ar�yla yolland�...."
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function

Private Sub Command1_Click()
Dim mesaj As String
mesaj = "Say�n " + UCase(Text1.Text) + ";" + "<p>" + "Ba�ak�ehir Livinglab mail listesine kay�d�n�z al�nm��t�r.L�tfen " + "<a href=" + Chr(34) + "http://eepurl.com/2bscf" + Chr(34) + ">buradan</a>" + " linke t�klayarak bilgilerinizi g�ncelleyiniz." + "<p/>"
    Dim retVal          As String
    Dim objControl      As Control
    Dim hata
    'Validate first
    For Each objControl In Me.Controls
        If TypeOf objControl Is TextBox Then
            If Trim$(objControl.Text) = vbNullString Then
               hata = MsgBox("Dikkat T�m Alanlar� Doldurunuz....", vbCritical, "Hata!!!!!")
                
                Exit Sub
            End If
        End If
    Next
    
    'Send
    Frame1.Enabled = False
    

    Command1.Enabled = False
    Label4.Caption = "Yollan�yor..."
    retVal = SendMail(Trim$(Text2.Text), _
        Trim$("Ba�ak�ehir LivingLab Mail Listesi Kayd�"), _
        Trim$("Ba�ak�ehir LivingLab") & "<" & Trim$("fromMail") & ">", _
        Trim$(mesaj), _
        Trim$("server"), _
        CInt(Trim$("port")), _
        Trim$("fromMail"), _
        Trim$("pass"), _
        CBool(False)) 'ssl kontrol
    Frame1.Enabled = True
    
    Command1.Enabled = True
    Label4.Caption = IIf(retVal = "Email ba�ar�yla yolland�....", "Email ba�ar�yla yolland�....", retVal)
    
End Sub


