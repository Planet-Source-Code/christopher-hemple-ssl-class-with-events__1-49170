VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SSL Class Example : Built On ""SSL Client"" By Jason K. Resch "
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstDebug 
      Height          =   2400
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtConnect 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "paypal.com"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox txtSendConnect 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "GET https://www.paypal.com/"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblVbcrlf 
      Caption         =   "&& VbCrLf"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label lblConnect 
      Caption         =   "Connect To :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3030
      Width           =   1335
   End
   Begin VB.Label lblConnectSend 
      Caption         =   "Send on Connect :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2685
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents clsSSL As clsSSL
Attribute clsSSL.VB_VarHelpID = -1

Private Sub clsSSL_SSLClose()
    'Display a Debug Message
    lstDebug.AddItem "--SSL Closed--"
End Sub

Private Sub clsSSL_SSLConnect()
    'Display a Debug Message
    lstDebug.AddItem "--Connected To SSL--"
    'Send the Data
    clsSSL.SendSSL txtSendConnect.Text & vbCrLf
End Sub

Private Sub clsSSL_SSLConnecting()
    'Display a Debug Message
    lstDebug.AddItem "--Connecting To SSL--"
End Sub

Private Sub clsSSL_SSLData(sData As String)
    'Display a Debug Message
    lstDebug.AddItem "--Got SSL Data--"
    'Set the Debug Text
    txtDebug.Text = sData
    'Close The Ssl Socket
    clsSSL.CloseSSL
End Sub

Private Sub clsSSL_SSLSendingData(sData As String)
    'Display a Debug Message
    lstDebug.AddItem "--Sending Data--"
End Sub

Private Sub cmdConnect_Click()
    'Start Connecting
    clsSSL.ConnectSSL txtConnect.Text
End Sub

Private Sub Form_Initialize()
    'Display a Debug Message
    lstDebug.AddItem "--Init'ed SSL Class--"
    'Create the SSL Class
    Set clsSSL = New clsSSL
End Sub

Private Sub Form_Terminate()
    'Display a Debug Message
    lstDebug.AddItem "--Terminated SSL Class--"
    'Remove the SSL Class To Free memory
    Set clsSSL = Nothing
End Sub
