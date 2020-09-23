VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   Caption         =   "Yahoo Status"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Check Yahoo ID"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2280
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Yahoo ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim StatVar As String
    If Inet1.StillExecuting Then Exit Sub
    Label1 = "Status : EXECUTING"
    DoEvents
    StatVar = Inet1.OpenURL("http://opi.yahoo.com/online?u=" & Text1 & "&m=j")
    DoEvents
    If InStr(StatVar, "NOT ONLINE") Then
        Label1 = "Status : NOT ONLINE"
    ElseIf InStr(StatVar, "ONLINE") Then
        Label1 = "Status : ONLINE"
    Else
        Label1 = "Status : ERROR"
    End If
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelStart = Len(Text1)
End Sub
