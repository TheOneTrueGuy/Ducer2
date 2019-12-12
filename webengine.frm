VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{60D62220-313C-11D1-9264-0000C00D66F5}#1.0#0"; "TEGODNLD.OCX"
Object = "{C1B9D260-7B70-11D2-9267-0000C00D66F5}#1.0#0"; "TEGOUPLD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TEGODNLDLib.TegoDnld TegoDnld1 
      Height          =   570
      Left            =   4485
      TabIndex        =   5
      Top             =   3750
      Visible         =   0   'False
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1005
      _StockProps     =   0
   End
   Begin TEGOUPLDLib.TegoUpld TegoUpld1 
      Height          =   570
      Left            =   5655
      TabIndex        =   4
      Top             =   3645
      Visible         =   0   'False
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   1005
      _StockProps     =   0
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7095
      Top             =   3765
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6450
      Top             =   3750
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   8535
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2790
      TabIndex        =   1
      Top             =   150
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Engage"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   2235
   End
   Begin VB.Timer Timer1 
      Left            =   9495
      Top             =   3750
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7845
      Top             =   3690
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "starting....."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   105
      TabIndex        =   3
      Top             =   1245
      Width           =   10050
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   15
      TabIndex        =   2
      Top             =   915
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim starter As String
Dim scripter As String, strFormData As String
Dim pop As String
Dim donn As String
' ***_____________________
Public gPath As String
Public gHtmlOcxHwnd
Public gProgramIsRunningInsideIE As Boolean
' ###______________________



Private Sub Command1_Click()

Dim start1 As String, strFormData As String
start1 = starter
'strFormData = "from=" & fromer$ & "&to=" & adrs & "&fylname=" & idN & "&jleft=" & jleft & "&jtop=" & jtop & _
'"&mcount=" & mcount & "&duration=" & duration & "&bottom=" & bmporjpg & "&moves=" & moves$
strFormData = "hwnd=" & LTrim(Str(Me.hWnd)) & "&donn=" & donn & _
    "&ipa=" & Winsock1.LocalIP

Inet1.Execute start1, "Post", strFormData, _
            "Content-Type: application/x-www-form-urlencoded"

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
starter = "http://duecer.myonlinehost.com/asp/starter.asp"
scripter = "http://wduecer.myonlinehost.com/asp/poposter.asp"
Debug.Print Me.hWnd
gPath = App.Path
If Right(gPath, 1) <> "\" Then gPath = gPath + "\"
gHtmlOcxHwnd = Val(TegoDnld1.TextFileToString(gPath + "~OcxHwnd.txt"))
If gHtmlOcxHwnd = 0 Then
   gProgramIsRunningInsideIE = False
'   xit.Caption = "E&xit"
Else
   gProgramIsRunningInsideIE = True
  
'   xit.Caption = "E&xit and return to previous page"
End If
Dim t As Integer, xyz As Integer
For t = 1 To 72
xyz = Int(Rnd * 10)
donn = donn & LTrim(Str(xyz))
Next t


End Sub

Private Sub Form_Unload(Cancel As Integer)
If gProgramIsRunningInsideIE = True Then
   'Result = TegoDnld1.GenerateCustomEvent(gHtmlOcxHwnd, "PREVIOUS PAGE")
   TegoDnld1.Execute "http://duecer.myonlinehost.com/thanx.htm"
Else: End

End If
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim dat As String
If State = 12 Then
dat$ = Inet1.GetChunk(1024, icString)
Label2.Caption = dat$
End If
End Sub
