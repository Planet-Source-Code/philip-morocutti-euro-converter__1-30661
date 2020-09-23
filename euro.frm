VERSION 5.00
Begin VB.Form €uro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "€uro"
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   Icon            =   "euro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox eurol 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D54600&
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Valuta in Euro"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox euro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D54600&
      Height          =   285
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Valuta in Euro"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox lirel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Valuta in Lire"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox lire 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   600
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "Valuta in Lire"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   245
      Left            =   0
      Top             =   0
      Width           =   5320
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "€uro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0007EAFC&
      Height          =   1740
      Left            =   1845
      TabIndex        =   9
      Top             =   195
      Width           =   3345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "€uro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D54600&
      Height          =   1740
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   3345
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   5320
      ToolTipText     =   "Riduci a icona"
      Top             =   10
      Width           =   120
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   225
      Left            =   5320
      Top             =   10
      Width           =   120
   End
   Begin VB.Image Image1 
      Height          =   230
      Left            =   5425
      ToolTipText     =   "Chiudi"
      Top             =   10
      Width           =   230
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0007EAFC&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   230
      Left            =   5425
      Top             =   10
      Width           =   230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00D54600&
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Image Image7 
      Height          =   540
      Left            =   15
      Picture         =   "euro.frx":636A
      Top             =   1265
      Width           =   1710
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Euro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D54600&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Euro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D54600&
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3180
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2460
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00EDEFEF&
      X1              =   2835
      X2              =   2835
      Y1              =   3840
      Y2              =   2400
   End
   Begin VB.Image Image5 
      Height          =   1245
      Left            =   15
      Picture         =   "euro.frx":68FD
      Top             =   15
      Width           =   1710
   End
   Begin VB.Image Image2 
      Height          =   3970
      Left            =   1800
      Picture         =   "euro.frx":79CC
      Stretch         =   -1  'True
      Top             =   230
      Width           =   3840
   End
   Begin VB.Image Image8 
      Height          =   1245
      Left            =   1680
      Picture         =   "euro.frx":826A
      Top             =   15
      Width           =   5280
   End
End
Attribute VB_Name = "€uro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright © 2002 Philip Morocutti

'E-mail:    philip@morocutti.f2s.com
'Web:       http://www.morocutti.f2s.com

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub euro_Change()
On Error Resume Next
If euro = "" Then euro = 0
Dim corrente As Double
corrente = Format(euro * 1936.27, "#0")
lirel = corrente
End Sub

Private Sub euro_Click()
euro.SelStart = 0
euro.SelLength = Len(euro.Text)
euro.SetFocus
End Sub

Private Sub euro_GotFocus()
euro.SelStart = 0
euro.SelLength = Len(euro.Text)
euro.SetFocus
End Sub

Private Sub eurol_Click()
eurol.SelStart = 0
eurol.SelLength = Len(eurol.Text)
eurol.SetFocus
End Sub

Private Sub eurol_GotFocus()
eurol.SelStart = 0
eurol.SelLength = Len(eurol.Text)
eurol.SetFocus
End Sub

Private Sub Form_Load()
lire.SelStart = 0
lire.SelLength = Len(lire.Text)
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Image3_Click()
Me.WindowState = 1
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.WindowState = 0 Then
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End If
End Sub

Private Sub lire_Change()
On Error Resume Next
If lire = "" Then lire = 0
Dim corrente As Double
corrente = Format(lire / 1936.27, "#0.00")
eurol = corrente
End Sub

Private Sub lire_Click()
lire.SelStart = 0
lire.SelLength = Len(lire.Text)
lire.SetFocus
End Sub

Private Sub lire_GotFocus()
lire.SelStart = 0
lire.SelLength = Len(lire.Text)
lire.SetFocus
End Sub

Private Sub lirel_Click()
lirel.SelStart = 0
lirel.SelLength = Len(lirel.Text)
lirel.SetFocus
End Sub

Private Sub lirel_GotFocus()
lirel.SelStart = 0
lirel.SelLength = Len(lirel.Text)
lirel.SetFocus
End Sub
