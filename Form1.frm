VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "HASH TABLES - Advanced Algorithms"
   ClientHeight    =   7530
   ClientLeft      =   4350
   ClientTop       =   2550
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10545
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   7350
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1320
      Width           =   2940
   End
   Begin VB.CommandButton Command7 
      Caption         =   "SHA-160 (SHA1)"
      Height          =   450
      Left            =   5475
      TabIndex        =   15
      Top             =   1305
      Width           =   1785
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   705
      Left            =   2085
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1365
      Width           =   2595
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MD5 -128 bits"
      Height          =   450
      Left            =   180
      TabIndex        =   12
      Top             =   1320
      Width           =   1785
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SHA-256 Optimized"
      Height          =   450
      Left            =   210
      TabIndex        =   10
      Top             =   3435
      Width           =   1785
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4380
      Width           =   3105
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   990
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   165
      Width           =   960
   End
   Begin VB.TextBox txtORG 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2145
      TabIndex        =   6
      Text            =   "Enter HERE the text you want to calculate its Hash Values."
      Top             =   180
      Width           =   7980
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SHA-1024"
      Height          =   450
      Left            =   7845
      TabIndex        =   3
      Top             =   3435
      Width           =   1785
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   2100
      Left            =   7125
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4380
      Width           =   3105
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SHA-512"
      Height          =   450
      Left            =   4140
      TabIndex        =   1
      Top             =   3435
      Width           =   1785
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   1875
      Left            =   3615
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4380
      Width           =   3105
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H000080FF&
      Height          =   810
      Left            =   885
      TabIndex        =   20
      Top             =   6810
      Width           =   9060
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":011E
      ForeColor       =   &H000080FF&
      Height          =   750
      Left            =   105
      TabIndex        =   19
      Top             =   2625
      Width           =   10035
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   3450
      Left            =   75
      Top             =   3270
      Width           =   10335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":02A2
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   750
      Width           =   8640
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1710
      Left            =   90
      Top             =   675
      Width           =   10335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   5490
      TabIndex        =   17
      Top             =   1845
      Width           =   1785
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   300
      Left            =   225
      TabIndex        =   14
      Top             =   1905
      Width           =   1785
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   210
      TabIndex        =   11
      Top             =   3990
      Width           =   1785
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "LOOP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   405
      TabIndex        =   7
      Top             =   225
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   7830
      TabIndex        =   5
      Top             =   3990
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4125
      TabIndex        =   4
      Top             =   3990
      Width           =   1785
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As Variant
Dim X As Long, y As Long, z As Long

Dim result As String, total As Long





Private Sub Command2_Click()

Dim origem As String
    
    origem = Trim(txtORG.Text)
    total = Combo1.List(Combo1.ListIndex)
    
    a = Timer
    For X = 1 To total
        result = SHA512(origem)
    Next
    

    Label2.Caption = Format(Timer - a, "00:00:00")
    Text2.Text = UCase(result)


End Sub

Private Sub Command3_Click()

Dim origem As String
    
    origem = Trim(txtORG.Text)
    total = Combo1.List(Combo1.ListIndex)
    
    a = Timer
    For X = 1 To total
        result = SHA1024(origem)
    Next
    

    Label3.Caption = Format(Timer - a, "00:00:00")
    Text3.Text = UCase(result)


End Sub

Private Sub Command5_Click()

Dim origem As String
    
    origem = Trim(txtORG.Text)
    total = Combo1.List(Combo1.ListIndex)
    
    a = Timer
    For X = 1 To total
        result = SHA256o(origem)
    Next
    

    Label6.Caption = Format(Timer - a, "00:00:00")
    Text5.Text = UCase(result)


End Sub

Private Sub Command6_Click()


Dim origem As String
    
    origem = Trim(txtORG.Text)
    total = Combo1.List(Combo1.ListIndex)
    
    a = Timer
    For X = 1 To total
        result = MD5_Calc(origem)
    Next
    

    Label7.Caption = Format(Timer - a, "00:00:00")
    Text6.Text = UCase(result)

End Sub

Private Sub Command7_Click()

Dim origem As String
    
    origem = Trim(txtORG.Text)
    total = Combo1.List(Combo1.ListIndex)
    
    a = Timer
    For X = 1 To total
        result = Sha1(origem)
    Next
    

    Label8.Caption = Format(Timer - a, "00:00:00")
    Text7.Text = UCase(result)

End Sub

Private Sub Form_Load()

    Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    For X = 1 To 10000 Step 100
        Combo1.AddItem X
    Next
    
    Combo1.ListIndex = 20
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
    
End Sub

Private Sub Form_Terminate()

    End
    
End Sub
