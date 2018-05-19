VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000E&
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   LinkTopic       =   "Form2"
   ScaleHeight     =   7725
   ScaleWidth      =   15435
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   1080
      ScaleHeight     =   6585
      ScaleWidth      =   11985
      TabIndex        =   0
      Top             =   360
      Width           =   12015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hybrid Cars"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   8160
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2295
         Left            =   3120
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   5895
      End
      Begin VB.Line Line2 
         X1              =   9000
         X2              =   12000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sports Cars"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   8040
         TabIndex        =   4
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Line Line4 
         X1              =   6000
         X2              =   6000
         Y1              =   3840
         Y2              =   6600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Hyper Cars"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Commercial Cars"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   3495
      End
      Begin VB.Line Line3 
         X1              =   6000
         X2              =   6000
         Y1              =   0
         Y2              =   1560
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture("car1.jpg")
Picture1.Picture = LoadPicture("car-logos.jpg")
Picture1.ScaleMode = 3
Picture1.AutoRedraw = True
Picture1.PaintPicture Picture1.Picture, _
0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
0, 0, Picture1.Picture.Width / 26.46, _
Picture1.Picture.Height / 26.46
Picture1.Picture = Picture1.Image
End Sub
Private Sub Label1_Click()
Me.Hide
Form3.Show
Form3.Text1.Text = ""
End Sub
Private Sub Label2_Click()
Me.Hide
Form5.Show
Form5.Text1.Text = ""
End Sub
Private Sub Label3_Click()
Me.Hide
Form6.Show
Form6.Text1.Text = ""
End Sub
Private Sub Label4_Click()
Me.Hide
Form7.Show
Form7.Text1.Text = ""
End Sub
