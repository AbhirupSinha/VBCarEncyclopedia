VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H8000000E&
   Caption         =   "SEARCH SPORTS CARS"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   7755
   ScaleWidth      =   15435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   3480
      Width           =   7695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sqlStr As String

Private Sub Command1_Click()
searchvar = Text1.Text
sqlsrch = "select * from Sports_Cars where Car_Name=" & "'" & searchvar & "'"
rs.Close
rs.Open (sqlsrch), conn, adOpenStatic, adLockReadOnly
If rs.Fields(0) <> "" Then
Form4.Label1.Caption = rs.Fields("Car_Name")
Form4.Label2.Caption = rs.Fields("Description")
Form4.Text1.Text = rs.Fields("Wheelbase")
Form4.Text2.Text = rs.Fields("Length")
Form4.Text3.Text = rs.Fields("Width")
Form4.Text4.Text = rs.Fields("Height")
Form4.Text5.Text = rs.Fields("Manufacturer")
Form4.Text6.Text = "NA"
Form4.Text7.Text = rs.Fields("Production_Years")
Form4.Text8.Text = rs.Fields("Base_Price")
Form4.Text9.Text = rs.Fields("Kerb_Weight")
Form4.Text10.Text = rs.Fields("Top_Speed")
Form4.Text11.Text = "NA"
Form4.Text12.Text = "NA"
Form4.Picture1.Picture = LoadPicture(rs.Fields("Picture"))
Form4.Picture1.ScaleMode = 3
Form4.Picture1.AutoRedraw = True
Form4.Picture1.PaintPicture Form4.Picture1.Picture, _
0, 0, Form4.Picture1.ScaleWidth, Form4.Picture1.ScaleHeight, _
0, 0, Form4.Picture1.Picture.Width / 26.46, _
Form4.Picture1.Picture.Height / 26.46
Form4.Picture1.Picture = Form4.Picture1.Image
Form4.Show
Me.Hide
Else
MsgBox ("No records found!")
rs.Close
sqlStr = "select * from Sports_Cars"
rs.Open (sqlStr), conn, adOpenDynamic, adLockOptimistic
End If
End Sub
Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=cars.mdb;Persist Security Info=False"
conn.Open
sqlStr = "select * from Sports_Cars"
rs.Open (sqlStr), conn, adOpenDynamic, adLockOptimistic
End Sub


