VERSION 5.00
Begin VB.Form frmMovie 
   Caption         =   "Movie"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemake 
      Caption         =   "Remake"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrequal 
      Caption         =   "Prequal"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSequal 
      Caption         =   "Sequal"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox picPoster 
      BorderStyle     =   0  'None
      Height          =   4755
      Left            =   120
      ScaleHeight     =   4755
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   120
      Width           =   3210
   End
   Begin VB.Label lblHW 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "* click the text you wish to be read to you *"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label lblPlot 
      Height          =   1935
      Left            =   3480
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      Height          =   1335
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrequal_Click()
Dim i As Long

For i = 1 To Max_Movies
    If Movie(i).Name = MovieScreen.SCPrequal Then
    Call LoadMovie(i)
    Exit Sub
    End If
Next
End Sub

Private Sub cmdRemake_Click()
Dim i As Long

For i = 1 To Max_Movies
    If Movie(i).Name = MovieScreen.SCRemakeName Then
        If Movie(i).YearMade = MovieScreen.SCRemakeYear Then
            Call LoadMovie(i)
            Exit Sub
        End If
    End If
Next

End Sub

Private Sub cmdSequal_Click()
Dim i As Long

For i = 1 To Max_Movies
    If Movie(i).Name = MovieScreen.SCSequal Then
    Call LoadMovie(i)
    Exit Sub
    End If
Next

End Sub

Private Sub Form_Load()
'movie info
End Sub

Private Sub lblInfo_Click()

Call ReadToMe(lblInfo.Caption)

End Sub

Private Sub lblPlot_Click()

Call ReadToMe(lblPlot.Caption)

End Sub
