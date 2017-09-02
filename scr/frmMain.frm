VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "client"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNYW 
      Caption         =   "Not yet watched"
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View Movie"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Movie"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDirector 
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstIndex 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "to"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "IMDb rating:"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Director:"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Year:"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSearch_Click()
Dim i As Long
Dim GoodMovie As Boolean

lstIndex.Clear
cmdView.Caption = "Repop List"

For i = 1 To Max_Movies
    GoodMovie = True
    
    If Movie(i).Name = "" Then
        GoodMovie = False
    End If
    
    If txtName.Text <> "" Then
        If Movie(i).Name <> txtName.Text Then
            GoodMovie = False
        End If
    End If
    
    If txtMin.Text <> "" And txtMax.Text <> "" Then
        If Movie(i).IMDBRating < txtMin.Text Or Movie(i).IMDBRating > txtMax.Text Then
            GoodMovie = False
        End If
    End If
    
    If txtYear.Text <> "" Then
        If Movie(i).YearMade <> txtYear.Text Then
            GoodMovie = False
        End If
    End If
    
    If chkNYW.Value = 1 Then
        If Movie(i).Watched = True Then
            GoodMovie = False
        End If
    Else
        If Movie(i).Watched = False Then
            GoodMovie = False
        End If
    End If
    
    If GoodMovie = True Then
        lstIndex.AddItem i & ": " & Movie(i).Name
    End If

Next

End Sub

Private Sub cmdView_Click()
Dim i As Long

If cmdView.Caption = "Repop List" Then
    lstIndex.Clear
    For i = 1 To Max_Movies
        If Movie(i).Name <> "" Then
            lstIndex.AddItem i & ": " & Movie(i).Name
        End If
    Next
    cmdView.Caption = "View Movie"
    Exit Sub
End If

If lstIndex.ListIndex < 0 Then Exit Sub

If Movie(lstIndex.ListIndex + 1).Picture = "" Then Exit Sub

Call LoadMovie(lstIndex.ListIndex + 1)

End Sub

Private Sub Form_Load()
Dim i As Long
'looking up horror movies
Call LoadMovies

For i = 1 To Max_Movies
    If Movie(i).Name <> "" Then
        lstIndex.AddItem i & ": " & Movie(i).Name
    End If
Next

End Sub

Private Sub txtMax_Change()

If txtMax.Text = "" Then Exit Sub

If IsNumeric(txtMax.Text) = False Then
    txtMax.Text = "1"
End If

If txtMax.Text > 10 Then txtMax.Text = "10"

End Sub

Private Sub txtMin_Change()

If txtMin.Text = "" Then Exit Sub

If IsNumeric(txtMin.Text) = False Then
    txtMin.Text = "1"
End If

If txtMin.Text > 10 Then txtMin.Text = "10"

End Sub
