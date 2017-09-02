Attribute VB_Name = "modLogic"

Public Sub SetPicture(ByVal index As Long)

frmMovie.picPoster.Picture = LoadPicture(App.Path & "/images/" & Movie(index).Picture & ".jpg")


End Sub

Public Sub LoadMovie(ByVal MovieIndex As Long)
frmMovie.Show

frmMovie.picPoster.Picture = Nothing

Call SetPicture(MovieIndex)

With frmMovie
    .lblInfo.Caption = Movie(MovieIndex).Name & " is a " & Movie(MovieIndex).YearMade & " " & Movie(MovieIndex).Genre & " movie directed by " & Movie(MovieIndex).Director & ". It was given an " & Movie(MovieIndex).Rating & " rating due to " & Movie(MovieIndex).RatingReasons
    .lblInfo.Caption = .lblInfo.Caption & vbCrLf & vbCrLf & " On IMDb, it received a rating of about " & Movie(MovieIndex).IMDBRating & "."
    .lblPlot.Caption = "General plot: " & vbCrLf & Movie(MovieIndex).Plot
End With

If Movie(MovieIndex).Watched = True Then
    frmMovie.lblHW.Caption = "Matthew has watched this."
Else
    frmMovie.lblHW.Caption = "Matthew has not watched this yet."
End If

MovieScreen.SCRemakeName = Movie(MovieIndex).RemakeName

If Movie(MovieIndex).RemakeName <> "None" Then
    MovieScreen.SCRemakeYear = Movie(MovieIndex).RemakeYear
    frmMovie.cmdRemake.Visible = True
    If Movie(MovieIndex).Comments(1) = "Remake" Then
        frmMovie.cmdRemake.Caption = "Original"
    Else
        frmMovie.cmdRemake.Caption = "Remake"
    End If
Else
    frmMovie.cmdRemake.Visible = False
End If

MovieScreen.SCPrequal = Movie(MovieIndex).Prequal
MovieScreen.SCSequal = Movie(MovieIndex).Sequal

If MovieScreen.SCPrequal <> "None" Then
    frmMovie.cmdPrequal.Visible = True
Else
    frmMovie.cmdPrequal.Visible = False
End If

If MovieScreen.SCSequal <> "None" Then
    frmMovie.cmdSequal.Visible = True
Else
    frmMovie.cmdSequal.Visible = False
End If

End Sub

Public Sub ReadToMe(ByVal Text As String)
Dim Msg, sapi
Msg = Text
Set sapi = CreateObject("sapi.spvoice")

sapi.speak Msg

End Sub

