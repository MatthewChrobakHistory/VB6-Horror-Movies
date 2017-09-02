Attribute VB_Name = "modCAT"
Option Explicit

Public Const Max_Movies As Long = 150

Public Movie(1 To Max_Movies) As MovieRec
Public MovieScreen As ScreenRec

Private Type MovieRec
    Name As String
    YearMade As String
    Director As String
    IMDBRating As Byte
    Comments(1 To 5) As String
    Picture As String
    Rating As String
    RatingReasons As String
    Genre As String
    Plot As String
    Sequal As String
    Prequal As String
    Watched As Boolean
    RemakeName As String
    RemakeYear As String
End Type


Private Type ScreenRec
    SCSequal As String
    SCPrequal As String
    SCRemakeName As String
    SCRemakeYear As String
End Type
