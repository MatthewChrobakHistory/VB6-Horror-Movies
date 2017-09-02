Attribute VB_Name = "modMovies"
'setting movie information
Option Explicit

Public Sub LoadMovies()
Dim i As Long

For i = 1 To Max_Movies
    Select Case i
    
        Case 1
            Movie(1).Name = "Friday the 13th"
            Movie(1).YearMade = "1980"
            Movie(1).Director = "Sean S. Cunningham"
            Movie(1).IMDBRating = "6.4"
            Movie(1).Comments(1) = "test"
            Movie(1).Picture = 1
            Movie(1).Rating = "R"
            Movie(1).RatingReasons = "intense gore, sexual content, and language."
            Movie(1).Genre = "Horror"
            Movie(1).Plot = "Camp counselors are stalked and murdered by an unknown assailant while trying to reopen a summer camp that was the site of a child's drowning."
            Movie(1).Prequal = "None"
            Movie(1).Sequal = "Friday the 13th Part 2"
            Movie(1).Watched = True
            Movie(1).RemakeName = "Friday the 13th"
            Movie(1).RemakeYear = "2009"
        Case 2
            Movie(2).Name = "Friday the 13th Part 2"
            Movie(2).YearMade = "1981"
            Movie(2).Director = "Steve Miner"
            Movie(2).IMDBRating = "5.8"
            Movie(2).Picture = 2
            Movie(2).Rating = "R"
            Movie(2).RatingReasons = "intense gore, sexual content, and language."
            Movie(2).Genre = "Horror"
            Movie(2).Plot = "Mrs. Voorhees is dead, and Camp Crystal Lake is shut down, but a camp next to the infamous place is stalked by an unknown assailant. Is it Mrs. Voorhees' son Jason who didn't drown in the lake some 30 years before?"
            Movie(2).Prequal = "Friday the 13th"
            Movie(2).Sequal = "Friday the 13th Part 3"
            Movie(2).Watched = True
            Movie(2).RemakeName = "None"
        Case 3
            Movie(3).Name = "Friday the 13th Part 3"
            Movie(3).YearMade = "1982"
            Movie(3).Director = "Steve Minor"
            Movie(3).IMDBRating = "5.3"
            Movie(3).Picture = 3
            Movie(3).Rating = "R"
            Movie(3).RatingReasons = "intense gore, sexual content, and language."
            Movie(3).Genre = "Horror"
            Movie(3).Plot = "Having escaped, Jason Voorhees is back, hockey mask and all, to continue his murderous rampage across Camp Crystal Lake."
            Movie(3).Prequal = "Friday the 13th Part 2"
            Movie(3).Sequal = "Friday the 13th: The Final Chapter"
            Movie(3).Watched = True
            Movie(3).RemakeName = "None"
        Case 4
            Movie(4).Name = "Friday the 13th: The Final Chapter"
            Movie(4).YearMade = "1984"
            Movie(4).Director = "Joseph Zito"
            Movie(4).IMDBRating = "5.5"
            Movie(4).Picture = 4
            Movie(4).Rating = "R"
            Movie(4).RatingReasons = "intense gore, sexual content, and language."
            Movie(4).Genre = "Horror"
            Movie(4).Plot = "After being mortally wounded and taken to the morgue, murderer Jason Voorhees spontaneously revives and embarks on a killing spree as he makes his way back to his home at Camp Crystal Lake."
            Movie(4).Prequal = "Friday the 13th Part 3"
            Movie(4).Sequal = "Friday the 13th: A New Beginning"
            Movie(4).Watched = False
            Movie(4).RemakeName = "None"
        Case 5
            Movie(5).Name = "Friday the 13th: A New Beginning"
            Movie(5).YearMade = "1985"
            Movie(5).Director = "Danny Steinmann"
            Movie(5).IMDBRating = "4.3"
            Movie(5).Picture = 5
            Movie(5).Rating = "R"
            Movie(5).RatingReasons = "intense gore, sexual content, and language."
            Movie(5).Genre = "Horror"
            Movie(5).Plot = "Still haunted by his gruesome past, Tommy Jarvis - the boy who killed Jason Voorhees - wonders if somehow he is connected to brutal slayings occurring in and around the secluded halfway house where he now lives."
            Movie(5).Prequal = "Friday the 13th: The Final Chapter"
            Movie(5).Sequal = "Friday the 13th Part 6"
            Movie(5).Watched = False
            Movie(5).RemakeName = "None"
        Case 6
            Movie(6).Name = "Friday the 13th Part 6"
            Movie(6).YearMade = "1986"
            Movie(6).Director = "Tom McLoughlin"
            Movie(6).IMDBRating = "5.5"
            Movie(6).Picture = 6
            Movie(6).Rating = "R"
            Movie(6).RatingReasons = "intense gore, sexual content, and language."
            Movie(6).Genre = "Horror"
            Movie(6).Plot = "Tommy Jarvis goes to the graveyard to get rid of Jason Voorhees' body, but he accidentally brings him back to life. Jason wants revenge and Tommy must defeat him once and for all!"
            Movie(6).Prequal = "Friday the 13th: A New Beginning"
            Movie(6).Sequal = "Friday the 13th: The New Blood"
            Movie(6).Watched = False
            Movie(6).RemakeName = "None"
        Case 7
            Movie(7).Name = "Friday the 13th: The New Blood"
            Movie(7).YearMade = "1988"
            Movie(7).Director = "John Carl Buechler"
            Movie(7).IMDBRating = "4.8"
            Movie(7).Picture = 7
            Movie(7).Rating = "R"
            Movie(7).RatingReasons = "intense gore, sexual content, and language."
            Movie(7).Genre = "Horror"
            Movie(7).Plot = "Years after Tommy Jarvis chained him underwater at Camp Crystal Lake, the hulking killer Jason Voorhees returns to the camp grounds when he is released accidentally by a teenager with psychic powers."
            Movie(7).Prequal = "Friday the 13th Part 6"
            Movie(7).Sequal = "Friday the 13th: Jason Takes Manhattan"
            Movie(7).Watched = False
            Movie(7).RemakeName = "None"
        Case 8
            Movie(8).Name = "Friday the 13th: Jason Takes Manhattan"
            Movie(8).YearMade = "1989"
            Movie(8).Director = "Rob Hedden"
            Movie(8).IMDBRating = "4.2"
            Movie(8).Picture = 8
            Movie(8).Rating = "R"
            Movie(8).RatingReasons = "intense gore, sexual content, and language."
            Movie(8).Genre = "Horror"
            Movie(8).Plot = "A passing boat bound for New York pulls Jason Voorhees along for the ride. Look out New York, here comes hell in a hockey mask."
            Movie(8).Prequal = "Friday the 13th: The New Blood"
            Movie(8).Sequal = "Jason Goes to Hell: The Final Friday"
            Movie(8).Watched = False
            Movie(8).RemakeName = "None"
        Case 9
            Movie(9).Name = "Jason Goes to Hell: The Final Friday"
            Movie(9).YearMade = "1993"
            Movie(9).Director = "Adam Marcus"
            Movie(9).IMDBRating = "4.2"
            Movie(9).Picture = 9
            Movie(9).Rating = "R"
            Movie(9).RatingReasons = "intense gore, sexual content, and language."
            Movie(9).Genre = "Horror"
            Movie(9).Plot = "Serial killer Jason Voorhees' supernatural origins are revealed."
            Movie(9).Prequal = "Friday the 13th: Jason Takes Manhattan"
            Movie(9).Sequal = "Jason X"
            Movie(9).Watched = True
            Movie(9).RemakeName = "None"
        Case 10
            Movie(10).Name = "Jason X"
            Movie(10).YearMade = "2001"
            Movie(10).Director = "James Isaac"
            Movie(10).IMDBRating = "4.3"
            Movie(10).Picture = 10
            Movie(10).Rating = "R"
            Movie(10).RatingReasons = "intense gore, sexual content, and language."
            Movie(10).Genre = "Horror"
            Movie(10).Plot = "Jason Voorhees returns with a new look, a new machete, and his same murderous attitude as he is awakened on a spaceship in the 25th century."
            Movie(10).Prequal = "Jason Goes to Hell: The Final Friday"
            Movie(10).Sequal = "Freddy vs. Jason"
            Movie(10).Watched = True
            Movie(10).RemakeName = "None"
        Case 11
            Movie(11).Name = "Freddy vs. Jason"
            Movie(11).YearMade = "2003"
            Movie(11).Director = "Ronny Yu"
            Movie(11).IMDBRating = "5.7"
            Movie(11).Picture = 11
            Movie(11).Rating = "R"
            Movie(11).RatingReasons = "intense gore, sexual content, and language."
            Movie(11).Genre = "Horror"
            Movie(11).Plot = "Freddy Krueger and Jason Voorhees return to terrorize the teenage population. Except this time, they're out to get each other, too."
            Movie(11).Prequal = "Jason X"
            Movie(11).Sequal = "None"
            Movie(11).Watched = True
            Movie(11).RemakeName = "None"
        Case 12
            Movie(12).Name = "Friday the 13th"
            Movie(12).YearMade = "2009"
            Movie(12).Director = "Marcus Nispel"
            Movie(12).IMDBRating = "5.5"
            Movie(12).Picture = 12
            Movie(12).Rating = "R"
            Movie(12).RatingReasons = "intense gore, sexual content, and language."
            Movie(12).Genre = "Horror"
            Movie(12).Plot = "A group of young adults discover a boarded up Camp Crystal Lake, where they soon encounter Jason Voorhees and his deadly intentions."
            Movie(12).Prequal = "None"
            Movie(12).Sequal = "None"
            Movie(12).Watched = True
            Movie(12).RemakeName = "Friday the 13th"
            Movie(12).RemakeYear = "1980"
            Movie(12).Comments(1) = "Remake"
        Case 13
            Movie(13).Name = "A Nightmare on Elm Street"
            Movie(13).YearMade = "1984"
            Movie(13).Director = "Wes Craven"
            Movie(13).IMDBRating = "7.5"
            Movie(13).Picture = 13
            Movie(13).Rating = "R"
            Movie(13).RatingReasons = "gore, sexual content, and some language."
            Movie(13).Genre = "Horror"
            Movie(13).Plot = "In the dreams of his victims, a spectral child murderer stalks the children of the members of the lynch mob that killed him."
            Movie(13).Prequal = "None"
            Movie(13).Sequal = "A Nightmare on Elm Street 2: Freddy's Revenge"
            Movie(13).Watched = True
            Movie(13).RemakeName = "A Nightmare on Elm Street"
            Movie(13).RemakeYear = "2010"
        Case 14
            Movie(14).Name = "A Nightmare on Elm Street 2: Freddy's Revenge"
            Movie(14).YearMade = "1985"
            Movie(14).Director = "Jack Sholder"
            Movie(14).IMDBRating = "5.1"
            Movie(14).Picture = 14
            Movie(14).Rating = "R"
            Movie(14).RatingReasons = "gore, sexual content, and some language."
            Movie(14).Genre = "Horror"
            Movie(14).Plot = "A teenage boy is haunted in his dreams by Freddy Krueger who is out to possess him to continue his murdering in the real world."
            Movie(14).Prequal = "A Nightmare on Elm Street"
            Movie(14).Sequal = "A Nightmare on Elm Street 3: Dream Warriors"
            Movie(14).Watched = True
            Movie(14).RemakeName = "None"
        Case 15
            Movie(15).Name = "A Nightmare on Elm Street 3: Dream Warriors"
            Movie(15).YearMade = "1987"
            Movie(15).Director = "Chuck Russell"
            Movie(15).IMDBRating = "6.3"
            Movie(15).Picture = 15
            Movie(15).Rating = "R"
            Movie(15).RatingReasons = "gore, and some language."
            Movie(15).Genre = "Horror"
            Movie(15).Plot = "Survivors of undead serial killer Freddy Krueger - who stalks his victims in their dreams - learn to take control of their own dreams in order to fight back."
            Movie(15).Prequal = "A Nightmare on Elm Street 2: Freddy's Revenge"
            Movie(15).Sequal = "A Nightmare on Elm Street 4: The Dream Master"
            Movie(15).Watched = True
            Movie(15).RemakeName = "None"
        Case 16
            Movie(16).Name = "A Nightmare on Elm Street 4: The Dream Master"
            Movie(16).YearMade = "1988"
            Movie(16).Director = "Renny Harln"
            Movie(16).IMDBRating = "5.4"
            Movie(16).Picture = 16
            Movie(16).Rating = "R"
            Movie(16).RatingReasons = "unknown reasons."
            Movie(16).Genre = "Horror"
            Movie(16).Plot = "Freddy Krueger returns once again to terrorize the dreams of the remaining Dream Warriors, as well as those of a young woman who may know the way to defeat him for good."
            Movie(16).Prequal = "A Nightmare on Elm Street 3: Dream Warriors"
            Movie(16).Sequal = "A Nightmare on Elm Street 5: The Dream Child"
            Movie(16).Watched = False
            Movie(16).RemakeName = "None"
        Case 17
            Movie(17).Name = "A Nightmare on Elm Street 5: The Dream Child"
            Movie(17).YearMade = "1989"
            Movie(17).Director = "Stephen Hopkins"
            Movie(17).IMDBRating = "4.9"
            Movie(17).Picture = 17
            Movie(17).Rating = "R"
            Movie(17).RatingReasons = "unknown reasons."
            Movie(17).Genre = "Horror"
            Movie(17).Plot = "Alice, having survived the previous installment of the Nightmare series, finds the deadly dreams of Freddy Krueger starting once again. This time, the taunting murderer is striking through the sleeping mind of Alice's unborn child. His intention is to be 'born again' into the real world. The only one who can stop Freddy is his dead mother, but can Alice free her spirit in time to save her own son?"
            Movie(17).Prequal = "A Nightmare on Elm Street 4: The Dream Master"
            Movie(17).Sequal = "Freddy's Dead: The Final Nightmare"
            Movie(17).Watched = False
            Movie(17).RemakeName = "None"
        Case 18
            Movie(18).Name = "Freddy's Dead: The Final Nightmare"
            Movie(18).YearMade = "1991"
            Movie(18).Director = "Rachel Talalay"
            Movie(18).IMDBRating = "4.7"
            Movie(18).Picture = 18
            Movie(18).Rating = "R"
            Movie(18).RatingReasons = "unknown reasons."
            Movie(18).Genre = "Horror"
            Movie(18).Plot = "Freddy Krueger returns once again to haunt both the dreams of his daughter and Springwood's last surviving teenager."
            Movie(18).Prequal = "A Nightmare on Elm Street 5: The Dream Child"
            Movie(18).Sequal = "New Nightmare"
            Movie(18).Watched = False
            Movie(18).RemakeName = "None"
        Case 19
            Movie(19).Name = "New Nightmare"
            Movie(19).YearMade = "1994"
            Movie(19).Director = "Wes Craven"
            Movie(19).IMDBRating = "6.3"
            Movie(19).Picture = 19
            Movie(19).Rating = "R"
            Movie(19).RatingReasons = "unknown reasons."
            Movie(19).Genre = "Horror"
            Movie(19).Plot = "A demonic force has chosen Freddy Krueger as its portal to the real world. Can Heather play the part of Nancy one last time and trap the evil trying to enter our world?"
            Movie(19).Prequal = "Freddy's Dead: The Final Nightmare"
            Movie(19).Sequal = "Freddy vs. Jason"
            Movie(19).Watched = False
            Movie(19).RemakeName = "None"
        Case 20
            With Movie(20)
                .Name = "A Nightmare on Elm Street"
                .YearMade = "2010"
                .Director = "Samuel Bayer"
                .IMDBRating = "5.1"
                .Picture = 20
                .Rating = "R"
                .RatingReasons = "intense gore, sexual content, and language."
                .Genre = "Horror"
                .Plot = "A re-imagining of the horror icon Freddy Krueger, a serial-killer who wields a glove with four blades embedded in the fingers and kills people in their dreams, resulting in their real death in reality."
                .Prequal = "None"
                .Sequal = "None"
                .Watched = True
                .RemakeName = "A Nightmare on Elm Street"
                .RemakeYear = "1984"
                .Comments(1) = "Remake"
            End With
        Case 21
            With Movie(21)
                .Name = "The Burning"
                .YearMade = "1981"
                .Director = "Tony Maylam"
                .IMDBRating = "6.2"
                .Picture = 21
                .Rating = "R"
                .RatingReasons = "gore, and sexual content."
                .Genre = "Thriller"
                .Plot = "A former summer camp caretaker, horribly burned from a prank gone wrong, lurks around an upstate New York summer camp bent on killing the teenagers responsible for his disfigurement."
                .Prequal = "None"
                .Sequal = "None"
                .Watched = True
                .RemakeName = "None"
            End With
        Case 22
            With Movie(22)
                .Name = "The Grudge"
                .YearMade = "2004"
                .Director = "Takashi Shimizu"
                .IMDBRating = "5.8"
                .Picture = 22
                .Rating = "PG-13"
                .RatingReasons = "frightening scenes."
                .Genre = "Horror"
                .Plot = "An American nurse living and working in Tokyo is exposed to a mysterious supernatural curse, one that locks a person in a powerful rage before claiming their life and spreading to another victim."
                .Prequal = "None"
                .Sequal = "The Grudge 2"
                .Watched = True
                .RemakeName = "Ju On: The Grudge"
                .RemakeYear = "2002"
                .Comments(1) = "Remake"
            End With
        Case 23
            With Movie(23)
                .Name = "The Grudge 2"
                .YearMade = "2006"
                .Director = "Takashi Shimizu"
                .IMDBRating = "4.7"
                .Picture = 23
                .Rating = "PG-13"
                .RatingReasons = "frightening scenes."
                .Genre = "Horror"
                .Plot = "In Tokyo, a young woman (Tamblyn) is exposed to the same mysterious curse that afflicted her sister (Gellar). The supernatural force, which fills a person with rage before spreading to its next victim, brings together a group of previously unrelated people who attempt to unlock its secret to save their lives."
                .Prequal = "The Grudge"
                .Sequal = "The Grudge 3"
                .Watched = True
                .RemakeName = "Ju On: The Grudge 2"
                .RemakeYear = "2003"
                .Comments(1) = "Remake"
            End With
        Case 24
            With Movie(24)
                .Name = "The Grudge 3"
                .YearMade = "2009"
                .Director = "Toby Wilkins"
                .IMDBRating = "4.5"
                .Picture = 24
                .Rating = "PG-13"
                .RatingReasons = "frightening scenes."
                .Genre = "Horror"
                .Plot = "A young Japanese woman who holds the key to stopping the evil spirit of Kayako, travels to the haunted Chicago apartment from the sequel, to stop the curse of Kayako once and for all and save a family who are currently being haunted by her malicious spirit."
                .Prequal = "The Grudge 2"
                .Sequal = "None"
                .Watched = True
                .RemakeName = "None"
            End With
        Case 25
            With Movie(25)
                .Name = "Ju On"
                .YearMade = "2000"
                .Director = "Takashi Shimizu"
                .IMDBRating = "6.8"
                .Picture = 25
                .Rating = "PG-13"
                .RatingReasons = "frightening scenes"
                .Genre = "Horror"
                .Plot = "Not given"
                .Prequal = "None"
                .Sequal = "Ju On: The Grudge"
                .Watched = True
                .RemakeName = "None"
            End With
        Case 26
            With Movie(26)
                .Name = "Ju On: The Grudge"
                .YearMade = "2002"
                .Director = "Takashi Shimizu"
                .IMDBRating = "6.6"
                .Picture = 26
                .Rating = "R"
                .RatingReasons = "frightening scenes."
                .Genre = "Horror"
                .Plot = "A mysterious and vengeful spirit marks and pursues anybody who dares enter the house in which it resides."
                .Prequal = "Ju On"
                .Sequal = "Ju On: The Grudge 2"
                .Watched = False
                .RemakeName = "The Grudge"
                .RemakeYear = "2004"
            End With
        Case 27
            With Movie(27)
                .Name = "Ju On: The Grudge 2"
                .YearMade = "2003"
                .Director = "Takashi Shimizu"
                .IMDBRating = "6.2"
                .Picture = 27
                .Rating = "R"
                .RatingReasons = "frightening scenes."
                .Genre = "Horror"
                .Plot = "While driving , the pregnant horror-movie actress Kyôko Harase and her fiancé are in a car crash caused by the Toshio's friend. Kyôko loses her baby and her fiancé winds up in a coma. Kyôko was cursed together with a television crew when they shot a show in the haunted house where Kayako was brutally murdered by her husband years ago. While each member of the team dies or disappears, Kyôko is informed that she has a three-and-a-half-month-old fetus in her womb."
                .Prequal = "Ju On: The Grudge"
                .Sequal = "None"
                .Watched = False
                .RemakeName = "The Grudge 2"
                .RemakeYear = "2006"
            End With
        Case 28
            With Movie(28)
                .Name = "Halloween"
                .YearMade = "1978"
                .Director = "John Carpenter"
                .IMDBRating = "7.9"
                .Picture = 28
                .Rating = "R"
                .RatingReasons = "violence, sexual content, and some language."
                .Genre = "Thriller"
                .Plot = "A psychotic murderer institutionalized since childhood for the murder of his sister, escapes and stalks a bookish teenage girl and her friends while his doctor chases him through the streets."
                .Prequal = "None"
                .Sequal = "Halloween II"
                .Watched = False
                .RemakeName = "Halloween"
                .RemakeYear = "2007"
            End With
        Case 29
            With Movie(29)
                .Name = "Halloween II"
                .YearMade = "1981"
                .Director = "Rick Rosenthal"
                .IMDBRating = "6.4"
                .Picture = 29
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "Thriller"
                .Plot = "Laurie Strode is rushed to the hospital, while Sheriff Brackett and Dr. Loomis hunt the streets for Michael Myers, who has found Laurie at the Haddonfield Hospital."
                .Prequal = "Halloween"
                .Sequal = "Halloween III: Season of the Witch"
                .Watched = False
                .RemakeName = "Halloween 2"
                .RemakeYear = "2009"
            End With
        Case 30
            With Movie(30)
                .Name = "Halloween III: Season of the Witch"
                .YearMade = "1982"
                .Director = "Tommy Lee Wallace"
                .IMDBRating = "4.1"
                .Picture = 30
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "Unknown"
                .Plot = "A large Halloween mask-making company has plans to kill millions of American children with something sinister hidden in Halloween masks."
                .Prequal = "Halloween II"
                .Sequal = "Halloween 4: The Return of Michael Myers"
                .Watched = False
                .RemakeName = "None"
            End With
        Case 31
            With Movie(31)
                .Name = "Halloween 4: The Return of Michael Myers"
                .YearMade = "1988"
                .Director = "Dwight H. Little"
                .IMDBRating = "4.7"
                .Picture = 31
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "unknown"
                .Plot = "Ten years after his original massacre, the invalid Michael Myers awakens and returns to Haddonfield to kill his seven-year-old niece on Halloween. Can Dr. Loomis stop him?"
                .Prequal = "Halloween III: Season of the Witch"
                .Sequal = "Halloween 5: The Revenge of Michael Myers"
                .Watched = False
                .RemakeName = "None"
            End With
        Case 32
            With Movie(32)
                .Name = "Halloween 5: The Revenge of Michael Myers"
                .YearMade = "1989"
                .Director = "Dominique Othenin-Girard"
                .IMDBRating = "4.8"
                .Picture = 32
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "unknown"
                .Plot = "It's one year later after the events of Halloween 4. Michael survives the shootings and on October 31st he returns with a vengeance. Lurking and stalking, Jamie, Rachel, and Rachel's friends, Michael forms a plan to lure Jamie out of the children's hospital where events lead up to the confrontation at the Myers house"
                .Prequal = "Halloween 4: The Return of Michael Myers"
                .Sequal = "Halloween: The Curse of Michael Myers"
                .Watched = False
                .RemakeName = "None"
            End With
        Case 33
            With Movie(33)
                .Name = "Halloween: The Curse of Michael Myers"
                .YearMade = "1995"
                .Director = "Joe Chappelle"
                .IMDBRating = "4.6"
                .Picture = 33
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "unknown"
                .Plot = "Six years ago, Michael Myers terrorized the town of Haddonfield, Illinois. He and his niece, Jamie Lloyd, have disappeared. Jamie was kidnapped by a bunch of evil druids who protect Michael Myers. And now, six years later, Jamie has escaped after giving birth to Michael's child. She runs to Haddonfield to get Dr. Loomis to help her again. Meanwhile, the family that adopted Laurie Strode is living in the Myers house and are being stalked by Myers. It's the curse of Thorn that Michael is possessed by that makes him kill his family. And it's up to Tommy Doyle, the boy from Halloween, and Dr. Loomis, to stop them all."
                .Prequal = "Halloween 5: The Revenge of Michael Myers"
                .Sequal = "Halloween H20: 20 Years Later"
                .Watched = False
                .RemakeName = "None"
            End With
        Case 34
            With Movie(34)
                .Name = "Halloween H20: 20 Years Later"
                .YearMade = "1998"
                .Director = "Steve Miner"
                .IMDBRating = "5.4"
                .Picture = 34
                .Rating = "R"
                .RatingReasons = "unknown reasons."
                .Genre = "unknown"
                .Plot = "Laurie Strode, now the dean of a Northern California private school with an assumed name, must battle the Shape one last time and now the life of her own son hangs in the balance."
                .Prequal = "Halloween: The Curse of Michael Myers"
                .Sequal = "Halloween: Resurrection"
                .Watched = False
                .RemakeName = "None"
            End With
        Case 35
            With Movie(35)
                .Name = "Halloween"
                .YearMade = "2007"
                .Director = "Rob Zombie"
                .IMDBRating = "5.5"
                .Picture = 35
                .Rating = "R"
                .RatingReasons = ""
                .Genre = "Thriller"
                .Plot = ""
                .Prequal = "None"
                .Sequal = "Halloween 2"
                .Watched = True
                .RemakeName = "Halloween"
                .RemakeYear = "1978"
                .Comments(1) = "Remake"
            End With
        Case 36
            With Movie(36)
                .Name = "Halloween 2"
                .YearMade = "2009"
                .Director = "Rob Zombie"
                
            End With
                
    End Select
Next

End Sub
