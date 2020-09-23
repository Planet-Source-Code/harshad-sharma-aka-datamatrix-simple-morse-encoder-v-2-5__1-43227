Attribute VB_Name = "Readme"
'---------------------------------------------------------------------------------
'   M O R S E   E N C O D E R   v 2.5
'---------------------------------------------------------------------------------
' WHODUNNIT:
'   The whole code presented here is written by    : Harshad "DatamatriX" Sharma
'   (except where mentioned otherwise)
'
' WHAT IT DOES:
'   1. Takes in a string of charachers
'   2. Encodes it into a series of dots (.) and dashes (-) based on
'      International Morse Code
'   3. Plays back the sound (dah-di-dit) from the computer speakers
'
' PURPOSE: To provide a template for building CW apps
'   Basically intended for those interested in HAM.
'   This code is just an example showing ease of programming in VB
'
' DISCLAIMER: This code is provided as-is, without warranty express or
'   implied. The author takes no responsibility for problems arising
'   from the use of this code. You can use it in way you like, but while
'   distributing either forward the set of files as you got them, or if
'   you have modified any file(s) state it that way.
'---------------------------------------------------------------------------------
' CONTACT: If you have any questions or comments, please feel free to
'   My e-mail  : harshad.sharma@bigfoot.com
'   My homepage: http://www20.brinkster.com/tanhadil
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
' W H A T ' S   N E W   I N   V E R S I O N   2 . 5
'---------------------------------------------------------------------------------
' 1. A totally new and (possibly) better interface
' 2. Algorithms fully rewritten for improved efficiency and accurate timing.
' 3. Two way translation capable.
' 4. Corrected the "Inverse Speed" itch!
' 5. Pause function added.

'---------------------------------------------------------------------------------
' W H A T   S T I L L   L U R K S   I N   M Y   M I N D . . .
'---------------------------------------------------------------------------------
' 1. Correct the Frequency Bug :   The output freq of Internal Speaker and
'                                  Soundcard differ considerably
' 2. Provide a better Help System
' 3. Study for my exams
' 4. Get some sleep
' 5. Get some more sleep

'---------------------------------------------------------------------------------
' T H A N K S  1.0E6 . . . (Thanks a Million)
'---------------------------------------------------------------------------------
' I HAVE COLLECTED THE HELPFUL LETTERS AND COMMENTS THAT I RECEIVED AND WHICH HAVE
' PLAYED A HUGE ROLE IN ENCOURAGING ME TO KEEP ON IMPROVING THIS PROJECT...
'---------------------------------------------------------------------------------
' NOTE: If by mistake (which I usually do) I have left out your mail or comment,
' please feel free to inform me... I'll be glad to put it here!
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
' Below is an excerpt from a mail that Jess Hancock sent me...
' He has been very kind and helpful in making my concepts about the morse code
' much more clear than before.
' THANKS JESS!!!
'->>>>----------------------------------------------------------------------------
' The "standard" for a Morse Code word is PARIS and is composed of 50 time units
' including a 7 TU word separation .
'
' Dot = 1 time unit = delta T = dt
' Dash = 3 x Dot = 3 dt
' Space between dot/dash = 1 dt
' Space between characters = 3 dt
' Space between words = 7 dt
'
' Adding dt for all the dits, dahs and spaces gives a total of 50 dt for PARIS.
' 1 (standard) word = 50 dt
' 1 word/min = 50 dt/min = 50 dt / 60 sec
' dt = 1WPM x 60 / 50 = 1.2 sec
'
' For higher speeds, say N wpm, dt must decrease proportionally.
' dt = 1.2/(N WPM) seconds = 1200 / (N WPM) milliseconds
'
' Assume the average ham code speed = 15 WPM so  then ,
' dt (at 15 WPM) =  80 milliseconds,
' at 5 wpm dt = 240 ms
' at 15 wpm dt = 80 ms
' at 30wpm dt = 40 ms
'
' Therefore, a timer needs to have an accurate millisecond time resolution.
'
' You can measure the speed at which your program is sending code by counting the
' number of times the word PARIS is sent in a minute.
'
' 73, Jess, w4pqk
'---------------------------------------------------------------------------------
' Well, I am trying to make this new app as much better as possible... but please
' excuse me if some of the fine settings are unavailable at this very moment... I
' shall try to add them later.
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
' I must also thank David Drane for giving my spirits that extra boost with this
' mail... I now firmly believe that this little app. can be put to some real use..
' THANKS DAVID!!!
'->>>>----------------------------------------------------------------------------
' I downloaded a previous version of your
' software a while ago and it did'nt work with windows 98 se, but I see you
' have fixed that problem since, and it works well. I have since integrated
' your programming into a program of my own. it sends and receives and decodes
' received morse code, and can reply to certain key words, I also integrated
' another program into it to allow it to use my hand held uhf cb to tranceive
' morse via the com port
'---------------------------------------------------------------------------------
' Apart from encouragement, I hope this mail also certifies that
' Simple Morse Encoder can run nicely on Win 98 SE.
'---------------------------------------------------------------------------------

'---------------------------------------------------------------------------------
' H A R S H A D    G R E E T Z: (Harshad Sharma / DatamatriX / datamatr)
'---------------------------------------------------------------------------------
' - All @ sdf.lonestar.org
'   especially vesku, ssinct, hapiworm, disturbd, mungo, lauryn
'   and all the wonderful people there who inspire me everytime!
'
' - My Little Brat Brother: Pritam
' - My First Love: Meghna
' - My Cousins: Amit, Vicky and Nicky
' - Also my friends here: Joy, Ashu, Bonney and Tanmoy
' - Not to forget... all at PlanetSourceCode.com
' - YOU!!!
' Okay, Okay... if you din't find your name, blame the coffee which could not
' keep me awake enough to remember all of my friends... please let me know...
' I'll add you here!!!
'---------------------------------------------------------------------------------
