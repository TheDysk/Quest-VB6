
# Quest



## <u>History</u>

**Quest** is a text adventure game written by Doug Urquhart, Keith Sheppard and Jerry McCarthy while working at ICL in the late 1970s and early 1980s. The game was released for free in 1983.

A brief history, written by Doug Urquhard, is included in the file **"ICL Quest Adventure Game - History.txt"** 

It was originally written to run on the ICL System 10 Mainframe and over times was ported to ICL System 25, DRS 20, CPM, DOS and Windows.

I first coming across the DOS version game in 1986, which was written in Microsoft BASIC 1.0. 

I spent over 35 years tracking down the game files. I eventually located Doug Urquhard and made contact. He kindly emailed me the Windows version of the game. 

The Windows version had been written in Visual Basic 3 and therefore as a 16bit executable, would not run on modern 32bit and 64bit windows.

I tested the game on WIndows98 in a Virtual Machine and, with Doug's permission, uploaded the files to the Interactive Fiction Archive in the early 2020's.

**https://unbox.ifarchive.org/?url=/if-archive/games/pc/Quest.zip**

Around the same time I discovered I had, in 1986, printed a copy of the BASIC Source Code so set about recreating a digital copy of the source code.

A few years later I decided to try and convert the VB3 game into VB6 with the aim of creating a 32bit executable that would run on Modern Windows.

I located a VB3 executable decompiler application and extracted the VB3 project files, ported these into VB6, resolved any conflicts and added some additional  features such as a Debug windows that shows all the records that have been read. The debug feature is included to aid in any future conversion to other programming languages or computer systems.

The decompiled project files do not contain the original variables names. The original variables names were replaced by sudo-variable names by the original VB3 compiler. I have compared the VB6 code to the original BASIC listing and have resolved most of the variable names. There are some outstanding that were added by the original VB3 developer and require further analysis to identify their purpose and assign sensible names. 

Due to differences between the way BASIC V1.0 and Visual BASIC handles variables, some variables have had to be renames as the same name was being used for both String and Integer variables. Something that BASIC V1.o allows but VB does not. These have been listed in the **"Variables names changed to avoid conflicts.txt"** file.

## <u>Version History</u>

**25/01/2026 - V1.00** - Game ported from Visual Basic 3 to Visual Basic 6. Replacement of Sudo-variable names with correct names from the original BASIC source code. Addition of Debug functionality. 32bit executable compiled 

**31/01/2026 - V1.10** - Added text justification options for game text, moved game command line onto main form, and moved the "Show Debug Form" option onto the main form. 

## <u>Running the Game</u>

The run the game download the zip file called **"QUEST (32bit) - game files.zip"** and extract the Executable and the two datafiles to any folder. Run the exe file and enjoy.

