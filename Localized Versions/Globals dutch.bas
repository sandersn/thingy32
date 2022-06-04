Attribute VB_Name = "Globals"
Option Explicit
'(started 13 Sept 1999)
'update from Jan2K, have added a number of vars to aid me in the new features:triple hex and double byte line/section breaks.
'globals:the variables that have to be visible to ALL forms...yes I hate having to do this but you have
'to have something like this I suppose
'23 Jan 2000
    'I have added many string constants that cover all text in thingy32. After I debug these, I will change these to Spanish and release a
'Spanish version.
'29 Mar 2001
    'added some more file types for game boy(color and advance)
Public trplmax As Integer   'the triple-bytes. I don't know what people will use them for; it's a request.
Public trpl(1 To 3000) As String
Public trplhex(1 To 3000) As String

Public dblmax As Integer   'the kanji holders
Public dbl(1 To 3000) As String
Public dblhex(1 To 3000) As String

Public commax As Integer
Public combo(1 To 300) As String   'beats me(I think singles)
Public comhex(1 To 300) As String

Public maxmix As Integer
Public newlhex As String    'the newline signal
Public newlhexdbl As String    'the auxiliary double-byte for newlines.

'String Constants section:
'these are the Constants that determine what language you are using. Translate them into xxx language and thingy32 should
'be in that language. I didn't want to use a .rc file(they're not native to VB) so this is the alternative.
        'These strings used in frmThingy32:
    'shortcut keys
Public Const JumpKey = "G"
Public Const TableSwitchKey = "T"
Public Const EditSelKey = " "
Public Const SearchKey = "Z"
Public Const ReSearchKey = "R"
Public Const DumpKey = "D"
Public Const InsertKey = "I"
Public Const HideHexKey = "H"
Public Const OptionsKey = "O"
Public Const MoreWidthKey = "."
Public Const LessWidthKey = ","
    'shortcut labels(all are visual elements)
Public Const JumpLabel = "G = Ga naar..."
Public Const OptionsLabel = "O = Opties"
Public Const EditSelLabel = "Space = Kiezen om te bewerken"
Public Const SearchLabel = "Z = Zoeken"
Public Const ReSearchLabel = "R = Opnieuw-zoeken"
Public Const DumpLabel = "D = Dump"
Public Const InsertLabel = "I = Invoegen"
Public Const PositionLabel = "Positie: "
Public Const MoreWidthLabel = "Meer Brete = >"
Public Const LessWidthLabel = "< = Minder Brete"
    'menus
Public Const OpenNewFile = "Open &Nieuwe File..."
Public Const OpenNewTable = "&Open Nieuw Table..."
Public Const ReloadTable = "&Herlaad Deze Tabel"
Public Const HideHex = "&Verberg Hex"
Public Const TableSwitch = "Kies tussen &Tabel Files"
Public Const JumpManual = "&Spring handmatig..."
Public Const DumpManual = "&Dump handmatig met deze locatie als start gebruiken"
Public Const InsertManual = "&Invoegen mat huidige locatie als start punt"
Public Const AddNewBookMark = "&Gebruik huidige locatie als nieuwe Bookmark..."
Public Const AddNewDumpMark = "&Gebruik huidige locatie als nieuwe Dumpmark..."
Public Const AddNewInsertMark = "&Gebruik huidige locatie als nieuwe Insertmark..."
    'misc messages
Public Const LabelTooltip = "Click op de labels om de actie uit te voeren" 'visual element
Public Const OptionsTooltip = "Click voor opties menu"
Public Const ReadyMsg = "Klaar voor gebruik"
Public Const UsingTableMsg = "Nu Tabel # in gebruik"
Public Const SelectingMsg = "Selecteren"
Public Const CurrentSearch = "Huidige zoekactie: "
Public Const SelDumpEnd = "Druk op " & DumpKey & " aan het einde van de dump"
Public Const SelInsertEnd = "Druk op " & InsertKey & " aan het einde van de dump"
        'misc--load
Public Const ErrorLoadingDataFromCommand = "Niet-bestaande data file of pad in shortcut! Je moet de file normaal laden." & vbCrLf & _
"Waarschuwing: Je moet de tabel files ook handmatig laden, zelfs als dat pad WEL bestaat."
Public Const ErrorLoadingDataTitle = "Error door file load"
Public Const ErrorLoadingTableFromCommand = "Niet-besstaande tabel file of pad in shortcut! Je moet de file normaal laden."
Public Const ErrorLoadingTableTitle = "Error door tabel load"
Public Const ErrorLoadingTable2Title = "Error door 2e tabel load"
    'app info
Public Const InfoTitle = "thingy32 Applicatie Informatie:"
Public Const InfoName = "Officieele naam: "
Public Const InfoVersion = "Versie: "
Public Const InfoComments = "Commentaar voor deze versie: "
Public Const InfoCopyright = "Copyright Informatie: "
Public Const InfoTrademarks = "Trademark Informatie: "
    'program start sequence(in LoadFiles mainly)
Public Const OpenDataFile = "Open Data File"
Public Const DataFileTypes = "Alle Files(*.*)|*.*|Snes Files(*.smc,*.swc,*.sfc,*.fig, and others)|*.smc;*.swc;*.sfc;*.fig;*.058;*.078;*.1;*.2;*.3|Nes Files(*.nes)|*.nes|Genesis/Megadrive/32x Files(*.smd,*.bin,*.32x)|*.smd;*.bin;*.32x|Master System Files(*.sms)|*.sms|Game Gear Files(*.gg)|*.gg|Game Boy Files(*.gb,*.gbc,*.gba)|*.gb;*.gbc;*.gba|Tekst Files(*.txt)|*.txt"
Public Const OpenTableFile1 = "Open Tabel File nr1"
Public Const TableFileTypes = "Alle Files(*.*)|*.*|Tabel Files(*.tbl)|*.tbl"
Public Const OpenTableFile2 = "Open Tabel File nr2(optioneel)"
Public Const CancelData = "Je hebt op Cancel gedrukt, wil je afsluiten?"
Public Const CancelTable = "Je hebt geen tabel file geload, wil je afsluiten, druk op Abort. Om terug te keren om een tabel file te laden druk op Retry. Om de standaart ANSI codering to te passen, druk op Ignore."
    'various extra things in code
Public Const GetBookmark = "Beschrijving?" & vbCrLf & vbCrLf & "(Edit de tabel file direct om de bookmarks te verwijderen.)"
Public Const GetBookmarkTitle = "Geef Bookmark beschrijving"
Public Const GetDumpmark = "Beschrijving?" & vbCrLf & vbCrLf & "(Edit de tabel file direct om de dumpmarks te verwijderen.)"
Public Const GetDumpmarkTitle = "Geef Dumpmark beschrijving"
Public Const GetInsertmark = "Beschrijving?" & vbCrLf & vbCrLf & "(Edit de tabel file direct om de insertmarks te verwijderen.)"
Public Const GetInsertmarkTitle = "Geef Insertmark beschrijving"
Public Const UseLegacyInsertMarkYesNo = "Weet je zeker dat je een thingformaat bookmark wil maken?."
Public Const InsertionMethod = "Invoeg Methode"
Public Const GetEndLocation = "Specificeer een ein locatie:"
Public Const GetEndLocationtitle = "Eind Locatie"
Public Const GetOutputFile = "Uitvoer file?"
Public Const OutputFileType = "Alle Files(*.*)|*.*|Tekst Files(*.txt)|*.txt|Data Files(*.dat)|*.dat|Dump Files(*.dump)|*.dump|Sonic the Hedgehog Files(*.sonic)|*.sonic"
Public Const DumpStart = "Start: "
Public Const DumpEnd = " Einde: "
        'these next strings are used in frmEdit:
    'instructions
Public Const TypeTextHere = "Typ je tekst hier: druk op ESC of Enter om te saven en terug te keren naar de hexeditor."
Public Const TypeHexHere = "Typ je hex charakter hier:"
Public Const ChangeRelSearch = "Verander naar relatieve zoekactie"
Public Const ChangeTblSearch = "Verander naar tabel zoekactie"
Public Const AlphaOrder = "Alfabet sorteren"
    'messages
Public Const AskYes = "Ask:Y"
Public Const AskNo = "Ask:N"
Public Const Bit16Yes = "16bit:Y"
Public Const Bit16No = "16bit:N"
Public Const DoneMsg = "Klaar!"
Public Const SuggestMsg1 = "Ik suggereer de " 'first and second half
Public Const SuggestMsg2 = " tile... invoeren?"
Public Const SuggestTitle = "Suggestie"
    '2 different titles for the form
Public Const TitleEdit = "Edit File Data"
Public Const TitleSearch = "Typ zoekwoord (waarschuwing: Zoeken gaat langzaam)"
        'these next string are used in frmJump:
    '3 diff titles for the form
Public Const TitleJump = "Kies Jump Locatie"
Public Const TitleDump = "Kies Dump Methode"
Public Const TitleInsert = "Kies Inserteer Methode"
    'visual elements
Public Const OK = "OK"
Public Const Cancel = "Annuleren"
Public Const AddCurrent = "&Gebruik huidig adress als bookmark"
Public Const ManualJump = "&Handmatig adress"
Public Const ManualDumpOrInsert = "Vind het einde &Handmatig"
Public Const BookmarkLabel = "&Bookmarks"
Public Const DumpmarkLabel = "&Dumpmarks"
Public Const InsertmarkLabel = "&Insertmarks"
    'messages
Public Const ManualMsg = "Typ handmatig adress in Decimal of Hex; Hex moet &H bevatten " & vbCrLf & vbCrLf & "(vrbld. &H100A)"
Public Const ManualTitle = "Handmatig adress"
'Relative search messages:
Public Const NeedMoreThanTwoValues = "Je moet minstens 2 waarden invoeren voor een relatieve zoekactie!"

Public Function RelSearch(RelData() As Integer, entry As Integer, startPos As Long, intFileno As Integer) As Long
'this returns a 0 based address if it finds the search data(contained in the array RelData) or -1 if it doesn't
'entry is number of data, startPos is address at which to start searching and intFileno is file to search.
Dim skipFlag(1 To 40) As Boolean '10 bigger than the RelData array, just in case
Dim searchLength As Integer
Dim bytesRead As Long
Dim relCount As Integer
Dim buffer As String
Dim buffer1 As Integer, buffer2 As Integer
Dim tempbuffer As String
Dim pos As Long
Dim offset As Long
Dim I As Integer
Dim c As Long
    RelSearch = -1  'init to false just in case we don't find anything
    For I = 1 To entry Step 1
        If RelData(I) = 32767 Then
            skipFlag(I) = True
        End If
    Next I
'first format a skipFlag array that contains a list of the wildcards
    Do While (skipFlag(entry) = True And entry > 1)
        ' No sense having wildcards at the end
        entry = entry - 1
    Loop
    If entry < 2 Then
        MsgBox NeedMoreThanTwoValues, vbInformation
        RelSearch = -1
        Exit Function
    End If
    
  searchLength = entry
  relCount = entry - 1
  buffer = String$(30000, " ")
    
    ' make relative search table
Dim rel() As Integer
Dim first() As Integer
Dim second() As Integer
ReDim rel(0 To 1) As Integer
ReDim first(0 To 1) As Integer
ReDim second(0 To 1) As Integer

  pos = 1
    For c = 0 To relCount - 1 Step 1
        first(c) = pos - 1
        Do While (skipFlag(pos + 1) = True)
          pos = pos + 1
          relCount = relCount - 1
        Loop
        second(c) = pos
        rel(c) = RelData(second(c) + 1) - RelData(first(c) + 1)
        If RelData(second(c) + 1) < RelData(first(c) + 1) Then rel(c) = rel(c) + 256
        pos = pos + 1
        ReDim Preserve rel(0 To c + 2) As Integer
        ReDim Preserve first(0 To c + 2) As Integer
        ReDim Preserve second(0 To c + 2) As Integer
    Next c

'put out the info on the relsearch table just created.
Dim msg As String
  msg = msg & "Search:" & vbCrLf
  For c = 0 To relCount - 1 Step 1
    msg = msg & "     Byte " & first(c) + 1 & " to byte " & second(c) + 1 & " :" & rel(c) & vbCrLf
Next c
msg = msg & vbCrLf
MsgBox msg, vbInformation
msg = ""
Get #intFileno, startPos, buffer   'read 30K
bytesRead = 30000   'we always read 30Kbytes because VB lets you read past the EOF somehow

If (LOF(intFileno) < searchLength) Then   'and make sure it's long enough...(but no VB equivalent)
    MsgBox "File not long enough!", vbCritical
    Exit Function
End If
  pos = 0
  Do
'somewhat working, still a little suspicious as to whether it's getting the last comparison of the two bytes.
    Do While ((pos + searchLength - 1) < bytesRead)

        For c = 0 To relCount - 1 Step 1
          buffer1 = Asc(Mid$(buffer, (pos + first(c) + 1), 1))
          buffer2 = Asc(Mid$(buffer, (pos + second(c) + 1), 1))
          If (buffer2 < buffer1) Then buffer2 = buffer2 + 256
          If (buffer2 - buffer1 <> rel(c)) Then Exit For ' no match
        Next
        If (c = relCount) Then
          RelSearch = startPos + offset + pos 'only get the first match. maybe we'll change this later if people want me to.
          Exit Function
        End If
        pos = pos + 1
    Loop

    For c = 0 To searchLength - 2 Step 1
      Mid(buffer, c + 1, 1) = Mid$(buffer, (bytesRead + 1 - searchLength + c), 1)
    Next c

    tempbuffer = Space$(30000 - searchLength)
    bytesRead = 30000 - searchLength
    Get #intFileno, , tempbuffer
    buffer = Left(buffer, searchLength) & tempbuffer    'this sims the pointer operation that was in the C version.
    pos = 0
    offset = offset + 30000 - searchLength
Loop Until EOF(intFileno)
'Jair's terms of use:
'TERMS OF USE
'-------------------
'This program is distributed with its source code. You may use, distribute, and modify it freely.
'Only two restrictions: These terms of use must stay the same, and you must always include the source code with the program.
End Function
