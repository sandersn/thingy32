Attribute VB_Name = "Globals"
'--------------------------------------------------------------------------------
'The new strings are marked with *** and an explanation if needed
'Also please translate this for the web site:
'Download thingy32 version 0.19 in russian: Open the .zip file into the same directory as
'thingy32 English. Run Thingy32Rus.exe instead of Thingy32.exe. Note: you need to install
'Thingy32 in English first.

'�������� thingy32 ������ 0.19 �� �������: ���������� .zip ���� � ���� ����� ���������� ��� �
'thingy32 ���������� ������. ��������� Thingy32Rus.exe
'������ Thingy32.exe. ����������: ��� ����� ������� ���������� Thingy32 ��
'����������.


Option Explicit
'(started 13 Sept 1999)
'update from Jan2K, have added a number of vars to aid me in the new features:triple hex and double byte line/section breaks.
'globals:the variables that have to be visible to ALL forms...yes I hate having to do this but you have
'to have something like this I suppose
'23 Jan 2000
    'I have added many string constants that cover all text in thingy32. After I debug these, I will change these to Spanish and release a
'Spanish version.
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
Public Const JumpKey = "J"
Public Const TableSwitchKey = "T"
Public Const EditSelKey = " "
Public Const SearchKey = "S"
Public Const ReSearchKey = "R"
Public Const DumpKey = "D"
Public Const InsertKey = "I"
Public Const HideHexKey = "H"
Public Const OptionsKey = "O" 'ok
Public Const MoreWidthKey = "."
Public Const LessWidthKey = ","
    'shortcut labels(all are visual elements)
Public Const JumpLabel = "J = ������� �"
Public Const OptionsLabel = "O = ��������� �����������" 'ok
Public Const EditSelLabel = "������ = ������������� ����������"
Public Const SearchLabel = "S = �����"
Public Const ReSearchLabel = "R = ��������� �����"
Public Const DumpLabel = "D = ����"
Public Const InsertLabel = "I = ��������"
Public Const PositionLabel = "���������: "
Public Const MoreWidthLabel = "��������� ������ = >"
Public Const LessWidthLabel = "< = ��������� ������"
    'menus *** the & mean _ on a menu(the hotkey)
    'choose which letters you want to use, but do not use the same letter twice
Public Const OpenNewFile = "&������� ����� ����"
Public Const OpenNewTable = "������� ����� &�������..."
Public Const ReloadTable = "&������������� ������� �������"
Public Const HideHex = "&������ Hex"
Public Const TableSwitch = "����&������� �������"
Public Const JumpManual = "������ ����&���..."
Public Const DumpManual = "������ &���� ������� � ������� �������"
Public Const InsertManual = "������ &������� ������� � ������� �������"
Public Const AddNewBookMark = "�������� ������� ������� ��� ����� &�������..."
Public Const AddNewDumpMark = "�������� ������� ������� ��� ����� ���&���� �����..."
Public Const AddNewInsertMark = "�������� ������� ������� ��� ����� ������� ������&�:..."
    'misc messages
Public Const LabelTooltip = "������� �� ������� ��� ���������� ��������������� �������"   'visual element
Public Const OptionsTooltip = "������� ��� ������ ���� �����"  'visual element
Public Const ReadyMsg = "������"
Public Const UsingTableMsg = "������ ���������� ������� #"
Public Const SelectingMsg = "��������"
Public Const CurrentSearch = "������� �����: "
Public Const SelDumpEnd = "����� ��������� �����"
Public Const SelInsertEnd = "����� ��������� �������"
        'misc--load 'ok
    Public Const ErrorLoadingDataFromCommand = "��� ����� ������ ��� ������������ ���� ������! �� �� ������� ��������� ���� ���������." & vbCrLf & "���������: �� ������� ��������� ���� ������(�) ��� ������� ����������� ����"
    Public Const ErrorLoadingDataTitle = "������ �������� ����� ������"
    Public Const ErrorLoadingTableFromCommand = "��� ����� ������� ��� ������������ ���� ������! �� �� ������� ��������� ���� ���������."
    Public Const ErrorLoadingTableTitle = "������ �������� �������"
    Public Const ErrorLoadingTable2Title = "������  �������� ������� #2"
    'app info
Public Const InfoTitle = "thingy32 ���������� ����������:"
Public Const InfoName = "����������� ��������: "
Public Const InfoVersion = "������: "
Public Const InfoComments = "����������� ��� ���� ������: "
Public Const InfoCopyright = "��������� �����: "
Public Const InfoTrademarks = "�������� �����: "
    'program start sequence(in LoadFiles mainly)
Public Const OpenDataFile = "������� ���� ������"
Public Const DataFileTypes = "��� �����(*.*)|*.*|����� Snes(*.smc,*.swc,*.sfc,*.fig, and others)|*.smc;*.swc;*.sfc;*.fig;*.058;*.078,*.1;*.2;*.3|����� Nes(*.nes)|*.nes|����� Genesis/Megadrive(*.smd, *.bin)|*.smd,*.bin|����� Master System(*.sms)|*.sms|����� Game Gear(*.gg)|*.gg|����� GameBoy(*.gbX)|*.gb,*.gbc,*.gba|��������� �����(*.txt)|*.txt"
Public Const OpenTableFile1 = "������� ���� ������� #1"
Public Const TableFileTypes = "��� �����(*.*)|*.*|����� ������(*.tbl)|*.tbl"
Public Const OpenTableFile2 = "������� ���� ������� #2(�������������)"
Public Const CancelData = "�� ������ ������. �� ������ ����� ?"
Public Const CancelTable = "�� �� ��������� ���� �������.���� �� ������ �����, ������� �����. ��� �������� �������� ����� �������, ������� ������.��� �������� ����� ������ �������� ����������� ANSI ���������, ������� ������������."
    'various extra things in code
Public Const GetBookmark = "��������?" & vbCrLf & vbCrLf & "(������������ ���� ������� ��������������� ��� ������� �������; ��� ���������� � �������.)"
Public Const GetBookmarkTitle = "���� �������� �������"
Public Const GetOutputFile = "���� ������? �������: �� ������ �� �������� ������ ��� �������� �����������"
Public Const OutputFileType = "��� �����(*.*)|*.*|��������� �����(*.txt)|*.txt|����� ������(*.dat)|*.dat|���� �����(*.dump)|*.dump|����� Sonic the Hedgehog(*.sonic)|*.sonic"
'begin ok
Public Const GetDumpmark = "��������?" & vbCrLf & vbCrLf & "(������������� ���� ������� �������� ��� �������� � ������� �����; ��� ���������� � �������.)"
Public Const GetDumpmarkTitle = "���� �������� ������� �����"
Public Const GetInsertmark = "��������?" & vbCrLf & vbCrLf & "(������������� ���� ������� �������� ��� �������� � ������� �������; ��� ���������� � �������.)"
Public Const GetInsertmarkTitle = "���� �������� ������� �������"
Public Const UseLegacyInsertMarkYesNo = "�� ������� ��� ������ ������� legacy-������ thingy ��������? �� �� ����� �������� ���������� � ����� ������ � �������������� �� ��� ����� �������. ������� ���� ���� ������� ����� �������, �� ����� ����������� ������ � ������ ������������."
Public Const InsertionMethod = "����� �������"
Public Const GetEndLocation = "���������� ����� ���������� �������� ����������:"
Public Const GetEndLocationtitle = "����� ����������"
Public Const DumpStart = "������: "
Public Const DumpEnd = "  �����: "
'end ok
        'these next strings are used in frmEdit:
    'instructions
Public Const TypeTextHere = "������� ���� ����� �����: ������� ESC ��� Enter ��� ���������� � �������� � ������� HEX P�������."
Public Const TypeHexHere = "������� ���� HEX ������� �����:"
Public Const ChangeRelSearch = "������� � �������������� ������"
Public Const ChangeTblSearch = "������� � ���������� ������"
Public Const AlphaOrder = "���������� �������"
    'messages
Public Const AskYes = "������:Y"
Public Const AskNo = "������:N"
Public Const Bit16Yes = "16bit:Y"
Public Const Bit16No = "16bit:N"
Public Const DoneMsg = "�����!"
Public Const SuggestMsg1 = "� ��������� ������������ "  'first and second half
Public Const SuggestMsg2 = " �������� � ?"
Public Const SuggestTitle = "�����������"
    '2 different titles for the form
Public Const TitleEdit = "������������� ���� ������"
Public Const TitleSearch = "������ ������ ��� ������(�������� ����� ���������)"
        'these next string are used in frmJump:
    '3 diff titles for the form
Public Const TitleJump = "�������� ����� ��������"
Public Const TitleDump = "�������� ����� �����"
Public Const TitleInsert = "�������� ����� �������"
    'visual elements
Public Const OK = "��� ��"
Public Const Cancel = "������"
Public Const AddCurrent = "&���������� ������� ��������� ��� �������"
Public Const ManualJump = "&������ �����"
Public Const ManualDumpOrInsert = "������� ����� &�������"
Public Const BookmarkLabel = "&�������:"
Public Const DumpmarkLabel = "&������� �����:"
Public Const InsertmarkLabel = "&������� �������:"
    'messages
Public Const ManualMsg = "������� ������ ����� � ���������� ��� Hex �����; Hex ������ ���������� � �������� &H" & vbCrLf & vbCrLf & "(i.e. &H100A)"
Public Const ManualTitle = "������ �����"
    'Relative search messages:
Public Const NeedMoreThanTwoValues = "���������� �� ������ ���� ��������� ��� �������������� ������"

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
          RelSearch = startPos + offset + pos 'only get the first match. maybe we 'll change this later if people want me to.
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





   

