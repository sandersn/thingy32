Attribute VB_Name = "Globals"
'--------------------------------------------------------------------------------
'The new strings are marked with *** and an explanation if needed
'Also please translate this for the web site:
'Download thingy32 version 0.19 in russian: Open the .zip file into the same directory as
'thingy32 English. Run Thingy32Rus.exe instead of Thingy32.exe. Note: you need to install
'Thingy32 in English first.

'Скачайте thingy32 версии 0.19 на русском: Распакуйте .zip файл в туже самую директорию что и
'thingy32 Английская версия. Запустите Thingy32Rus.exe
'вместо Thingy32.exe. Примечание: вам нужно сначала установить Thingy32 на
'Английском.


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
Public Const JumpLabel = "J = Перейти к"
Public Const OptionsLabel = "O = Параметры выделенного" 'ok
Public Const EditSelLabel = "Пробел = Редактировать выделенное"
Public Const SearchLabel = "S = Поиск"
Public Const ReSearchLabel = "R = Повторный Поиск"
Public Const DumpLabel = "D = Дамп"
Public Const InsertLabel = "I = Вставить"
Public Const PositionLabel = "Положение: "
Public Const MoreWidthLabel = "Увеличить ширину = >"
Public Const LessWidthLabel = "< = Уменьшить ширину"
    'menus *** the & mean _ on a menu(the hotkey)
    'choose which letters you want to use, but do not use the same letter twice
Public Const OpenNewFile = "&Открыть новый файл"
Public Const OpenNewTable = "Открыть новую &таблицу..."
Public Const ReloadTable = "&Перезагрузить текущую таблицу"
Public Const HideHex = "&Скрыть Hex"
Public Const TableSwitch = "Пере&ключить таблицы"
Public Const JumpManual = "Данный пере&ход..."
Public Const DumpManual = "Данный &Дамп начиная с текущей позиции"
Public Const InsertManual = "Данная &Вставка начиная с текущей позиции"
Public Const AddNewBookMark = "Добавить текущую позицию как новую &заметку..."
Public Const AddNewDumpMark = "Добавить текущую позицию как новую зам&етку дампа..."
Public Const AddNewInsertMark = "Добавить текущую позицию как новую заметки вставк&и:..."
    'misc messages
Public Const LabelTooltip = "Нажмите на отметки для выполнения соответственной функции"   'visual element
Public Const OptionsTooltip = "Нажмите для вызова меню опций"  'visual element
Public Const ReadyMsg = "Готово"
Public Const UsingTableMsg = "Сейчас используем Таблицу #"
Public Const SelectingMsg = "Выбираем"
Public Const CurrentSearch = "Текущий Поиск: "
Public Const SelDumpEnd = "Выбор окончания Дампа"
Public Const SelInsertEnd = "Выбор окончания Вставки"
        'misc--load 'ok
    Public Const ErrorLoadingDataFromCommand = "Нет файла данных или неправильный путь ярлыка! Вы не сможете загрузить файл правильно." & vbCrLf & "Замечание: Вы сможете загрузить файл таблиц(ы) при наличии правильного пути"
    Public Const ErrorLoadingDataTitle = "Ошибка загрузки файла данных"
    Public Const ErrorLoadingTableFromCommand = "Нет файла таблицы или неправильный путь ярлыка! Вы не сможете загрузить файл правильно."
    Public Const ErrorLoadingTableTitle = "Ошибка загрузки Таблицы"
    Public Const ErrorLoadingTable2Title = "Ошибка  загрузки Таблицы #2"
    'app info
Public Const InfoTitle = "thingy32 Информация Приложения:"
Public Const InfoName = "Официальное название: "
Public Const InfoVersion = "Версия: "
Public Const InfoComments = "Комментарий для этой версии: "
Public Const InfoCopyright = "Авторское право: "
Public Const InfoTrademarks = "Торговая марка: "
    'program start sequence(in LoadFiles mainly)
Public Const OpenDataFile = "Открыть Файл Данных"
Public Const DataFileTypes = "Все Файлы(*.*)|*.*|Файлы Snes(*.smc,*.swc,*.sfc,*.fig, and others)|*.smc;*.swc;*.sfc;*.fig;*.058;*.078,*.1;*.2;*.3|Файлы Nes(*.nes)|*.nes|Файлы Genesis/Megadrive(*.smd, *.bin)|*.smd,*.bin|Файлы Master System(*.sms)|*.sms|Файлы Game Gear(*.gg)|*.gg|Файлы GameBoy(*.gbX)|*.gb,*.gbc,*.gba|Текстовые Файлы(*.txt)|*.txt"
Public Const OpenTableFile1 = "Открыть Файл Таблицы #1"
Public Const TableFileTypes = "Все Файлы(*.*)|*.*|Файлы Таблиц(*.tbl)|*.tbl"
Public Const OpenTableFile2 = "Открыть Файл Таблицы #2(необязательно)"
Public Const CancelData = "Вы нажали Отмену. Вы хотите выйти ?"
Public Const CancelTable = "Вы не загрузили файл таблицы.Если вы хотите выйти, нажмите Отказ. Для возврата изагруки файла таблицы, нажмите Повтор.Для загрузки файла данных испльзуя стандартную ANSI кодировку, нажмите Игнорировать."
    'various extra things in code
Public Const GetBookmark = "Описание?" & vbCrLf & vbCrLf & "(Редактируйте файл таблицы непосредственно для удаленя Заметок; Они находяться в скобках.)"
Public Const GetBookmarkTitle = "Дать описапие Заметок"
Public Const GetOutputFile = "Файл вывода? заметка: Вы можите НЕ нажимать Отмена или пограмма завершиться"
Public Const OutputFileType = "Все Файлы(*.*)|*.*|Текстовые Файлы(*.txt)|*.txt|Файлы Данных(*.dat)|*.dat|Дамп Файлы(*.dump)|*.dump|Файлы Sonic the Hedgehog(*.sonic)|*.sonic"
'begin ok
Public Const GetDumpmark = "Описание?" & vbCrLf & vbCrLf & "(Редактировать файл таблицы напрямую для удаления к заметке дампа; она находиться в скобках.)"
Public Const GetDumpmarkTitle = "Дать описание заметке дампа"
Public Const GetInsertmark = "Описание?" & vbCrLf & vbCrLf & "(Редактировать файл таблицы напрямую для удаления к заметке вставки; она находиться в скобках.)"
Public Const GetInsertmarkTitle = "Дать описание заметке вставки"
Public Const UseLegacyInsertMarkYesNo = "Вы уверены что хотите создать legacy-формат thingy закладок? Он не точно отмечает размешение в файле данных и распологаеться во всём файле вставки. Пеэтому если файл вставки очень большой, он может презаписать данные в данных безвозвратно."
Public Const InsertionMethod = "Метод вставки"
Public Const GetEndLocation = "Пожалуйста точно опредилите конечное размешение:"
Public Const GetEndLocationtitle = "Конец Размешения"
Public Const DumpStart = "Начало: "
Public Const DumpEnd = "  Конец: "
'end ok
        'these next strings are used in frmEdit:
    'instructions
Public Const TypeTextHere = "Введите свой текст здесь: Нажмите ESC или Enter для сохранения и возврата в главный HEX Pедактор."
Public Const TypeHexHere = "Введите свои HEX символы здесь:"
Public Const ChangeRelSearch = "Перейти к сравнительному поиску"
Public Const ChangeTblSearch = "Перейти к табличному поиску"
Public Const AlphaOrder = "Алфавитный Порядок"
    'messages
Public Const AskYes = "Вопрос:Y"
Public Const AskNo = "Вопрос:N"
Public Const Bit16Yes = "16bit:Y"
Public Const Bit16No = "16bit:N"
Public Const DoneMsg = "Конец!"
Public Const SuggestMsg1 = "Я предлагаю использовать "  'first and second half
Public Const SuggestMsg2 = " Положить в ?"
Public Const SuggestTitle = "Предложение"
    '2 different titles for the form
Public Const TitleEdit = "Редактировать Файл Данных"
Public Const TitleSearch = "Ввести строку для поиска(Внимание поиск медленный)"
        'these next string are used in frmJump:
    '3 diff titles for the form
Public Const TitleJump = "Выберите точку перехода"
Public Const TitleDump = "Выберите метод Дампа"
Public Const TitleInsert = "Выберите метод вставки"
    'visual elements
Public Const OK = "Все Ок"
Public Const Cancel = "Отмена"
Public Const AddCurrent = "&Определить текушее положение как Заметку"
Public Const ManualJump = "&Данный адрес"
Public Const ManualDumpOrInsert = "Указать конец &Данного"
Public Const BookmarkLabel = "&Заметки:"
Public Const DumpmarkLabel = "&Заметки Дампа:"
Public Const InsertmarkLabel = "&Заметки Вставки:"
    'messages
Public Const ManualMsg = "Введите Данный Адрес в Десятичной или Hex форме; Hex должен начинаться с префикса &H" & vbCrLf & vbCrLf & "(i.e. &H100A)"
Public Const ManualTitle = "Данный Адрес"
    'Relative search messages:
Public Const NeedMoreThanTwoValues = "Необходимо не меньше двух выражений для сравнительного поиска"

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





   

