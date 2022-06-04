Attribute VB_Name = "Globals"
Option Explicit
'(started 13 Sept 1999)
'update from Jan2K, have added a number of vars to aid me in the new features:triple hex and double byte line/section breaks.
'globals:the variables that have to be visible to ALL forms...yes I hate having to do this but you have
'to have something like this I suppose
'23 Jan 2000
    'I have added many string constants that cover all text in thingy32. After I debug these, I will change these to Spanish and release a
'Spanish version.
    '24 Jan 2000: Ummm well, maybe my request was *too* successful. Now I have both an american spanish version and a european spanish
'version. :) Oh well. I'll do at least the american version tonight to see how it goes, then distribute a patch that differentiates the two from
'the english version, the distrib the patches in the original .zip.
    '16 Feb: A note from previous text, only 1 spanish version exists after all, and I'm having to provide the .exe in a separate zip because the patches
'don't work(they're the same size as the .exe) Now I'm inserting a French version!
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
Public Const JumpKey = "S"
Public Const TableSwitchKey = "C"
Public Const EditSelKey = " "
Public Const SearchKey = "B"
Public Const ReSearchKey = "R"
Public Const DumpKey = "E"
Public Const InsertKey = "I"
Public Const HideHexKey = "H"
Public Const OptionsKey = "O" '***
Public Const MoreWidthKey = "."
Public Const LessWidthKey = ","
    'shortcut labels(all are visual elements)
        '** HideHexLabel y TableSwitchLabel son borrado
Public Const JumpLabel = "S = Saltar"
Public Const OptionsLabel = "O = Opciones"   '***
Public Const EditSelLabel = "Espacio = Seleccionar"
Public Const SearchLabel = "B = Buscar"
Public Const ReSearchLabel = "R = ReBuscar"
Public Const DumpLabel = "E = Extraer"
Public Const InsertLabel = "I = Insertar"
Public Const PositionLabel = "Posici�n: "
Public Const MoreWidthLabel = "Mayor anchura = >"
Public Const LessWidthLabel = "< = Menor anchura"
    'menus ***    -Why are there those &?-
Public Const OpenNewFile = "Abrir archivo &nuevo..."
Public Const OpenNewTable = "&Abrir nueva tabla..."
Public Const ReloadTable = "&Recargar tabla actual"
Public Const HideHex = "No mostrar &hex"
Public Const TableSwitch = "Cambiar &tabla"
Public Const JumpManual = "&Salto manual..."
Public Const DumpManual = "&Extracci�n manual a partir de la posici�n actual"
Public Const InsertManual = "&Inserci�n manual a partir de la posici�n actual"
Public Const AddNewBookMark = "A�adir posici�n actual como &Marcador de salto..."
Public Const AddNewDumpMark = "A�adir posici�n actual como Mar&cador de extracci�n..."
Public Const AddNewInsertMark = "A�adir posici�n actual como Marca&dor de inserci�n..."
    'misc messages
Public Const LabelTooltip = "Haz clic para ejecutar la funci�n respectiva"   'visual element
Public Const OptionsTooltip = "Haz clic para Men� de opciones"
Public Const ReadyMsg = "Listo"
Public Const UsingTableMsg = "Ahora usando tabla #"
Public Const SelectingMsg = "Seleccionando"
Public Const CurrentSearch = "B�squeda actual: "
Public Const SelDumpEnd = "Presiona " & DumpKey & " al final del texto a extraer"    '**
Public Const SelInsertEnd = "Presiona " & InsertKey & " al final de lo extra�do"    '**
        'misc--load ***
    Public Const ErrorLoadingDataFromCommand = "�Archivo de datos inexistente (especificado en l�nea de comando)! Cerci�rate de que el nombre de archivo no contiene espacios. Tendr�s que cargar el archivo manualmente." & vbCrLf & "Nota: Tambi�n tendr�s que cargar manualmente la(s) tabla(s), aun si �stas ten�an correcto el nombre."
    Public Const ErrorLoadingDataTitle = "Error al cargar archivo de datos"
    Public Const ErrorLoadingTableFromCommand = "�Archivo de tabla inexistente (especificado en l�nea de comando)! Cerci�rate de que el nombre de archivo no contiene espacios. Tendr�s que cargar la tabla manualmente."
    Public Const ErrorLoadingTableTitle = "Error al cargar tabla"
    Public Const ErrorLoadingTable2Title = "Error al cargar tabla #2"

    'app info
Public Const InfoTitle = "Informaci�n de la aplicaci�n thingy32 (No est� en espa�ol. Lo siento):"
Public Const InfoName = "Nombre oficial: "
Public Const InfoVersion = "Versi�n: "
Public Const InfoComments = "Comentarios de esta compilaci�n: "
Public Const InfoCopyright = "Informaci�n de copyright: "
Public Const InfoTrademarks = "Informaci�n de marca registrada: "
    'program start sequence(in LoadFiles mainly)
Public Const OpenDataFile = "Abrir archivo de datos"
Public Const DataFileTypes = "Todos los archivos(*.*)|*.*|Snes(*.smc,*.swc,*.sfc,*.fig, y otros)|*.smc;*.swc;*.sfc;*.fig;*.058;*.078;*.1;*.2;*.3|Nes(*.nes)|*.nes|Megadrive/Genesis/32x(*.smd, *.bin)|*.smd;*.bin;*.32x|Master System(*.sms)|*.sms|Game Gear(*.gg)|*.gg|Game Boy(*.gb)|*.gb|Texto(*.txt)|*.txt"
Public Const OpenTableFile1 = "Abrir tabla #1"
Public Const TableFileTypes = "Todos los archivos(*.*)|*.*|Archivos de tablas(*.tbl)|*.tbl"
Public Const OpenTableFile2 = "Abrir tabla #2 (opcional)"
Public Const CancelData = "�Est�s seguro de que quieres cancelar?"
Public Const CancelTable = "No has cargado ninguna tabla. Si deseas salir, presiona Cancelar. Para regresar y cargar un archivo de tabla presiona Reintentar. Para usar codificaci�n ANSI est�ndar, presiona Ignorar."
    'various extra things in code
Public Const GetBookmark = "�Descripci�n?" & vbCrLf & vbCrLf & "(Edita el archivo de tabla directamente para remover los marcadores; est�n entre par�ntesis.)"
Public Const GetBookmarkTitle = "Escribe descripci�n del marcador"
Public Const GetOutputFile = "�Archivo destino?"
Public Const OutputFileType = "Todos los archivos(*.*)|*.*|Archivos de texto(*.txt)|*.txt|Archivos de datos(*.dat)|*.dat|Archivos de extracci�n(*.extra)|*.extra|Archivos de Sonic the Hedgehog(*.sonic)|*.sonic"
'***
Public Const GetDumpmark = "�Descripci�n?" & vbCrLf & vbCrLf & "(Edita el archivo de tabla directamente para remover los marcadores de extracci�n; est�n entre corchetes.)"
Public Const GetDumpmarkTitle = "Escribe descripci�n del marcador de extracci�n"
Public Const GetInsertmark = "Description?" & vbCrLf & vbCrLf & "(Edita el archivo de tabla directamente para remover los marcadores de inserci�n; est�n entre llaves.)"
Public Const GetInsertmarkTitle = "Escribe descripci�n del marcador de inserci�n"
Public Const UseLegacyInsertMarkYesNo = "�Est�s seguro de querer crear un marcador de inserci�n? No especificar� la posici�n final en el archivo de datos e insertar� el archivo de texto completo. Si �ste es demasiado grande podr�a sobreescribir algo en el archivo de datos."
Public Const InsertionMethod = "M�todo de inserci�n"
Public Const GetEndLocation = "Especifica la posici�n final:"
Public Const GetEndLocationtitle = "Posici�n final"
Public Const DumpStart = "Inicio: "
Public Const DumpEnd = "  Final: "
'***

        'these next strings are used in frmEdit:
    'instructions
Public Const TypeTextHere = "Escribe tu texto aqu�: Presiona ESC o Enter para guardarlo y volver al editor hexadecimal principal."
Public Const TypeHexHere = "Escribe el n�mero hexadecimal aqu�:"
Public Const ChangeRelSearch = "Cambiar a b�squeda relativa"
Public Const ChangeTblSearch = "Cambiar a b�squeda de tabla"
Public Const AlphaOrder = "Orden alfab�tico"
    'messages
Public Const AskYes = "Preguntar:S�"
Public Const AskNo = "Preguntar:No"
Public Const Bit16Yes = "16bit:S�"
Public Const Bit16No = "16bit:No"
Public Const DoneMsg = "�Hecho!"
Public Const SuggestMsg1 = "Sugiero usar el t�tulo "  'first and second half
Public Const SuggestMsg2 = "...�lo ponemos?"
Public Const SuggestTitle = "Sugerencia"
    '2 different titles for the form
Public Const TitleEdit = "Editar archivo"
Public Const TitleSearch = "Escribe texto a buscar (advertencia: la b�squeda toma tiempo)"
        'these next string are used in frmJump:
    '3 diff titles for the form
Public Const TitleJump = "Escoge nueva direcci�n"
Public Const TitleDump = "Escoge m�todo de extracci�n"
Public Const TitleInsert = "Escoge m�todo de inserci�n"
    'visual elements
Public Const OK = "OK"
Public Const Cancel = "Cancelar"
Public Const AddCurrent = "&Agrega posici�n actual como nuevo marcador"
Public Const ManualJump = "&Direcci�n manual"
Public Const ManualDumpOrInsert = "&Final manual"
Public Const BookmarkLabel = "&Marcadores"
Public Const DumpmarkLabel = "Marcadores de &extracci�n"
Public Const InsertmarkLabel = "Marcadores de &inserci�n"
    'messages
Public Const ManualMsg = "Escribe direcci�n en decimal o hexadecimal; en caso de hexadecimal debes poner &H antes del n�mero" & vbCrLf & vbCrLf & "(v.g. &H100A)"
Public Const ManualTitle = "Direcci�n manual"
    'relative search messages
Public Const NeedMoreThanTwoValues = "�Debes tener al menos dos valores para la b�squeda relativa!"

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
  msg = msg & "Buscar:" & vbCrLf
  For c = 0 To relCount - 1 Step 1
    msg = msg & "     byte " & first(c) + 1 & " a byte " & second(c) + 1 & " :" & rel(c) & vbCrLf
Next c
msg = msg & vbCrLf
MsgBox msg, vbInformation
msg = ""
Get #intFileno, startPos, buffer   'read 30K
bytesRead = 30000   'we always read 30Kbytes because VB lets you read past the EOF somehow

If (LOF(intFileno) < searchLength) Then   'and make sure it's long enough...(but no VB equivalent)
    MsgBox "�Archivo demasiado peque�o!", vbCritical
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

