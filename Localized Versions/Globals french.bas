Attribute VB_Name = "Globals"
Option Explicit
'(started 13 Sept 1999)
'update from Jan2K, have added a number of vars to aid me in the new features:triple hex and double byte line/section breaks.
'globals:the variables that have to be visible to ALL forms...yes I hate having to do this but you have
'to have something like this I suppose
'23 Jan 2000
    'I have added many string constants that cover all text in thingy32. After I debug these, I will change these to Spanish and release a
'Spanish version.
    '16 Feb
'oops comments are gone, see Globals american spanish.bas for them. They weren't very important anyway.
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
Public Const TableSwitchKey = "T"
Public Const EditSelKey = " "
Public Const SearchKey = "C"
Public Const ReSearchKey = "R"
Public Const DumpKey = "E"
Public Const InsertKey = "I"
Public Const HideHexKey = "H"
Public Const OptionsKey = "O" '***
Public Const MoreWidthKey = "."
Public Const LessWidthKey = ","
    'shortcut labels(all are visual elements)
        '** HideHexLabel y TableSwitchLabel son borrado
Public Const JumpLabel = "S = Saut"
Public Const OptionsLabel = "O = Options"   '***
Public Const EditSelLabel = "Espace = Sélectionnez pour modifier"
Public Const SearchLabel = "C = Cherchez"
Public Const ReSearchLabel = "R = Recherchez"
Public Const DumpLabel = "E = Extration"
Public Const InsertLabel = "I = Insertion"
Public Const MoreWidthLabel = "Plus Large = >"
Public Const LessWidthLabel = "< = Moins Large"
Public Const PositionLabel = "Position: "
    'menus ***
Public Const OpenNewFile = "Ouvrir &Nouveau..."
Public Const OpenNewTable = "&Ouvrir Nouvelle Table..."
Public Const ReloadTable = "&Reload Current Table"
Public Const HideHex = "&Cacher Hex"
Public Const TableSwitch = "Changer &Fichier Table "
Public Const JumpManual = "&Sauter manuellement..."
Public Const DumpManual = "&Extraire manuellement en se servant de la position présente comme départ"
Public Const InsertManual = "&Inserer manuellement en se servant de la position présente comme départ"
Public Const AddNewBookMark = "&Ajouter la possition présente aux signets..."
Public Const AddNewDumpMark = "&Ajouter la possition présente aux marques d'extrations..."
Public Const AddNewInsertMark = "&Ajouter la possition présente aux marques d'insertion..."
    'misc messages
Public Const LabelTooltip = "Cliquez sur le raccourci pour exécuter la fonction respective"   'visual element
Public Const OptionsTooltip = "Cliquez sur le pour Menu de options"
Public Const ReadyMsg = "Prêt"
Public Const UsingTableMsg = "Utilse en ce moment la Table #"
Public Const SelectingMsg = "Sélectionné"
Public Const CurrentSearch = "Recherche en ce moment: "
Public Const SelDumpEnd = "Pesez " & DumpKey & " à la fin de la partit a extracter"    '**
Public Const SelInsertEnd = "Pesez " & InsertKey & " à la fin de la partit a extracter"    '**
        'misc--load ***
    Public Const ErrorLoadingDataFromCommand = "Fichier non existant ou le chemin est un racourci! Soyez sûr que le nom du fichier n'a aucun espace. Vous allez devoir ouvrir le fichier normalement." & vbCrLf & "Note: Vous allez devoir ouvrir les fichier table(s) manuellement aussi, même si le chemin est correct."
    Public Const ErrorLoadingDataTitle = "Erreur de chargement des donnés"
    Public Const ErrorLoadingTableFromCommand = "fichier table non existant ou le chemin est un racourci! Soyez sûr que le nom du fichier n'a aucun espace. Vous allez devoir ouvrir le fichier normalement."
    Public Const ErrorLoadingTableTitle = "Erreur dans le chargement de la table"
    Public Const ErrorLoadingTable2Title = "Erreur dans le chargement de la table #2"

    'app info
Public Const InfoTitle = "Information de l'Application thingy32:"
Public Const InfoName = "Nom Officiel: "
Public Const InfoVersion = "Version: "
Public Const InfoComments = "Commentaire pour cette construction : "
Public Const InfoCopyright = "Information sur les droits d'auteur : "
Public Const InfoTrademarks = "Information sur les droits de distribution : "
    'program start sequence(in LoadFiles mainly)
Public Const OpenDataFile = "Ouvrir les donnés des fichiers :"
Public Const DataFileTypes = "Tous les types Fichiers(*.*)|*.*|Snes fichiers(*.smc,*.swc,*.sfc,*.fig, and others)|*.smc;*.swc;*.sfc;*.fig;*.058;*.078;*.1;*.2;*.3|Nes fichiers(*.nes)|*.nes|Genesis/Megadrive/32x fichiers(*.smd, *.bin)|*.smd;*.bin;*.32x|Master System fichiers(*.sms)|*.sms|Game Gear fichiers(*.gg)|*.gg|Game Boy fichiers(*.gb)|*.gb|Fichier Text(*.txt)|*.txt"
Public Const OpenTableFile1 = "Ouvrir Table #1"
Public Const TableFileTypes = "Tous les types Fichiers(*.*)|*.*|Fichier Table(*.tbl)|*.tbl"
Public Const OpenTableFile2 = "Ouvrir Table #2(optionnel)"
Public Const CancelData = "Vous avez pesé sur Annuler. Voulez-vous quittez?"
Public Const CancelTable = "Vous n'avez pas chargé de table. Si vous voulez quittez, appuyez sur Annuler. Pour revenir et charger une table, appuyer sur Réessayer. Pour charger un fichier qui utilise la table ANSI, appuyez sur Ignoré."
    'various extra things in code
Public Const GetBookmark = "Description?" & vbCrLf & vbCrLf & "(Éditez le fichier table pour modifier un signet; ils sont entre  parenthèses.)"
Public Const GetBookmarkTitle = "Donné une description du signet"
Public Const GetOutputFile = "Fichier d'extration?"
Public Const OutputFileType = "Tous les fichiers(*.*)|*.*|Fichier de donner(*.txt)|*.txt|Fichier Texte(*.dat)|*.dat|Fichier extrait(*.dump)|*.dump|Fichier de Sonic the Hedgehog(*.sonic)|*.sonic"
'***
Public Const GetDumpmark = "Description?" & vbCrLf & vbCrLf & "(Éditez le fichier table directement pour enlever les marques d'extration; Ils sont entre parenthèses.)"
Public Const GetDumpmarkTitle = "Donnez une descrition de la marque d'extration"
Public Const GetInsertmark = "Description?" & vbCrLf & vbCrLf & "(Éditez le fichier table directement pour enlever les marques d'insertion; Ils sont entre crochet.)"
Public Const GetInsertmarkTitle = "Donnez une descrition de la marque d'insertion"
Public Const UseLegacyInsertMarkYesNo = "Êtes vous sûr de créer un livre de signet en format légal? Il ne spécifie pas la location dans le fichier et tout le fichier d'insetion. Si le fichier est trop grand, ils pourait écrire par dessus le fichier"
Public Const InsertionMethod = "Methode d'insertion"
Public Const GetEndLocation = "veuillez spécifiez votre location:"
Public Const GetEndLocationtitle = "Fin de la location"
Public Const DumpStart = "Départ: "
Public Const DumpEnd = "  Fin: "
'***
        'these next strings are used in frmEdit:
    'instructions
Public Const TypeTextHere = "Écrivez votre texte ici: Peser Echap ou Entrer pour sauvegarder et revenir à éditeur Hexadécimal."
Public Const TypeHexHere = "Marquez vos caractères hexa ici :"
Public Const ChangeRelSearch = "Change pour Recherche relative"
Public Const ChangeTblSearch = "Change pour recherche de Table"
Public Const AlphaOrder = "Ordre Alphabètique"
    'messages
Public Const AskYes = "Pregunta:O"
Public Const AskNo = "Pregunta:N"
Public Const Bit16Yes = "16bit:O"
Public Const Bit16No = "16bit:N"
Public Const DoneMsg = "Fait!"
Public Const SuggestMsg1 = "Je suggère d'utilisé "  'first and second half
Public Const SuggestMsg2 = " tire...Mettre dans?"
Public Const SuggestTitle = "Suggestion"
    '2 different titles for the form
Public Const TitleEdit = "Éditer le fichier de données"
Public Const TitleSearch = "Entrer les donnés à rechercher(Avertissement, la recherche peut être longue)"
        'these next string are used in frmJump:
    '3 diff titles for the form
Public Const TitleJump = "Choisissez la location du saut"
Public Const TitleDump = "Choisissez une Méthode d'extraction"
Public Const TitleInsert = "Choisissez un Méthode d'insertion"
    'visual elements
Public Const OK = "OK"
Public Const Cancel = "Annuler"
Public Const AddCurrent = "&Entrer la nouvelle location en tant que signet"
Public Const ManualJump = "&Adresse Manuel"
Public Const ManualDumpOrInsert = "&Localisation Manuel"
Public Const BookmarkLabel = "&Marque:"
Public Const DumpmarkLabel = "Marque d'&extraction:"
Public Const InsertmarkLabel = "Marque d'&insertion:"
    'messages
Public Const ManualMsg = "Entré l'adresse décimale ou hexadécimal; l'adresse Hexadécimal doit avoir le préfixe &H" & vbCrLf & vbCrLf & "(ex: &H100A)"
Public Const ManualTitle = "Adresse Manuel"
    'relsearch messages
Public Const NeedMoreThanTwoValues = "Vous avez besoin d'au moins deux valeur pour faire une recherche!"

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

