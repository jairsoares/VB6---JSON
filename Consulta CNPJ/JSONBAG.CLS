VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe Ajustada por Jair Soares
' Email: jairsoares.css@gmail.com
' CSS-Sistemas Software House
' www.css-sistemas.com
'--------------------------------------------

Option Explicit

Private Const CLASS_VERSION As String = "1.7"

Private Const LBRACE As String = "{"
Private Const RBRACE As String = "}"
Private Const LBRACKET As String = "["
Private Const RBRACKET As String = "]"
Private Const COLON As String = ":"
Private Const COMMA As String = ","
Private Const QUOTE As String = """"
Private Const PLUS As String = "+"
Private Const MINUS As String = "-"
Private Const RADIXPOINT As String = "."
Private Const ZERO As String = "0"
Private Const NINE As String = "9"
Private Const REVSOLIDUS As String = "\"

Private Const WHITE_SPACE As String = vbTab & vbLf & vbCr & " "

Private Const S_OK As Long = 0
Private Const VARIANT_ALPHABOOL As Long = &H2&
Private Const LOCALE_INVARIANT As Long = 127&

Private Declare Function HashData Lib "shlwapi" ( _
    ByVal pbData As Long, _
    ByVal cbData As Long, _
    ByVal pbHash As Long, _
    ByVal cbHash As Long) As Long

Private Declare Function StrSpn Lib "shlwapi" Alias "StrSpnW" ( _
    ByVal psz As Long, _
    ByVal pszSet As Long) As Long

Private Declare Function VariantChangeTypeEx Lib "oleaut32" ( _
    ByRef vargDest As Variant, _
    ByRef varSrc As Variant, _
    ByVal lcid As Long, _
    ByVal wFlags As Integer, _
    ByVal vt As VbVarType) As Long

Private TypeNameOfMe As String
Private Names As Collection
Private Values As Collection
Private CursorIn As Long
Private LengthIn As Long
Private TextOut As String
Private CursorOut As Long
Private NumberType As VbVarType

Private vbUS As String

Private mIsArray As Boolean
Private mDecimalMode As Boolean


Public Whitespace As Boolean
Attribute Whitespace.VB_VarDescription = "True to insert whitespace when reading serialized contents string."

Public Property Get Count() As Long
Attribute Count.VB_Description = "Count of Items in the list."
    Count = Values.Count
End Property

Public Property Get DecimalMode() As Boolean
Attribute DecimalMode.VB_Description = "Causes numbers to be parsed and stored as Decimal type instead of Double."
    DecimalMode = mDecimalMode
End Property

Public Property Let DecimalMode(ByVal RHS As Boolean)
    mDecimalMode = RHS
    If mDecimalMode Then
        NumberType = vbDecimal
    Else
        NumberType = vbDouble
    End If
End Property

Public Property Let IsArray(ByVal RHS As Boolean)
    If Values.Count > 0 Then
        Err.Raise 5, TypeNameOfMe, "Cannot change IsArray setting after items have been added"
    Else
        mIsArray = RHS
        If mIsArray Then Set Names = Nothing
    End If
End Property

Public Property Get IsArray() As Boolean
Attribute IsArray.VB_Description = "True if this object is a JSON array instead of a JSON object.  Must be set before first item is added."
    IsArray = mIsArray
End Property

'Default property.
Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_Description = "Retrieve an Item by Key or Index, change or add a new Item."
Attribute Item.VB_UserMemId = 0
    'Retrieval works either by key or index for "objects" but only
    'by index for "arrays."
    
    Dim PrefixedKey As String
    
    If IsNull(Key) Then Err.Raise 94, TypeNameOfMe, "Key must be String or an index)"
    If VarType(Key) = vbString Then
        If mIsArray Then
            Err.Raise 5, TypeNameOfMe, "Array values can only be acessed by index"
        End If
        
        PrefixedKey = PrefixHash(Key)
        If IsObject(Values.Item(PrefixedKey)) Then
            Set Item = Values.Item(PrefixedKey)
        Else
            Item = Values.Item(PrefixedKey)
        End If
    Else
        If IsObject(Values.Item(Key)) Then
            Set Item = Values.Item(Key)
        Else
            Item = Values.Item(Key)
        End If
    End If
End Property

Public Property Let Item(Optional ByVal Key As Variant = Null, ByVal RHS As Variant)

    Dim PrefixedKey As String
    
    With Values
        If mIsArray Then
            If VarType(Key) = vbString Then
                Err.Raise 5, TypeNameOfMe, "Array values can only be changed by index or added via Null"
            End If
            
            If IsNull(Key) Then
                .Add RHS
            Else
                If Key > .Count Then
                    .Add RHS
                Else
                    .Remove Key
                    .Add RHS, , Key
                End If
            End If
        Else
            If VarType(Key) <> vbString Then
                Err.Raise 5, TypeNameOfMe, "Object values can only be changed or added by key not by index"
            End If
            
            PrefixedKey = PrefixHash(Key)
            On Error Resume Next
            .Add RHS, PrefixedKey
            If Err Then
                On Error GoTo 0
                .Remove PrefixedKey
                .Add RHS, PrefixedKey
                Names.Remove PrefixedKey
            Else
                On Error GoTo 0
            End If

            Names.Add Key, PrefixedKey
        End If
    End With
End Property

Public Property Set Item(ByVal Key As Variant, ByVal RHS As Variant)
    Item(Key) = RHS
End Property

Public Property Get JSON() As String
Attribute JSON.VB_Description = "A string representing the serialized contents of the object."
    CursorOut = 1
    SerializeItem vbNullString, Me
    JSON = Left$(TextOut, CursorOut - 1)
    TextOut = ""
End Property

Public Property Let JSON(ByRef RHS As String)
    Clear
    
    CursorIn = 1
    LengthIn = Len(RHS)
    
    SkipWhitespace RHS
    
    Select Case Mid$(RHS, CursorIn, 1)
        Case LBRACE
            CursorIn = CursorIn + 1
            ParseObject RHS, CursorIn, Len(RHS)
        Case LBRACKET
            CursorIn = CursorIn + 1
            ParseArray RHS, CursorIn, Len(RHS)
        Case Else
            Error13A "either " & LBRACE & " or " & LBRACKET, CursorIn
    End Select
End Property

Public Property Get Name(ByVal Index As Long) As String
Attribute Name.VB_Description = "Retrieves list Item Name by Index."
    If mIsArray Then Err.Raise 5, TypeNameOfMe, "Array items do not have names"
    
    Name = Names.Item(Index)
End Property

Public Property Get Version() As String()
    Version = Split(CLASS_VERSION)
End Property


Public Function AddNewArray(Optional ByVal Key As Variant = Null) As JsonBag
Attribute AddNewArray.VB_Description = "Create new ""array"" type Item and add it to the list, returning a reference to it."
    Dim NewArray As JsonBag
    
    Set NewArray = New JsonBag
    NewArray.IsArray = True
    Set Item(Key) = NewArray
    Set AddNewArray = NewArray
End Function

Public Function AddNewObject(Optional ByVal Key As Variant = Null) As JsonBag
Attribute AddNewObject.VB_Description = "Create new ""object"" type Item and add it to the list, returning a reference to it."
    Dim NewObject As JsonBag
    
    Set NewObject = New JsonBag
    Set Item(Key) = NewObject
    Set AddNewObject = NewObject
End Function

Public Sub Clear()
Attribute Clear.VB_Description = "Clears all data and sets IsArray to False."
    Set Names = New Collection
    Set Values = New Collection
    mIsArray = False
End Sub

Public Function Exists(ByVal Key As Variant) As Boolean
Attribute Exists.VB_Description = "Returns True if item specified by Key or Index is present."
    Dim Name As String
    
    On Error Resume Next
    Name = Names.Item(Key)
    Exists = Err.Number = 0
    Err.Clear
End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Iterates over the Item names."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    If mIsArray Then Err.Raise 5, TypeNameOfMe, "Arrays must be iterated using index values"
    
    Set NewEnum = Names.[_NewEnum]
End Function

Public Sub Remove(ByVal Key As Variant)
Attribute Remove.VB_Description = "Removes Item specified by Key or Index."

    Dim PrefixedKey As String
    
    If VarType(Key) = vbString Then
        If mIsArray Then Err.Raise 5, TypeNameOfMe, "Must remove by index for arrays"
        
        PrefixedKey = PrefixHash(Key)
        On Error Resume Next
        Names.Remove PrefixedKey
        If Err Then
            Exit Sub
        End If
        On Error GoTo 0
        Values.Remove PrefixedKey
    Else
        If Key < Values.Count Then
            Values.Remove Key
            If Not IsArray Then Names.Remove Key
        End If
    End If
End Sub



Friend Sub ParseArray( _
    ByRef Text As String, _
    ByRef StartCursor As Long, _
    ByVal TextLength As Long)
    
    Dim ArrayValue As Variant
    
    CursorIn = StartCursor
    LengthIn = TextLength

    mIsArray = True
    Do
        SkipWhitespace Text
        Select Case Mid$(Text, CursorIn, 1)
            Case COMMA
                CursorIn = CursorIn + 1
            Case RBRACKET
                CursorIn = CursorIn + 1
                Exit Do
            Case Else
                ParseValue Text, ArrayValue
                Values.Add ArrayValue
        End Select
    Loop
    StartCursor = CursorIn
End Sub

Friend Sub ParseObject( _
    ByRef Text As String, _
    ByRef StartCursor As Long, _
    ByVal TextLength As Long)
    
    Dim Char As String
    Dim ItemName As String
    Dim Value As Variant
    Dim FoundFirstItem As Boolean
    
    CursorIn = StartCursor
    LengthIn = TextLength
    
    Do
        SkipWhitespace Text
        Char = Mid$(Text, CursorIn, 1)
        CursorIn = CursorIn + 1
        Select Case Char
            Case QUOTE
                ItemName = ParseName(Text)
                ParseValue Text, Value
                Item(ItemName) = Value
                FoundFirstItem = True
            Case COMMA
                If Not FoundFirstItem Then
                    Err.Raise 13, TypeNameOfMe, "Found "","" before first item at character " & CStr(CursorIn - 1)
                End If
            Case RBRACE
                Exit Do
            Case Else
                Error13A ", or }", CursorIn - 1
        End Select
    Loop
    StartCursor = CursorIn
End Sub


Private Sub Cat(ByRef NewText As String)
    Const TEXT_CHUNK As Long = 512
    Dim LenNew As Long
    
    LenNew = Len(NewText)
    If LenNew > 0 Then
        If CursorOut + LenNew - 1 > Len(TextOut) Then
            If LenNew > TEXT_CHUNK Then
                TextOut = TextOut & Space$(LenNew + TEXT_CHUNK)
            Else
                TextOut = TextOut & Space$(TEXT_CHUNK)
            End If
        End If
        Mid$(TextOut, CursorOut, LenNew) = NewText
        CursorOut = CursorOut + LenNew
    End If
End Sub

Private Sub Error13A(ByVal Symbol As String, ByVal Position As Long)
    Err.Raise 13, TypeNameOfMe, "Expected " & Symbol & " at character " & CStr(Position)
End Sub

Private Sub Error13B(ByVal Position As Long)
    Err.Raise 13, TypeNameOfMe, "Bad string character escape at character " & CStr(Position)
End Sub

Private Function ParseName(ByRef Text As String) As String
    ParseName = ParseString(Text)
    
    SkipWhitespace Text
    If Mid$(Text, CursorIn, 1) <> COLON Then
        Error13A COLON, CursorIn
    End If
    CursorIn = CursorIn + 1
End Function

Private Function ParseNumber(ByRef Text As String) As Variant
    Dim SaveCursor As Long
    Dim BuildString As String
    Dim BuildCursor As Long
    Dim Char As String
    Dim GotDecPoint As Boolean
    Dim GotExpSign As Boolean
    
    SaveCursor = CursorIn
    BuildString = Space$(LengthIn - CursorIn + 1)

    BuildCursor = 1
    Mid$(BuildString, 1, 1) = Mid$(Text, CursorIn, 1)
    
    For CursorIn = CursorIn + 1 To LengthIn
        Char = LCase$(Mid$(Text, CursorIn, 1))
        Select Case Char
            Case RADIXPOINT
                If GotDecPoint Then
                    Err.Raise 13, TypeNameOfMe, "Second decimal point at character " & CStr(CursorIn)
                End If
                If Mid$(BuildString, BuildCursor, 1) = MINUS Then
                    Err.Raise 13, TypeNameOfMe, "Digit expected at character " & CStr(CursorIn)
                End If
                GotDecPoint = True
            Case ZERO To NINE
            Case "e"
                CursorIn = CursorIn + 1
                Exit For
            Case Else
                Exit For
        End Select
        BuildCursor = BuildCursor + 1
        Mid$(BuildString, BuildCursor, 1) = Char
    Next
    
    If Char = "e" Then
        BuildCursor = BuildCursor + 1
        Mid$(BuildString, BuildCursor, 1) = Char
        
        For CursorIn = CursorIn To LengthIn
            Char = Mid$(Text, CursorIn, 1)
            Select Case Char
                Case PLUS, MINUS
                    If GotExpSign Then
                        Err.Raise 13, TypeNameOfMe, "Second exponent sign at character " & CStr(CursorIn)
                    End If
                    GotExpSign = True
                Case ZERO To NINE
                    'Do nothing.
                Case Else
                    Exit For
            End Select
            BuildCursor = BuildCursor + 1
            Mid$(BuildString, BuildCursor, 1) = Char
        Next
    End If
    
    If CursorIn > LengthIn Then
        Err.Raise 13, TypeNameOfMe, "Ran off end of string while parsing a number"
    End If
    
    ParseNumber = Left$(BuildString, BuildCursor)
    If VariantChangeTypeEx(ParseNumber, ParseNumber, LOCALE_INVARIANT, 0, NumberType) <> S_OK Then
        Err.Raise 6, TypeNameOfMe, "Number overflow or parse error at character " & CStr(SaveCursor)
    End If
End Function

Private Function ParseString(ByRef Text As String) As String
    Dim BuildCursor As Long
    Dim Char As String
    
    ParseString = Space$(LengthIn - CursorIn + 1)
    
    For CursorIn = CursorIn To LengthIn
        Char = Mid$(Text, CursorIn, 1)
        Select Case Char
            Case vbNullChar To vbUS
                Err.Raise 13, TypeNameOfMe, "Invalid string character at " & CStr(CursorIn)
            Case REVSOLIDUS
                CursorIn = CursorIn + 1
                If CursorIn > LengthIn Then
                    Error13B CursorIn
                End If
                Char = LCase$(Mid$(Text, CursorIn, 1))
                Select Case Char
                    Case QUOTE, REVSOLIDUS, "/"
                        'Do nothing.
                    Case "b"
                        Char = vbBack
                    Case "f"
                        Char = vbFormFeed
                    Case "n"
                        Char = vbLf
                    Case "r"
                        Char = vbCr
                    Case "t"
                        Char = vbTab
                    Case "u"
                        CursorIn = CursorIn + 1
                        If LengthIn - CursorIn < 3 Then
                            Error13B CursorIn
                        End If
                        On Error Resume Next
                        Char = ChrW$(CLng("&H0" & Mid$(Text, CursorIn, 4)))
                        If Err Then
                            On Error GoTo 0
                            Error13B CursorIn
                        End If
                        On Error GoTo 0
                        CursorIn = CursorIn + 3
                    Case Else
                        Error13B CursorIn
                End Select
            Case QUOTE
                CursorIn = CursorIn + 1
                Exit For

        End Select
        BuildCursor = BuildCursor + 1
        Mid$(ParseString, BuildCursor, 1) = Char
    Next
    
    If CursorIn > LengthIn Then
        Error13A QUOTE, LengthIn + 1
    End If
    ParseString = Left$(ParseString, BuildCursor)
End Function

Private Sub ParseValue(ByRef Text As String, ByRef Value As Variant)
    Dim SubBag As JsonBag
    Dim Token As String
    
    SkipWhitespace Text
    Select Case Mid$(Text, CursorIn, 1)
        Case QUOTE
            CursorIn = CursorIn + 1
            Value = ParseString(Text)
        Case LBRACE
            CursorIn = CursorIn + 1
            Set SubBag = New JsonBag
            SubBag.DecimalMode = DecimalMode
            SubBag.ParseObject Text, CursorIn, LengthIn
            Set Value = SubBag
        Case LBRACKET
            CursorIn = CursorIn + 1
            Set SubBag = New JsonBag
            SubBag.DecimalMode = DecimalMode
            SubBag.ParseArray Text, CursorIn, LengthIn
            Set Value = SubBag
        Case MINUS, ZERO To NINE
            Value = ParseNumber(Text)
        Case Else
            Token = LCase$(Mid$(Text, CursorIn, 4))
            If Token = "null" Then
                Value = Null
                CursorIn = CursorIn + 4
            ElseIf Token = "true" Then
                Value = True
                CursorIn = CursorIn + 4
            Else
                Token = LCase$(Mid$(Text, CursorIn, 5))
                If Token = "false" Then
                    Value = False
                    CursorIn = CursorIn + 5
                Else
                    Err.Raise 13, TypeNameOfMe, "Bad value at character " & CStr(CursorIn)
                End If
            End If
    End Select
End Sub

Private Function PrefixHash(ByVal KeyString As String) As String
    Dim Hash As Long
    
    HashData StrPtr(KeyString), 2 * Len(KeyString), VarPtr(Hash), 4
    PrefixHash = Right$("0000000" & Hex$(Hash), 8) & KeyString
End Function

Private Sub SerializeItem( _
    ByVal ItemName As String, _
    ByRef Item As Variant, _
    Optional ByVal Level As Integer)
    
    Const TEXT_CHUNK As Long = 64
    Dim Indent As String
    Dim Anonymous As Boolean
    Dim Name As Variant
    Dim ItemIndex As Long
    Dim TempItem As Variant
    Dim ItemBag As JsonBag
    Dim SubBag As JsonBag
    Dim ItemText As String
    Dim ArrayItem As Variant
    
    If Whitespace Then
        Indent = Space$(4 * Level)
    End If
    
    Anonymous = StrPtr(ItemName) = 0
    If Not Anonymous Then
        If Whitespace Then Cat Indent
        Cat SerializeString(ItemName) & COLON
    End If
    
    Select Case VarType(Item)
        Case vbEmpty, vbNull
            If Whitespace And Anonymous Then Cat Indent
            Cat "null"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbBoolean
            If Whitespace And Anonymous Then Cat Indent
            If VariantChangeTypeEx(TempItem, _
                                   Item, _
                                   LOCALE_INVARIANT, _
                                   VARIANT_ALPHABOOL, _
                                   vbString) <> S_OK Then
                Err.Raise 51, TypeNameOfMe, ItemName & ", value " & CStr(Item) & " failed to serialize"
            End If
            Cat LCase$(TempItem)
        Case vbString
            If Whitespace And Anonymous Then Cat Indent
            Cat SerializeString(Item)
        Case vbObject
            Set ItemBag = Item
            If ItemBag.IsArray Then
                If Whitespace And Anonymous Then Cat Indent
                Cat LBRACKET
                If ItemBag.Count < 1 Then
                    Cat RBRACKET
                Else
                    If Whitespace Then Cat vbNewLine
                    With ItemBag
                        For ItemIndex = 1 To .Count
                            If IsObject(.Item(ItemIndex)) Then
                                Set TempItem = .Item(ItemIndex)
                            Else
                                TempItem = .Item(ItemIndex)
                            End If
                            SerializeItem vbNullString, TempItem, Level + 1
                            Cat COMMA
                            If Whitespace Then Cat vbNewLine
                        Next
                    End With
                    If Whitespace Then
                        CursorOut = CursorOut - 3
                        Cat vbNewLine & Indent & RBRACKET
                    Else
                        Mid$(TextOut, CursorOut - 1) = RBRACKET
                    End If
                End If
            Else
                If Whitespace And Anonymous Then Cat Indent
                Cat LBRACE
                If ItemBag.Count < 1 Then
                    Cat RBRACE
                Else
                    If Whitespace Then Cat vbNewLine
                    For Each Name In ItemBag
                        If IsObject(ItemBag.Item(Name)) Then
                            Set TempItem = ItemBag.Item(Name)
                        Else
                            TempItem = ItemBag.Item(Name)
                        End If
                        SerializeItem Name, TempItem, Level + 1
                        Cat COMMA
                        If Whitespace Then Cat vbNewLine
                    Next
                    If Whitespace Then
                        CursorOut = CursorOut - 3
                        Cat vbNewLine & Indent & RBRACE
                    Else
                        Mid$(TextOut, CursorOut - 1) = RBRACE
                    End If
                End If
            End If
        Case Else
            Err.Raise 51, TypeNameOfMe, ItemName & ", unknown/unsupported type = " & CStr(VarType(Item))
    End Select
End Sub

Private Function SerializeString(ByVal Text As String) As String
    Dim BuildString As String
    Dim BuildCursor As Long
    Dim TextCursor As Long
    Dim Char As String
    Dim intChar As Integer

    BuildString = Space$(3 * Len(Text) \ 2)
    BuildCursor = 1
    StringCat BuildString, BuildCursor, QUOTE
    For TextCursor = 1 To Len(Text)
        Char = Mid$(Text, TextCursor, 1)
        Select Case Char
            Case QUOTE, REVSOLIDUS
                StringCat BuildString, BuildCursor, REVSOLIDUS & Char
            Case vbBack
                StringCat BuildString, BuildCursor, REVSOLIDUS & "b"
            Case vbFormFeed
                StringCat BuildString, BuildCursor, REVSOLIDUS & "f"
            Case vbLf
                StringCat BuildString, BuildCursor, REVSOLIDUS & "n"
            Case vbCr
                StringCat BuildString, BuildCursor, REVSOLIDUS & "r"
            Case vbTab
                StringCat BuildString, BuildCursor, REVSOLIDUS & "t"
            Case " " To "!", "#" To LBRACKET, RBRACKET To "~"
                StringCat BuildString, BuildCursor, Char
            Case Else
                intChar = AscW(Char)
                Select Case intChar
                    Case 0 To &H1F, &H7F To &H9F, &H34F, &H200B To &H200F, _
                         &H2028 To &H202E, &H2060, &HFE01 To &HFE0F, _
                         &HFEFF, &HFFFD, &HD800 To &HDFFF
                        StringCat BuildString, BuildCursor, _
                               REVSOLIDUS & "u" & Right$("000" & Hex$(intChar), 4)
                    Case Else
                        StringCat BuildString, BuildCursor, Char
                End Select
        End Select
    Next
    StringCat BuildString, BuildCursor, QUOTE
    SerializeString = Left$(BuildString, BuildCursor - 1)
End Function

Private Sub SkipWhitespace(ByRef Text As String)
    CursorIn = CursorIn + StrSpn(StrPtr(Text) + 2 * (CursorIn - 1), StrPtr(WHITE_SPACE))
End Sub

Private Sub StringCat(ByRef TextOut As String, ByRef CursorOut, ByRef NewText As String)
    Const TEXT_CHUNK As Long = 64
    Dim LenNew As Long
    
    LenNew = Len(NewText)
    If LenNew > 0 Then
        If CursorOut + LenNew - 1 > Len(TextOut) Then
            If LenNew > TEXT_CHUNK Then
                TextOut = TextOut & Space$(LenNew + TEXT_CHUNK)
            Else
                TextOut = TextOut & Space$(TEXT_CHUNK)
            End If
        End If
        Mid$(TextOut, CursorOut, LenNew) = NewText
        CursorOut = CursorOut + LenNew
    End If
End Sub

Private Sub Class_Initialize()
    TypeNameOfMe = TypeName(Me)
    vbUS = ChrW$(&H1F&)
    DecimalMode = False
    
    Clear
End Sub
