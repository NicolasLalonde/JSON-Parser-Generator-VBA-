Attribute VB_Name = "JSON"
'STAND ALONE MODULE
'
'LIBRARIES:
' Microsoft Scripting Runtime


'INTERFACE:
' stringify(Scripting.Dictionary/Collection) As String: renders the dictionary into a JSON String
'   Raises Error 32101 if trying to convert an unsupported object
' parse(String, optional string, optional string) As Scripting.Dictionary: renders a JSON string into a Scripting dictionary, uses the first and second string as the newlines and tabs in a beautifier respectively
'   Raises Error 32100 if trying to convert a bad/non-JSON file

'GOAL:
' Provide an interface for simply converting to and from JSON strings / Scripting.Dictionary


'For scripting.dictionary you must enable microsoft scripting runtime, click Tools->References, scroll down to microsoft scripting runtime and enable

'The parser is based on the json standard which can be found on json.org
'The basic idea is to have 2 public methods, similar to those available in JavaScript
'one that converts a json string into a vba dictionary (parse)
'and one that does the reverse (stringify)

'the parsing method has 2 stages:

'Lexical analysis, where we parse the elements of the json
'     string into a collection of tokens, which is much easier to
'     manipulate, organise, and check for syntax
'     Here I decided that the string would be parsed into
'     the 8 different types of tokens listed in the TokenType enum
'     where the value token is always followed by its corresponding value
'     the token type is decided by the first character of the token
'     and the position is set to the end of the token if it was
'     successfully parsed, a bad token will generate an error
'     string position is passed and manipulated between functions

'Syntactic analysis, where we iterate over the array of tokens
'     and build the corresponding structure recursively with
'     2 methods, buildObject and buildArray, whose syntax rules
'     are different, an object element consists of
'     [string][colon] [value|string] seperated by [comma]
'     while for arrays an element is just [value|string] seperated by [comma]
'     buildObject is called when a startobject token is encountered, and
'     quits when a end object is encountered, buildArray works in a similar fashion
'     after parsing an element, the build methods check for an end token
'     or a comma, and if neither are encountered generates an error
'     also generates an error if an end token is found right after a comma
'     array position is passed and manipulated between functions


'The stringify method functions recursively as well, stringifying
'dictionaries into objects and then other instances of collections into arrays
'and converting primitive types (numbers, booleans and strings) to their JSON standard representation

'while collections CAN contain keyed entries, it's impossible to loop
'over them so it's assumed they are arrays, as only the values are iterable

'relies on small stringifier functions which are called according to the object type


'because json does not support dates, dates are converted to strings, and must
'be converted back explicitly to dates after the json is parsed
'dates do not follow the ISO standard, they are the string representation of vba dates

Private Enum TokenType
  ttStartObject = 1
  ttEndObject = 2
  ttStartArray = 3
  ttEndArray = 4
  ttColon = 5
  ttComma = 6
  ttValue = 7
  ttString = 8
End Enum

Private TabCharacter As String
Private NewLineCharacter As String


Public Function parse(text As String) As Variant
  Dim tokens As Collection
    Set tokens = tokenize(text)
  
  If tokens.item(1) = ttStartObject Then
    Set parse = buildObject(tokens)
  ElseIf tokens.item(1) = ttStartArray Then
    Set parse = buildArray(tokens)
  End If
End Function

Public Function stringify(element As Variant, Optional tabBeautifier As String = vbTab, Optional newLineBeautifier As String = vbNewLine) As String
  TabCharacter = tabBeautifier
  NewLineCharacter = newLineBeautifier
  stringify = elementToString(element)
End Function


'============================ LEXICAL ANALYSIS ==========================================================================================================





'tokenize the json string into tokens of type
'start object, end object, start array, end array,
'colon, comma, string, number, true, false, null
'variable character tokens include strings and numbers
'Takes normal ANSI string (watch out for the encoding)
Private Function tokenize(convertString As String, Optional startPosition As Long = 1) As Collection
  Dim jsonString As String
    jsonString = convertString
  Dim tokenArray As New Collection
  Dim position As Long
    position = startPosition
  Dim character As String
  Dim charCode As Long
  
  'For loop bounds are precompiled, cant change them while inside the loop so
  'we use a while loop instead
  While position <= Len(jsonString) 'reminder that vba strings are 1-indexed
    character = Mid(jsonString, position, 1)
    
    'start object
    If character = "{" Then
      tokenArray.add (ttStartObject)
      
    'end object
    ElseIf character = "}" Then
      tokenArray.add (ttEndObject)
      
    'start array
    ElseIf character = "[" Then
      tokenArray.add (ttStartArray)
      
    'end array
    ElseIf character = "]" Then
      tokenArray.add (ttEndArray)
    
    'colon
    ElseIf character = ":" Then
      tokenArray.add (ttColon)
    
    'comma
    ElseIf character = "," Then
      tokenArray.add (ttComma)
    
    'string
    ElseIf character = """" Then
      tokenArray.add (ttString)
      tokenArray.add parseString(jsonString, position)
      
    'number
    ElseIf isJSONFirstNumerical(character) Then
      tokenArray.add (ttValue)
      tokenArray.add parseNumber(jsonString, position)
      
    'boolean
    ElseIf character = "f" Or character = "t" Then
      tokenArray.add (ttValue)
      tokenArray.add parseBoolean(jsonString, position)
      
    'null
    ElseIf character = "n" Then
      tokenArray.add (ttValue)
      tokenArray.add parseNull(jsonString, position)
      
    'check for invalid characters (anything not whitespace)
    Else
      charCode = AscW(character)
      'whitespace characters specified by the json standard
      If charCode <> &H9 And charCode <> &HA And charCode <> &HD And charCode <> &H20 Then
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & _
                vbNewLine & "Error parsing token at character: """ & character & _
                """ at position: " & position
      End If
      
    End If
    
    
    position = position + CLng(1)
  Wend
  Set tokenize = tokenArray
End Function

Private Function isJSONFirstNumerical(character As String) As Boolean
  If InStr("-0123456789", character) = 0 Then
    isJSONFirstNumerical = False
  Else
    isJSONFirstNumerical = True
  End If
End Function


'side-effect: increments position to the end of the token
Private Function parseString(jsonString As String, ByRef position As Long) As String
'keep an escape flag, go until you meet a " and not in escape flag, replace escape characters by their proper values
  Dim returnString As String
  Dim lastPosition As Long
    lastPosition = position + CLng(1) 'start at character after quote
  Dim character As String
    character = Mid(jsonString, lastPosition, 1)
  
  While character <> """"
    If character = "\" Then
      'escape sequence
      lastPosition = lastPosition + CLng(1)
      character = Mid(jsonString, lastPosition, 1)
      
      If character = """" Then 'quotation mark
        returnString = returnString & character
        
      ElseIf character = "\" Then 'reverse solidus
        returnString = returnString & character
        
      ElseIf character = "/" Then 'solidus
        returnString = returnString & character
        
      ElseIf character = "b" Then 'backspace
        returnString = returnString & vbBack
        
      ElseIf character = "f" Then 'formfeed
        returnString = returnString & vbFormFeed
      
      ElseIf character = "n" Then 'newline
        returnString = returnString & vbNewLine
        
      ElseIf character = "r" Then 'carriage return
        returnString = returnString & vbCr
        
      ElseIf character = "t" Then 'horizontal tab
        returnString = returnString & vbTab
      
      ElseIf character = "u" Then 'hex representation
        lastPosition = lastPosition + CLng(1) 'get past the 'u'
        returnString = returnString & ChrW(CLng("&H" & Mid(jsonString, lastPosition, 4))) 'convert 4 next characters  (as hex) to a unicode char
        lastPosition = lastPosition + CLng(3)
        
      Else 'generate error
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & _
                               "At position: " & position & vbNewLine & _
                               """\" & character & """ is not a valid JSON escape sequence, " & _
                               "valid escape sequences are "" \"" "", "" \\ "", "" \/ "", "" \b "", "" \f "", "" \n "", "" \r "", "" \t "", and "" \uHHHH "" where the Hs are hexadecimal characters."
        
      End If
    
    Else 'any other unicode character
      returnString = returnString & character
    End If
  
    lastPosition = lastPosition + CLng(1)
    character = Mid(jsonString, lastPosition, 1)
  Wend
  position = lastPosition
  parseString = returnString

End Function

'side-effect: increments position to the end of the token
Private Function parseNumber(jsonString As String, ByRef position As Long) As Double
  Dim lastPosition As Long
    lastPosition = position
    
  'get the whole number, last position = bad number
  While InStr("-+.0123456789eE", Mid(jsonString, lastPosition, 1)) <> 0
    lastPosition = lastPosition + CLng(1)
  Wend

  On Error GoTo numError
    parseNumber = CDbl(Mid(jsonString, position, (lastPosition - position)))
    position = lastPosition - CLng(1) 'last character part of number
  
Exit Function
numError:
  Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & _
                       vbNewLine & "Error at position: " & position & _
                       vbNewLine & "Tried parsing a number token, instead got: """ & _
                       Mid(jsonString, position, (lastPosition - position)) & """"
End Function


'side-effect: increments position to the end of the token
Private Function parseBoolean(jsonString As String, ByRef position As Long) As Boolean
  If InStr(position, jsonString, "true") = position Then
    position = position + CLng(3)
    parseBoolean = True
    
  ElseIf InStr(position, jsonString, "false") = position Then
    position = position + CLng(4)
    parseBoolean = False
    
  Else
    Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & _
                        vbNewLine & "Error at position: " & position & _
                        vbNewLine & "Tried parsing boolean (""true"" or ""false"") token, instead got: """ & _
                        Mid(jsonString, position, 6) & "..."""
  End If
  
End Function


'side-effect: increments position to the end of the token
Private Function parseNull(jsonString As String, ByRef position As Long) As Variant
  If InStr(position, jsonString, "null") = position Then
    position = position + CLng(3)
    parseNull = Null
  Else
    Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & _
                      vbNewLine & "Error at position: " & position & _
                      vbNewLine & "Tried parsing ""null"" token, instead got: """ & _
                      Mid(jsonString, position, 5) & "..."""
  End If
End Function


 
'==================================== SYNTACTIC ANALYSIS ==========================================================

'side-effects: increments the passed index until the object is done
Private Function buildObject(tokenArray As Collection, Optional ByRef index As Long = 1) As Scripting.Dictionary
  
  Dim returnObject As New Scripting.Dictionary
  Dim currentToken As TokenType
    currentToken = tokenArray.item(index)
  Dim currentKey As String
  Dim currentValue As Variant
  
  If currentToken <> ttStartObject Then
    Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
    
  Else
    index = index + CLng(1)
    currentToken = tokenArray.item(index)
    
    While currentToken <> ttEndObject
    
      'parse key
      If currentToken <> ttString Then
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      Else
        index = index + CLng(1)
        currentKey = tokenArray.item(index)
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
      End If
    
      'parse colon
      If currentToken <> ttColon Then
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      Else
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
      End If
      
      'parse value
      If currentToken = ttStartObject Then
        Set currentValue = buildObject(tokenArray, index)
        currentToken = tokenArray.item(index)
      ElseIf currentToken = ttStartArray Then
        Set currentValue = buildArray(tokenArray, index)
        currentToken = tokenArray.item(index)
      ElseIf currentToken = ttString Or currentToken = ttValue Then
        index = index + CLng(1)
        currentValue = tokenArray.item(index)
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
      Else
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      End If
      
      returnObject.add currentKey, currentValue
      
      'parse comma
      If currentToken = ttComma Then
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
        If currentToken = ttEndObject Then Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      ElseIf currentToken <> ttEndObject Then
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      End If
    Wend
    
    index = index + CLng(1)
    
  End If
  
  Set buildObject = returnObject
  
End Function


'side-effects: increments the passed index until the object is done
Private Function buildArray(tokenArray As Collection, Optional ByRef index As Long = 1) As Collection
  
  Dim returnObject As New Collection
  Dim currentToken As TokenType
    currentToken = tokenArray.item(index)
  Dim currentKey As String
  Dim currentValue As Variant
  
  If currentToken <> ttStartArray Then
    Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
    
  Else
    index = index + CLng(1)
    currentToken = tokenArray.item(index)
    
    While currentToken <> ttEndArray
    
      'parse value
      If currentToken = ttStartObject Then
        Set currentValue = buildObject(tokenArray, index)
        currentToken = tokenArray.item(index)
      ElseIf currentToken = ttStartArray Then
        Set currentValue = buildArray(tokenArray, index)
        currentToken = tokenArray.item(index)
      ElseIf currentToken = ttString Or currentToken = ttValue Then
        index = index + CLng(1)
        currentValue = tokenArray.item(index)
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
      Else
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      End If
      
      returnObject.add currentValue
      
      'parse comma
      If currentToken = ttComma Then
        index = index + CLng(1)
        currentToken = tokenArray.item(index)
        If currentToken = ttEndObject Then Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      ElseIf currentToken <> ttEndArray Then
        Err.Raise 32100, description:="Invalid JSON Format:" & vbNewLine & "Unexpected token"
      End If
    Wend
    
    index = index + CLng(1)
    
  End If
  
  Set buildArray = returnObject
  
End Function





'=====================================================================================================================
'========================================= STRINGIFY =================================================================
'=====================================================================================================================

Private Function elementToString(element As Variant, Optional tabs As Long = 0) As String
  If TypeOf element Is Scripting.Dictionary Then
    Dim dict As Scripting.Dictionary
      Set dict = element
    elementToString = dictToString(dict, tabs)
  ElseIf isIterable(element) Then
    elementToString = arrayToString(element, tabs)
  'ElseIf TypeOf element Is JSONFormattable Then
  '  Dim obj As JSONFormattable
  '    Set obj = element
  '  elementToString = convertibleObjectToString(obj, tabs)
  ElseIf TypeName(element) = "Boolean" Then
    Dim bool As Boolean
      bool = element
    elementToString = booleanToString(bool)
  ElseIf TypeName(element) = "String" Then
    Dim str As String
      str = element
    elementToString = stringToString(str)
  ElseIf TypeName(element) = "Date" Then
    Dim adate As Date
      adate = element
    elementToString = dateToString(adate)
  ElseIf IsNumeric(element) Then
    elementToString = CStr(element)
  Else
    Dim obj2 As Object
      Set obj2 = element
    elementToString = objectToString(obj2)
  End If
End Function

Public Function isIterable(arr As Variant) As Boolean
  isIterable = True
  On Error GoTo notIterable
    For Each item In arr 'will immediately raise an exception for non-iterable objects
      Exit Function ' prevent iterating over the entire object
    Next item
    Exit Function 'catch empty collections
notIterable:
  isIterable = False
End Function


Private Function dictToString(dict As Scripting.Dictionary, Optional tabs As Long = 0) As String
  Dim returnString As String
    returnString = "{" & NewLineCharacter
    Dim element As Variant
  For Each element In dict.Keys
    If returnString <> "{" & NewLineCharacter Then returnString = returnString & ", " & NewLineCharacter 'add comma if not first element
    returnString = returnString & stringTabs(tabs + 1) & stringToString(CStr(element)) & ": " & elementToString(dict.item(element), tabs + 1)
  Next element
  dictToString = returnString & NewLineCharacter & stringTabs(tabs) & "}"
End Function


Private Function arrayToString(arr As Variant, Optional tabs As Long = 0) As String
  Dim returnString As String
    returnString = "[" & NewLineCharacter
    Dim element As Variant
  For Each element In arr
    If returnString <> "[" & NewLineCharacter Then returnString = returnString & "," & NewLineCharacter 'add comma if not first element
    returnString = returnString & stringTabs(tabs + 1) & elementToString(element, tabs + 1)
  Next element
  arrayToString = returnString & NewLineCharacter & stringTabs(tabs) & "]"
End Function

'Private Function convertibleObjectToString(obj As JSONFormattable, Optional tabs As Integer = 0) As String
'  convertibleObjectToString = dictToString(obj.dictFormat, tabs)
'End Function


Private Function stringToString(exString As String) As String
  Dim returnString As String
    returnString = """"
  Dim charPosition As Long
  Dim character As String
  Dim nlFlag As Boolean
  For charPosition = 1 To CLng(Len(exString))
    character = Mid(exString, charPosition, 1)
    
    'VBA converts newlines to carriage return + line feed if using windows
    ' which are two characters so we have to do a little wizardry
    If nlFlag Then 'skip the second character of newline since we already processed it
      nlFlag = False
      GoTo Continue
    End If
    If Len(vbNewLine) = 2 And Mid(exString, charPosition, 2) = vbNewLine Then
      character = vbNewLine
      nlFlag = True
    End If
    
    'if it's a special escape character
    If InStr("""\/" & vbBack & vbFormFeed & vbNewLine & vbCr & vbTab, character) Then
      returnString = returnString + "\"
      If InStr("""\/", character) Then
        returnString = returnString + character
      ElseIf character = vbBack Then
        returnString = returnString + "b"
      ElseIf character = vbFormFeed Then
        returnString = returnString + "f"
      ElseIf character = vbNewLine Then
        returnString = returnString + "n"
      ElseIf character = vbCr Then
        returnString = returnString + "r"
      ElseIf character = vbTab Then
        returnString = returnString + "t"
      End If
    Else
      returnString = returnString + character
    End If
Continue:
  Next charPosition
  stringToString = returnString & """"
End Function

Private Function booleanToString(bool As Boolean) As String
  If bool Then
    booleanToString = "true"
  Else
    booleanToString = "false"
  End If
End Function


Private Function dateToString(exdate As Date) As String
  dateToString = """" & Format(exdate, "yyyy-mm-dd") & """"
End Function


Private Function objectToString(obj As Object) As String
  If obj Is Nothing Then
    objectToString = "null"
  Else
    Err.Raise 32101, description:="Error converting to JSON: " & vbNewLine & _
                              "Cannot convert this object to a JSON format, can only convert these objects: Scripting.Dictionary, Collection, String, Date, Numerical formats(Integers, Long, etc.), Boolean, Null Objects"
  End If
End Function







'=============================== HELPERS =======================================

'returns a string containing a specified number of tabs (beautifying a json string)
Private Function stringTabs(numberOfTabs As Long) As String
  Dim i As Long
  Dim returnString As String

  For i = 1 To numberOfTabs
    returnString = returnString + TabCharacter
  Next
  stringTabs = returnString
End Function


'for lexical analysis debugging
Private Function printStack(tokens As Collection) As String
  Dim i As Long
  Dim returnString As String
  Dim token As Variant
  i = 1
  While i <= tokens.count
    returnString = returnString & i & ". "
    token = tokens.item(i)
    If token = ttStartObject Then
      returnString = returnString & "{"
    ElseIf token = ttEndObject Then
      returnString = returnString & "}"
    ElseIf token = ttStartArray Then
      returnString = returnString & "["
    ElseIf token = ttEndArray Then
      returnString = returnString & "]"
    ElseIf token = ttComma Then
      returnString = returnString & ","
    ElseIf token = ttColon Then
      returnString = returnString & ":"
    ElseIf token = ttString Or token = ttValue Then
      i = i + 1
      token = tokens.item(i)
      returnString = returnString & "(value) & " & i & ". " & token
    End If
    i = i + 1
    returnString = returnString & vbNewLine
  Wend
  printStack = returnString
End Function

