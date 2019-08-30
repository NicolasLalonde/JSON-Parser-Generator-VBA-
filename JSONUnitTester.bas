Attribute VB_Name = "JSONUnitTester"
Public Sub testJSON()
  Debug.Print vbNewLine & "   Testing JSON module"

  If Not testParseString() Then Debug.Print "X Error Parsing Strings"
  
  If Not testParseStringEscapes() Then Debug.Print "X Error Parsing Escape Strings"
  
  If Not testParseBoolean() Then Debug.Print "X Error Parsing Booleans"
  
  If Not testParseInteger() Then Debug.Print "X Error Parsing Integers"
  
  If Not testParseFloat() Then Debug.Print "X Error Parsing Floats"
  
  If Not testParseNegative() Then Debug.Print "X Error Parsing Negatives"
  
  If Not testParseExponent() Then Debug.Print "X Error Parsing Exponents"
  
  If Not testParseNull() Then Debug.Print "X Error Parsing Null"
  
  If Not testParseEmptyObject() Then Debug.Print "X Error Parsing Empty Object"
  
  If Not testParseEmptyArray() Then Debug.Print "X Error Parsing Empty Array"
  
  If Not testParseBaseArray() Then Debug.Print "X Error Parsing Base Array"
  
  If Not testParseNestedArray() Then Debug.Print "X Error Parsing Nested Array"
  
  If Not testParseObjectInArray() Then Debug.Print "X Error Parsing Object In Array"
  
  If Not testStringifyString() Then Debug.Print "X Error Stringifying String"
  
  If Not testStringifyEscapeString() Then Debug.Print "X Error Stringifying Escaped String"
  
  If Not testStringifyInteger() Then Debug.Print "X Error Stringifying Integer"
  
  If Not testStringifyFloat() Then Debug.Print "X Error Stringifying Float"
  
  If Not testStringifyBoolean() Then Debug.Print "X Error Stringifying Boolean"
  
  If Not testStringifyDate() Then Debug.Print "X Error Stringifying Date"
  
  If Not testStringifyCollection() Then Debug.Print "X Error Stringifying Collection"
  
  If Not testStringifyArray() Then Debug.Print "X Error Stringifying Array"
  
  If Not testParseIgnoresWhitespace() Then Debug.Print "X Error Parsing Whitespace"
  
  If Not testParseBigFile() Then Debug.Print "X Error Parsing Big JSON Files"
  
  Debug.Print "   Done JSON module"
End Sub

Private Function testParseString() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": ""DEF""}")
  testParseString = (dict("ABC") = "DEF")
End Function

Private Function testParseStringEscapes() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": ""\"" \\ \/ \b \f \n \r \t \u00DF""}")
  testParseStringEscapes = (dict("ABC") = """ \ / " & vbBack & " " & vbFormFeed & " " & vbNewLine & " " & vbCr & " " & vbTab & " " & "ß")
End Function

Private Function testParseBoolean() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": true}")
  Dim dict2 As Scripting.Dictionary
  Set dict2 = JSON.parse("{""ABC"": false}")
  testParseBoolean = (dict("ABC") = True) And (dict2("ABC") = False)
End Function

Private Function testParseInteger() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": 123}")
  testParseInteger = (dict("ABC") = 123)
End Function

Private Function testParseFloat() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": 123.1}")
  testParseFloat = (dict("ABC") = 123.1)
End Function

Private Function testParseNegative() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": -123}")
  testParseNegative = (dict("ABC") = -123)
End Function

Private Function testParseExponent() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": 10E2}")
  Dim dict2 As Scripting.Dictionary
  Set dict2 = JSON.parse("{""ABC"": 10e2}")
  Dim dict3 As Scripting.Dictionary
  Set dict3 = JSON.parse("{""ABC"": 10e+2}")
  testParseExponent = (dict("ABC") = 1000) And (dict2("ABC") = 1000) And (dict3("ABC") = 1000)
End Function

Private Function testParseNull() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": null}")
  testParseNull = IsNull(dict("ABC"))
End Function

Private Function testParseEmptyObject() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": {}}")
  testParseEmptyObject = TypeOf dict("ABC") Is Scripting.Dictionary And dict("ABC").count = 0
End Function

Private Function testParseEmptyArray() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{""ABC"": []}")
  testParseEmptyArray = TypeOf dict("ABC") Is Collection And dict("ABC").count = 0
End Function

Private Function testParseBaseArray() As Boolean
  Dim dict As Collection
  Set dict = JSON.parse("[""ABC""]")
  testParseBaseArray = dict(1) = "ABC"
End Function

Private Function testParseNestedArray() As Boolean
  Dim dict As Collection
  Set dict = JSON.parse("[[""ABC"", 24]]")
  testParseNestedArray = dict(1)(1) = "ABC" And dict(1)(2) = 24
End Function

Private Function testParseObjectInArray() As Boolean
  Dim dict As Collection
  Set dict = JSON.parse("[{""ABC"": 24}]")
  testParseObjectInArray = dict(1)("ABC") = 24
End Function

Private Function testStringifyString() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", "DEF"
  testStringifyString = JSON.stringify(dict, "", "") = "{""ABC"": ""DEF""}"
End Function

Private Function testStringifyEscapeString() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", """ \ / " & vbBack & " " & vbFormFeed & " " & vbNewLine & " " & vbCr & " " & vbTab & " " & "ß"
  testStringifyEscapeString = JSON.stringify(dict, "", "") = "{""ABC"": ""\"" \\ \/ \b \f \n \r \t ß""}"
End Function

Private Function testStringifyInteger() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", 1
  testStringifyInteger = JSON.stringify(dict, "", "") = "{""ABC"": 1}"
End Function

Private Function testStringifyFloat() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", 1.5
  testStringifyFloat = JSON.stringify(dict, "", "") = "{""ABC"": 1.5}"
End Function

Private Function testStringifyBoolean() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", True
  testStringifyBoolean = JSON.stringify(dict, "", "") = "{""ABC"": true}"
End Function

Private Function testStringifyDate() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary
  dict.add "ABC", CDate(Now())
  testStringifyDate = JSON.stringify(dict, "", "") = "{""ABC"": """ & Format(Now(), "yyyy-mm-dd") & """}"
End Function

Private Function testStringifyCollection() As Boolean
  Dim dict As Collection
  Set dict = New Collection
  dict.add 1
  dict.add 2
  dict.add 3
  testStringifyCollection = JSON.stringify(dict, "", "") = "[1,2,3]"
End Function

Private Function testStringifyArray() As Boolean
  Dim dict(3) As Integer
  For i = 0 To 3
    dict(i) = i
  Next i
  testStringifyArray = JSON.stringify(dict, "", "") = "[0,1,2,3]"
End Function

Private Function testParseIgnoresWhitespace() As Boolean
  Dim dict As Scripting.Dictionary
  Set dict = JSON.parse("{" & vbNewLine & vbTab & """ABC"":  1 " & vbNewLine & "}")
  
  testParseIgnoresWhitespace = JSON.stringify(dict, "", "") = "{""ABC"": 1}"
End Function

Private Function testParseBigFile() As Boolean
  'test parsing for big json file (ie bigger than an integer in length)
  Dim dict As Scripting.Dictionary
  Dim bigStr As String
  bigStr = "a"
  'get a string which is just 2^17 "a"s in a row
  For i = 0 To 16
    bigStr = bigStr & bigStr
  Next i
  Set dict = JSON.parse("{""ABC"": """ & bigStr & """}")
  testParseBigFile = (dict("ABC") = bigStr)
End Function
