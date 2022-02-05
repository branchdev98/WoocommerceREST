VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   17685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Update Data"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   8175
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   14895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "List Product"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function
Public Function HTMLEntititesDecode(p_strText As String) As String
Dim strTemp As String
strTemp = p_strText
strTemp = Replace(strTemp, "&quot;", """")
strTemp = Replace(strTemp, "&amp;", "&")
strTemp = Replace(strTemp, "&apos;", "'")
strTemp = Replace(strTemp, "&lt;", "<")
strTemp = Replace(strTemp, "&gt;", ">")
strTemp = Replace(strTemp, "&nbsp;", "")
strTemp = Replace(strTemp, "&iexcl;", "¡")
strTemp = Replace(strTemp, "&cent;", "¢")
strTemp = Replace(strTemp, "&pound;", "£")
strTemp = Replace(strTemp, "&curren;", "¤")
strTemp = Replace(strTemp, "&yen;", "¥")
strTemp = Replace(strTemp, "&brvbar;", "¦")
strTemp = Replace(strTemp, "&sect;", "§")
strTemp = Replace(strTemp, "&uml;", "¨")
strTemp = Replace(strTemp, "&copy;", "©")
strTemp = Replace(strTemp, "&ordf;", "ª")
strTemp = Replace(strTemp, "&laquo;", "«")
strTemp = Replace(strTemp, "&not;", "¬")
strTemp = Replace(strTemp, "*", "")
strTemp = Replace(strTemp, "&reg;", "®")
strTemp = Replace(strTemp, "&macr;", "¯")
strTemp = Replace(strTemp, "&deg;", "°")
strTemp = Replace(strTemp, "&plusmn;", "±")
strTemp = Replace(strTemp, "&sup2;", "²")
strTemp = Replace(strTemp, "&sup3;", "³")
strTemp = Replace(strTemp, "&acute;", "´")
strTemp = Replace(strTemp, "&micro;", "µ")
strTemp = Replace(strTemp, "&para;", "¶")
strTemp = Replace(strTemp, "&middot;", "·")
strTemp = Replace(strTemp, "&cedil;", "¸")
strTemp = Replace(strTemp, "&sup1;", "¹")
strTemp = Replace(strTemp, "&ordm;", "º")
strTemp = Replace(strTemp, "&raquo;", "»")
strTemp = Replace(strTemp, "&frac14;", "¼")
strTemp = Replace(strTemp, "&frac12;", "½")
strTemp = Replace(strTemp, "&frac34;", "¾")
strTemp = Replace(strTemp, "&iquest;", "¿")
strTemp = Replace(strTemp, "\u00c0", "À")
strTemp = Replace(strTemp, "\u00c1", "Á")
strTemp = Replace(strTemp, "\u00c2", "Â")
strTemp = Replace(strTemp, "\u00c3", "Ã")
strTemp = Replace(strTemp, "\u00c4", "Ä")
strTemp = Replace(strTemp, "\u00c5", "Å")
strTemp = Replace(strTemp, "\u00c6", "Æ")
strTemp = Replace(strTemp, "\u00c7", "Ç")
strTemp = Replace(strTemp, "\u00c8", "È")
strTemp = Replace(strTemp, "\u00c9", "É")
strTemp = Replace(strTemp, "\u00ca", "Ê")
strTemp = Replace(strTemp, "\u00cb", "Ë")
strTemp = Replace(strTemp, "\u00cc", "Ì")
strTemp = Replace(strTemp, "\u00cd", "Í")
strTemp = Replace(strTemp, "\u00ce", "Î")
strTemp = Replace(strTemp, "\u00cf", "Ï")
strTemp = Replace(strTemp, "\u00d0", "Ð")
strTemp = Replace(strTemp, "&Ntilde;", "Ñ")
strTemp = Replace(strTemp, "\u00d2", "Ò")
strTemp = Replace(strTemp, "\u00d3", "Ó")
strTemp = Replace(strTemp, "\u00d4", "Ô")
strTemp = Replace(strTemp, "\u00d5", "Õ")
strTemp = Replace(strTemp, "&Ouml;", "Ö")
strTemp = Replace(strTemp, "&times;", "×")
strTemp = Replace(strTemp, "&Oslash;", "Ø")
strTemp = Replace(strTemp, "&Ugrave;", "Ù")
strTemp = Replace(strTemp, "\u00da", "Ú")
strTemp = Replace(strTemp, "&Ucirc;", "Û")
strTemp = Replace(strTemp, "&Uuml;", "Ü")
strTemp = Replace(strTemp, "&Yacute;", "Ý")
strTemp = Replace(strTemp, "&THORN;", "Þ")
strTemp = Replace(strTemp, "&szlig;", "ß")
strTemp = Replace(strTemp, "\u00e0", "à")
strTemp = Replace(strTemp, "\u00e1", "á")
strTemp = Replace(strTemp, "\u00e2", "â")
strTemp = Replace(strTemp, "\u00e3", "ã")
strTemp = Replace(strTemp, "\u00e4", "ä")
strTemp = Replace(strTemp, "\u00e5", "å")
strTemp = Replace(strTemp, "\u00e6", "æ")
strTemp = Replace(strTemp, "\u00e7", "ç")
strTemp = Replace(strTemp, "\u00e8", "è")
strTemp = Replace(strTemp, "\u00e9", "é")
strTemp = Replace(strTemp, "\u00ea", "ê")
strTemp = Replace(strTemp, "\u00eb", "ë")
strTemp = Replace(strTemp, "\u00ec", "ì")
strTemp = Replace(strTemp, "\u00ed", "í")
strTemp = Replace(strTemp, "\u00ee", "î")
strTemp = Replace(strTemp, "\u00ef", "ï")
strTemp = Replace(strTemp, "\u00f0", "ð")
strTemp = Replace(strTemp, "\u00f1", "ñ")
strTemp = Replace(strTemp, "\u00f2", "ò")
strTemp = Replace(strTemp, "\u00f3", "ó")
strTemp = Replace(strTemp, "\u00f4", "ô")
strTemp = Replace(strTemp, "\u00f5", "õ")
strTemp = Replace(strTemp, "\u00f6", "ö")
strTemp = Replace(strTemp, "\u00f7", "÷")
strTemp = Replace(strTemp, "\u00f8", "ø")
strTemp = Replace(strTemp, "\u00f9", "ù")
strTemp = Replace(strTemp, "\u00fa", "ú")
strTemp = Replace(strTemp, "\u00fb", "û")
strTemp = Replace(strTemp, "\u00fc", "ü")
strTemp = Replace(strTemp, "\u00fd", "ý")
strTemp = Replace(strTemp, "\u00fe", "þ")
strTemp = Replace(strTemp, "\u00ff", "ÿ")
strTemp = Replace(strTemp, "&OElig;", "Œ")
strTemp = Replace(strTemp, "&oelig;", "œ")
strTemp = Replace(strTemp, "&Scaron;", "Š")
strTemp = Replace(strTemp, "&scaron;", "š")
strTemp = Replace(strTemp, "&Yuml;", "Ÿ")
strTemp = Replace(strTemp, "&fnof;", "ƒ")
strTemp = Replace(strTemp, "&circ;", "ˆ")
strTemp = Replace(strTemp, "&tilde;", "˜")
strTemp = Replace(strTemp, "&thinsp;", "")
strTemp = Replace(strTemp, "&zwnj;", "")
strTemp = Replace(strTemp, "&zwj;", "")
strTemp = Replace(strTemp, "&lrm;", "")
strTemp = Replace(strTemp, "&rlm;", "")
strTemp = Replace(strTemp, "&ndash;", "–")
strTemp = Replace(strTemp, "&mdash;", "—")
strTemp = Replace(strTemp, "&lsquo;", "‘")
strTemp = Replace(strTemp, "&rsquo;", "’")
strTemp = Replace(strTemp, "&sbquo;", "‚")
strTemp = Replace(strTemp, "&ldquo;", "“")
strTemp = Replace(strTemp, "&rdquo;", "”")
strTemp = Replace(strTemp, "&bdquo;", "„")
strTemp = Replace(strTemp, "&dagger;", "†")
strTemp = Replace(strTemp, "&Dagger;", "‡")
strTemp = Replace(strTemp, "&bull;", "•")
strTemp = Replace(strTemp, "&hellip;", "…")
strTemp = Replace(strTemp, "&permil;", "‰")
strTemp = Replace(strTemp, "&lsaquo;", "‹")
strTemp = Replace(strTemp, "&rsaquo;", "›")
strTemp = Replace(strTemp, "&euro;", "€")
strTemp = Replace(strTemp, "&trade;", "™")
HTMLEntititesDecode = strTemp
End Function



Private Sub Command1_Click()



' To authentication over HTTPS using query params, put the query params in the URL.

Dim success As Long


success = http.QuickGetSb("https://shop.provedoriageral.com.br/wp-json/wc/v3/products?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", sbResponseBody)
If (success = 0) Then
    Debug.Print http.LastErrorText
    Exit Sub
End If

Debug.Print "Response Body:"

Debug.Print sbResponseBody.GetAsString()
Dim respStatusCode As Long
respStatusCode = http.LastStatus
Debug.Print "Response Status Code = " & respStatusCode
If (respStatusCode >= 400) Then
    Debug.Print "Response Header:"
    Debug.Print http.LastHeader
    Debug.Print "Failed."
    Exit Sub
End If

End Sub



Private Sub Command2_Click()
 Set Req = New WinHttp.WinHttpRequest
  With Req
   .open "GET", "https://shop.provedoriageral.com.br/wp-json/wc/v3/products?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", False
   
       ' .open "GET", "https://shop.provedoriageral.com.br/wp-json/wc/v3/products?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", async:=False
      '  .SetRequestHeader "Content-Type", "application/json"
     '   .SetRequestHeader "Accept", "*/*"
        
        .send
        
        'Note: Normally you don't include all of this whitespace, but
        'we'll use it in this example:
        
      '  Label1.Caption = CStr(.Status) & " " & .StatusText & vbNewLine _
                       & .GetAllResponseHeaders() & vbNewLine _
                       & String$(40, "-") & vbNewLine _
                       & .ResponseText
                       .waitForResponse
        Text1.Text = HTMLEntititesDecode(.responseText)
        
    End With
    

End Sub

Private Sub Command3_Click()
Dim http As New ChilkatHttp
Dim success As Long

' Implements the following CURL command:

' curl -X PUT https://shop.provedoriageral.com.br/wp-json/wc/v3/products/61809 \
'     -u consumer_key:consumer_secret \
'     -H "Content-Type: application/json" \
'     -d '{
'   "regular_price": "24.54"
' }'

' Use the following online tool to generate HTTP code from a CURL command
' Convert a cURL Command to HTTP Source Code

http.BasicAuth = 1
http.login = "ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349"
http.password = "cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992"

' Use this online tool to generate code from sample JSON:
' Generate Code to Create JSON

' The following JSON is sent in the request body.

' {
'   "regular_price": "24.54"
' }

Dim json As New ChilkatJsonObject
success = json.UpdateString("regular_price", "44.78")

http.SetRequestHeader "Content-Type", "application/json"

Dim sbRequestBody As New ChilkatStringBuilder
success = json.EmitSb(sbRequestBody)

Dim resp As ChilkatHttpResponse
Set resp = http.PTextSb("PUT", "https://shop.provedoriageral.com.br/wp-json/wc/v3/products/61809?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", sbRequestBody, "utf-8", "application/json", 0, 0)
If (http.LastMethodSuccess = 0) Then
    Debug.Print http.LastErrorText
    Exit Sub
End If

Dim sbResponseBody As New ChilkatStringBuilder
success = resp.GetBodySb(sbResponseBody)
Dim jResp As New ChilkatJsonObject
success = jResp.LoadSb(sbResponseBody)
jResp.EmitCompact = 0

Debug.Print "Response Body:"
Debug.Print jResp.Emit()

Dim respStatusCode As Long
respStatusCode = resp.StatusCode
Debug.Print "Response Status Code = " & respStatusCode
If (respStatusCode >= 400) Then
    Debug.Print "Response Header:"
    Debug.Print resp.Header
    Debug.Print "Failed."

    Exit Sub
End If


Debug.Print "Example Completed."
End Sub

Private Sub Command4_Click()
Set Req = New WinHttp.WinHttpRequest
  With Req
   .open "PUT", "https://shop.provedoriageral.com.br/wp-json/wc/v3/products/61809?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", False
   
       ' .open "GET", "https://shop.provedoriageral.com.br/wp-json/wc/v3/products?consumer_key=ck_d1b69049aab3c4f6f9b2b3723e22b2e066c00349&consumer_secret=cs_4fdbde619f769b35ddccf0c6eb6a538ffe89c992", async:=False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Accept", "*/*"
      
        Dim JSONString As String
        
        JSONString = "{""regular_price"":""44.78"", ""description"":""testtesttest"", ""stock_quantity"":""40""}"
        .send JSONString
        
        'Note: Normally you don't include all of this whitespace, but
        'we'll use it in this example:
        
      '  Label1.Caption = CStr(.Status) & " " & .StatusText & vbNewLine _
                       & .GetAllResponseHeaders() & vbNewLine _
                       & String$(40, "-") & vbNewLine _
                       & .ResponseText
                       .waitForResponse
        Text1.Text = HTMLEntititesDecode(.responseText)
        
    End With
    
End Sub
