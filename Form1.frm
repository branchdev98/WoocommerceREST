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

Public Function HTMLEntititesDecode(p_strText As String) As String
Dim strTemp As String
strTemp = p_strText
strTemp = Replace(strTemp, "&quot;", """")
strTemp = Replace(strTemp, "&amp;", "&")
strTemp = Replace(strTemp, "&apos;", "'")
strTemp = Replace(strTemp, "&lt;", "<")
strTemp = Replace(strTemp, "&gt;", ">")
strTemp = Replace(strTemp, "&nbsp;", "")
strTemp = Replace(strTemp, "&iexcl;", "?")
strTemp = Replace(strTemp, "&cent;", "?")
strTemp = Replace(strTemp, "&pound;", "?")
strTemp = Replace(strTemp, "&curren;", "?")
strTemp = Replace(strTemp, "&yen;", "?")
strTemp = Replace(strTemp, "&brvbar;", "?")
strTemp = Replace(strTemp, "&sect;", "?")
strTemp = Replace(strTemp, "&uml;", "?")
strTemp = Replace(strTemp, "&copy;", "?")
strTemp = Replace(strTemp, "&ordf;", "?")
strTemp = Replace(strTemp, "&laquo;", "?")
strTemp = Replace(strTemp, "&not;", "?")
strTemp = Replace(strTemp, "*", "")
strTemp = Replace(strTemp, "&reg;", "?")
strTemp = Replace(strTemp, "&macr;", "?")
strTemp = Replace(strTemp, "&deg;", "?")
strTemp = Replace(strTemp, "&plusmn;", "?")
strTemp = Replace(strTemp, "&sup2;", "?")
strTemp = Replace(strTemp, "&sup3;", "?")
strTemp = Replace(strTemp, "&acute;", "?")
strTemp = Replace(strTemp, "&micro;", "?")
strTemp = Replace(strTemp, "&para;", "?")
strTemp = Replace(strTemp, "&middot;", "?")
strTemp = Replace(strTemp, "&cedil;", "?")
strTemp = Replace(strTemp, "&sup1;", "?")
strTemp = Replace(strTemp, "&ordm;", "?")
strTemp = Replace(strTemp, "&raquo;", "?")
strTemp = Replace(strTemp, "&frac14;", "?")
strTemp = Replace(strTemp, "&frac12;", "?")
strTemp = Replace(strTemp, "&frac34;", "?")
strTemp = Replace(strTemp, "&iquest;", "?")
strTemp = Replace(strTemp, "\u00c0", "?")
strTemp = Replace(strTemp, "\u00c1", "?")
strTemp = Replace(strTemp, "\u00c2", "?")
strTemp = Replace(strTemp, "\u00c3", "?")
strTemp = Replace(strTemp, "\u00c4", "?")
strTemp = Replace(strTemp, "\u00c5", "?")
strTemp = Replace(strTemp, "\u00c6", "?")
strTemp = Replace(strTemp, "\u00c7", "?")
strTemp = Replace(strTemp, "\u00c8", "?")
strTemp = Replace(strTemp, "\u00c9", "?")
strTemp = Replace(strTemp, "\u00ca", "?")
strTemp = Replace(strTemp, "\u00cb", "?")
strTemp = Replace(strTemp, "\u00cc", "?")
strTemp = Replace(strTemp, "\u00cd", "?")
strTemp = Replace(strTemp, "\u00ce", "?")
strTemp = Replace(strTemp, "\u00cf", "?")
strTemp = Replace(strTemp, "\u00d0", "?")
strTemp = Replace(strTemp, "&Ntilde;", "?")
strTemp = Replace(strTemp, "\u00d2", "?")
strTemp = Replace(strTemp, "\u00d3", "?")
strTemp = Replace(strTemp, "\u00d4", "?")
strTemp = Replace(strTemp, "\u00d5", "?")
strTemp = Replace(strTemp, "&Ouml;", "?")
strTemp = Replace(strTemp, "&times;", "?")
strTemp = Replace(strTemp, "&Oslash;", "?")
strTemp = Replace(strTemp, "&Ugrave;", "?")
strTemp = Replace(strTemp, "\u00da", "?")
strTemp = Replace(strTemp, "&Ucirc;", "?")
strTemp = Replace(strTemp, "&Uuml;", "?")
strTemp = Replace(strTemp, "&Yacute;", "?")
strTemp = Replace(strTemp, "&THORN;", "?")
strTemp = Replace(strTemp, "&szlig;", "?")
strTemp = Replace(strTemp, "\u00e0", "?")
strTemp = Replace(strTemp, "\u00e1", "?")
strTemp = Replace(strTemp, "\u00e2", "?")
strTemp = Replace(strTemp, "\u00e3", "?")
strTemp = Replace(strTemp, "\u00e4", "?")
strTemp = Replace(strTemp, "\u00e5", "?")
strTemp = Replace(strTemp, "\u00e6", "?")
strTemp = Replace(strTemp, "\u00e7", "?")
strTemp = Replace(strTemp, "\u00e8", "?")
strTemp = Replace(strTemp, "\u00e9", "?")
strTemp = Replace(strTemp, "\u00ea", "?")
strTemp = Replace(strTemp, "\u00eb", "?")
strTemp = Replace(strTemp, "\u00ec", "?")
strTemp = Replace(strTemp, "\u00ed", "?")
strTemp = Replace(strTemp, "\u00ee", "?")
strTemp = Replace(strTemp, "\u00ef", "?")
strTemp = Replace(strTemp, "\u00f0", "?")
strTemp = Replace(strTemp, "\u00f1", "?")
strTemp = Replace(strTemp, "\u00f2", "?")
strTemp = Replace(strTemp, "\u00f3", "?")
strTemp = Replace(strTemp, "\u00f4", "?")
strTemp = Replace(strTemp, "\u00f5", "?")
strTemp = Replace(strTemp, "\u00f6", "?")
strTemp = Replace(strTemp, "\u00f7", "?")
strTemp = Replace(strTemp, "\u00f8", "?")
strTemp = Replace(strTemp, "\u00f9", "?")
strTemp = Replace(strTemp, "\u00fa", "?")
strTemp = Replace(strTemp, "\u00fb", "?")
strTemp = Replace(strTemp, "\u00fc", "?")
strTemp = Replace(strTemp, "\u00fd", "?")
strTemp = Replace(strTemp, "\u00fe", "?")
strTemp = Replace(strTemp, "\u00ff", "?")
strTemp = Replace(strTemp, "&OElig;", "?")
strTemp = Replace(strTemp, "&oelig;", "?")
strTemp = Replace(strTemp, "&Scaron;", "?")
strTemp = Replace(strTemp, "&scaron;", "?")
strTemp = Replace(strTemp, "&Yuml;", "?")
strTemp = Replace(strTemp, "&fnof;", "?")
strTemp = Replace(strTemp, "&circ;", "?")
strTemp = Replace(strTemp, "&tilde;", "?")
strTemp = Replace(strTemp, "&thinsp;", "")
strTemp = Replace(strTemp, "&zwnj;", "")
strTemp = Replace(strTemp, "&zwj;", "")
strTemp = Replace(strTemp, "&lrm;", "")
strTemp = Replace(strTemp, "&rlm;", "")
strTemp = Replace(strTemp, "&ndash;", "?")
strTemp = Replace(strTemp, "&mdash;", "?")
strTemp = Replace(strTemp, "&lsquo;", "?")
strTemp = Replace(strTemp, "&rsquo;", "?")
strTemp = Replace(strTemp, "&sbquo;", "?")
strTemp = Replace(strTemp, "&ldquo;", "?")
strTemp = Replace(strTemp, "&rdquo;", "?")
strTemp = Replace(strTemp, "&bdquo;", "?")
strTemp = Replace(strTemp, "&dagger;", "?")
strTemp = Replace(strTemp, "&Dagger;", "?")
strTemp = Replace(strTemp, "&bull;", "?")
strTemp = Replace(strTemp, "&hellip;", "?")
strTemp = Replace(strTemp, "&permil;", "?")
strTemp = Replace(strTemp, "&lsaquo;", "?")
strTemp = Replace(strTemp, "&rsaquo;", "?")
strTemp = Replace(strTemp, "&euro;", "?")
strTemp = Replace(strTemp, "&trade;", "?")
HTMLEntititesDecode = strTemp
End Function





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
