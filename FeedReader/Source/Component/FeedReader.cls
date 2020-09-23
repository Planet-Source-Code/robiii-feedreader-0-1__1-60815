VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeedReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'AUTHOR         : Rob Janssen
'DATE           : May 30th, 2005
'Version        : 0.1

'DESCRIPTION    : This class will read an RSS (0.9x/2.0), ATOM (0.3) or RDF (1.0) feed from
'                 a given URL and return all items in it with some basic properties.
'                 It can be easily extended to return lots of other properties. This object
'                 supports proxy servers.

'REQUIREMENTS   : This class requires VBScript 5.6 > 0 to be installed  (http://msdn.microsoft.com/library/default.asp?url=/downloads/list/webdev.asp)
'                 Also it requires any version of MSXML to be installed (http://msdn.microsoft.com/XML/XMLDownloads/default.aspx)
'                 Also, it tries to create a WinHTTPRequest, XMLHTTP or ServerXMLHTTP object to retrieve the feed.

Private oRE As Object   'Our Regular Expression Object

'Some Consts for the WinHTTP object
Private Const HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0
Private Const HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1

'Feed types
Public Enum eFeedType
    AutoDetect = 0
    RSS = 1
    RDF = 2
    Atom = 3
End Enum

'Internal variables
Private m_lLastHTTPStatus As Long
Private m_sLastHTTPStatusText As String
Private m_lLastErrorCode As Long
Private m_sLastErrorDesc As String
Private m_colItems As Collection
Private m_lFeedType As eFeedType
Private m_strProxyUser As String
Private m_strProxyPass As String

Private Sub Class_Initialize()
    Clear  'No Errors... yet ;-)
    'Initial values for proxy settings
    m_strProxyUser = ""
    m_strProxyPass = ""
    
    'Create our regexp object
    Set oRE = fCreateObject("VBScript.RegExp")
    If Not oRE Is Nothing Then
        With oRE
            .IgnoreCase = True
            .Global = True
        End With
    End If
End Sub

Private Sub Class_Terminate()
    'Free the RegExp object
    Set oRE = Nothing
End Sub

'Main function to read a feed
Public Function ReadFeed(sURL As String, Optional sUserName As String = "", Optional sPassWord As String = "", Optional lType As eFeedType = AutoDetect, Optional bAllowHTML As Boolean = False) As Boolean
    Dim oXMLDoc As Object
    Dim bParseOK As Boolean
    Dim arrObject() As String
    Dim T As Long
    
    On Error GoTo errReadFeed
    
    bParseOK = False    'Not OK (yet)
    If Not (oRE Is Nothing) Then    'Is the RegExp created ok?
        'We'l' be trying to create any of these objects
        arrObject = Split("WinHttp.WinHttpRequest.5|MSXML2.XMLHTTP|MSXML2.ServerXMLHTTP|Microsoft.XMLHTTP", "|")
        'If there were any errors, clear them
        Clear
        'If we allready have some items, clear the collection
        If Not m_colItems Is Nothing Then Set m_colItems = Nothing
        Set m_colItems = New Collection
    
        'Try to retrieve a feed with one of the above objects. If an error
        'occurs or parsing of a document fails, try the next object.
        T = 0
        Set oXMLDoc = fCreateObject("MSXML.DomDocument")
        If Not oXMLDoc Is Nothing Then  'Could object be created?
            While (Len(oXMLDoc.XML) = 0) And (T <= UBound(arrObject))
                oXMLDoc.LoadXML ReadHTTP(sURL, sUserName, sPassWord, arrObject(T))
                T = T + 1
            Wend
            'We should have read a feed, OR all objects failed to do their thing...
            If Len(oXMLDoc.XML) > 0 Then bParseOK = ParseDoc(oXMLDoc, lType, bAllowHTML)    'Try to parse the document
            Set oXMLDoc = Nothing
        End If
    End If
    'And return a boolean. True is all went ok, False if the last status <> 200, or parsing failed.
    ReadFeed = (m_lLastHTTPStatus = 200) And (bParseOK)
    Exit Function

errReadFeed:
    'Tell user something went wrong
    LogError Err.Number, Err.Description
End Function

'Main function to get the Feed from a HTTP server
Private Function ReadHTTP(sURL As String, sUserName As String, sPassWord As String, sObj As String) As String
    Dim oHTTP As Object
    
    On Error GoTo errReadHTTP
    Set oHTTP = fCreateObject(sObj) 'Try to create the desired object
    If Not oHTTP Is Nothing Then
        'The WinHTTPRequest object function a tiny bit differently from the rest...
        If StrComp(sObj, "WinHTTP.WinHTTPRequest.5", vbTextCompare) = 0 Then
            oHTTP.Open "GET", sURL, False
            'Set proxy settings and credentials for authorising if required
            If (Len(m_strProxyUser) > 0) Or (Len(m_strProxyPass) > 0) Then oHTTP.SetCredentials m_strProxyUser, m_strProxyPass, HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
            If (Len(sUserName) > 0) Or (Len(sPassWord) > 0) Then oHTTP.SetCredentials sUserName, sPassWord, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
        Else
            'The user should run proxycfg when using a proxy server. It SHOULD use IE settings, but if
            'this fails, try running proxycfg
            oHTTP.Open "GET", sURL, False, sUserName, sPassWord
        End If
        'Send the HTTP request
        oHTTP.Send
        
        'Store status
        m_sLastHTTPStatusText = oHTTP.StatusText
        m_lLastHTTPStatus = oHTTP.Status
        
        'Return the XML text (minus some crap)
        ReadHTTP = StripCrap(oHTTP.ResponseText)
    End If
    Set oHTTP = Nothing
    Exit Function

errReadHTTP:
    'Tell user something went wrong
    LogError Err.Number, Err.Description
End Function

'Main function to parse a feed and populate the items collection
Private Function ParseDoc(ByRef oXMLDoc As Object, lType As eFeedType, bAllowHTML As Boolean) As Boolean
    Dim sNodes As String
    Dim sTitle As String
    Dim sDesc As String
    Dim sURL As String
    Dim oNodes As Object
    Dim oNode As Object
    Dim oItem As FeedItem
    
    'Each type of feed has a different structure
    Const cRSSItems = "/rss/channel/item"
    Const cRDFItems = "/rdf:RDF/item"
    Const cAtomItems = "/feed/entry"
    
    On Error GoTo errParseDoc
    
    'Try to autodetect the type of feed by selecting a root node
    If lType = AutoDetect Then
        If Not oXMLDoc.SelectSingleNode(cRSSItems) Is Nothing Then
            lType = RSS
        ElseIf Not oXMLDoc.SelectSingleNode(cRDFItems) Is Nothing Then
            lType = RDF
        ElseIf Not oXMLDoc.SelectSingleNode(cAtomItems) Is Nothing Then
            lType = Atom
        End If
    End If
    
    'Store the feedtype
    m_lFeedType = lType
    
    'Depending on the type of feed we'll be needing different nodes from the document...
    Select Case lType
        Case RSS
            sNodes = cRSSItems
            sTitle = "title"
            sDesc = "description"
            sURL = "link"
        Case RDF
            sNodes = cRDFItems
            sTitle = "title"
            sDesc = "description"
            sURL = "link"
        Case Atom
            sNodes = cAtomItems
            sTitle = "title"
            sDesc = "content"
            sURL = "link[@rel='alternate']/@href"
        Case Else
            ParseDoc = False
            Exit Function
    End Select
    
    'Select all items from the feed
    Set oNodes = oXMLDoc.SelectNodes(sNodes)
    'Loop through all items in the feed and populate our collection
    For Each oNode In oNodes
        Set oItem = New FeedItem
        oItem.URL = GetNodeValue(oNode, sURL)
        oItem.Title = GetNodeValue(oNode, sTitle)
        If bAllowHTML Then  'Allow HTML in descriptions?
            oItem.Description = GetNodeValue(oNode, sDesc)
        Else
            oItem.Description = GetNodeValue(oNode, StripHTML(sDesc))
        End If
        m_colItems.Add oItem
    Next
    Set oNodes = Nothing
    ParseDoc = True 'Parse was ok...
    Exit Function

errParseDoc:
    'Tell user something went wrong
    LogError Err.Number, Err.Description
    ParseDoc = False    'We failed to parse the document
End Function

'Status Properties
Public Property Get LastHTTPStatus() As Long
    LastHTTPStatus = m_lLastHTTPStatus
End Property

Public Property Get LastHTTPStatusText() As String
    LastHTTPStatusText = m_sLastHTTPStatusText
End Property

Public Property Get LastErrorCode() As Long
    LastErrorCode = m_lLastErrorCode
End Property

Public Property Get LastErrorDesc() As String
    LastErrorDesc = m_sLastErrorDesc
End Property

'Proxy server properties
Public Property Get ProxyUser() As String
    ProxyUser = m_strProxyUser
End Property

Public Property Get ProxyPass() As String
    ProxyPass = m_strProxyPass
End Property

Public Property Let ProxyUser(sUserName As String)
    m_strProxyUser = sUserName
End Property

Public Property Let ProxyPass(sPassWord As String)
    m_strProxyPass = sPassWord
End Property

'Feed properties
Public Property Get FeedType() As eFeedType
    FeedType = m_lFeedType
End Property

Public Property Get FeedTypeString() As String
    Dim sTMP As String
    
    Select Case m_lFeedType
        Case eFeedType.RSS:     sTMP = "RSS"
        Case eFeedType.RDF:     sTMP = "RDF"
        Case eFeedType.Atom:    sTMP = "Atom"
        Case Else:
            sTMP = "Unknown / None"
    End Select
    FeedTypeString = sTMP
End Property

'The collection of Feed items
Public Property Get FeedItems() As Collection
    Set FeedItems = m_colItems
End Property

'Return the value of a Node as a String. If the node does not exist, return an empty string
Private Function GetNodeValue(ByRef oNode As Object, ByVal sName As String) As String
    On Error GoTo errGetNodeValue
    If oNode.SelectSingleNode(sName) Is Nothing Then
        GetNodeValue = ""
    Else
        GetNodeValue = oNode.SelectSingleNode(sName).NodeTypedValue
    End If
    Exit Function
    
errGetNodeValue:
    'Tell user something went wrong
    LogError Err.Number, Err.Description
End Function

'Try to create an object and handle any errors if any (do NOT cause a run-time error)
Private Function fCreateObject(sObject As String) As Object
    On Error GoTo errCreateObject
    Set fCreateObject = CreateObject(sObject)
    Exit Function
    
errCreateObject:
    'Tell user something went wrong
    LogError 429, "Error creating object '" & sObject & "'"
    Set fCreateObject = Nothing
End Function

'Remove DOCTYPE declaration and non-ASCII chars from document
Private Function StripCrap(sXML As String) As String
    oRE.Pattern = "<!DOCTYPE(.|\n)+?>|[^\x20-\xFF]"
    StripCrap = oRE.Replace(sXML, "")
End Function

'Strip HTML tags from a string
Public Function StripHTML(sString As String) As String
    oRE.Pattern = "<(.|\n)+?>"
    StripHTML = oRE.Replace(sString, "")
End Function

'Logs an error
Private Sub LogError(lErrCode As Long, sErrDesc As String)
    'Store last error code and description
    m_lLastErrorCode = lErrCode
    m_sLastErrorDesc = sErrDesc
    
    'The messagebox should be replaced by some code to log to a file or eventlog. This is just for debugging
    MsgBox "Error: " & lErrCode & vbCrLf & vbCrLf & sErrDesc, vbCritical + vbApplicationModal + vbOKOnly
End Sub

'Replace %1, %2, ..., %n in a string by its corrosponding parameter from the array
Private Function ReplParms(ByVal sString As String, ParamArray sParams()) As String
    Dim T As Long
    
    For T = 0 To UBound(sParams)
        sString = Replace(sString, "%" & T + 1, sParams(T))
    Next
    ReplParms = sString
End Function

'Clear all error and status variables and reset feed type
Public Sub Clear()
    m_lLastHTTPStatus = 0
    m_sLastHTTPStatusText = ""
    m_lLastErrorCode = 0
    m_sLastErrorDesc = ""
    m_lFeedType = 0
End Sub
