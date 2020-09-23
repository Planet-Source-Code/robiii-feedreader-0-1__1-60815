VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "FeedReader DEMO"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowBrowser 
      Appearance      =   0  'Flat
      Caption         =   "Show Browser"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3720
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser MSBrowser 
      CausesValidation=   0   'False
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   5318
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve!"
      Default         =   -1  'True
      Height          =   285
      Left            =   8640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "http://site.combat.cc/rssout.asp?type=rss&ln=en"
      Top             =   120
      Width           =   7455
   End
   Begin VB.ListBox lstItems 
      Appearance      =   0  'Flat
      Height          =   2955
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lbItemURL 
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   3600
      UseMnemonic     =   0   'False
      Width           =   6015
   End
   Begin VB.Label lbItemDesc 
      Height          =   1815
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      UseMnemonic     =   0   'False
      Width           =   6015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbItemTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      UseMnemonic     =   0   'False
      Width           =   6015
   End
   Begin VB.Label lbFeedURL 
      AutoSize        =   -1  'True
      Caption         =   "Feed URL:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   780
   End
   Begin VB.Label lbStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===================================================================================================
'The default sample (RSS)feed is:
'http://site.combat.cc/rssout.asp?type=rss&ln=en
'
'For sample ATOM or RDF feeds:
'http://site.combat.cc/rssout.asp?type=atom&ln=en
'http://site.combat.cc/rssout.asp?type=rdf&ln=en
'
'ALL our feeds from combat.cc are VALID feeds (at the time of writing). To validate our feeds go to:
'http://feedvalidator.org/check.cgi?url=http://site.combat.cc/rssout.asp?type=rss
'http://feedvalidator.org/check.cgi?url=http://site.combat.cc/rssout.asp?type=rdf
'http://feedvalidator.org/check.cgi?url=http://site.combat.cc/rssout.asp?type=atom
'===================================================================================================

'The FeedReader object
Private oFR As CombatFeedReader.FeedReader

Private Sub Form_Load()
    'Create a feedreader object
    Set oFR = New FeedReader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Free the FeedReader Object
    Set oFR = Nothing
End Sub

Private Sub cmdRetrieve_Click()
    Dim oItem As CombatFeedReader.FeedItem
    
    Me.MousePointer = vbHourglass
    
    'Clear all previous items
    lstItems.Clear
    lbItemTitle.Caption = ""
    lbItemDesc.Caption = ""
    lbItemURL.Caption = ""
    MSBrowser.Visible = False
    
    'Get a sample RSS feed. NOTE: You can pass a username and password for feeds that you need to authenticate for.
    oFR.ReadFeed txtURL.Text
    
    'Show the status
    lbStatus.Caption = "Status: " & oFR.FeedItems.Count & " items in the feed. Feed type is: " & oFR.FeedTypeString
    
    'Populate the listbox
    For Each oItem In oFR.FeedItems
        lstItems.AddItem oItem.Title
    Next
    
    Me.MousePointer = vbNormal
End Sub

Private Sub lstItems_Click()
    Dim oItem As CombatFeedReader.FeedItem
    
    'Show the item properties in the appropriate labels
    Set oItem = oFR.FeedItems.Item(lstItems.ListIndex + 1)
    lbItemTitle.Caption = oItem.Title
    lbItemDesc.Caption = oItem.Description
    lbItemURL.Caption = oItem.URL
    
    'Open the URL in the browser pane?
    If chkShowBrowser.Value = vbChecked Then
        MSBrowser.Visible = True
        MSBrowser.Navigate2 oItem.URL
    End If
    
    Set oItem = Nothing
End Sub

'Misc. un-interesting code
Private Sub chkShowBrowser_Click()
    If chkShowBrowser.Value <> vbChecked Then
        MSBrowser.Visible = False
        MSBrowser.Navigate2 "about:blank"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MSBrowser.Move 120, 3960, Me.ScaleWidth - 240, Me.ScaleHeight - 4080
    lbStatus.Width = Me.ScaleWidth - 240
    lbItemTitle.Width = Me.ScaleWidth - 3840
    lbItemDesc.Width = Me.ScaleWidth - 3840
    lbItemURL.Width = Me.ScaleWidth - 3840
    txtURL.Width = Me.ScaleWidth - 2400
    cmdRetrieve.Left = Me.ScaleWidth - 1200
End Sub

Private Sub MSBrowser_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub MSBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Cancel = True
End Sub
