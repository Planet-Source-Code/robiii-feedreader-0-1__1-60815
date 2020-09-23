VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeedItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_sURL As String
Private m_sTitle As String
Private m_sDesc As String

Public Property Get URL() As String
    URL = m_sURL
End Property

Public Property Get Title() As String
    Title = m_sTitle
End Property

Public Property Get Description() As String
    Description = m_sDesc
End Property

Friend Property Let URL(sURL As String)
    m_sURL = sURL
End Property

Friend Property Let Title(sTitle As String)
    m_sTitle = sTitle
End Property

Friend Property Let Description(sDesc As String)
    m_sDesc = sDesc
End Property

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

Private Sub Clear()
    'Set variables to initial values...
    m_sURL = ""
    m_sTitle = ""
    m_sDesc = ""
End Sub
