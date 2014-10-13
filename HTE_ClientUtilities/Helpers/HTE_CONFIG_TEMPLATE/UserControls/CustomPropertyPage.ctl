VERSION 5.00
Begin VB.UserControl CustomPropertyPage 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3075
   ScaleWidth      =   5655
   ToolboxBitmap   =   "CustomPropertyPage.ctx":0000
End
Attribute VB_Name = "CustomPropertyPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements HTE_GPS.PropertyPage
Dim m_Callback As HTE_GPS.PropPageCallback
Dim bChanged As Boolean
Dim m_Settings As MSXML2.DOMDocument30

Private Property Get PropertyPage_Changed() As Boolean
    PropertyPage_Changed = bChanged
End Property

Private Sub PropertyPage_Exit()
    Set m_Callback = Nothing
End Sub

Private Property Get PropertyPage_LicenseKey() As String
    PropertyPage_LicenseKey = vbNullString
End Property

Private Property Get PropertyPage_Name() As String
    PropertyPage_Name = App.EXEName & "." & UserControl.Name
End Property

Private Property Let PropertyPage_PropertyCallback(RHS As HTE_GPS.PropPageCallback)
    Set m_Callback = RHS
End Property

Private Property Get PropertyPage_PropertyCallback() As HTE_GPS.PropPageCallback
     Set PropertyPage_PropertyCallback = m_Callback
End Property

Private Function PropertyPage_SaveChanges() As Boolean
Dim bReturn As Boolean
    If Not m_Callback Is Nothing Then
        If Not m_Settings Is Nothing Then
            bReturn = m_Callback.SaveChanges(m_Settings.xml)
            bChanged = Not bReturn
            PropertyPage_SaveChanges = bReturn
        End If
    End If
End Function

Private Property Let PropertyPage_Settings(ByVal RHS As String)
    loadLocalSettings (RHS)
    bChanged = False
End Property

Private Property Get PropertyPage_Settings() As String
    If Not m_Settings Is Nothing Then
        PropertyPage_Settings = m_Settings.xml
    End If
End Property

Private Function loadLocalSettings(ByVal sXML As String) As Boolean
    Set m_Settings = New MSXML2.DOMDocument30
    loadLocalSettings = m_Settings.loadXML(sXML)
    If loadLocalSettings Then
        '*TODO - ADD CUSTOM PROPERTY GET FUNCTIONS HERE
        '*TODO - WHEN A PROPERTY CHANGES YOU WILL NEED TO FLAG YOUR CHANGED VARIABLE
        '* Example: bChanged = bChanged Or letProperty(m_Settings, cInBufferSize, txtInBufferSize.Text)
    End If
End Function
