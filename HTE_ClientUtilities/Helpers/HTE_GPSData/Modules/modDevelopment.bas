Attribute VB_Name = "modDevelopment"
Option Explicit

'But are located, obviously in different project folders....
Private m_bInDevelopment As Boolean
 
Public Function InDevelopment() As Boolean
   ' Debug.Assert code not run in an EXE.  Therefore m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function

Public Function GetParentPath() As String
Dim sPath As String
Dim hackPath() As String
Dim X As Long

    sPath = vbNullString
    hackPath = Split(App.Path, "\")
    For X = UBound(hackPath) - (1 + Abs(InDevelopment)) To 0 Step -1
        sPath = String$(Abs(X <> 0), "\") & hackPath(X) & sPath
    Next
    GetParentPath = sPath

End Function

