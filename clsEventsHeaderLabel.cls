VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsHeaderLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Events for the Header Label

Private m_modernListbox As clsModernListbox
Public WithEvents mylabel As MSForms.Label
Attribute mylabel.VB_VarHelpID = -1

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
End Property

Private Sub mylabel_Click()
    Call m_modernListbox.Sort(mylabel.name)
End Sub

Private Sub myLabel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_modernListbox.SetHeaderHoverHighlight(mylabel)
End Sub

