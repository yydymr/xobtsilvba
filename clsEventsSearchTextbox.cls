VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventsSearchTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM
' Description: Event for Search Text box.

Private m_modernListbox As clsModernListbox
Public WithEvents myTextbox As MSForms.TextBox
Attribute myTextbox.VB_VarHelpID = -1

Public Property Set parentListBox(ByVal argModernListBox As clsModernListbox)
    Set m_modernListbox = argModernListBox
End Property

Private Sub myTextbox_Change()
    Call m_modernListbox.FilterRows(Trim(myTextbox.Value))
End Sub

'This is to simply clear the hover from the listbox row when the
' mouse is over the search label.
Private Sub myTextbox_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call m_modernListbox.ClearHover
End Sub
