VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormClass 
   Caption         =   "Movie Database"
   ClientHeight    =   8055
   ClientLeft      =   15
   ClientTop       =   0
   ClientWidth     =   15090
   OleObjectBlob   =   "UserFormClass.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserFormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ExcelMacroMastery.com
' Author: Paul Kelly
' YouTube Video: https://youtu.be/QYW1SlKfKdM

' Declare the ListBox variable at the top of the UserForm code
Private WithEvents m_modernListbox As clsModernListbox
Attribute m_modernListbox.VB_VarHelpID = -1
Private m_displayData As Boolean

'控制最小化、恢复
Private m_InitialFormWidth As Single
Private m_InitialFormHeight As Single
Private m_InitialListWidth As Single
Private m_InitialListHeight As Single
Private m_InitialSearchWidth As Single
Private m_InitialQueryLeft As Single
Private m_InitialProcessLeft As Single
Private m_InitialProcessTop As Single



' LISTBOX EVENT
' Event when item select in the clsModernListbox
Private Sub m_modernListBox_ItemSelected(dataRow As Long)
   'MsgBox "Item selected:" & row
   
   If m_displayData = False Then Exit Sub
   
   Dim form As New formEdit
   Dim arr As Variant
   arr = m_modernListbox.GetRow(dataRow)
   Call form.Fill(arr)
   form.show
End Sub

' USERFORM EVENTS
Private Sub UserForm_Initialize()
    
    Set m_modernListbox = Instantiate_clsModernListbox
    Set m_modernListbox.parentFrame = FrameListBox

    OptionMulti.Value = True
    CheckBoxHover.Value = True
    checkboxScrollbars.Value = False
    checkboxSelectOn.Value = True
    textboxRecords.Value = 15
    checkboxAutoHeight.Value = True
    CheckBoxAutoWidth.Value = True
    
    With m_modernListbox
        
        .HoverOn = CheckBoxHover.Value
        
        .multiSelect = IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
        .columnWidths = "550;150;75;100;75;75;100"
        .ScrollBars = IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
        .recordsPerPage = textboxRecords.Value
        .AutomaticHeight = checkboxAutoHeight.Value
        .AutomaticWidth = CheckBoxAutoWidth.Value
        .HeaderFieldsFromString = "Title;Director;Year;Genre;IMDb;Duration;Budget;Box Office"
        Dim rg As Range: Set rg = shMovies.Range("A2:H31")
        .List = rg.Value

    
    End With
    
    ' Show controls
    Call ShowControlSection
    
    ' Set Userform to screen size
    Me.left = 0
    Me.Height = Application.Height
    Me.top = 0
    Me.width = Application.width
    
      Call SetFormStyle(Me)
    m_InitialFormWidth = Me.width
    m_InitialFormHeight = Me.Height

    ' 捕获 ListView 的初始尺寸
    m_InitialListWidth = m_modernListbox.width
    m_InitialListHeight = m_modernListbox.Height

    ' 捕获需要移动的控件的初始位置/尺寸
'    m_InitialSearchWidth = Me.txtSearch.width
'    m_InitialQueryLeft = Me.btnQuery.left
'    m_InitialProcessLeft = Me.btnProcess.left
'    m_InitialProcessTop = Me.btnProcess.top
'
'

End Sub

' HELPER
Private Sub ShowControlSection(Optional turnOn As Boolean = True)
    Dim c As Control
    For Each c In Controls
        If TypeName(c) = "Frame" And c.tag = "temp" Then
            c.Visible = turnOn
        End If
    Next c
End Sub

' USERFORM CONTROLS EVENTS

' Display all the selected record in a message box
Private Sub buttonSelectedData_Click()
    Dim data As Variant
    data = m_modernListbox.SelectedItems()
    If IsEmpty(data) Then
        MsgBox "No records selected"
    Else
        MsgBox arrayToString(data)
    End If
    
End Sub

' Display the first select record in a message box
Private Sub buttonSelectedOne_Click()
    Dim data As Variant
    data = m_modernListbox.SelectedItem()
    If IsEmpty(data) Then
        MsgBox "No record selected"
    Else
            MsgBox arrayRowToString(data, 1)
    End If

End Sub

Private Sub buttonUpdateHeight_Click()
    If Trim(TextBoxHeight.Value) = "" Then Exit Sub
    m_modernListbox.Height = Trim(TextBoxHeight.Value)
End Sub

Private Sub buttonUpdateWidth_Click()
    If Trim(textboxWidth.Value) = "" Then Exit Sub
    m_modernListbox.width = Trim(textboxWidth.Value)
End Sub

Private Sub buttonColumnWidths_Click()
    If Trim(TextBoxWidths.Value) = "" Then Exit Sub
    m_modernListbox.columnWidths = Trim(TextBoxWidths.Value)
End Sub

Private Sub buttonAddHeaders_Click()
    If Trim(textboxHeaders.Value) = "" Then Exit Sub
    m_modernListbox.HeaderFieldsFromString = Trim(textboxHeaders.Value)
End Sub

Private Sub checkboxAutoHeight_Click()
    m_modernListbox.AutomaticHeight = checkboxAutoHeight.Value
End Sub

Private Sub CheckBoxAutoWidth_Click()
    m_modernListbox.AutomaticWidth = CheckBoxAutoWidth.Value
End Sub

' When turned on the ListBox will display an UserForm with the record details when you click on a record.
' It fired the m_modernListBox_ItemSelected() above when a record is clicked
Private Sub checkboxDisplaySelect_Click()
    m_displayData = checkboxDisplaySelect.Value
End Sub

' Selects the record number specificed in the textbox.
' Note this is the row number of the data and not the row number on the screen
Private Sub textboxSelectItem_Change()
    If Len(Trim(textboxSelectItem.Value)) = 0 Then Exit Sub
    
    Call m_modernListbox.SetSelected(textboxSelectItem.Value, checkboxSelectOn.Value)
End Sub

' Turn on/off hover functionality
Private Sub CheckBoxHover_Click()
    m_modernListbox.HoverOn = CheckBoxHover.Value
End Sub

' Turn on/off scrollbars
Private Sub checkboxScrollbars_Click()
    m_modernListbox.ScrollBars = IIf(checkboxScrollbars.Value = True, fmScrollBarsBoth, fmScrollBarsNone)
End Sub

' set the number of records displayed on a page
Private Sub textboxRecords_Change()
    If Len(Trim(textboxRecords.Value)) = 0 Then Exit Sub
    m_modernListbox.recordsPerPage = Trim(textboxRecords.Value)
End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not m_modernListbox Is Nothing Then
        Call m_modernListbox.ClearHover
    End If
End Sub

Private Sub Frame1_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not m_modernListbox Is Nothing Then
        Call m_modernListbox.ClearHover
    End If
End Sub

' Allow Multiple selections
Private Sub OptionMulti_Click()
    m_modernListbox.multiSelect = IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
End Sub

Private Sub OptionSingle_Click()
    m_modernListbox.multiSelect = IIf(OptionSingle, fmMultiSelectSingle, fmMultiSelectExtended)
End Sub


' PROPERTIES OF THE USERFORM

' Returns the
Public Function getSelectedItems() As Variant
    getSelectedItems = m_modernListbox.SelectedItems
End Function

Public Function getSelectedItem() As Variant
    getSelectedItem = m_modernListbox.SelectedItem
End Function

' Prevents the UserForm from being unloaded when X is clicked.
' This is so we can access the selections after
Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    ' Hide the Userform and set cancelled to true
    Hide
    
End Sub


Private Sub UserForm_Resize()
    ' 当窗体大小发生变化时（包括最大化和从最大化恢复），此事件会触发

    Dim deltaWidth As Single
    Dim deltaHeight As Single

    ' 使用错误处理，防止在初始化完成前触发 Resize 事件导致错误
    On Error Resume Next
    ' 如果初始宽度未记录，则退出，避免出错
    If m_InitialFormWidth = 0 Then Exit Sub

    ' 1. 计算窗体宽度和高度的变化量
    deltaWidth = Me.width - m_InitialFormWidth
    deltaHeight = Me.Height - m_InitialFormHeight

    ' 2. 将变化量应用到需要调整的控件上

    ' --- 让 ListView 同时拉伸宽度和高度 ---
    m_modernListbox.width = m_InitialListWidth + deltaWidth
    m_modernListbox.Height = m_InitialListHeight + deltaHeight
End Sub
