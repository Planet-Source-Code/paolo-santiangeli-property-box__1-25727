VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPropertyBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const LVM_FIRST As Long = &H1000
'Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
'Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
'Private Const LVHT_NOWHERE As Long = &H1
'Private Const LVHT_ONITEMICON As Long = &H2
'Private Const LVHT_ONITEMLABEL As Long = &H4
'Private Const LVHT_ONITEMSTATEICON As Long = &H8


Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
'Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
'Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
'Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)


'Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
                                     LVHT_ONITEMLABEL Or _
                                   LVHT_ONITEMSTATEICON)

Private Const LVIR_BOUNDS = 0
Private Const LVIR_ICON = 1
Private Const LVIR_LABEL = 2
Private Const LVIR_SELECTBOUNDS = 3

'Private Type POINTAPI
'   X As Long
'   Y As Long
'End Type

Private Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' Private Type LVHITTESTINFO
'    pt As POINTAPI
'    flags As Long
'    iItem As Long
'    iSubItem  As Long  'ie3+ only .. was NOT in win95.
'                       'Valid only for LVM_SUBITEMHITTEST
'End Type

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Enum eDataType
       omString
       omDate
       omTime
       omBoolean
       omNumber
End Enum

Private mCurrentItem As Long
Private mDirty As Boolean

Public Event cmbClick(pItem As ListItem)

Private WithEvents mLV As MSComctlLib.ListView
Attribute mLV.VB_VarHelpID = -1
Private WithEvents mTXT As VB.TextBox
Attribute mTXT.VB_VarHelpID = -1
Private WithEvents mList As VB.ListBox
Attribute mList.VB_VarHelpID = -1
Private WithEvents mCmd As VB.PictureBox
Attribute mCmd.VB_VarHelpID = -1
Public Sub Init(pLV As MSComctlLib.ListView)
Dim state As Long

    Set mLV = pLV
    Set mLV.Parent.Font = mLV.Font
    
    Set mCmd = mLV.Parent.Controls.Add("VB.PictureBox", "cmdCombo")
        mCmd.ZOrder 0
        mCmd.Appearance = vbFlat
        mCmd.BorderStyle = 0
        mCmd.BackColor = RGB(210, 210, 210)
        mCmd.Picture = LoadPicture(App.Path & "\cmd.ico")
    
    Set mList = mLV.Parent.Controls.Add("VB.ListBox", "cmdList")
        mList.ZOrder 0
        mList.Appearance = vbFlat
        mList.Height = 75
        
    Set mTXT = mLV.Parent.Controls.Add("VB.TextBox", "txtValue")
        
        mTXT.ZOrder 0
        mTXT.Appearance = vbFlat
        mTXT.BorderStyle = 0
        
        
    Set mTXT.Font = mLV.Font
    Set mList.Font = mLV.Font
            
        
        
End Sub
Sub Display(pIndex As Long, pSubitem As Long)

Const lvBorder = 3
Const lvGrid = 1
Dim pMode As String
Dim pFont As New StdFont

Dim pRect As Rect
Dim APIItemIndex As Long
    
    If mList.Visible = True Then mList.Visible = False
    
    APIItemIndex = pIndex - 1
    
    pRect.Top = pSubitem
    pRect.Left = LVIR_LABEL
    
    a = SendMessage(mLV.hWnd, LVM_GETSUBITEMRECT, APIItemIndex, pRect)
    
    pMode = "cmb"
    
    Select Case pMode
        
        Case "txt"
                
                mCmd.Visible = False
            
            With mTXT
                
                .Visible = True
                .ZOrder 0
                
                .BackColor = RGB(230, 230, 230)
                
                .Top = mLV.Top + pRect.Top + lvBorder + lvGrid
                .Left = mLV.Left + pRect.Left + lvBorder + lvGrid
                
                .Height = (pRect.Bottom) - (pRect.Top) - lvGrid
                .Width = (pRect.Right - pRect.Left) - lvGrid
        
            
                .Text = mLV.ListItems(pIndex).SubItems(pSubitem)
        
            End With
        
            
        Case "cmb"
                
                mCmd.Visible = True
            
                With mTXT
                    
                    .ZOrder 0
                    .Visible = True
                    .BackColor = RGB(230, 230, 230)
    
                    .Top = mLV.Top + pRect.Top + lvBorder + lvGrid
                    .Left = mLV.Left + pRect.Left + lvBorder + lvGrid
                    
                    .Height = (pRect.Bottom) - (pRect.Top) - lvGrid
                    .Width = (pRect.Right - pRect.Left) - lvGrid - .Height
            
                    Set .Font = mLV.Font
                    Set mLV.Parent.Font = mLV.Font
                
                    .Text = mLV.ListItems(pIndex).SubItems(pSubitem)
            
                End With
                
                mCmd.ZOrder 0
                mCmd.Height = mTXT.Height
                mCmd.Width = mTXT.Height
                mCmd.Top = mTXT.Top
                mCmd.Left = mTXT.Left + mTXT.Width
        
    End Select
    
    
End Sub
Private Sub Class_Terminate()
    Set mTXT = Nothing
    Set mCmd = Nothing
    Set mList = Nothing
End Sub

Private Sub mCmd_Click()
Dim tmp As Variant
    
    If mList.Visible = True Then
       
       mList.Visible = False
    
    Else
    
        mList.Top = mTXT.Top + mTXT.Height
        mList.Width = mTXT.Width + mCmd.Width
        mList.Left = mTXT.Left
        
        tmp = mLV.ListItems(mCurrentItem).Tag
    
        mList.Clear
        
            For i = 0 To UBound(tmp)
                mList.AddItem tmp(i)
                DoEvents
            Next
        
        mList.Visible = True
    
    End If
    
End Sub

Private Sub mList_Click()

    mTXT = mList.Text

End Sub

Private Sub mList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mList.Visible = False

End Sub
Private Sub mLV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    mTXT.Visible = True
    mTXT.ZOrder 0
        
        If mDirty Then
            mLV.ListItems(mCurrentItem).SubItems(1) = mTXT.Text
        End If
    
    mCurrentItem = Item.Index
    Display mCurrentItem, 1

End Sub
Private Sub mLV_KeyPress(KeyAscii As Integer)
    
    mTXT.Text = Chr(KeyAscii)
    mTXT.SelStart = 1
    
    If mTXT.Visible Then
        
        mTXT.SetFocus
    
    Else
        
        mTXT.Visible = True
        mTXT.SetFocus
    
    End If

End Sub
Private Sub mTXT_Change()

mDirty = True

End Sub

Private Sub mTXT_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            
            mLV.ListItems.Item(mCurrentItem).SubItems(1) = mTXT.Text

    End If

End Sub
Public Sub Resize()
    
    mTXT.Visible = False
    mCmd.Visible = False
    
    mLV.Width = mLV.Parent.ScaleWidth - mLV.Left
    
    mLV.ColumnHeaders(2).Width = mLV.Parent.ScaleWidth - mLV.ColumnHeaders(1).Width - 50
    mLV.Height = mLV.Parent.Height

End Sub
Sub AddProperty(pText As String, pValue As String, ListValues As Variant, Optional pType As eDataType = omString)

Dim pItem As ListItem
Dim pTestWidth As Long
    
    'Calculate optimal width
    pTestWidth = getTextWidth(pText) + 30
    If pTestWidth > mLV.ColumnHeaders(1).Width Then
       mLV.ColumnHeaders(1).Width = pTestWidth + 30
    End If
    
    'Calculate optimal width
    pTestWidth = getTextWidth(pValue) + 40
    If pTestWidth > mLV.ColumnHeaders(2).Width Then
       mLV.ColumnHeaders(2).Width = pTestWidth + 40
    End If
    
    
    Set pItem = mLV.ListItems.Add(, , pText, 0, 1)
        pItem.SubItems(1) = pValue
        pItem.Tag = ListValues

    Set pItem = Nothing

    
End Sub
Private Function getTextWidth(pString As String) As Long

    getTextWidth = mLV.Parent.TextWidth(pString)

End Function

Private Sub mTXT_LostFocus()
            
            mLV.ListItems.Item(mCurrentItem).SubItems(1) = mTXT.Text

End Sub

