Attribute VB_Name = "modMenu"
'The only purpose of this module is to let you set your menus
'to have radio-buttons instead of check-buttons.
'By Paul Berlin 2003
Option Explicit

Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenu Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const MFT_RADIOCHECK = &H200
Private Const MFT_STRING = &H0

Public Sub SetMenuRadio(ByVal hMenu As Long, ByVal index As Long)
  'This sub sets the specified menu and menu item index to a radio check-box
  Dim mii As MENUITEMINFO

  With mii
    .cbSize = Len(mii)
    .fMask = MIIM_TYPE
    .fType = MFT_RADIOCHECK Or MFT_STRING
    .dwTypeData = Space(256)
    .cch = 256
  End With
  
  GetMenuItemInfo hMenu, index, 1, mii
  
  With mii
    .fType = .fType Or MFT_RADIOCHECK
  End With
  
  SetMenuItemInfo hMenu, index, 1, mii

End Sub

Public Function GetSubMenuHandle(ByVal hMenu As Long, ByVal index As Long) As Long
  'This sub returns the submenu handle of the specified menu and menu item index, if any
  Dim mii As MENUITEMINFO

  With mii
    .cbSize = Len(mii)
    .fMask = MIIM_SUBMENU
    
    Call GetMenuItemInfo(hMenu, index, 1, mii)
    GetSubMenuHandle = .hSubMenu
    
  End With
  
End Function

Public Function GetMenuHandle(ByVal hwnd As Long) As Long
  GetMenuHandle = GetMenu(hwnd)
End Function
