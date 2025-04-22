' UIFunctions.bas - UI Functions Module
Attribute VB_Name = "UIFunctions"
Option Explicit

#If VBA7 Then
    ' For 64-bit Office/VBA7
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function DeleteMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    ' For 32-bit Office
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
    Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

' Constants for system menu manipulation
Private Const MF_BYPOSITION = &H400&
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000
Private Const WS_MAXIMIZEBOX = &H10000

' Function to show or hide the Minimize button on a UserForm
Public Function ShowMinimizeButton(UF As Object, HideButton As Boolean) As Boolean
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim hMenu As LongPtr
#Else
    Dim hWnd As Long
    Dim hMenu As Long
#End If
    Dim WindowStyle As Long
    
    On Error GoTo ErrorHandler
    
    ' Get the window handle for the UserForm
    hWnd = FindWindow("ThunderDFrame", UF.Caption)
    If hWnd = 0 Then Exit Function
    
    ' Get the system menu for the window
    hMenu = GetSystemMenu(hWnd, 0)
    If hMenu = 0 Then Exit Function
    
    ' Hide the minimize button if requested
    If HideButton Then
        DeleteMenu hMenu, 6, MF_BYPOSITION ' 6 is the position of the Minimize menu item
    End If
    
    ShowMinimizeButton = True
    Exit Function
    
ErrorHandler:
    ShowMinimizeButton = False
End Function

' Function to show or hide the Maximize button on a UserForm
Public Function ShowMaximizeButton(UF As Object, HideButton As Boolean) As Boolean
#If VBA7 Then
    Dim hWnd As LongPtr
    Dim hMenu As LongPtr
#Else
    Dim hWnd As Long
    Dim hMenu As Long
#End If
    Dim WindowStyle As Long
    
    On Error GoTo ErrorHandler
    
    ' Get the window handle for the UserForm
    hWnd = FindWindow("ThunderDFrame", UF.Caption)
    If hWnd = 0 Then Exit Function
    
    ' Get the system menu for the window
    hMenu = GetSystemMenu(hWnd, 0)
    If hMenu = 0 Then Exit Function
    
    ' Hide the maximize button if requested
    If HideButton Then
        DeleteMenu hMenu, 5, MF_BYPOSITION ' 5 is the position of the Maximize menu item
    End If
    
    ShowMaximizeButton = True
    Exit Function
    
ErrorHandler:
    ShowMaximizeButton = False
End Function

' Function to make a UserForm resizable
Public Function MakeFormResizable(UF As Object, Sizable As Boolean) As Boolean
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    Dim WindowStyle As Long
    
    On Error GoTo ErrorHandler
    
    ' Get the window handle for the UserForm
    hWnd = FindWindow("ThunderDFrame", UF.Caption)
    If hWnd = 0 Then Exit Function
    
    ' Get the current window style
    WindowStyle = GetWindowLong(hWnd, GWL_STYLE)
    
    ' Add or remove the necessary styles to make the form resizable
    If Sizable Then
        WindowStyle = WindowStyle Or WS_THICKFRAME Or WS_MAXIMIZEBOX
    Else
        WindowStyle = WindowStyle And Not (WS_THICKFRAME Or WS_MAXIMIZEBOX)
    End If
    
    ' Set the new window style
    SetWindowLong hWnd, GWL_STYLE, WindowStyle
    
    ' Redraw the window with the new style
    DrawMenuBar hWnd
    
    MakeFormResizable = True
    Exit Function
    
ErrorHandler:
    MakeFormResizable = False
End Function

' Function to check if a UserForm is currently resizable
Public Function IsFormResizable(UF As Object) As Boolean
#If VBA7 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
    Dim WindowStyle As Long
    
    On Error GoTo ErrorHandler
    
    ' Get the window handle for the UserForm
    hWnd = FindWindow("ThunderDFrame", UF.Caption)
    If hWnd = 0 Then Exit Function
    
    ' Get the current window style
    WindowStyle = GetWindowLong(hWnd, GWL_STYLE)
    
    ' Check if the resizable style is set
    IsFormResizable = ((WindowStyle And WS_THICKFRAME) = WS_THICKFRAME)
    Exit Function
    
ErrorHandler:
    IsFormResizable = False
End Function 