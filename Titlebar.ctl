VERSION 5.00
Begin VB.UserControl Titlebar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "Titlebar.ctx":0000
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ToolboxBitmap   =   "Titlebar.ctx":000F
   Begin VB.Label Caption1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Caption2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0083180A&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image TitleIcon 
      Height          =   315
      Left            =   840
      Picture         =   "Titlebar.ctx":0321
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   315
   End
   Begin VB.Image CloseButton 
      Height          =   315
      Left            =   1320
      Picture         =   "Titlebar.ctx":06AB
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image MaximizeButton 
      Height          =   315
      Left            =   960
      Picture         =   "Titlebar.ctx":0B29
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image MinimizeButton 
      Height          =   315
      Left            =   600
      Picture         =   "Titlebar.ctx":0F9D
      Top             =   1560
      Width           =   315
   End
   Begin VB.Image TitleLeft 
      Height          =   450
      Left            =   360
      Picture         =   "Titlebar.ctx":1410
      Top             =   2040
      Width           =   150
   End
   Begin VB.Image TitleRight 
      Height          =   450
      Left            =   1440
      Picture         =   "Titlebar.ctx":1807
      Top             =   2040
      Width           =   150
   End
   Begin VB.Image Left 
      Height          =   795
      Left            =   480
      Picture         =   "Titlebar.ctx":1C07
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   60
   End
   Begin VB.Image Right 
      Height          =   795
      Left            =   1440
      Picture         =   "Titlebar.ctx":1F35
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   60
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   600
      Picture         =   "Titlebar.ctx":2263
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   795
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   1440
      Picture         =   "Titlebar.ctx":2591
      Top             =   3480
      Width           =   60
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   480
      Picture         =   "Titlebar.ctx":28CB
      Top             =   3480
      Width           =   60
   End
   Begin VB.Image NormalButton 
      Height          =   315
      Left            =   3360
      Picture         =   "Titlebar.ctx":2C04
      Top             =   1680
      Width           =   315
   End
   Begin VB.Image Title 
      Height          =   450
      Left            =   600
      Picture         =   "Titlebar.ctx":30E3
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   795
   End
End
Attribute VB_Name = "Titlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' ### ### #####         ###### ### ###### ###    ######
'  #####  ## ###        ###### ### ###### ###    ###
'  #####  #####  ######   ##   ###   ##   ###### ######
' ### ### ###    ######   ##   ###   ##   ###### ######
'
'     Copyright © 2002 by Doug Sheffer
'
'     Distributed freely so long that this notice stays at the top
'
'     Please include authors name in your resulting application

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Dim bTransparent As Boolean

Dim PropBag As PropertyBag

Public Event Closed()
Public Event Normal()
Public Event Maximized()
Public Event Minimized()

Public Sub RePos()
    'This repositions the different controls on the form when it is resized
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    'Titlebar
    With TitleLeft
        .Left = 0
        .Top = 0
    End With
    
    With Title
        .Left = TitleLeft.Width
        .Top = 0
        .Width = X - TitleLeft.Width - TitleRight.Width
    End With
    
    With TitleRight
        .Left = Title.Left + Title.Width
        .Top = 0
    End With
    
    'Borders
    With BottomLeft
        .Left = 0
        .Top = Y - .Height
    End With
    
    With BottomRight
        .Left = X - .Width
        .Top = Y - .Height
    End With
    
    With Left
        .Left = 0
        .Top = TitleLeft.Top + TitleLeft.Height
        .Height = BottomLeft.Top - .Top
    End With
    
    With Right
        .Left = X - .Width
        .Top = TitleRight.Top + TitleRight.Height
        .Height = BottomRight.Top - .Top
    End With
    
    With Bottom
        .Left = BottomLeft.Width
        .Top = Y - Bottom.Height
        .Width = X - BottomLeft.Width - BottomRight.Width
    End With
    
    'Buttons
    With CloseButton
        .Left = Right.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    With MaximizeButton
        .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    With NormalButton
    .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
        
    End With
    
    
    
    With MinimizeButton
        .Left = MaximizeButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Icon
    With TitleIcon
        .Left = Left.Left + Left.Width + 2
        .Top = (Title.Height - .Height) / 2
    End With
    
    'Titlebar Caption
    With Caption1
        .Left = TitleIcon.Left + TitleIcon.Width + 3
        .Top = ((Title.Height - 13) / 2) - 1
        .Width = MinimizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        .Height = 13
    End With
    
    With Caption2
        .Left = TitleIcon.Left + TitleIcon.Width + 2
        .Top = ((Title.Height - 13) / 2) + 1
        .Width = MinimizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        .Height = 13
    End With
    
    'Checks if it should have transparent corners
    If bTransparent = True Then
        ReTrans
    End If
End Sub

Public Sub TransparentEdges()
    'This is used as a safe guard set when the application starts,
    'otherwise the control would set the corners transparent at design time
    bTransparent = True
    RePos
End Sub

Public Sub ReTrans()
    Dim Add As Long
    Dim Sum As Long
    
    Dim X As Single
    Dim Y As Single
    
    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small
    
    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels
    
    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn UserControl.ContainerHwnd, Sum, True   'Sets corners transparent
End Sub

Private Sub CloseButton_Click()
    RaiseEvent Closed   'User knows when users presses button
    
End Sub

Private Sub MaximizeButton_Click()
    RaiseEvent Maximized 'User knows when users presses button
    MaximizeButton.Visible = False
    NormalButton.Visible = True
End Sub

Private Sub MinimizeButton_Click()
    RaiseEvent Minimized    'User knows when users presses button
End Sub

Private Sub NormalButton_Click()
RaiseEvent Normal 'User knows when users presses button
MaximizeButton.Visible = True
    NormalButton.Visible = False
    
End Sub

Private Sub UserControl_Initialize()

    bTransparent = False    'So we do not set the corners transparent while still in design mode
    RePos   'Reposition
End Sub

Private Sub UserControl_Resize()
    RePos   'Reposition
End Sub

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub TitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Private Sub Caption2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Lets user move parent form
    Call ReleaseCapture: Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Property Get Caption() As String
    Caption = Caption1.Caption  'Retrieves caption
End Property

Public Property Let Caption(Caption As String)
    Caption1.Caption = Caption  'Sets caption
    Caption2.Caption = Caption  'Sets caption
    RePos   'Reposition
End Property

Public Property Get ShowMaximize() As Boolean
    ShowMaximize = MaximizeButton.Visible   'Returns whether it is visible
End Property

Public Property Let ShowMaximize(Visible As Boolean)
    MaximizeButton.Visible = Visible 'Sets whether it is visible
End Property
Public Property Get ShowNormal() As Boolean
    ShowMaximize = NormalButton.Visible   'Returns whether it is visible
End Property

Public Property Let ShowNormal(Visible As Boolean)
    NormalButton.Visible = Visible   'Sets whether it is visible
End Property
Public Property Get ShowMinimize() As Boolean
    ShowMinimize = MinimizeButton.Visible  'Returns whether it is visible
End Property

Public Property Let ShowMinimize(Visible As Boolean)
    MinimizeButton.Visible = Visible    'Sets whether it is visible
End Property

Public Property Get ShowClose() As Boolean
    ShowClose = CloseButton.Visible 'Returns whether it is visible
End Property

Public Property Let ShowClose(Visible As Boolean)
    CloseButton.Visible = Visible   'Sets whether it is visible
End Property

Public Sub Icon(IconToUse As Object)
    TitleIcon.Picture = IconToUse   'Sets icon
End Sub

Public Function DefaultBackgroundColor() As String
    DefaultBackgroundColor = &HEAF1F1   'Returns a common off-white Windows XP color
End Function
