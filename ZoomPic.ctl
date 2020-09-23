VERSION 5.00
Begin VB.UserControl ZoomPicCtl 
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ToolboxBitmap   =   "ZoomPic.ctx":0000
   Begin VB.CommandButton cmdZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   435
      Picture         =   "ZoomPic.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   75
      Width           =   375
   End
   Begin VB.CommandButton cmdZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   60
      Picture         =   "ZoomPic.ctx":049C
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Width           =   375
   End
   Begin VB.Timer timToolbar 
      Interval        =   100
      Left            =   1125
      Top             =   4560
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Height          =   4605
      Left            =   375
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   195
      Width           =   6285
      Begin VB.PictureBox picTemp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5970
         ScaleHeight     =   375
         ScaleWidth      =   345
         TabIndex        =   3
         Top             =   4290
         Width           =   345
      End
      Begin VB.HScrollBar hscPicture 
         Height          =   255
         LargeChange     =   20
         Left            =   0
         SmallChange     =   5
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4290
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.VScrollBar vscPicture 
         Height          =   4305
         LargeChange     =   20
         Left            =   5970
         SmallChange     =   5
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1440
         Left            =   0
         ScaleHeight     =   96
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   96
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6675
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   4785
      Width           =   360
   End
End
Attribute VB_Name = "ZoomPicCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' I have indeed learnt alot from this great site
' Most of this code can be traced to this site.
' I am really gratefull to all members of PSC.

Option Explicit

Private Const SM_CXBORDER               As Long = 5                 '*' Width of non-sizable borders
Private Const SM_CYBORDER               As Long = 6                 '*' Height of non-sizable borders
Private Const SM_CXDLGFRAME             As Long = 7                 '*' Width of dialog box borders
Private Const SM_CYDLGFRAME             As Long = 8                 '*' Height of dialog box borders
Private Const SM_CYVTHUMB               As Long = 9                 '*' Height of scroll box on vertical scroll bar
Private Const SM_CXHTHUMB               As Long = 10                '*' Width of scroll box on horizontal scroll bar

Private Type OriginalImage
    Height                              As Long
    Width                               As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

Private Type ScrollPositions
    HorizontalScrollMax                 As Long
    HorizontalScrollPosition            As Long
    VerticalScrollMax                   As Long
    VerticalScrollPosition              As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Declare Function GetCursorPos Lib "user32" ( _
                          lpPoint As POINTAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" ( _
                          ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" ( _
                          ByVal hwnd As Long, _
                          lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" ( _
                          ByVal hwndLock As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" ( _
                          ByVal xPoint As Long, _
                          ByVal yPoint As Long) As Long

Private m_dblPercentage                 As Double
Private m_strFileName                   As String
Private m_udtOriginal                   As OriginalImage
Private m_bolAllowOut                   As Boolean
Private m_bolAllowIn                    As Boolean
Private m_bolUseQuickBar                As Boolean
Private m_stpLastPosition               As ScrollPositions
Private m_bolInDrag                     As Boolean
Private m_XDrag                         As Long
Private m_YDrag                         As Long

Event Click()
Event DblClick()
Event KeyDown( _
      KeyCode As Integer, _
      Shift As Integer)
Event KeyPress( _
      KeyAscii As Integer)
Event KeyUp( _
      KeyCode As Integer, _
      Shift As Integer)
Event MouseDown( _
      Button As Integer, _
      Shift As Integer, _
      X As Single, _
      Y As Single)
Event MouseMove( _
      Button As Integer, _
      Shift As Integer, _
      X As Single, _
      Y As Single)
Event MouseUp( _
      Button As Integer, _
      Shift As Integer, _
      X As Single, _
      Y As Single)
Event Paint()
Event Resize()
Event Scroll()
Event ZoomChanged( _
      ByVal ZoomPercent As Long)
Event ZoomInClick()
Event ZoomOutClick()

Public Property Get AllowZoomIn() As Boolean

    AllowZoomIn = m_bolAllowIn

End Property

Public Property Let AllowZoomIn(Value As Boolean)

    m_bolAllowIn = Value
    PropertyChanged "AllowZoomIn"

End Property

Public Property Let AllowZoomOut(Value As Boolean)

    m_bolAllowOut = Value
    PropertyChanged "AllowZoomOut"

End Property

Public Property Get AllowZoomOut() As Boolean

    AllowZoomOut = m_bolAllowOut

End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)

    picView.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"

End Property

Public Property Get AutoRedraw() As Boolean

    AutoRedraw = picView.AutoRedraw

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = picView.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    picBuffer.BackColor = New_BackColor
    picMain.BackColor = New_BackColor
    picTemp.BackColor = New_BackColor
    picView.BackColor = New_BackColor
    PropertyChanged "BackColor"

End Property

Private Sub CheckForScrolls()

    On Error Resume Next

        If (picView.Width < picMain.Width - GetSystemMetrics(SM_CXHTHUMB)) Then
            hscPicture.Value = hscPicture.Min
            hscPicture.Visible = False
            picView.Left = (picMain.Width - picView.Width) / 2
          Else
            With hscPicture
                .Visible = True
                .ZOrder
                .Min = 0
                .Max = -(picView.Width - (picMain.Width - GetSystemMetrics(SM_CXHTHUMB)) + 4)
                .Value = (.Max - .Min) / 2
            End With
        End If

        If (picView.Height < picMain.Height - GetSystemMetrics(SM_CYVTHUMB)) Then
            vscPicture.Value = vscPicture.Min
            vscPicture.Visible = False
            picView.Top = (picMain.Height - picView.Height) / 2
          Else
            With vscPicture
                .Visible = True
                .ZOrder
                .Min = 0
                .Max = -(picView.Height - (picMain.Height - GetSystemMetrics(SM_CYVTHUMB)) + 4)
                .Value = (.Max - .Min) / 2
            End With
        End If

        picTemp.Visible = (hscPicture.Visible And vscPicture.Visible)
        picTemp.ZOrder

        picTemp.Move hscPicture.Width, vscPicture.Height, vscPicture.Width, hscPicture.Height

        With hscPicture

            .Left = 0

            If vscPicture.Visible Then
                .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
                .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - _
                         (2 * GetSystemMetrics(SM_CXBORDER))
              Else
                .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
                .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - GetSystemMetrics(SM_CXDLGFRAME)
            End If
        End With

        With vscPicture

            .Top = 0

            If hscPicture.Visible Then
                .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
                .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - _
                          (2 * GetSystemMetrics(SM_CYBORDER))
              Else
                .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
                .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - GetSystemMetrics(SM_CYDLGFRAME)
            End If
        End With

    On Error GoTo 0

End Sub

Public Sub Cls()

    picView.Cls

End Sub

Private Sub cmdZoom_Click(Index As Integer)

    On Error Resume Next

        If Index = 0 Then
            RaiseEvent ZoomInClick
          Else
            RaiseEvent ZoomOutClick
        End If
        picView.SetFocus

End Sub

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property

Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property

Public Property Get HasDC() As Boolean

    HasDC = picView.HasDC

End Property

Public Property Get hDC() As Long

    hDC = picView.hDC

End Property

Private Sub hscPicture_Change()

    On Error Resume Next

        picView.Left = hscPicture.Value

    On Error GoTo 0

End Sub

Private Sub hscPicture_Scroll()

    On Error Resume Next

        RaiseEvent Scroll

    On Error GoTo 0

End Sub

Public Property Get hwnd() As Long

    hwnd = picView.hwnd

End Property

Public Property Get Image() As Picture

    Set Image = picView.Image

End Property

Public Sub LoadImage(FilePath As String)

    On Error Resume Next

        If FilePath = vbNullString Then
            Exit Sub
        End If

        m_strFileName = FilePath

        picBuffer.Picture = LoadPicture(FilePath)

        Call ShowPicture

        m_udtOriginal.Height = picView.Height
        m_udtOriginal.Width = picView.Width

        Zoom = 100
    On Error GoTo 0

End Sub

Public Property Get MousePointer() As Integer

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)

    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property

Public Sub PaintPicture( _
                        ByVal Picture As Picture, _
                        ByVal X1 As Single, _
                        ByVal Y1 As Single, _
                        Optional ByVal Width1 As Variant, _
                        Optional ByVal Height1 As Variant, _
                        Optional ByVal X2 As Variant, _
                        Optional ByVal Y2 As Variant, _
                        Optional ByVal Width2 As Variant, _
                        Optional ByVal Height2 As Variant, _
                        Optional ByVal Opcode As Variant)

    On Error Resume Next

        picView.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode

    On Error GoTo 0

End Sub

Private Sub picMain_DblClick()

    On Error Resume Next

        RaiseEvent DblClick

    On Error GoTo 0

End Sub

Private Sub picMain_Resize()

    On Error Resume Next

        With hscPicture

            .Left = 0

            If vscPicture.Visible Then
                .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
                .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - _
                         (2 * GetSystemMetrics(SM_CXBORDER))
              Else
                .Top = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - GetSystemMetrics(SM_CYDLGFRAME)
                .Width = ScaleX(UserControl.Width, vbTwips, vbPixels) - GetSystemMetrics(SM_CXDLGFRAME)
            End If
        End With

        With vscPicture

            .Top = 0

            If hscPicture.Visible Then
                .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
                .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - hscPicture.Height - _
                          (2 * GetSystemMetrics(SM_CYBORDER))
              Else
                .Left = ScaleX(UserControl.Width, vbTwips, vbPixels) - vscPicture.Width - GetSystemMetrics(SM_CXDLGFRAME)
                .Height = ScaleY(UserControl.Height, vbTwips, vbPixels) - GetSystemMetrics(SM_CYDLGFRAME)
            End If
        End With

        picTemp.Move hscPicture.Width, vscPicture.Height, vscPicture.Width, hscPicture.Height

        CheckForScrolls

    On Error GoTo 0

End Sub

Public Property Get Picture() As StdPicture

    Set Picture = picBuffer.Picture

End Property

Public Property Set Picture(Value As StdPicture)

    Set picBuffer.Picture = Value

    Call ShowPicture

    m_udtOriginal.Height = picView.Height
    m_udtOriginal.Width = picView.Width

    Zoom = 100

End Property

Private Sub picView_Click()

    RaiseEvent Click

End Sub

Private Sub picView_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub picView_KeyDown( _
                            KeyCode As Integer, _
                            Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub picView_KeyPress( _
                             KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub picView_KeyUp( _
                          KeyCode As Integer, _
                          Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub picView_MouseDown( _
                              Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    On Error Resume Next

        RaiseEvent MouseDown(Button, Shift, X, Y)

    On Error GoTo 0

End Sub

Private Sub picView_MouseMove( _
                              Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

  Dim lngCurrentX                         As Long
  Dim lngCurrentY                         As Long
  Dim rctCurrentMain                      As RECT
  Dim pntCursor                           As POINTAPI

    On Error Resume Next

        If m_bolInDrag Then

            Call GetWindowRect(picMain.hwnd, rctCurrentMain)
            Call GetCursorPos(pntCursor)

            lngCurrentX = pntCursor.X - rctCurrentMain.Left
            lngCurrentY = pntCursor.Y - rctCurrentMain.Top

            If hscPicture.Visible = True Then
                If lngCurrentX < m_XDrag Then
                    hscPicture.Value = hscPicture.Value + (Abs(m_XDrag - lngCurrentX))
                  ElseIf lngCurrentX > m_XDrag Then
                    hscPicture.Value = hscPicture.Value - (Abs(lngCurrentX - m_XDrag))
                End If
            End If

            If vscPicture.Visible = True Then
                If lngCurrentY < m_YDrag Then
                    vscPicture.Value = vscPicture.Value + (Abs(m_YDrag - lngCurrentY))
                  ElseIf lngCurrentY > m_YDrag Then
                    vscPicture.Value = vscPicture.Value - (Abs(lngCurrentY - m_YDrag))
                End If
            End If

            m_YDrag = lngCurrentY
            m_XDrag = lngCurrentX

        End If

    On Error GoTo 0

End Sub

Private Sub picView_MouseUp( _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    On Error Resume Next

        RaiseEvent MouseUp(Button, Shift, X, Y)

        m_bolInDrag = False

    On Error GoTo 0

End Sub

Private Sub picView_Paint()

    RaiseEvent Paint

End Sub

Public Sub RecallPosition()

    If hscPicture.Max = m_stpLastPosition.HorizontalScrollMax And _
       vscPicture.Max = m_stpLastPosition.VerticalScrollMax Then

        hscPicture.Value = m_stpLastPosition.HorizontalScrollPosition
        vscPicture.Value = m_stpLastPosition.VerticalScrollPosition

    End If

End Sub

Private Sub SetZoom()

    On Error Resume Next

        If m_strFileName = vbNullString Or picBuffer.Picture.Handle = 0 Then
            Exit Sub
        End If

        LockWindowUpdate picView.hwnd

        picView.Width = m_udtOriginal.Width * (m_dblPercentage / 100)
        picView.Height = m_udtOriginal.Height * (m_dblPercentage / 100)

        Call picView.PaintPicture(picBuffer.Picture, 0, 0, picView.Width, picView.Height)

        Call CheckForScrolls

        LockWindowUpdate 0

    On Error GoTo 0

End Sub

Private Sub ShowPicture()

    On Error Resume Next

        picView.Visible = True
        picView.Cls

        Set picView.Picture = picBuffer.Picture

        If picView.Width < picMain.Width Then
            picView.Left = (picMain.Width - picView.Width - GetSystemMetrics(SM_CYVTHUMB)) / 2
          Else
            picView.Left = (picMain.Width - picView.Width) / 2
        End If

        If picView.Height < picMain.Height Then
            picView.Top = (picMain.Height - picView.Height - GetSystemMetrics(SM_CYVTHUMB)) / 2
          Else
            picView.Top = (picMain.Height - picView.Height) / 2
        End If

        Call CheckForScrolls

        DoEvents

        Call CheckForScrolls

    On Error GoTo 0

End Sub

Public Sub StorePosition()

    With m_stpLastPosition
        .HorizontalScrollMax = hscPicture.Max
        .HorizontalScrollPosition = hscPicture.Value
        .VerticalScrollMax = vscPicture.Max
        .VerticalScrollPosition = vscPicture.Value
    End With

End Sub

Private Sub timToolbar_Timer()

    On Error Resume Next

        If Not m_bolUseQuickBar Then
            Exit Sub
        End If

        If Not (cmdZoom(0).Enabled = m_bolAllowIn) Then
            cmdZoom(0).Enabled = m_bolAllowIn
        End If

        If Not (cmdZoom(1).Enabled = m_bolAllowOut) Then
            cmdZoom(1).Enabled = m_bolAllowOut
        End If

        If picBuffer.Picture.Handle = 0 Then

            Exit Sub

        End If

        If OnUserControl Then
        
            cmdZoom(0).Visible = True
            cmdZoom(1).Visible = True
        
        Else
            
            cmdZoom(0).Visible = False
            cmdZoom(1).Visible = False
        
        End If

    On Error GoTo 0

End Sub

Function OnUserControl() As Boolean
    Dim pntMouse As POINTAPI
    Dim lngHwnd  As Long
    Dim Ctrl As Variant
  
        Call GetCursorPos(pntMouse)
  
        lngHwnd = WindowFromPoint(pntMouse.X, pntMouse.Y)
  
        If lngHwnd = UserControl.hwnd Then
            
            OnUserControl = True
        
        Else
            For Each Ctrl In UserControl.Controls
      
                On Error GoTo NoHwnd
      
                If lngHwnd = Ctrl.hwnd Then
        
                    OnUserControl = True
                    Exit For
      
                End If

NoHwnd:
      
                On Error GoTo 0
    
            Next
  
        End If
End Function

Public Sub UnloadImage()

    On Error Resume Next

        m_strFileName = vbNullString
        m_udtOriginal.Height = 0
        m_udtOriginal.Width = 0

        picBuffer.Picture = Nothing
        picView.Picture = Nothing

        picView.Cls

        cmdZoom(0).Enabled = False
        cmdZoom(1).Enabled = False
    On Error GoTo 0

End Sub

Public Property Get UseQuickBar() As Boolean

    UseQuickBar = m_bolUseQuickBar

End Property

Public Property Let UseQuickBar(Value As Boolean)

    m_bolUseQuickBar = Value
    PropertyChanged "UseQuickBar"

End Property

Private Sub UserControl_Initialize()

    picTemp.ZOrder

End Sub

Private Sub UserControl_InitProperties()

    On Error Resume Next

        timToolbar.Enabled = UserControl.Ambient.UserMode

    On Error GoTo 0

End Sub

Private Sub Usercontrol_KeyDown( _
                                KeyCode As Integer, _
                                Shift As Integer)

    On Error Resume Next

        If (((KeyCode = 39) Or (KeyCode = 37)) And (hscPicture.Value)) Then
            hscPicture.SetFocus
            Exit Sub
        End If

        If (((KeyCode = 38) Or (KeyCode = 40)) And (vscPicture.Value)) Then
            vscPicture.SetFocus
            Exit Sub
        End If

    On Error GoTo 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next

        timToolbar.Enabled = UserControl.Ambient.UserMode

        picView.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
        picView.BackColor = PropBag.ReadProperty("BackColor", &H8000000C)
        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
        UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
        AllowZoomIn = PropBag.ReadProperty("AllowZoomIn", False)
        AllowZoomOut = PropBag.ReadProperty("AllowZoomOut", False)
        UseQuickBar = PropBag.ReadProperty("UseQuickBar", False)

    On Error GoTo 0

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

        picMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

        picTemp.ZOrder

        RaiseEvent Resize

    On Error GoTo 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next

        Call PropBag.WriteProperty("AutoRedraw", picView.AutoRedraw, True)
        Call PropBag.WriteProperty("BackColor", picView.BackColor, &H8000000C)
        Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
        Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call PropBag.WriteProperty("AllowZoomIn", m_bolAllowIn, False)
        Call PropBag.WriteProperty("AllowZoomOut", m_bolAllowOut, False)
        Call PropBag.WriteProperty("UseQuickBar", m_bolUseQuickBar, False)

    On Error GoTo 0

End Sub

Private Sub vscPicture_Change()

    On Error Resume Next

        picView.Top = vscPicture.Value

    On Error GoTo 0

End Sub

Public Property Let Zoom(New_Zoom As Double)

    m_dblPercentage = New_Zoom

    SetZoom

    RaiseEvent ZoomChanged(New_Zoom)

End Property

Public Property Get Zoom() As Double

    Zoom = m_dblPercentage

End Property

