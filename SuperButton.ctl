VERSION 5.00
Begin VB.UserControl SuperButton 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2115
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   141
   ToolboxBitmap   =   "SuperButton.ctx":0000
   Begin VB.Timer tmrMousePos 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   780
      Top             =   1200
   End
End
Attribute VB_Name = "SuperButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***SuperButton*** by John Sugas 2002. This code uses type lib. TypeLibJS.
'Personal Use ONLY.... Don't sell this code. All disclaimers apply.
Option Explicit
Private Declare Function SendMessageAsAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public Enum EEButtonStyle
    eeFlat = 1
    eeRaised = 2
End Enum
Public Enum EEEdgeStyle
    eeNone = 0
    eeDrawn1 = 1
    eeDrawn2 = 2
    eeClient = 3
    eeModal = 4
End Enum

Public Enum EEFontShadow
    eNone = 0
    eFSRight = 1
    eFSLeft = 2
    eFSAll = 3
End Enum
Const sEmpty = ""

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseEnter()
Event MouseExit()
Event ReadProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,ReadProperties
Attribute ReadProperties.VB_Description = "Occurs when a user control or user document is asked to read its data from a file."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Event WriteProperties(PropBag As PropertyBag) 'MappingInfo=UserControl,UserControl,-1,WriteProperties
Attribute WriteProperties.VB_Description = "Occurs when a user control or user document is asked to write its data to a file."
'main vars
Private fIsShown As Boolean, objContainer As Object, fDownWard As Boolean, hTT As Long
Private bW As Long, bH As Long, bL As Integer, bT As Integer, TI As TOOLINFO, iWinStyle As Long
Private exTop As Long, exLeft As Long, exWidth As Long, exHeight As Long, rc1 As RECT
Private BaseDC As Long, BaseBitmap As Long, BaseMemBitmap As Long, BaseColor As Long
Private BaseParent As Long, BaseWidth As Long, BaseHeight As Long
Private BackColorDC As Long, hbmpBCOld As Long
Private PicDC As Long, PicWidth As Long, PicHeight As Long
Private PicUpDC As Long, PicUpWidth As Long, PicUpHeight As Long
Private PicDnDc As Long, PicDnWidth As Long, PicDnHeight As Long
Private fMouseIn As Boolean, fMouseUp As Boolean, iMove As Long
Private iPicLeft As Long, iPicTop As Long, fHavePicture As Boolean
Private iDnPicLeft As Long, iDnPicTop As Long, fHavePictureDn As Boolean
Private iUpPicLeft As Long, iUpPicTop As Long, fHavePictureUp As Boolean
Private hIconMain As Long, hIconUp As Long, hIconDn As Long
'Api Text vars
Private iTextHeightRet As Long, iTextWidthRet As Long
Private ilBrushStyle As Long, ilHatchStyle As Long, iDT_FLAGS As Long
'Default Property Values:
Const m_def_FontPenWidth = 1
Const m_def_FontPathAntiAlias = 0
Const m_def_EdgeType = 1
Const m_def_MouseRightClickEnable = 0
Const m_def_BackStyle = 0
Const m_def_FontCharSpacing = 0
Const m_def_FontShadow = 0
Const m_def_FontOrientation = 0
Const m_def_FontEscapement = 0
Const m_def_FontShadowOffsetX = 3
Const m_def_FontShadowOffsetY = 3
Const m_def_FontUseBrush = 0
Const m_def_FontColorPen = vbRed
Const m_def_FontColorBrush = vbBlue
Const m_def_ToolTipText = ""
Const m_def_Caption = ""
Const m_def_FontUsePen = 0
Const m_def_CaptionLeft = -1
Const m_def_CaptionTop = -1
Const m_def_FontColor = vbBlack
Const m_def_FontShadowLeft = -1
Const m_def_FontShadowRight = -1
Const m_def_FontTransParent = 1
Const m_def_FontROP = vbCopyPen
Const m_def_BackColor = vbWhite
Const m_def_PictureOffsetX = 0
Const m_def_PictureOffsetY = 0
Const m_def_MaskColor = vbWhite
Const m_def_AutoSize = 0
Const m_def_EdgeWidth = 1
Const m_def_ButtonStyle = 2
Const m_def_ToolTipShowDelay = 500
Const m_def_ToolTipTimeShown = 2000
Const m_def_ToolTipBackColor = &HF1CF30
Const m_def_ToolTipFontColor = vbBlack
Const m_def_PicDownMove = True
Const m_def_MaskColorUpPic = vbWhite
Const m_def_MaskColorDownPic = vbBlack
Const m_def_CaptionCenterX = 1
Const m_def_CaptionCenterY = 1
Const m_def_ToolTipMaxWidth = 100

'Property Variables:
Dim m_FontPenWidth As Long
Dim m_FontPathAntiAlias As Boolean
Dim m_EdgeType As EEEdgeStyle
Dim m_MouseRightClickEnable As Boolean
Dim m_BackStyle As eeBackStyle
Dim m_FontCharSpacing As Long
Dim m_FontShadow As EEFontShadow
Dim m_FontOrientation As Long
Dim m_FontEscapement As Long
Dim m_FontShadowOffsetX As Long
Dim m_FontShadowOffsetY As Long
Dim m_FontUseBrush As Boolean
Dim m_FontColorPen As OLE_COLOR
Dim m_FontColorBrush As OLE_COLOR
Dim m_ToolTipText As String
Dim m_ToolTipFontColor As OLE_COLOR
Dim m_ToolTipBackColor As OLE_COLOR
Dim m_Caption As String
Dim m_FontUsePen As Boolean
Dim m_CaptionCenterX As Boolean
Dim m_CaptionCenterY As Boolean
Dim m_FontColor As OLE_COLOR
Dim m_FontShadowLeft  As OLE_COLOR
Dim m_FontShadowRight  As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_FontROP As DrawModeConstants 'RasterOpConstants
Dim m_FontTransParent As Boolean
Dim m_MaskColorUpPic As OLE_COLOR
Dim m_MaskColorDownPic As OLE_COLOR
Dim m_PicDownMove As Boolean
Dim m_ToolTipShowDelay As Long
Dim m_ToolTipTimeShown As Long
Dim m_ToolTipMaxWidth As Long
Dim m_PictureUp As StdPicture
Dim m_PictureDown As StdPicture
Dim m_CaptionLeft As Long
Dim m_CaptionTop As Long
Dim m_PictureOffsetX As Long
Dim m_PictureOffsetY As Long
Dim m_MaskColor As OLE_COLOR
Dim m_AutoSize As Boolean
Dim m_EdgeWidth As Long
Dim m_ButtonStyle As EEButtonStyle
Dim m_Picture As StdPicture


Private Sub DrawButtonDown()
    Dim i As Long, rcEdge As RECT
    
    PaintBaseCoat
    If m_EdgeType <> 0 Then
        CopyRect rcEdge, rc1 'get copy of userctrl dimensions
        If m_EdgeType < 3 Then  'only for the drawn edges
            Call InflateRect(rcEdge, 2, 2) 'cover the very edge of ctrl
            Call DrawEdge(UserControl.hdc, rcEdge, EDGE_SUNKEN, BF_RECT)
        End If
        If m_EdgeType = eeDrawn1 Then
            For i = 1 To EdgeWidth
                Call InflateRect(rcEdge, -1, -1)
                Call DrawEdge(UserControl.hdc, rcEdge, BDR_SUNKENINNER, BF_RECT)
                Call InflateRect(rcEdge, -1, -1)
                Call DrawEdge(UserControl.hdc, rcEdge, BDR_SUNKENINNER, BF_RECT)
            Next
        ElseIf m_EdgeType = eeDrawn2 Then
            For i = 1 To EdgeWidth
                Call InflateRect(rcEdge, -3, -3)
                Call DrawEdge(UserControl.hdc, rcEdge, EDGE_SUNKEN, BF_RECT)
            Next
        Else
            GetClientRect UserControl.hwnd, rcEdge
            For i = 1 To EdgeWidth
                Call DrawEdge(UserControl.hdc, rcEdge, EDGE_SUNKEN, BF_RECT) 'EDGE_SUNKEN
                Call InflateRect(rcEdge, -1, -1)
            Next
        End If
    End If
    
    iMove = 0
    If m_PicDownMove Then iMove = 1 'move 1x1 pixel for a downward move
    If fHavePictureDn Then
        If PictureDown.Type = vbPicTypeBitmap Then
            DrawState UserControl.hdc, 0, 0, hIconDn, 0, iDnPicLeft + iMove, iDnPicTop + iMove, PicDnWidth, PicDnHeight, DST_ICON Or DSS_NORMAL
'            BitBlt UserControl.hdc, iDnPicLeft + iMove, iDnPicTop + iMove, PicDnWidth, PicDnHeight, PicDnDc, 0, 0, vbSrcInvert
'            BitBlt UserControl.hdc, iDnPicLeft + iMove, iDnPicTop + iMove, PicDnWidth, PicDnHeight, hdcDnMask, 0, 0, vbSrcAnd
'            BitBlt UserControl.hdc, iDnPicLeft + iMove, iDnPicTop + iMove, PicDnWidth, PicDnHeight, PicDnDc, 0, 0, vbSrcInvert
        ElseIf PictureDown.Type = vbPicTypeIcon Then
            DrawIcon UserControl.hdc, iDnPicLeft + iMove, iDnPicTop + iMove, PictureDown
        End If
    Else
        If fHavePicture Then
            If Picture.Type = vbPicTypeBitmap Then
                DrawState UserControl.hdc, 0, 0, hIconMain, 0, iPicLeft + iMove, iPicTop + iMove, PicWidth, PicHeight, DST_ICON Or DSS_NORMAL
            ElseIf Picture.Type = vbPicTypeIcon Then
                DrawIcon UserControl.hdc, iPicLeft + iMove, iPicTop + iMove, Picture
            End If
        End If
    End If

    DrawTextOnDc m_CaptionLeft + iMove, m_CaptionTop + iMove
    Refresh
End Sub

Private Sub DrawButtonUp()
    Dim i As Long, rcEdge As RECT
    
    PaintBaseCoat
    If m_EdgeType <> 0 Then
        CopyRect rcEdge, rc1
        If m_EdgeType < 3 Then  'only for the drawn edges
            Call InflateRect(rcEdge, 2, 2) 'cover the very edge of ctrl
            Call DrawEdge(UserControl.hdc, rcEdge, EDGE_SUNKEN, BF_RECT)
        End If
        If m_EdgeType = eeDrawn1 Then
            For i = 1 To EdgeWidth
                Call InflateRect(rcEdge, -1, -1)
                Call DrawEdge(UserControl.hdc, rcEdge, BDR_RAISEDINNER, BF_RECT)
                Call InflateRect(rcEdge, -1, -1)
                Call DrawEdge(UserControl.hdc, rcEdge, BDR_RAISEDINNER, BF_RECT)
            Next
        ElseIf m_EdgeType = eeDrawn2 Then
            For i = 1 To EdgeWidth
                Call InflateRect(rcEdge, -3, -3)
                Call DrawEdge(UserControl.hdc, rcEdge, EDGE_RAISED, BF_RECT)
            Next
        Else
            GetClientRect UserControl.hwnd, rcEdge
            For i = 1 To EdgeWidth
                Call DrawEdge(UserControl.hdc, rcEdge, EDGE_RAISED, BF_RECT)
                Call InflateRect(rcEdge, -1, -1)
            Next
        End If
    End If
    
    If Not fMouseUp Then Exit Sub 'don't need to draw pictures only edges
    If fHavePictureUp Then
        If PictureUp.Type = vbPicTypeBitmap Then
            DrawState UserControl.hdc, 0, 0, hIconUp, 0, iUpPicLeft, iUpPicTop, PicUpWidth, PicUpHeight, DST_ICON Or DSS_NORMAL
'            BitBlt UserControl.hdc, iUpPicLeft, iUpPicTop, PicUpWidth, PicUpHeight, PicUpDC, 0, 0, vbSrcInvert
'            BitBlt UserControl.hdc, iUpPicLeft, iUpPicTop, PicUpWidth, PicUpHeight, hdcUpMask, 0, 0, vbSrcAnd
'            BitBlt UserControl.hdc, iUpPicLeft, iUpPicTop, PicUpWidth, PicUpHeight, PicUpDC, 0, 0, vbSrcInvert
        ElseIf PictureUp.Type = vbPicTypeIcon Then
            DrawIcon UserControl.hdc, iUpPicLeft, iUpPicTop, PictureUp
        End If
    Else  'only main pic... do the animation if set... pic is moved 1x1 pixel, move back
        If fHavePicture Then
            If Picture.Type = vbPicTypeBitmap Then
                DrawState UserControl.hdc, 0, 0, hIconMain, 0, iPicLeft, iPicTop, PicWidth, PicHeight, DST_ICON Or DSS_NORMAL
            ElseIf Picture.Type = vbPicTypeIcon Then
                DrawIcon UserControl.hdc, iPicLeft, iPicTop, Picture
            End If
        End If
    End If
    DrawTextOnDc m_CaptionLeft, m_CaptionTop
    Refresh
End Sub
Private Sub PaintBaseCoat()
    UserControl.Cls
    If m_BackStyle = bsOpaque Then
        StretchBlt UserControl.hdc, UserControl.ScaleLeft, UserControl.ScaleTop, _
                UserControl.ScaleWidth, UserControl.ScaleHeight, _
                            BackColorDC, 0, 0, _
                             bW, bH, vbSrcCopy
    Else
        StretchBlt UserControl.hdc, 0, 0, exWidth, exHeight, _
                            BaseDC, 0, 0, _
                             bW, bH, vbSrcCopy
    End If
End Sub

Private Sub DrawButton()
    PaintBaseCoat
    If m_EdgeType <> eeNone Then _
        If ButtonStyle = eeRaised Or fMouseIn Then DrawButtonUp

    If fHavePicture Then
        If Picture.Type = vbPicTypeBitmap Then
            If UserControl.Enabled Then
                DrawState UserControl.hdc, 0, 0, hIconMain, 0, iPicLeft, iPicTop, PicWidth, PicHeight, DST_ICON Or DSS_NORMAL
            Else
                DrawState UserControl.hdc, 0, 0, hIconMain, 0, iPicLeft, iPicTop, PicWidth, PicHeight, DST_ICON Or DSS_DISABLED  'DST_BITMAP
            End If
        ElseIf Picture.Type = vbPicTypeIcon Then
            If UserControl.Enabled Then
                DrawIcon UserControl.hdc, iPicLeft, iPicTop, Picture
            Else
                DrawState UserControl.hdc, 0, 0, Picture, 0, iPicLeft, iPicTop, PicWidth, PicHeight, DST_ICON Or DSS_DISABLED
            End If
        End If
    End If
    DrawTextOnDc m_CaptionLeft, m_CaptionTop
    Refresh
End Sub

Private Sub DrawTextOnDc(x As Long, y As Long)
    Dim rc As RECT, HFONT As Long, wTextParams As DRAWTEXTPARAMS, rcBack As RECT, rcLeft As RECT
    Dim hOldFont As Long, lf As LOGFONT, bFaceName() As Byte, j As Integer, lBG As Long
    Dim lTC As Long, lBc As Long, aColor(1 To 3) As Long, i As Integer, rcRight As RECT
    Dim usepen As Long, oldpen As Long, useBrush As Long, oldBrush As Long, iDoPath As Integer
    Dim tLogBrush As LOGBRUSH, oldROP As Long
    
    If m_Caption$ = sEmpty$ Then Exit Sub

    aColor(3) = m_FontShadowLeft 'put colors in array for ease of use in loop
    aColor(2) = m_FontShadowRight
    aColor(1) = m_FontColor 'lColorMain
    ilBrushStyle = BS_SOLID
    ilHatchStyle = HS_CROSS
    
    
    Dim iWorkHdc As Long
    iWorkHdc = UserControl.hdc

    Call SetGraphicsMode(iWorkHdc, GM_ADVANCED)
    
    With lf
        .lfHeight = -(UserControl.Font.Size / 75 * GetDeviceCaps(UserControl.hdc, LOGPIXELSY))
        .lfWidth = m_FontCharSpacing
        .lfEscapement = m_FontEscapement * 10 ' Set lfEscapement and lfOrientation to 10 * angle to rotate text
        .lfOrientation = m_FontOrientation * 10 ' for example, to rotate 45 degrees set the values to 450
        .lfWeight = UserControl.Font.Weight ' ilfWeight
        .lfItalic = Abs(UserControl.FontItalic)  ' blfItalic
        .lfUnderline = Abs(UserControl.FontUnderline)  ' blfUnderline
        .lfStrikeOut = Abs(UserControl.FontStrikethru)  ' blfStrikeOut
        .lfCharSet = UserControl.Font.Charset
        .lfOutPrecision = OUT_TT_PRECIS
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfQuality = ANTIALIASED_QUALITY
        .lfPitchAndFamily = (DEFAULT_PITCH Or FF_DONTCARE) ' FF_SWISS
    End With
    bFaceName = StrConv(UserControl.Font.Name & Chr$(0), vbFromUnicode)
    For j = 0 To UBound(bFaceName)
        lf.lfFaceName(j) = bFaceName(j)
    Next j
    HFONT = CreateFontIndirectANSII(lf)
    hOldFont = SelectObject(iWorkHdc, HFONT)
    wTextParams.cbSize = Len(wTextParams)
    iDT_FLAGS = DT_CENTER Or DT_NOCLIP Or DT_VCENTER Or DT_WORDBREAK
    
    oldROP = SetROP2(iWorkHdc, m_FontROP)
    lTC = SetTextColor(iWorkHdc, FontColor)
    lBc = SetBkColor(iWorkHdc, BackColor)
    lBG = SetBkMode(iWorkHdc, TRANSPARENT)

    If m_FontUsePen Then
        iDoPath = 1  'iDoPath used in case switch later
    Else

    End If
    If m_FontUseBrush Then
        iDoPath = iDoPath + 2
        tLogBrush.lbColor = m_FontColorBrush
        tLogBrush.lbHatch = ilHatchStyle
        tLogBrush.lbStyle = ilBrushStyle 'BS_HATCHED ' BS_SOLID
    Else
    
    End If
    
    If iDoPath > 0 Then BeginPath iWorkHdc 'UserControl.hdc
    
    'init rect to main text size
    With rcBack
        .Left = x: .Top = y
        .Right = x + iTextWidthRet
        .Bottom = y + iTextHeightRet
    End With
    CopyRect rc, rcBack
    'get any extremes from shadow text
    If m_FontShadow <> eNone Then
        If m_FontShadow = eFSRight Or m_FontShadow = eFSAll Then
            With rcRight
                .Left = x + m_FontShadowOffsetX
                .Top = y + m_FontShadowOffsetY
                .Right = x + iTextWidthRet + m_FontShadowOffsetX
                .Bottom = y + iTextHeightRet + m_FontShadowOffsetY
            End With
            UnionRect rc, rc, rcRight
        End If
        If m_FontShadow = eFSLeft Or m_FontShadow = eFSAll Then
            With rcLeft
                .Left = x + -m_FontShadowOffsetX
                .Top = y + -m_FontShadowOffsetY
                .Right = x + iTextWidthRet + -m_FontShadowOffsetX
                .Bottom = y + iTextHeightRet + -m_FontShadowOffsetY
            End With
            UnionRect rc, rc, rcLeft
        End If
    End If
    If FontTransParent = False Then
        Dim rectBrush As Long, rectBrushOld As Long
        rectBrush = CreateSolidBrush(BackColor)
        rectBrushOld = SelectObject(iWorkHdc, rectBrush)
        InflateRect rc, 3, 3
        Rectangle iWorkHdc, rc.Left, rc.Top, rc.Right, rc.Bottom
        SetRectEmpty rc
        Call SelectObject(iWorkHdc, rectBrushOld)
        Call DeleteObject(rectBrush)
    End If
    ' draw the text transp so bg boxes don't overlap
    
    For j = 3 To 1 Step -1
        If j = 3 Then  'left shadow
            If m_FontShadow <> eFSAll Then _
                If m_FontShadow = eFSRight Or m_FontShadow = eNone Then GoTo ByPass
            CopyRect rc, rcLeft
        End If
        If j = 2 Then 'right shadow
            If m_FontShadow <> eFSAll Then _
                If m_FontShadow = eFSLeft Or m_FontShadow = eNone Then GoTo ByPass
            CopyRect rc, rcRight
        End If
        If j = 1 Then  'main text
            CopyRect rc, rcBack
        End If

        Call SetTextColor(iWorkHdc, aColor(j))
        If UserControl.Enabled Then
            Call DrawTextExANSII(iWorkHdc, m_Caption, Len(m_Caption), rc, iDT_FLAGS, wTextParams)
        Else
            DrawStateByString iWorkHdc, 0, 0, Caption, Len(Caption), rc.Left, rc.Top, 0, 0, DST_TEXT Or DSS_DISABLED
        End If
ByPass:
    Next j

    EndPath iWorkHdc
    If iDoPath <> 0 Then
        Select Case iDoPath
            Case 0
            Case 1
                usepen = CreatePen(PS_Solid, m_FontPenWidth, m_FontColorPen)
                oldpen = SelectObject(iWorkHdc, usepen)
                StrokeAndFillPath iWorkHdc
                Call SelectObject(iWorkHdc, oldpen)
                Call DeleteObject(usepen)
            Case 2
                useBrush = CreateBrushIndirect(tLogBrush)
                oldBrush = SelectObject(iWorkHdc, useBrush)
                StrokeAndFillPath iWorkHdc
                Call SelectObject(iWorkHdc, oldBrush)
                Call DeleteObject(useBrush)
            Case 3
                usepen = CreatePen(PS_Solid, m_FontPenWidth, m_FontColorPen)
                oldpen = SelectObject(iWorkHdc, usepen)
                useBrush = CreateSolidBrush(m_FontColorBrush)
                oldBrush = SelectObject(iWorkHdc, useBrush)
                StrokeAndFillPath iWorkHdc
                Call SelectObject(iWorkHdc, oldBrush)
                Call DeleteObject(useBrush)
                Call SelectObject(iWorkHdc, oldpen)
                Call DeleteObject(usepen)
        End Select
    End If

    Call SetTextColor(iWorkHdc, lTC)
    Call SetBkMode(iWorkHdc, lBG)
    Call SetROP2(iWorkHdc, oldROP)
    Call SetBkColor(iWorkHdc, lBc)
    Call SelectObject(iWorkHdc, hOldFont)
    DeleteObject HFONT
    Refresh
End Sub
Private Sub CenterTheCaption()
    GetTextDimensions
    If CaptionCenterX Then
        CaptionLeft = (UserControl.ScaleWidth - iTextWidthRet) \ 2
    End If
    If CaptionCenterY Then
        CaptionTop = (UserControl.ScaleHeight - iTextHeightRet) \ 2
    End If
    DrawButton
End Sub
Private Sub GetTextDimensions()
    Dim ReturnSL As SIZEL
    iTextHeightRet = 0
    iTextWidthRet = 0
    If m_Caption = sEmpty$ Then Exit Sub
    Call GetTextExtentPoint32(UserControl.hdc, m_Caption, Len(m_Caption), ReturnSL)
    iTextHeightRet = ReturnSL.cy
    iTextWidthRet = ReturnSL.cx
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    UserControl.AutoRedraw = True
    UserControl.BackColor = vbWhite
    
End Sub

Private Sub UserControl_Show()
    CenterTheCaption
    If hTT = 0 Then
        Call InitCommonControls 'have to call or won't work
        hTT = CreateWindowEx(0, "tooltips_class32", 0&, TTS_NOPREFIX Or TTS_ALWAYSTIP, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, objContainer.hwnd, 0&, App.hInstance, ByVal 0&)
        Dim lStyle As Long
        If hTT Then
           lStyle = GetWindowLong(hTT, GWL_STYLE)
           SetWindowLong hTT, GWL_STYLE, lStyle
        End If
        DoEvents
    End If
    fIsShown = True
    RaiseEvent Show
End Sub
Private Sub InitButton()
    Dim oldMode As Long, oldAutoRedraw As Long
    Set objContainer = Extender.Container
    oldMode = objContainer.ScaleMode
    objContainer.ScaleMode = vbPixels
    oldAutoRedraw = objContainer.AutoRedraw
    objContainer.AutoRedraw = True
    
    rc1.Left = 0:    rc1.Top = 0 'rect for drawing
    rc1.Right = UserControl.ScaleWidth
    rc1.Bottom = UserControl.ScaleHeight
    'picture positions
    iPicLeft = ((UserControl.ScaleWidth - PicWidth) \ 2) + PictureOffsetX
    iPicTop = ((UserControl.ScaleHeight - PicHeight) \ 2) + PictureOffsetY
    iDnPicLeft = ((UserControl.ScaleWidth - PicDnWidth) \ 2) + PictureOffsetX
    iDnPicTop = ((UserControl.ScaleHeight - PicDnHeight) \ 2) + PictureOffsetY
    iUpPicLeft = ((UserControl.ScaleWidth - PicUpWidth) \ 2) + PictureOffsetX
    iUpPicTop = ((UserControl.ScaleHeight - PicUpHeight) \ 2) + PictureOffsetY
    CreateBase   'create DC for the base picture
    CreateBackColorDC 'backcolor DC... prob's with ctrl b.c. and masking bitmaps
    CreateDC  'create DC for the static button picture
    CreateUpDC
    CreateDnDC
    CenterTheCaption
    DrawButton  'draw it.....
    objContainer.ScaleMode = oldMode
    objContainer.AutoRedraw = oldAutoRedraw
End Sub
Private Function CreateBase() As Long

    If IsBaseCreated Then DestroyBase

    BaseParent = objContainer.hdc
    BaseWidth = bW
    BaseHeight = bH

    ' Create a memory device context to use
    BaseDC = CreateCompatibleDC(BaseParent)

    ' Tell'em it's a picture (so drawings can be done on the DC)
    BaseMemBitmap = CreateCompatibleBitmap(BaseParent, BaseWidth, BaseHeight)
    BaseBitmap = SelectObject(BaseDC, BaseMemBitmap)

    Call SetBkColor(BaseDC, GetBkColor(BaseParent))
    BaseColor = GetTextColor(BaseParent)
    Call SetBkMode(BaseDC, GetBkMode(BaseParent))
    
    StretchBlt BaseDC, 0, 0, bW, bH, _
                            objContainer.hdc, exLeft, exTop, _
                             bW, bH, vbSrcCopy
        
    CreateBase = BaseDC
End Function
Private Sub DestroyBase()
    If Not IsBaseCreated Then Exit Sub
    Call SelectObject(BaseDC, BaseBitmap)
    Call DeleteObject(BaseMemBitmap)
    Call DeleteDC(BaseDC)
    BaseDC = 0
    UserControl.Cls
End Sub
Private Function IsBaseCreated() As Boolean
    If BaseDC <> 0 Then IsBaseCreated = True Else: IsBaseCreated = False
End Function
Private Function CreateBackColorDC() As Long
    Dim useBrush As Long, oldBrush As Long, rcBrh As RECT, hbmpBC As Long
    If IsBackColorDCCreated Then DestroyBackColorDC
    
    BackColorDC = CreateCompatibleDC(BaseParent) 'create dc
    hbmpBC = CreateCompatibleBitmap(BaseParent, bW, bH) 'create bmp
    hbmpBCOld = SelectObject(BackColorDC, hbmpBC)
    rcBrh.Bottom = UserControl.ScaleHeight: rcBrh.Right = UserControl.ScaleWidth
    useBrush = CreateSolidBrush(BackColor)
    oldBrush = SelectObject(BackColorDC, useBrush)
    FillRect BackColorDC, rcBrh, BS_SOLID 'fill dc with backcolor
    Call SelectObject(BackColorDC, oldBrush)
    Call DeleteObject(useBrush)
    Call DeleteObject(hbmpBC)
    CreateBackColorDC = BackColorDC
End Function
Private Sub DestroyBackColorDC()
    If Not IsBackColorDCCreated Then Exit Sub
    Call SelectObject(BackColorDC, hbmpBCOld)
    Call DeleteDC(BackColorDC)
    BackColorDC = 0
End Sub
Private Function IsBackColorDCCreated() As Boolean
    If BackColorDC <> 0 Then IsBackColorDCCreated = True Else: IsBackColorDCCreated = False
End Function
Private Function CreateDC() As Long
    If IsDcCreated Then DestroyDC
    If Not fHavePicture Then Exit Function

    If Picture.Type = vbPicTypeBitmap Then
        Dim PicBitmap As Long, PicBitmapOld As Long, tIconInfo As ICONINFO
        Dim AndBmp As Long, AndMask As Long, AndBmpOld As Long
        Dim hdcMask As Long, hbmpMask As Long, hbmpMaskOld As Long
        Dim cPlanes As Long, cPixelBits As Long
        cPlanes = GetDeviceCaps(UserControl.hdc, 14)
        cPixelBits = GetDeviceCaps(UserControl.hdc, 12)
        'create bitmap from the picture
        PicDC = CreateCompatibleDC(0&)
        PicBitmap = CreateCompatibleBitmap(PicDC, PicWidth, PicHeight)
        PicBitmapOld = SelectObject(PicDC, Picture.Handle)
        Call BitBlt(PicDC, 0, 0, PicWidth, PicHeight, Picture.Handle, 0, 0, vbSrcCopy)
        
        AndMask = CreateCompatibleDC(0&)
        AndBmp = CreateBitmap(PicWidth, PicHeight, cPlanes, cPixelBits, 0&)
        AndBmpOld = SelectObject(AndMask, AndBmp)
        Call BitBlt(AndMask, 0, 0, PicWidth, PicHeight, PicDC, 0, 0, vbSrcCopy)
        
        'Create mask next
        hdcMask = CreateCompatibleDC(0&)
        ' Create bitmap (monochrome by default)
        hbmpMask = CreateCompatibleBitmap(hdcMask, PicWidth, PicHeight)
        ' Select it into DC
        hbmpMaskOld = SelectObject(hdcMask, hbmpMask)
        ' Set background of source to the mask color
        Call SetBkColor(PicDC, MaskColor)
        ' Copy color bitmap to monochrome DC to create mono mask
        Call BitBlt(hdcMask, 0, 0, PicWidth, PicHeight, PicDC, 0, 0, vbSrcCopy) 'vbSrcCopy
        
        'pic & mask can't be selected into dc when creating icon
        Call SelectObject(PicDC, PicBitmapOld)
        Call DeleteDC(PicDC)
        ' Invert background of image to create AND Mask
        Call SetBkColor(AndMask, vbBlack)
        Call SetTextColor(AndMask, vbWhite)
        Call BitBlt(AndMask, 0, 0, PicWidth, PicHeight, hdcMask, 0, 0, vbSrcAnd)
        Call SelectObject(hdcMask, hbmpMaskOld)
        Call DeleteDC(hdcMask)
        Call SelectObject(AndMask, AndBmpOld)
        Call DeleteDC(AndMask)
        With tIconInfo
            .fIcon = True
            .hbmColor = AndBmp
            .hbmMask = hbmpMask
        End With
        hIconMain = CreateIconIndirect(tIconInfo)
        Call DeleteObject(hbmpMask)
        Call DeleteObject(AndBmp)
        Call DeleteObject(PicBitmap)
    End If
    CreateDC = PicDC
End Function
Private Sub DestroyDC()
    If Not IsDcCreated Then Exit Sub
    'delete the icon created
    DestroyIcon hIconMain
    PicDC = 0
End Sub
Private Function IsDcCreated() As Boolean
    If PicDC <> 0 Then IsDcCreated = True Else: IsDcCreated = False
End Function

Private Function CreateUpDC() As Long
    If IsUpDcCreated Then DestroyUpDC
    If Not fHavePictureUp Then Exit Function
    
    If PictureUp.Type = vbPicTypeBitmap Then
        Dim hbmpUpPic As Long, hbmpUpPicOld As Long, tIconInfo As ICONINFO
        Dim AndBmp As Long, AndMask As Long, AndBmpOld As Long
        Dim hdcUpMask As Long, hbmpUpMask As Long, hbmpUpMaskOld As Long
        Dim cPlanes As Long, cPixelBits As Long
        cPlanes = GetDeviceCaps(UserControl.hdc, 14)
        cPixelBits = GetDeviceCaps(UserControl.hdc, 12)
        PicUpDC = CreateCompatibleDC(0&)
        hbmpUpPic = CreateCompatibleBitmap(PicUpDC, PicUpWidth, PicUpHeight)
        hbmpUpPicOld = SelectObject(PicUpDC, PictureUp.Handle)
        Call BitBlt(PicUpDC, 0, 0, PicUpWidth, PicUpHeight, PictureUp.Handle, 0, 0, vbSrcCopy)
        
        AndMask = CreateCompatibleDC(0&)
        AndBmp = CreateBitmap(PicUpWidth, PicUpHeight, cPlanes, cPixelBits, 0&)
        AndBmpOld = SelectObject(AndMask, AndBmp)
        Call BitBlt(AndMask, 0, 0, PicUpWidth, PicUpHeight, PicUpDC, 0, 0, vbSrcCopy)
        
        hdcUpMask = CreateCompatibleDC(0&)
        hbmpUpMask = CreateCompatibleBitmap(hdcUpMask, PicUpWidth, PicUpHeight)
        hbmpUpMaskOld = SelectObject(hdcUpMask, hbmpUpMask)
        Call SetBkColor(PicUpDC, MaskColorUpPic)
        Call BitBlt(hdcUpMask, 0, 0, PicUpWidth, PicUpHeight, PicUpDC, 0, 0, vbSrcCopy)
        
        'pic & mask can't be selected into dc when creating icon
        Call SelectObject(PicUpDC, hbmpUpPicOld)
        Call DeleteDC(PicUpDC)
        ' Invert background of image to create AND Mask
        Call SetBkColor(AndMask, vbBlack)
        Call SetTextColor(AndMask, vbWhite)
        Call BitBlt(AndMask, 0, 0, PicUpWidth, PicUpHeight, hdcUpMask, 0, 0, vbSrcAnd)
        Call SelectObject(hdcUpMask, hbmpUpMaskOld)
        Call DeleteDC(hdcUpMask)
        Call SelectObject(AndMask, AndBmpOld)
        Call DeleteDC(AndMask)
        With tIconInfo
            .fIcon = True
            .hbmColor = AndBmp
            .hbmMask = hbmpUpMask
        End With
        hIconUp = CreateIconIndirect(tIconInfo)
        Call DeleteObject(hbmpUpMask)
        Call DeleteObject(AndBmp)
        Call DeleteObject(hbmpUpPic)
    End If
    CreateUpDC = PicUpDC
End Function
Private Sub DestroyUpDC()
    If Not IsUpDcCreated Then Exit Sub
    DestroyIcon hIconUp
    PicUpDC = 0
End Sub
Private Function IsUpDcCreated() As Boolean
    If PicUpDC <> 0 Then IsUpDcCreated = True Else: IsUpDcCreated = False
End Function
Private Function CreateDnDC() As Long
    If IsDnDcCreated Then DestroyDnDC
    If Not fHavePictureDn Then Exit Function
    
    If PictureDown.Type = vbPicTypeBitmap Then
        Dim hbmpDnPic As Long, hbmpDnPicOld As Long, tIconInfo As ICONINFO
        Dim AndBmp As Long, AndMask As Long, AndBmpOld As Long
        Dim hdcDnMask As Long, hbmpDnMask As Long, hbmpDnMaskOld As Long
        Dim cPlanes As Long, cPixelBits As Long
        cPlanes = GetDeviceCaps(UserControl.hdc, 14)
        cPixelBits = GetDeviceCaps(UserControl.hdc, 12)
        PicDnDc = CreateCompatibleDC(0&)
        hbmpDnPic = CreateCompatibleBitmap(PicDnDc, PicDnWidth, PicDnHeight)
        hbmpDnPicOld = SelectObject(PicDnDc, PictureDown.Handle)
        Call BitBlt(PicDnDc, 0, 0, PicDnWidth, PicDnHeight, PictureDown.Handle, 0, 0, vbSrcCopy)
        
        AndMask = CreateCompatibleDC(0&)
        AndBmp = CreateBitmap(PicDnWidth, PicDnHeight, cPlanes, cPixelBits, 0&)
        AndBmpOld = SelectObject(AndMask, AndBmp)
        Call BitBlt(AndMask, 0, 0, PicDnWidth, PicDnHeight, PicDnDc, 0, 0, vbSrcCopy)
        
        hdcDnMask = CreateCompatibleDC(0&)
        hbmpDnMask = CreateCompatibleBitmap(hdcDnMask, PicDnWidth, PicDnHeight)
        hbmpDnMaskOld = SelectObject(hdcDnMask, hbmpDnMask)
        Call SetBkColor(PicDnDc, MaskColorDownPic)
        Call BitBlt(hdcDnMask, 0, 0, PicDnWidth, PicDnHeight, PicDnDc, 0, 0, vbSrcCopy)
        
        'pic & mask can't be selected into dc when creating icon
        Call SelectObject(PicDnDc, hbmpDnPicOld)
        Call DeleteDC(PicDnDc)
        ' Invert background of image to create AND Mask
        Call SetBkColor(AndMask, vbBlack)
        Call SetTextColor(AndMask, vbWhite)
        Call BitBlt(AndMask, 0, 0, PicDnWidth, PicDnHeight, hdcDnMask, 0, 0, vbSrcAnd)
        Call SelectObject(hdcDnMask, hbmpDnMaskOld)
        Call DeleteDC(hdcDnMask)
        Call SelectObject(AndMask, AndBmpOld)
        Call DeleteDC(AndMask)
        With tIconInfo
            .fIcon = True
            .hbmColor = AndBmp
            .hbmMask = hbmpDnMask
        End With
        hIconDn = CreateIconIndirect(tIconInfo)
        Call DeleteObject(hbmpDnMask)
        Call DeleteObject(AndBmp)
        Call DeleteObject(hbmpDnPic)
    End If
    CreateDnDC = PicDnDc
End Function
Private Sub DestroyDnDC()
    If Not IsDnDcCreated Then Exit Sub
'    Call SelectObject(hdcDnMask, hbmpDnMaskOld)
'    Call DeleteDC(hdcDnMask)
'    Call DeleteObject(hbmpDnMask)
'    Call SelectObject(PicDnDc, hbmpDnPicOld)
'    Call DeleteDC(PicDnDc)
    DestroyIcon hIconDn
    PicDnDc = 0
End Sub
Private Function IsDnDcCreated() As Boolean
    If PicDnDc <> 0 Then IsDnDcCreated = True Else: IsDnDcCreated = False
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,
Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Picture = New_Picture
    fHavePicture = False
    If Not New_Picture Is Nothing Then
        PicWidth = ScaleX(New_Picture.Width, vbHimetric, vbPixels)
        PicHeight = ScaleY(New_Picture.Height, vbHimetric, vbPixels)
        fHavePicture = True
        PropertyChanged "Picture"
    End If
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureUp() As StdPicture
    Set PictureUp = m_PictureUp
End Property

Public Property Set PictureUp(ByVal New_PictureUp As StdPicture)
    Set m_PictureUp = New_PictureUp
    fHavePictureUp = False
    If Not New_PictureUp Is Nothing Then
        PicUpWidth = ScaleX(New_PictureUp.Width, vbHimetric, vbPixels)
        PicUpHeight = ScaleY(New_PictureUp.Height, vbHimetric, vbPixels)
        fHavePictureUp = True
        PropertyChanged "PictureUp"
    End If
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get PictureDown() As StdPicture
    Set PictureDown = m_PictureDown
End Property

Public Property Set PictureDown(ByVal New_PictureDown As StdPicture)
    Set m_PictureDown = New_PictureDown
    fHavePictureDn = False
    If Not New_PictureDown Is Nothing Then
        PicDnWidth = ScaleX(New_PictureDown.Width, vbHimetric, vbPixels)
        PicDnHeight = ScaleY(New_PictureDown.Height, vbHimetric, vbPixels)
        fHavePictureDn = True
        PropertyChanged "PictureDown"
    End If
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    InitButton
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And m_MouseRightClickEnable = False Then Exit Sub
    DrawButtonDown
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If GetCapture() <> UserControl.hwnd Then
         'not really needed except to prevent extender tooltips showing
         SetCapture UserControl.hwnd
    End If
    If Not fMouseIn Then
         fMouseIn = True
         'If ButtonStyle = eeFlat Then DrawButton
         fMouseUp = True
         DrawButtonUp
         tmrMousePos.Enabled = True
         RaiseEvent MouseEnter
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And m_MouseRightClickEnable = False Then Exit Sub
    fMouseUp = True
    DrawButtonUp
    RaiseEvent MouseUp(Button, Shift, x, y)
    doWorkAround
End Sub
Private Sub tmrMousePos_Timer()
    If Not IsMouseOver Then
         ReleaseCapture
         RaiseEvent MouseExit
         tmrMousePos.Enabled = False
         fMouseIn = False
         fMouseUp = False
         DrawButton
    End If
End Sub
Private Function IsMouseOver() As Boolean
    Dim pt As POINTL
    On Error Resume Next
    GetCursorPos pt
    If WindowFromPoint(pt.x, pt.y) = UserControl.hwnd Then
         IsMouseOver = True
    End If
End Function

' Create a tiny mouse movement to work around a mysterious phenomenom that if one clicks
' a few times in the same spot the Mouse_up may not respond. 'this function   By Herman Liu
Private Sub doWorkAround()
    Dim typPoint As POINTL
    ClientToScreen UserControl.hwnd, typPoint
    GetCursorPos typPoint
    If fDownWard = True Then
         typPoint.x = typPoint.x + 3
         typPoint.y = typPoint.y + 3
         fDownWard = False
    Else
         typPoint.x = typPoint.x - 3
         typPoint.y = typPoint.y - 3
         fDownWard = True
    End If
    SetCursorPos typPoint.x, typPoint.y
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,HasDC
Friend Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determines whether a unique display context is allocated for the control."
    HasDC = UserControl.HasDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Friend Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

Private Sub UserControl_Resize()
    Dim oldMode As Long, oldAutoRedraw As Long
    Set objContainer = Extender.Container
    oldMode = objContainer.ScaleMode
    objContainer.ScaleMode = vbPixels
    oldAutoRedraw = objContainer.AutoRedraw
    objContainer.AutoRedraw = True

    DoEvents
    If Not fIsShown Then
        UserControl_Show
        DoEvents
    End If
    If fHavePicture Then 'set the size if picture present so it isn't clipped
        If AutoSize Or UserControl.Width < PicWidth + 5 Or _
                        UserControl.Height < PicHeight + 5 Then
            Size (PicWidth + EdgeWidth * 5) * Screen.TwipsPerPixelX, (PicHeight + EdgeWidth * 5) * Screen.TwipsPerPixelY
        Else
            
        End If
    End If
    If UserControl.Height < 5 Or UserControl.Width < 5 Then
        UserControl.Height = 5
        If m_Caption$ <> sEmpty$ Then
            UserControl.Width = iTextWidthRet + 4
        Else
            UserControl.Width = 5
        End If
    End If

    exTop = Extender.Top
    exLeft = Extender.Left
    exWidth = Extender.Width
    exHeight = Extender.Height
    bW = UserControl.ScaleWidth
    bH = UserControl.ScaleHeight
    bL = UserControl.ScaleLeft
    bT = UserControl.ScaleTop
    
    CenterTheCaption
    
    InitButton

    objContainer.ScaleMode = oldMode
    objContainer.AutoRedraw = oldAutoRedraw
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Size
Public Sub Size(ByVal Width As Single, ByVal Height As Single)
Attribute Size.VB_Description = "Changes the width and height of a User Control."
    UserControl.Size Width, Height
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    ToolTipShowDelay = PropBag.ReadProperty("ToolTipShowDelay", m_def_ToolTipShowDelay)
    ToolTipTimeShown = PropBag.ReadProperty("ToolTipTimeShown", m_def_ToolTipTimeShown)
    ToolTipFontColor = PropBag.ReadProperty("ToolTipFontColor", m_def_ToolTipFontColor)
    ToolTipBackColor = PropBag.ReadProperty("ToolTipBackColor", m_def_ToolTipBackColor)
    ToolTipMaxWidth = PropBag.ReadProperty("ToolTipMaxWidth", m_def_ToolTipMaxWidth)
    ToolTipText = PropBag.ReadProperty("ToolTipText", Extender.ToolTipText)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    FontShadowLeft = PropBag.ReadProperty("FontShadowLeft", m_def_FontShadowLeft)
    FontShadowRight = PropBag.ReadProperty("FontShadowRight", m_def_FontShadowRight)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    FontROP = PropBag.ReadProperty("FontROP", m_def_FontROP)
    FontTransParent = PropBag.ReadProperty("FontTransParent", m_def_FontTransParent)
    m_FontUsePen = PropBag.ReadProperty("FontUsePen", m_def_FontUsePen)
    m_FontColorPen = PropBag.ReadProperty("FontColorPen", m_def_FontColorPen)
    m_FontColorBrush = PropBag.ReadProperty("FontColorBrush", m_def_FontColorBrush)
    m_FontUseBrush = PropBag.ReadProperty("FontUseBrush", m_def_FontUseBrush)
    m_FontOrientation = PropBag.ReadProperty("FontOrientation", m_def_FontOrientation)
    m_FontEscapement = PropBag.ReadProperty("FontEscapement", m_def_FontEscapement)
    m_FontShadowOffsetX = PropBag.ReadProperty("FontShadowOffsetX", m_def_FontShadowOffsetX)
    m_FontShadowOffsetY = PropBag.ReadProperty("FontShadowOffsetY", m_def_FontShadowOffsetY)
    m_FontShadow = PropBag.ReadProperty("FontShadow", m_def_FontShadow)
    m_FontCharSpacing = PropBag.ReadProperty("FontCharSpacing", m_def_FontCharSpacing)
    CaptionCenterX = PropBag.ReadProperty("CaptionCenterX", m_def_CaptionCenterX)
    CaptionCenterY = PropBag.ReadProperty("CaptionCenterY", m_def_CaptionCenterY)
    CaptionLeft = PropBag.ReadProperty("CaptionLeft", m_def_CaptionLeft)
    CaptionTop = PropBag.ReadProperty("CaptionTop", m_def_CaptionTop)
    Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
    EdgeWidth = PropBag.ReadProperty("EdgeWidth", m_def_EdgeWidth)
    AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set PictureUp = PropBag.ReadProperty("PictureUp", Nothing)
    Set PictureDown = PropBag.ReadProperty("PictureDown", Nothing)
    PictureOffsetX = PropBag.ReadProperty("PictureOffsetX", m_def_PictureOffsetX)
    PictureOffsetY = PropBag.ReadProperty("PictureOffsetY", m_def_PictureOffsetY)
    PicDownMove = PropBag.ReadProperty("PicDownMove", m_def_PicDownMove)
    MaskColor = PropBag.ReadProperty("MaskColor", m_def_MaskColor)
    MaskColorUpPic = PropBag.ReadProperty("MaskColorUpPic", m_def_MaskColorUpPic)
    MaskColorDownPic = PropBag.ReadProperty("MaskColorDownPic", m_def_MaskColorDownPic)
    BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_MouseRightClickEnable = PropBag.ReadProperty("MouseRightClickEnable", m_def_MouseRightClickEnable)
    EdgeType = PropBag.ReadProperty("EdgeType", m_def_EdgeType)
    m_FontPathAntiAlias = PropBag.ReadProperty("FontPathAntiAlias", m_def_FontPathAntiAlias)
    m_FontPenWidth = PropBag.ReadProperty("FontPenWidth", m_def_FontPenWidth)
End Sub

Private Sub UserControl_Terminate()
    DestroyDnDC
    DestroyUpDC
    DestroyDC
    DestroyBase
    DelToolTip
    If hTT Then DestroyWindow (hTT)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("FontShadowLeft", m_FontShadowLeft, m_def_FontShadowLeft)
    Call PropBag.WriteProperty("FontShadowRight", m_FontShadowRight, m_def_FontShadowRight)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
    Call PropBag.WriteProperty("EdgeWidth", m_EdgeWidth, m_def_EdgeWidth)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, m_def_MaskColor)
    Call PropBag.WriteProperty("ToolTipText", Extender.ToolTipText, "")
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("PictureUp", m_PictureUp, Nothing)
    Call PropBag.WriteProperty("PictureDown", m_PictureDown, Nothing)
    Call PropBag.WriteProperty("CaptionLeft", m_CaptionLeft, m_def_CaptionLeft)
    Call PropBag.WriteProperty("CaptionTop", m_CaptionTop, m_def_CaptionTop)
    Call PropBag.WriteProperty("PictureOffsetX", m_PictureOffsetX, m_def_PictureOffsetX)
    Call PropBag.WriteProperty("PictureOffsetY", m_PictureOffsetY, m_def_PictureOffsetY)
    Call PropBag.WriteProperty("ToolTipShowDelay", m_ToolTipShowDelay, m_def_ToolTipShowDelay)
    Call PropBag.WriteProperty("ToolTipTimeShown", m_ToolTipTimeShown, m_def_ToolTipTimeShown)
    Call PropBag.WriteProperty("PicDownMove", m_PicDownMove, m_def_PicDownMove)
    Call PropBag.WriteProperty("MaskColorUpPic", m_MaskColorUpPic, m_def_MaskColorUpPic)
    Call PropBag.WriteProperty("MaskColorDownPic", m_MaskColorDownPic, m_def_MaskColorDownPic)
    Call PropBag.WriteProperty("CaptionCenterX", m_CaptionCenterX, m_def_CaptionCenterX)
    Call PropBag.WriteProperty("CaptionCenterY", m_CaptionCenterY, m_def_CaptionCenterY)
    Call PropBag.WriteProperty("ToolTipMaxWidth", m_ToolTipMaxWidth, m_def_ToolTipMaxWidth)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    'Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("ToolTipFontColor", m_ToolTipFontColor, m_def_ToolTipFontColor)
    Call PropBag.WriteProperty("ToolTipBackColor", m_ToolTipBackColor, m_def_ToolTipBackColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("FontUsePen", m_FontUsePen, m_def_FontUsePen)
    Call PropBag.WriteProperty("FontColorPen", m_FontColorPen, m_def_FontColorPen)
    Call PropBag.WriteProperty("FontColorBrush", m_FontColorBrush, m_def_FontColorBrush)
    Call PropBag.WriteProperty("FontUseBrush", m_FontUseBrush, m_def_FontUseBrush)
    Call PropBag.WriteProperty("FontOrientation", m_FontOrientation, m_def_FontOrientation)
    Call PropBag.WriteProperty("FontEscapement", m_FontEscapement, m_def_FontEscapement)
    Call PropBag.WriteProperty("FontShadowOffsetX", m_FontShadowOffsetX, m_def_FontShadowOffsetX)
    Call PropBag.WriteProperty("FontShadowOffsetY", m_FontShadowOffsetY, m_def_FontShadowOffsetY)
    Call PropBag.WriteProperty("FontShadow", m_FontShadow, m_def_FontShadow)
    Call PropBag.WriteProperty("FontCharSpacing", m_FontCharSpacing, m_def_FontCharSpacing)
    Call PropBag.WriteProperty("FontTransParent", m_FontTransParent, m_def_FontTransParent)
    Call PropBag.WriteProperty("FontROP", m_FontROP, m_def_FontROP)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("MouseRightClickEnable", m_MouseRightClickEnable, m_def_MouseRightClickEnable)
    Call PropBag.WriteProperty("EdgeType", m_EdgeType, m_def_EdgeType)
    Call PropBag.WriteProperty("FontPathAntiAlias", m_FontPathAntiAlias, m_def_FontPathAntiAlias)
    Call PropBag.WriteProperty("FontPenWidth", m_FontPenWidth, m_def_FontPenWidth)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_FontColor = m_def_FontColor
    m_BackColor = m_def_BackColor
    m_ButtonStyle = m_def_ButtonStyle
    m_EdgeWidth = m_def_EdgeWidth
    m_AutoSize = m_def_AutoSize
    m_MaskColor = m_def_MaskColor
    Set m_Picture = LoadPicture("")
    Set m_PictureUp = LoadPicture("")
    Set m_PictureDown = LoadPicture("")
    m_CaptionLeft = m_def_CaptionLeft
    m_CaptionTop = m_def_CaptionTop
    m_PictureOffsetX = m_def_PictureOffsetX
    m_PictureOffsetY = m_def_PictureOffsetY
    m_ToolTipShowDelay = m_def_ToolTipShowDelay
    m_ToolTipTimeShown = m_def_ToolTipTimeShown
    m_PicDownMove = m_def_PicDownMove
    m_MaskColorUpPic = vbWhite
    m_MaskColorDownPic = vbBlack
    m_CaptionCenterX = True
    m_CaptionCenterY = True
    UserControl.ScaleMode = vbPixels
    m_ToolTipMaxWidth = m_def_ToolTipMaxWidth
    m_FontROP = m_def_FontROP
    m_FontShadowLeft = m_def_FontShadowLeft
    m_FontShadowRight = m_def_FontShadowRight
    m_FontTransParent = m_def_FontTransParent
    m_ToolTipText = Extender.ToolTipText
    m_ToolTipFontColor = m_def_ToolTipFontColor
    m_ToolTipBackColor = m_def_ToolTipBackColor
    m_Caption = m_def_Caption
    m_FontUsePen = m_def_FontUsePen
    m_FontColorPen = m_def_FontColorPen
    m_FontColorBrush = m_def_FontColorBrush
    m_FontUseBrush = m_def_FontUseBrush
    m_FontOrientation = m_def_FontOrientation
    m_FontEscapement = m_def_FontEscapement
    m_FontShadowOffsetX = m_def_FontShadowOffsetX
    m_FontShadowOffsetY = m_def_FontShadowOffsetY
    m_FontShadow = m_def_FontShadow
    m_FontCharSpacing = m_def_FontCharSpacing
    m_BackStyle = m_def_BackStyle
    m_MouseRightClickEnable = m_def_MouseRightClickEnable
    m_EdgeType = m_def_EdgeType
    m_FontPathAntiAlias = m_def_FontPathAntiAlias
    m_FontPenWidth = m_def_FontPenWidth
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,1
Public Property Get ButtonStyle() As EEButtonStyle
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As EEButtonStyle)
    m_ButtonStyle = New_ButtonStyle
    If fIsShown Then DrawButton
    PropertyChanged "ButtonStyle"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get EdgeWidth() As Long
    EdgeWidth = m_EdgeWidth
End Property

Public Property Let EdgeWidth(ByVal New_EdgeWidth As Long)
    m_EdgeWidth = New_EdgeWidth
    PropertyChanged "EdgeWidth"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    UserControl_Resize
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PictureOffsetX() As Long
    PictureOffsetX = m_PictureOffsetX
End Property

Public Property Let PictureOffsetX(ByVal New_PictureOffsetX As Long)
    m_PictureOffsetX = New_PictureOffsetX
    If fIsShown Then InitButton
    PropertyChanged "PictureOffsetX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get PictureOffsetY() As Long
    PictureOffsetY = m_PictureOffsetY
End Property

Public Property Let PictureOffsetY(ByVal New_PictureOffsetY As Long)
    m_PictureOffsetY = New_PictureOffsetY
    If fIsShown Then InitButton
    PropertyChanged "PictureOffsetY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
    ToolTipText = Extender.ToolTipText ' m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    Extender.ToolTipText = New_ToolTipText
    If New_ToolTipText = sEmpty Then
        DelToolTip
    Else
        SetToolTipLabel
    End If
    PropertyChanged "ToolTipText"
End Property

Private Sub SetToolTipLabel()
'Dissected tooltip stuff from CustomToolTips -  http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=9185&lngWId=1
    With TI
        .hwnd = UserControl.hwnd
        .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
'        If bCenter Then
'            .uFlags = .uFlags Or TTF_CENTERTIP
'        End If
        .uId = UserControl.hwnd
        .lpszText = m_ToolTipText
        .cbSize = Len(TI)
    End With
    'SendMessageLong hTT, WM_SETFONT, usercontrol.font, 1&
    SendMessageAsAny hTT, TTM_ADDTOOL, 0, TI
End Sub
Public Sub DelToolTip()
   Dim TI As TOOLINFO
   TI.hwnd = UserControl.hwnd
   TI.cbSize = Len(TI)
   TI.uId = UserControl.hwnd
   SendMessageAsAny hTT, TTM_DELTOOL, 0, TI
   'hTT = 0
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,
Public Property Get ToolTipFontColor() As OLE_COLOR
    ToolTipFontColor = m_ToolTipFontColor
End Property

Public Property Let ToolTipFontColor(ByVal New_ToolTipFontColor As OLE_COLOR)
    m_ToolTipFontColor = New_ToolTipFontColor
    SendMessageLong hTT, TTM_SETTIPTEXTCOLOR, New_ToolTipFontColor, 0&
    PropertyChanged "ToolTipFontColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,
Public Property Get ToolTipBackColor() As OLE_COLOR
    ToolTipBackColor = m_ToolTipBackColor
End Property

Public Property Let ToolTipBackColor(ByVal New_ToolTipBackColor As OLE_COLOR)
    m_ToolTipBackColor = New_ToolTipBackColor
    SendMessageLong hTT, TTM_SETTIPBKCOLOR, New_ToolTipBackColor, 0&
    PropertyChanged "ToolTipBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,2000
Public Property Get ToolTipShowDelay() As Long
    ToolTipShowDelay = m_ToolTipShowDelay
End Property

Public Property Let ToolTipShowDelay(ByVal New_ToolTipShowDelay As Long)
    m_ToolTipShowDelay = New_ToolTipShowDelay
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_INITIAL, New_ToolTipShowDelay
    PropertyChanged "ToolTipShowDelay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,2000
Public Property Get ToolTipTimeShown() As Long
    ToolTipTimeShown = m_ToolTipTimeShown
End Property

Public Property Let ToolTipTimeShown(ByVal New_ToolTipTimeShown As Long)
    m_ToolTipTimeShown = New_ToolTipTimeShown
    SendMessageLong hTT, TTM_SETDELAYTIME, TTDT_AUTOPOP, New_ToolTipTimeShown
    PropertyChanged "ToolTipTimeShown"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,100
Public Property Get ToolTipMaxWidth() As Long
    ToolTipMaxWidth = m_ToolTipMaxWidth
End Property

Public Property Let ToolTipMaxWidth(ByVal New_ToolTipMaxWidth As Long)
    m_ToolTipMaxWidth = New_ToolTipMaxWidth
    SendMessageLong hTT, TTM_SETMAXTIPWIDTH, 0, New_ToolTipMaxWidth
    PropertyChanged "ToolTipMaxWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get PicDownMove() As Boolean
    PicDownMove = m_PicDownMove
End Property

Public Property Let PicDownMove(ByVal New_PicDownMove As Boolean)
    m_PicDownMove = New_PicDownMove
    PropertyChanged "PicDownMove"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    If fIsShown Then InitButton
    PropertyChanged "MaskColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get MaskColorUpPic() As OLE_COLOR
    MaskColorUpPic = m_MaskColorUpPic
End Property

Public Property Let MaskColorUpPic(ByVal New_MaskColorUpPic As OLE_COLOR)
    m_MaskColorUpPic = New_MaskColorUpPic
    If fIsShown Then InitButton
    PropertyChanged "MaskColorUpPic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get MaskColorDownPic() As OLE_COLOR
    MaskColorDownPic = m_MaskColorDownPic
End Property

Public Property Let MaskColorDownPic(ByVal New_MaskColorDownPic As OLE_COLOR)
    m_MaskColorDownPic = New_MaskColorDownPic
    If fIsShown Then InitButton
    PropertyChanged "MaskColorDownPic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    CenterTheCaption
    If fIsShown Then DrawButton
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionLeft() As Long
    CaptionLeft = m_CaptionLeft
End Property

Public Property Let CaptionLeft(ByVal New_CaptionLeft As Long)
    m_CaptionLeft = New_CaptionLeft
    If fIsShown Then DrawButton
    PropertyChanged "CaptionLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get CaptionTop() As Long
    CaptionTop = m_CaptionTop
End Property

Public Property Let CaptionTop(ByVal New_CaptionTop As Long)
    m_CaptionTop = New_CaptionTop
    If fIsShown Then DrawButton
    PropertyChanged "CaptionTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CaptionCenterX() As Boolean
    CaptionCenterX = m_CaptionCenterX
End Property

Public Property Let CaptionCenterX(ByVal New_CaptionCenterX As Boolean)
    m_CaptionCenterX = New_CaptionCenterX
    If m_CaptionCenterX Then
        CaptionLeft = (UserControl.ScaleWidth - iTextWidthRet) \ 2
    End If
    If fIsShown Then DrawButton
    PropertyChanged "CaptionCenterX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CaptionCenterY() As Boolean
    CaptionCenterY = m_CaptionCenterY
End Property

Public Property Let CaptionCenterY(ByVal New_CaptionCenterY As Boolean)
    m_CaptionCenterY = New_CaptionCenterY
    If m_CaptionCenterY Then
        CaptionTop = (UserControl.ScaleHeight - iTextHeightRet) \ 2
    End If
    If fIsShown Then DrawButton
    PropertyChanged "CaptionCenterY"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    If fIsShown Then InitButton
    PropertyChanged "BackColor"
End Property
Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property
Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    If fIsShown Then DrawButton
    PropertyChanged "FontColor"
End Property
Public Property Get FontShadowLeft() As OLE_COLOR
    FontShadowLeft = m_FontShadowLeft
End Property
Public Property Let FontShadowLeft(ByVal New_FontShadowLeft As OLE_COLOR)
    m_FontShadowLeft = New_FontShadowLeft
    If fIsShown Then DrawButton
    PropertyChanged "FontShadowLeft"
End Property
Public Property Get FontShadowRight() As OLE_COLOR
    FontShadowRight = m_FontShadowRight
End Property
Public Property Let FontShadowRight(ByVal New_FontShadowRight As OLE_COLOR)
    m_FontShadowRight = New_FontShadowRight
    If fIsShown Then DrawButton
    PropertyChanged "FontShadowRight"
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    CenterTheCaption
    If fIsShown Then DrawButton
    PropertyChanged "Font"
End Property

Public Property Get FontTransParent() As Boolean
    FontTransParent = m_FontTransParent
End Property

Public Property Let FontTransParent(ByVal fFontTransParentA As Boolean)
    m_FontTransParent = fFontTransParentA
    If fIsShown Then DrawButton
    PropertyChanged "FontTransParent"
End Property

Public Property Get FontROP() As DrawModeConstants
    FontROP = m_FontROP
End Property

Public Property Let FontROP(ByVal iFontROPA As DrawModeConstants)
    m_FontROP = iFontROPA
    If fIsShown Then DrawButton
    PropertyChanged "FontROP"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontUsePen() As Boolean
    FontUsePen = m_FontUsePen
End Property

Public Property Let FontUsePen(ByVal New_FontUsePen As Boolean)
    m_FontUsePen = New_FontUsePen
    If fIsShown Then DrawButton
    PropertyChanged "FontUsePen"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbred
Public Property Get FontColorPen() As OLE_COLOR
    FontColorPen = m_FontColorPen
End Property

Public Property Let FontColorPen(ByVal New_FontColorPen As OLE_COLOR)
    m_FontColorPen = New_FontColorPen
    If fIsShown Then DrawButton
    PropertyChanged "FontColorPen"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblue
Public Property Get FontColorBrush() As OLE_COLOR
    FontColorBrush = m_FontColorBrush
End Property

Public Property Let FontColorBrush(ByVal New_FontColorBrush As OLE_COLOR)
    m_FontColorBrush = New_FontColorBrush
    If fIsShown Then DrawButton
    PropertyChanged "FontColorBrush"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FontUseBrush() As Boolean
    FontUseBrush = m_FontUseBrush
End Property

Public Property Let FontUseBrush(ByVal New_FontUseBrush As Boolean)
    m_FontUseBrush = New_FontUseBrush
    If fIsShown Then DrawButton
    PropertyChanged "FontUseBrush"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontOrientation() As Long
    FontOrientation = m_FontOrientation
End Property

Public Property Let FontOrientation(ByVal New_FontOrientation As Long)
    m_FontOrientation = New_FontOrientation
    CenterTheCaption
    If fIsShown Then DrawButton
    PropertyChanged "FontOrientation"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontEscapement() As Long
    FontEscapement = m_FontEscapement
End Property

Public Property Let FontEscapement(ByVal New_FontEscapement As Long)
    m_FontEscapement = New_FontEscapement
    CenterTheCaption
    If fIsShown Then DrawButton
    PropertyChanged "FontEscapement"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontShadowOffsetX() As Long
    FontShadowOffsetX = m_FontShadowOffsetX
End Property

Public Property Let FontShadowOffsetX(ByVal New_FontShadowOffsetX As Long)
    m_FontShadowOffsetX = New_FontShadowOffsetX
    If fIsShown Then DrawButton
    PropertyChanged "FontShadowOffsetX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontShadowOffsetY() As Long
    FontShadowOffsetY = m_FontShadowOffsetY
End Property

Public Property Let FontShadowOffsetY(ByVal New_FontShadowOffsetY As Long)
    m_FontShadowOffsetY = New_FontShadowOffsetY
    If fIsShown Then DrawButton
    PropertyChanged "FontShadowOffsetY"
End Property


'Public Property Get lHatchStyle() As Long
'    lHatchStyle = ilHatchStyle
'End Property
'
'Public Property Let lHatchStyle(ByVal ilHatchStyleA As Long)
'    ilHatchStyle = ilHatchStyleA
'End Property
'Public Property Get lBrushStyle() As Long
'    lBrushStyle = ilBrushStyle
'End Property
'
'Public Property Let lBrushStyle(ByVal ilBrushStyleA As Long)
'    ilBrushStyle = ilBrushStyleA
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontShadow() As EEFontShadow
    FontShadow = m_FontShadow
End Property

Public Property Let FontShadow(ByVal New_FontShadow As EEFontShadow)
    m_FontShadow = New_FontShadow
    If fIsShown Then DrawButton
    PropertyChanged "FontShadow"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FontCharSpacing() As Long
    FontCharSpacing = m_FontCharSpacing
End Property

Public Property Let FontCharSpacing(ByVal New_FontCharSpacing As Long)
    m_FontCharSpacing = New_FontCharSpacing
    CenterTheCaption
    If fIsShown Then DrawButton
    PropertyChanged "FontCharSpacing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackStyle() As eeBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As eeBackStyle)
    m_BackStyle = New_BackStyle
    If fIsShown Then InitButton
    PropertyChanged "BackStyle"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Let MouseRightClickEnable(ByVal New_MouseRightClickEnable As Boolean)
    m_MouseRightClickEnable = New_MouseRightClickEnable
    PropertyChanged "MouseRightClickEnable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get EdgeType() As EEEdgeStyle
    EdgeType = m_EdgeType
End Property

Public Property Let EdgeType(ByVal New_EdgeType As EEEdgeStyle)
    m_EdgeType = New_EdgeType
    'ext-styles snippet from C++ extended styles vbp...no author given...
    iWinStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
    If m_EdgeType = eeClient Then
        iWinStyle = iWinStyle Or WS_EX_CLIENTEDGE And Not WS_EX_DLGMODALFRAME
    ElseIf m_EdgeType = eeModal Then
        iWinStyle = iWinStyle Or WS_EX_DLGMODALFRAME And Not WS_EX_CLIENTEDGE
    Else
        iWinStyle = iWinStyle And Not WS_EX_CLIENTEDGE And Not WS_EX_DLGMODALFRAME
    End If
    SetWindowLong UserControl.hwnd, GWL_EXSTYLE, iWinStyle
    SetWindowPos UserControl.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    If fIsShown Then InitButton
    PropertyChanged "EdgeType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
'Public Property Get FontPathAntiAlias() As Boolean
'    FontPathAntiAlias = m_FontPathAntiAlias
'End Property
'
'Public Property Let FontPathAntiAlias(ByVal New_FontPathAntiAlias As Boolean)
'    m_FontPathAntiAlias = New_FontPathAntiAlias
'    if fIsShown then DrawButton
'    PropertyChanged "FontPathAntiAlias"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get FontPenWidth() As Long
    FontPenWidth = m_FontPenWidth
End Property

Public Property Let FontPenWidth(ByVal New_FontPenWidth As Long)
    m_FontPenWidth = New_FontPenWidth
    If fIsShown Then DrawButton
    PropertyChanged "FontPenWidth"
End Property

