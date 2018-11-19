Attribute VB_Name = "JConstDHL"
Option Explicit
Option Private Module

'/**
' * Colors
' *
' */
Public Const clrDHLRed As Long = 1115604        ' RGB(212, 5, 17)
Public Const clrDHLYellow As Long = 52479       ' RGB(255, 204, 0)
' (light to dark)
Public Const clrDHLYellow1 As Long = 13432319   ' RGB(255, 245, 204)
Public Const clrDHLYellow2 As Long = 9234687    ' RGB(255, 232, 140)
Public Const clrDHLYellow3 As Long = 5037055    ' RGB(255, 219, 76)
' (light to dark)
Public Const clrDHLGrey1 As Long = 15000804     ' RGB(228, 228, 228)
Public Const clrDHLGrey2 As Long = 12566463     ' RGB(191, 191, 191)
Public Const clrDHLGrey3 As Long = 10066329     ' RGB(153, 153, 153)
Public Const clrDHLGrey4 As Long = 6710886      ' RGB(102, 102, 102)
' (complementary colors)
Public Const clrDHLLightBlue As Long = 15915720 ' RGB(200, 218, 242)
Public Const clrDHLBlue As Long = 13204482      ' RGB(2, 124, 201)
Public Const clrDHLGreen As Long = 377614
' (fade)
Public Const clrDHLFadeBackground As Long = clrDHLGrey1
Public Const clrDHLFadeForeground As Long = clrDHLGrey4
' (Table style)
Public Const clrDHLViewBorderHorizontal As Long = clrDHLGrey2   ' Top & inner horizontal borders (lights)
Public Const clrDHLViewBorderBottom As Long = clrDHLGrey4       ' Header & table bottom borders (shadows)
Public Const clrDHLViewBorderVertical As Long = clrDHLGrey3     ' All vertical borders
Public Const clrDHLViewGradientStart As Long = clrDHLGrey1      ' Header gradient start color
Public Const clrDHLViewGradientEnd As Long = clrDHLGrey2        ' Header gradient end color

'/**
' * Paths
' *
' */
Public Const strPathToDHLQEFFolder As String = "I:\05 - Training Materials\03 - QEF formuláøe\QEF formuláøe\QEF formuláøe"
