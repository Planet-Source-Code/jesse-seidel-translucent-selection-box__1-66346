VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translucent Selection - By Jesse Seidel (DoctorFire)"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   4680
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   56
         Y1              =   32
         Y2              =   32
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   56
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   56
         X2              =   56
         Y1              =   32
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   32
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

 Private Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean
  

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type
  
 Dim down As Boolean
 Dim starty As Long
 Dim startx As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Picture1.Visible = True
    Picture1.Top = Y 'Set
    Picture1.Left = X
    starty = Y
    startx = X
    down = True 'Notify other subs that the mouse is down

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim tProperties As typeBlendProperties
    Dim lngBlend As Long
    
    tProperties.tBlendAmount = 255 - 65 'Set translucency level
    
    If down = True Then 'Checks weather the mouse is down so it doesn't resize when it doesn't need to
    
        If Y < starty Then
            Picture1.Top = Y
            Picture1.Height = starty - Y
            
            If X < startx Then
                Picture1.Left = X
                Picture1.Width = startx - X
            End If
            
        Else
            Picture1.Height = Y - Picture1.Top
            Picture1.Width = X - Picture1.Left
        End If
        
        'The following is for the outline
        Line1.Y1 = Picture1.ScaleHeight
        Line2.Y1 = Line1.Y1
        Line2.X1 = Picture1.ScaleWidth - 1
        Line2.X2 = Picture1.ScaleWidth - 1
        Line3.X2 = Picture1.ScaleWidth - 1
        Line4.X2 = Line3.X2
        Line4.Y1 = Line1.Y1 - 1
        Line4.Y2 = Line1.Y1 - 1
        'End of outline sizing
        
        CopyMemory lngBlend, tProperties, 4 'Blend colors
        Picture1.Cls 'Clear
        AlphaBlend Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hDC, Picture1.Left / Screen.TwipsPerPixelX, Picture1.Top / Screen.TwipsPerPixelY, Picture1.ScaleWidth, Picture1.ScaleHeight, lngBlend 'Blend together
        Picture1.Refresh 'Reduce flicker
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Picture1.Visible = False 'Hide Picture1
    down = False 'Make sure Picture1 doesn't resize when it doesn't need to
End Sub

Private Sub Form_Unload(Cancel As Integer)
 End
End Sub

