VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "cDropShadowBlt test"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCLS 
      Caption         =   "&CLS"
      Height          =   435
      Left            =   2175
      TabIndex        =   0
      Top             =   3045
      Width           =   1005
   End
   Begin VB.CommandButton cmdPaint 
      Caption         =   "&Paint"
      Height          =   435
      Left            =   3315
      TabIndex        =   1
      Top             =   3045
      Width           =   1005
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cDropShadow As New cDropShadowBlt

Private Sub Form_Load()
    
    Call m_cDropShadow.CreateFromHandle( _
         hImage:=LoadResPicture(101, vbResIcon), _
         MaskColor:=vbMagenta, _
         ShadowColor:=vbBlack, _
         Opacity:=50, _
         xOffset:=3, _
         yOffset:=3 _
         )
End Sub

Private Sub cmdCLS_Click()
    
    Call Me.Cls
End Sub

Private Sub cmdPaint_Click()
    
  Dim i As Long
  Dim j As Long
    
    For i = 0 To 5
        For j = 0 To 3
            Call m_cDropShadow.Paint(Me.hDC, i * 50 + 5, j * 50 + 5)
        Next j
    Next i
End Sub
