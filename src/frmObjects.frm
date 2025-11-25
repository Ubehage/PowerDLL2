VERSION 5.00
Begin VB.Form frmObjects 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNothing 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2370
      Picture         =   "frmObjects.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1170
      Width           =   540
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   330
      Width           =   600
   End
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      Height          =   345
      Left            =   675
      ScaleHeight     =   285
      ScaleWidth      =   510
      TabIndex        =   0
      Top             =   240
      Width           =   570
   End
End
Attribute VB_Name = "frmObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Friend Sub SizeIconPicture(Small As Boolean)
  Dim iWidth As Long
  Dim iHeight As Long
  If Small Then
    iWidth = (Screen.TwipsPerPixelX * 16)
    iHeight = (Screen.TwipsPerPixelY * 16)
  Else
    iWidth = (Screen.TwipsPerPixelX * 32)
    iHeight = (Screen.TwipsPerPixelY * 32)
  End If
  With picIcon
    .Width = ((.Width - .ScaleWidth) + iWidth)
    .Height = ((.Height - .ScaleHeight) + iHeight)
  End With
End Sub
