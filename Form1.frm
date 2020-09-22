VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMrAnimator 
      Interval        =   100
      Left            =   120
      Top             =   1080
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "Right"
      Height          =   1215
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin PicClip.PictureClip pctRight 
      Left            =   120
      Top             =   600
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "left"
      Height          =   1215
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin PicClip.PictureClip pctLeft 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A "How to: code" for beginners, By Atilla Duran
' -----------------------------------------------
'
' An example of How to animate a bitmap inside a button
' You can animate where ever you want to, my example is in a button.
'
'
' Do you have any questions, just mail me at atilla@oslo.online.no
'
Private Sub Form_Load()

    '----- Set the PictureClip objects
        
        'Load pictures
        pctLeft.Picture = LoadPicture(App.Path & "\left.bmp")
        pctRight.Picture = LoadPicture(App.Path & "\right.bmp")
    
        'Set Cols and Rows
        pctLeft.Cols = 2
        pctLeft.Rows = 2
        pctRight.Cols = 2
        pctRight.Rows = 2
    
       'Enable the timer, tmrMrAnimator
       tmrMrAnimator.Enabled = True
       
       'Set timers speer interval 100 = 0.10 sec
       tmrMrAnimator.Interval = 100
       
End Sub

Private Sub tmrMrAnimator_Timer()

    Static MrLooper As Byte
        
        'If counter is the last image then counter sets to first image
        'Else, just carry on counting
        If MrLooper = 3 Then
            MrLooper = 0
          Else
            MrLooper = MrLooper + 1
        End If
        
        'Load images into buttons
        cmdLeft.Picture = pctLeft.GraphicCell(MrLooper)
        cmdRight.Picture = pctRight.GraphicCell(MrLooper)

End Sub
