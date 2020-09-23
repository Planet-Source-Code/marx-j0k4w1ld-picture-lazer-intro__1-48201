VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   0
      Picture         =   "frmIntro.frx":0000
      ScaleHeight     =   3534.566
      ScaleMode       =   0  'User
      ScaleWidth      =   11312.52
      TabIndex        =   0
      Top             =   -240
      Visible         =   0   'False
      Width           =   11535
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   240
         Top             =   1440
      End
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

     Dim Proba, Proba2 As Integer
     Dim Boja2 As String

Private Function Zrak(slika As PictureBox, StartX As Integer, StartY As Integer, Levo As Integer, Desno As Integer, Boja As String)

     

Me.ScaleMode = vbPixels

With slika
     .ScaleMode = vbPixels
     .AutoRedraw = True
End With

For Proba2 = 0 To slika.ScaleWidth
    DoEvents

For Proba = 0 To slika.ScaleHeight
    Boja2 = slika.Point(Proba2, Proba)
   Line (StartX, StartY)-(Levo + Proba2, Desno + Proba), Boja2
Next
   Line (StartX, StartY)-(Levo + Proba2, Desno + slika.ScaleHeight), Boja
Next

For Proba2 = 0 To slika.ScaleHeight
   Line (StartX, StartY)-(Levo + slika.ScaleWidth, Desno + Proba2), Boja
Next

End Function



Private Sub Timer1_Timer()
Zrak Picture1, 565, 301, 0, 0, Me.BackColor    ' adjust the scale and position from where
                                                 ' laser will draw  your image
                                                ' best performance is if you place start position
                                                ' in down right corner of your picture
                                                'Also  adjust the size of the form , to be same
                                                ' to the size   of the picture , or perform some variation.
   If Timer1.interval = 1000 Then
'    frmAdministrator.Show    , if you are going further with your program
    Unload Me                  ' you can replace timer with command_click()
    
End If
End Sub
                                             
                                                

