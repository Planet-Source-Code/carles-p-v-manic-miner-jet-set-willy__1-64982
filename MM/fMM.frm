VERSION 5.00
Begin VB.Form fMM 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MANIC MINER"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "fMM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    If (App.LogMode <> 1) Then
        Call VBA.MsgBox("Please, compile me!", vbExclamation)
        End
    End If

    Call Me.Show                    ' Show window
    Call VBA.DoEvents
    
    Call InitializeFullScreen(Me)   ' Initialize full-screen module
    Call mMM.Initialize(Me)         ' Initialize MM
    Call mMM.StartGame              ' Start framing (loop)
    Call mMM.Terminate              ' Close MM (after mMM.StopGame)
    
    Call VB.Unload(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call mMM.StopGame               ' Stop framing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fMM = Nothing
End Sub

'========================================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mMM.KeyDown(KeyCode)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Call mMM.KeyUp(KeyCode)
End Sub

'========================================================================================

Private Sub Form_Paint()
    Call mMM.ScreenUpdate
End Sub

