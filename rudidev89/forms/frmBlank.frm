VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBlank 
   Caption         =   "Blank"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11910
   OleObjectBlob   =   "frmBlank.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
    AddMinMaxButton Me.Caption, MinButton:=True, MaxButton:=True
End Sub

Private Sub UserForm_Initialize()
    Call UserForm_Resize
End Sub

Private Sub UserForm_Resize()

    Img_head.Top = 0
    Img_head.Left = 0
    Img_head.Width = Me.Width
    
    Img_Logo.Top = 5
    Img_Logo.Left = 40
    
    frame_status.Top = 5
    frame_status.Left = Me.Width - 140
    frame_status.BackColor = Img_head.BackColor
    
    frame_Control.Top = 54
    frame_Control.Left = Me.Width - 58
    frame_Control.Height = Me.Height - 100
    
    frame_Container.Top = 54
    frame_Container.Left = 5
    frame_Container.Height = Me.Height - 100
    frame_Container.Width = Me.Width - 68
    
    Label_Status.Left = 0
    Label_Status.Width = Me.Width - Cmb_me.Width
    Label_Status.Top = Me.Height - 42
    
    Cmb_me.Top = Label_Status.Top
    Cmb_me.Left = Label_Status.Width
    
End Sub

Private Sub CBReFresh_Click()

End Sub

Private Sub cmdBaru_Click()

End Sub

Private Sub cmdBatal_Click()
End Sub

Private Sub cmdCari_Click()

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()

End Sub

Private Sub cmdHapus_Click()

End Sub

Private Sub cmdLihat_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSimpan_Click()

End Sub


