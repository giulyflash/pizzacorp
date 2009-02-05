VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pizza R Us"
   ClientHeight    =   2910
   ClientLeft      =   3465
   ClientTop       =   4680
   ClientWidth     =   7215
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame fraDivider 
      Height          =   30
      Left            =   -120
      TabIndex        =   5
      Top             =   1560
      Width           =   7575
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   0
      Picture         =   "frmIntro.frx":08CA
      ScaleHeight     =   1575
      ScaleWidth      =   7215
      TabIndex        =   4
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox cboDestination 
      Height          =   315
      ItemData        =   "frmIntro.frx":12902
      Left            =   1680
      List            =   "frmIntro.frx":12927
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1890
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destination:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============
'Pizza 'R' Us - Cashier and Head Office software system
'Jeff Erbrecht, Naveen Sidhu, Chris Chiu
'April 10th, 2007
'=============

'Intro form
'Displays the Pizza 'R' Us logo and allows the user to log in to the system

'The Go button; attempts to log in to the system
Private Sub cmdGo_Click()

    'Check if they selected a store or the head office
    If cboDestination.ListIndex >= 0 And cboDestination.ListIndex <= 9 Then
    
        'Check password
        If Not txtPassword.Text = "pizza" & (cboDestination.ListIndex + 1) Then
            Call MsgBox("Invalid password!", vbCritical, "Denied")
            Exit Sub
        End If
        
        'Cashier program
        CURRENTSTORE = cboDestination.ListIndex 'Store the current store
        
        'Show the cashier form
        Unload Me
        Load frmMain
        frmMain.Reload
        frmMain.Show
        
    Else
        
        'Check password
        If Not txtPassword.Text = "headpizza" Then
            Call MsgBox("Invalid password!", vbCritical, "Denied")
            Exit Sub
        End If
        
        'Head office
        CURRENTSTORE = -1 '-1 is treated as the head office
        Unload Me
        Load frmHeadOffice
        frmHeadOffice.Show
    End If
    
End Sub

Private Sub Form_Load()
    cboDestination.ListIndex = 0
    
    'Load the various names and variables into memory
    Call Main
    
End Sub
