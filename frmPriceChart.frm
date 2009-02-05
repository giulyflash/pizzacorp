VERSION 5.00
Begin VB.Form frmPriceChart 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price chart"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "frmPriceChart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save && Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame fraOtherCosts 
      BackColor       =   &H00000000&
      Caption         =   "Other costs"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   4920
      TabIndex        =   33
      Top             =   3600
      Width           =   4575
      Begin VB.TextBox txtTax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtSlicePopCombo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   36
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPopCost 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax: %"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slice-pop discount: $"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price (pop): $"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraSpecials 
      BackColor       =   &H00000000&
      Caption         =   "Specials"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   4920
      TabIndex        =   19
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   30
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   29
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   28
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   27
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   26
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkSpecialTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtSpecialName 
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboSpecials 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame fraMisc 
      BackColor       =   &H00000000&
      Caption         =   "Other names"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   4575
      Begin VB.TextBox txtTopName 
         Height          =   285
         Left            =   3000
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboTop 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox cboPop 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtPopName 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame fraSizes 
      BackColor       =   &H00000000&
      Caption         =   "Pizza"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtSliceCost 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   32
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtSizeHalfTopPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtSizeSpecialPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   9
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtSizeTopPrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtSizePrice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtSizeName 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Price (slice): $"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price    (topping): $"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price    (specialty): $"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price    (half-topping): $"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2400
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price: $"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmPriceChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Price Chart form
'Displays all prices and various names/labels, all of which can be edited and saved
'to the main text file


Private Sub cboPop_Click()
    'Displays the selected pop name in the text box
    txtPopName.Text = POPNAMES(cboPop.ListIndex)
End Sub

Private Sub cboSize_Click()
    'Updates the size text boxes to match those of the selected size
    txtSizeName.Text = PIZZASIZES(cboSize.ListIndex)
    txtSizePrice.Text = PIZZACOST(cboSize.ListIndex)
    txtSizeTopPrice.Text = TOPCOST(cboSize.ListIndex)
    txtSizeHalfTopPrice.Text = HALFTOPCOST(cboSize.ListIndex)
    txtSizeSpecialPrice.Text = SPECIALCOST(cboSize.ListIndex)
End Sub

Private Sub cboSpecials_Click()
    Dim I As Integer
    
    'Updates the special objects to match those of the selected Special
    txtSpecialName.Text = SPECIALNAMES(cboSpecials.ListIndex)
    For I = 0 To 7
        chkSpecialTop(I).Value = SPECIALTOPS(cboSpecials.ListIndex, I)
    Next I
End Sub

Private Sub cboTop_Click()
    'Displays the selected topping name in the topping text box
    txtTopName.Text = TOPNAMES(cboTop.ListIndex)
End Sub

Private Sub chkSpecialTop_Click(Index As Integer)
    'Edits the special toppings when a topping is clicked
    SPECIALTOPS(cboSpecials.ListIndex, Index) = chkSpecialTop(Index).Value
End Sub

'Closes the form
Private Sub cmdClose_Click()
    Unload Me
    frmHeadOffice.SetFocus 'Sometimes the entire app loses focus, so do this just in case
End Sub

'Writes the new data to main.txt and closes the form
Private Sub cmdGo_Click()
    Call UpdateNames
    cmdClose_Click
End Sub

'Updates the basic text box values
Private Sub Form_Activate()
    txtSliceCost.Text = SLICECOST
    txtPopCost.Text = POPCOST
    txtSlicePopCombo.Text = SLICECOMBO
    txtTax.Text = TAXAMT
End Sub

Private Sub Form_Load()
    Dim I As Integer 'For loop variable
    
    'Reset the sizes list
    cboSize.Clear
    For I = 0 To 3
        cboSize.AddItem PIZZASIZES(I)
    Next I
    cboSize.ListIndex = 0
    
    'Reset the toppings list
    cboTop.Clear
    For I = 0 To 7
        cboTop.AddItem TOPNAMES(I)
    Next I
    cboTop.ListIndex = 0
    
    'Reset the pop list
    cboPop.Clear
    For I = 0 To 3
        cboPop.AddItem POPNAMES(I)
    Next I
    cboPop.ListIndex = 0
    
    'Reset the specials list
    cboSpecials.Clear
    For I = 0 To 2
        cboSpecials.AddItem SPECIALNAMES(I)
    Next I
    cboSpecials.ListIndex = 0
    
    'Reset the special toppings checkboxes
    For I = 0 To 7
        chkSpecialTop(I).Caption = TOPNAMES(I)
    Next I
    
End Sub

'===============
'The following code section is for handling what happens when the user enters in
'a value into a text box. Each text box has its own handler, and aside from the
'difference in what text box the enter key was pressed in, the code is generally
'all the same.

'The general procedure from here on-out is:
    '- check if the Enter key was pressed (key #13)
    '- if so, store the contents of the respective text box to the appropriate variable
    '- update any other form objects with this new change (see: txtTopName; the checkboxes need to be updated if a topping name is changed)
    
'There's no point in repeating the exact same comments, so the comments in this section
'apply to all of the following procedures.
'===============

Private Sub txtPopCost_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        POPCOST = CSng(Val(txtPopCost.Text))
        txtPopCost.Text = POPCOST
    End If
End Sub

Private Sub txtPopName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboPop.List(cboPop.ListIndex) = txtPopName.Text
        POPNAMES(cboPop.ListIndex) = txtPopName.Text
    End If
End Sub

Private Sub txtSizeHalfTopPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        HALFTOPCOST(cboSize.ListIndex) = CSng(Val(txtSizeHalfTopPrice.Text))
        txtSizeHalfTopPrice.Text = HALFTOPCOST(cboSize.ListIndex)
    End If
End Sub

Private Sub txtSizeName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboSize.List(cboSize.ListIndex) = txtSizeName.Text
        PIZZASIZES(cboSize.ListIndex) = txtSizeName.Text
    End If
End Sub

Private Sub txtSizePrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        PIZZACOST(cboSize.ListIndex) = CSng(Val(txtSizePrice.Text))
        txtSizePrice.Text = PIZZACOST(cboSize.ListIndex)
    End If
End Sub

Private Sub txtSizeSpecialPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SPECIALCOST(cboSize.ListIndex) = CSng(Val(txtSizeSpecialPrice.Text))
        txtSizeSpecialPrice.Text = SPECIALCOST(cboSize.ListIndex)
    End If
End Sub

Private Sub txtSizeTopPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TOPCOST(cboSize.ListIndex) = CSng(Val(txtSizeTopPrice.Text))
        txtSizeTopPrice.Text = TOPCOST(cboSize.ListIndex)
    End If
End Sub

Private Sub txtSliceCost_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SLICECOST = CSng(Val(txtSliceCost.Text))
        txtSliceCost.Text = SLICECOST
    End If
End Sub

Private Sub txtSlicePopCombo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SLICECOMBO = CSng(Val(txtSlicePopCombo.Text))
        txtSlicePopCombo.Text = SLICECOMBO
    End If
End Sub

Private Sub txtSpecialName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboSpecials.List(cboSpecials.ListIndex) = txtSpecialName.Text
        SPECIALNAMES(cboSpecials.ListIndex) = txtSpecialName.Text
    End If
End Sub

Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        TAXAMT = CSng(Val(txtTax.Text))
        txtTax.Text = TAXAMT
    End If
End Sub

Private Sub txtTopName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cboTop.List(cboTop.ListIndex) = txtTopName.Text
        TOPNAMES(cboTop.ListIndex) = txtTopName.Text
        chkSpecialTop(cboTop.ListIndex).Caption = txtTopName.Text
    End If
End Sub
