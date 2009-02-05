VERSION 5.00
Begin VB.Form frmViewOrder 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Order"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmViewOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelOrder 
      BackColor       =   &H000000FF&
      Caption         =   "Delete Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ComboBox cboOrder 
      Height          =   315
      ItemData        =   "frmViewOrder.frx":08CA
      Left            =   1080
      List            =   "frmViewOrder.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   540
      Width           =   3255
   End
   Begin VB.ComboBox cboStore 
      Height          =   315
      ItemData        =   "frmViewOrder.frx":08CE
      Left            =   1080
      List            =   "frmViewOrder.frx":08F0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.ListBox lstOrder 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Order #:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Store #:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmViewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Order recaller
'Can go back and recall any order from any store, as well as delete any order


'Handles the form changes when a new order is selected
Private Sub cboOrder_Click()

    Dim I As Integer 'For loop variable
    Dim TMPS As New StringCollection 'Temporary string collection
    
    With STOREDATA(CInt(cboStore.Text) - 1)
    
        'Parse the order for displaying
        Call ParseOrder(.ORDERS(cboOrder.ListIndex), TMPS)
        
        'Display the contents of TMPS in the listbox
        lstOrder.Clear
        For I = 0 To TMPS.Count - 1
            lstOrder.AddItem TMPS.Item(I)
        Next I
        
        'Add the Cash and Change items to the listbox
        lstOrder.AddItem "CASH:                    " & ToMoney(.ORDERS(cboOrder.ListIndex).CASH, 10)
        lstOrder.AddItem "CHANGE:                  " & ToMoney(.ORDERS(cboOrder.ListIndex).CHANGE, 10)
        
        'Select the last item in the listbox
        lstOrder.ListIndex = lstOrder.ListCount - 1
    End With
    
End Sub

'Handles the form changes when a new store is selected
Private Sub cboStore_Click()

    Dim I As Integer 'For loop variable
    
    'Update the orders list; first clear the orders list
    cboOrder.Clear
    
    'Load the orders
    For I = 0 To UBound(STOREDATA(CInt(cboStore.Text) - 1).ORDERS)
        cboOrder.AddItem (I + 1) & ") " & STOREDATA(CInt(cboStore.Text) - 1).ORDERS(I).ID
    Next I
    
    'Select the first order
    cboOrder.ListIndex = 0
    cboOrder_Click
    
End Sub

'Closes the form
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelOrder_Click()
    Dim MSGRESULT As Integer, TMPINDEX As Integer
    MSGRESULT = MsgBox("Are you sure you want to delete this order? This action cannot be undone.", vbYesNo + vbCritical, "Delete order")
    
    If MSGRESULT = vbNo Then Exit Sub
    
    'Check if it's the only order in the store data
    If cboOrder.ListCount = 1 Then
        'Remove it from the list
        cboOrder.RemoveItem cboOrder.ListIndex
        
        'Delete the file
        Kill App.PATH & "\Data\store" & cboStore.Text & ".txt"
        
        'Take it out of the store list
        STOREDATA(CInt(cboStore.Text) - 1).EXISTS = False
        
        TMPINDEX = cboStore.ListIndex
        cboStore.RemoveItem cboStore.ListIndex
        cboStore.ListIndex = IIf(TMPINDEX >= cboStore.ListCount, cboStore.ListCount - 1, TMPINDEX)

        Exit Sub
    End If
    
    Call DelOrder(STOREDATA(CInt(cboStore.Text) - 1).ORDERS, cboOrder.ListIndex)
    
    TMPINDEX = cboOrder.ListIndex
    cboOrder.RemoveItem (cboOrder.ListIndex)
    cboOrder.ListIndex = IIf(TMPINDEX >= cboOrder.ListCount, cboOrder.ListCount - 1, TMPINDEX)
        
End Sub

'Saves any changes made to the store data file
Private Sub cmdSave_Click()
    Call UpdateFile(CInt(cboStore.Text) - 1, STOREDATA(CInt(cboStore.Text) - 1).ORDERS)
    Call MsgBox("Saved the data for store #" & cboStore.Text, vbInformation, "Saved")
End Sub

Private Sub Form_Activate()
    Dim I As Integer
    
    cboStore.Clear
    cboOrder.Clear
    
    'Load the stores
    For I = 0 To 9
        If STOREDATA(I).EXISTS Then cboStore.AddItem (I + 1)
    Next I
    
    If cboStore.ListCount = 0 Then
        Unload Me
        Exit Sub
    End If
    
    cboStore.ListIndex = 0
    
End Sub
