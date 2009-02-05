VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pizza Calc - Cashier"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelOrder 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete Selected Order"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   8640
      Width           =   2295
   End
   Begin VB.Frame fraKeypad 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   6960
      TabIndex        =   71
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdChange 
         BackColor       =   &H00FF8080&
         Caption         =   "ENTER"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   85
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1200
         TabIndex        =   83
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2280
         TabIndex        =   82
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   81
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   1200
         TabIndex        =   80
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   2280
         TabIndex        =   79
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   120
         TabIndex        =   78
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   1200
         TabIndex        =   77
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   2280
         TabIndex        =   76
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtKeypad 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdKeyPad 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   1200
         TabIndex        =   74
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdClearKeypad 
         Caption         =   "CLEAR"
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdBackspaceKeypad 
         Caption         =   "BKSP"
         Height          =   375
         Left            =   2280
         TabIndex        =   72
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblDollar 
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   255
         Left            =   2160
         TabIndex        =   86
         Top             =   3300
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdClearCalc 
      Caption         =   "CLEAR"
      Height          =   495
      Left            =   6000
      TabIndex        =   70
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdNewOrder 
      BackColor       =   &H0080FFFF&
      Caption         =   "NEW ORDER"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3240
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
      Height          =   5460
      Left            =   2640
      TabIndex        =   68
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton cmdCalc 
      BackColor       =   &H00FF8080&
      Caption         =   "CALC"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame fraPop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pop"
      Height          =   2055
      Left            =   6960
      TabIndex        =   44
      Top             =   3240
      Width           =   3255
      Begin VB.CommandButton cmdDelPop 
         Caption         =   "DEL"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   56
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdDelPop 
         Caption         =   "DEL"
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   55
         Top             =   1140
         Width           =   615
      End
      Begin VB.CommandButton cmdDelPop 
         Caption         =   "DEL"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   54
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtPop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   53
         Text            =   "0"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtPop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "0"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox txtPop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   51
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtPop 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   50
         Text            =   "0"
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdDelPop 
         Caption         =   "DEL"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   49
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cmdAddPop 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   48
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddPop 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddPop 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddPop 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraSlices 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Slices"
      Height          =   3015
      Left            =   6960
      TabIndex        =   32
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdSpecialSlice 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   63
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdSpecialSlice 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   64
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdSpecialSlice 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   65
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearSlice 
         Caption         =   "Clear"
         Height          =   375
         Left            =   1800
         TabIndex        =   66
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkNoSlice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Slices"
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   735
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   41
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkSliceTop 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddSlice 
         Caption         =   "ADD"
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdDelSlice 
         Caption         =   "DEL"
         Height          =   375
         Left            =   2400
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox cboSlices 
         Height          =   315
         ItemData        =   "frmMain.frx":08CA
         Left            =   120
         List            =   "frmMain.frx":08D1
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ListBox lstQueue 
      Height          =   5130
      ItemData        =   "frmMain.frx":08D8
      Left            =   120
      List            =   "frmMain.frx":08DA
      TabIndex        =   30
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame fraPizza 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pizza"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox chkNoPizza 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pizza"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   0
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.Frame fraSpecials 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   1560
         TabIndex        =   29
         Top             =   720
         Width           =   1455
         Begin VB.CommandButton cmdSpecial 
            Height          =   375
            Index           =   2
            Left            =   0
            TabIndex        =   59
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdSpecial 
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   60
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdSpecial 
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   61
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   0
            TabIndex        =   62
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraPizzaSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Size"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
         Begin VB.OptionButton optSize 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1200
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1200
         End
         Begin VB.OptionButton optSize 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdDelPizza 
         Caption         =   "DEL"
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cmdAddPizza 
         Caption         =   "ADD"
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
      Begin VB.Frame fraHalf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Half"
         Enabled         =   0   'False
         Height          =   2295
         Left            =   4920
         TabIndex        =   5
         Top             =   600
         Width           =   1695
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   28
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkTop2 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full"
         Height          =   2295
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Width           =   1695
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkTop 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.OptionButton optHalf 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Half-half"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optFull 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Full"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cboPizza 
         Height          =   315
         ItemData        =   "frmMain.frx":08DC
         Left            =   120
         List            =   "frmMain.frx":08E3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Image imgLogo 
      Height          =   3240
      Left            =   6960
      Picture         =   "frmMain.frx":08EA
      Top             =   5520
      Width           =   3240
   End
   Begin VB.Label lblQueue 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER QUEUE (mm/dd/yyyy)"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnuLn1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cashier form
'Processes all orders directly, can go back and change orders, and can delete orders


Option Explicit

Dim QUEUECOUNT As Integer   'Number of orders in the queue
Dim CURRENTORDER As Integer 'The current order number

Dim LOADING As Boolean     'If true, then procedures performed automatically from listbox modifications will not execute to maintain accuracy
Dim ADDNEW As Boolean      'If true, then clicking New Order will add to the ORDERS array
Dim CHANGETAKEN As Boolean 'If true, then change has been taken from the customer for the current order
Dim DOUPDATE As Boolean    'If true, then the UpdateFile procedure will be called on the form unload; is set to true as soon as CALC is clicked for the first time. It's false until then, because there's no need to save the orders if nothing has happened

'Handles the form changes when a different pizza number is selected
Private Sub cboPizza_Click()

    Dim I As Integer 'For loop variables
    
    'Load the pizza's properties
    With ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex)
        optSize(.SIZE).Value = True
        
        'Set the fullmode/half-half object properties
        If .FULLMODE Then
            optFull.Value = True
            optFull_Click
        Else
            optHalf.Value = True
            optHalf_Click
        End If
        
        'Show the selected toppings for the current pizza
        For I = 0 To 7
            chkTop(I).Value = IIf(.TOPS(I) = 1, Checked, Unchecked)
            chkTop2(I).Value = IIf(.TOPS2(I) = 1, Checked, Unchecked)
        Next I
        
    End With
    
    'Reset the keypad frame
    fraKeypad.Visible = False
    txtKeypad.Text = ""
    CHANGETAKEN = False

End Sub

'Handles the form changes when a different slice number is selected
Private Sub cboSlices_Click()

    Dim I As Integer 'For loop variable
    
    'Load the slice's properties
    With ORDERS(CURRENTORDER).SLICES(cboSlices.ListIndex)
        'Display the selected toppings
        For I = 0 To 7
            chkSliceTop(I).Value = IIf(.TOPS(I) = 1, Checked, Unchecked)
        Next I
    End With
    
End Sub

'Toggles certain form objects when the Pizza checkbox is clicked
Private Sub chkNoPizza_Click()

    Dim I As Integer 'For loop variable
    Dim SWITCH As Boolean 'The target enabled toggle to set things to
    
    'Set the enabled toggle
    SWITCH = IIf(chkNoPizza.Value = Checked, True, False)
    
    'Set all relevant form objects' .Enabled properties to SWITCH
    fraFull.Enabled = SWITCH
    optFull.Enabled = SWITCH
    fraHalf.Enabled = SWITCH And optHalf.Value
    optHalf.Enabled = SWITCH
    For I = 0 To 2
        cmdSpecial(I).Enabled = SWITCH
    Next I
    For I = 0 To 3
        optSize(I).Enabled = SWITCH
    Next I
    For I = 0 To 7
        chkTop(I).Enabled = SWITCH
        chkTop2(I).Enabled = SWITCH And optHalf.Value
    Next I
    cmdAddPizza.Enabled = SWITCH
    cmdDelPizza.Enabled = SWITCH
    cmdClear.Enabled = SWITCH
    cboPizza.Enabled = SWITCH
    ORDERS(CURRENTORDER).DOPIZZA = SWITCH
End Sub

'Toggles certain form objects when the Pizza checkbox is clicked
Private Sub chkNoSlice_Click()

    Dim I As Integer 'For loop variable
    Dim SWITCH As Boolean 'The target enabled toggle to set things to
    
    'Set the enabled toggle
    SWITCH = IIf(chkNoSlice.Value = Checked, True, False)
    
    'Set all relevant form objects' .Enabled properties to SWITCH
    For I = 0 To 7
        chkSliceTop(I).Enabled = SWITCH
    Next I
    For I = 0 To 2
        cmdSpecialSlice(I).Enabled = SWITCH
    Next I
    cmdClearSlice.Enabled = SWITCH
    cboSlices.Enabled = SWITCH
    cmdAddSlice.Enabled = SWITCH
    cmdDelSlice.Enabled = SWITCH
    ORDERS(CURRENTORDER).DOSLICES = SWITCH
End Sub

'Sets the internal variable whenever a topping is clicked
Private Sub chkSliceTop_Click(Index As Integer)
    ORDERS(CURRENTORDER).SLICES(cboSlices.ListIndex).TOPS(Index) = IIf(chkSliceTop(Index).Value = Checked, 1, 0)
End Sub

'Sets the internal variable whenever a topping is clicked
Private Sub chkTop_Click(Index As Integer)
    ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).TOPS(Index) = IIf(chkTop(Index).Value = Checked, 1, 0)
End Sub

'Sets the internal variable whenever a topping is clicked
Private Sub chkTop2_Click(Index As Integer)
    ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).TOPS2(Index) = IIf(chkTop2(Index).Value = Checked, 1, 0)
End Sub

'Adds a pizza to the current order
Private Sub cmdAddPizza_Click()
    Call AddPizza(ORDERS(CURRENTORDER))
    cboPizza.AddItem (UBound(ORDERS(CURRENTORDER).PIZZADATA) + 1)
    cboPizza.ListIndex = UBound(ORDERS(CURRENTORDER).PIZZADATA)
    cboPizza_Click
End Sub

'Adds a pop to the appropriate pop variable
Private Sub cmdAddPop_Click(Index As Integer)
    With ORDERS(CURRENTORDER)
        .POP(Index) = .POP(Index) + 1
        txtPop(Index).Text = .POP(Index)
    End With
End Sub

'Adds a slice to the current order
Private Sub cmdAddSlice_Click()
    Call AddSlice(ORDERS(CURRENTORDER))
    cboSlices.AddItem (UBound(ORDERS(CURRENTORDER).SLICES) + 1)
    cboSlices.ListIndex = UBound(ORDERS(CURRENTORDER).SLICES)
    cboSlices_Click
End Sub

'Backspaces a number in the keypad window
Private Sub cmdBackspaceKeypad_Click()
    If txtKeypad.Text <> "" Then txtKeypad.Text = Left(txtKeypad.Text, Len(txtKeypad.Text) - 1)
End Sub

'Calculates and displays all necessary values pertaining to the order
Private Sub cmdCalc_Click()

    Dim TMPS As New StringCollection 'Temporary collection of strings to add to the order summary box
    Dim I As Integer 'For loop variable
    
    'Clear the order summary window
    cmdClearCalc_Click
    
    'Parse the order into a string collection
    Call ParseOrder(ORDERS(CURRENTORDER), TMPS)
    
    'Display the output of the order parsing
    For I = 0 To TMPS.Count - 1
        lstOrder.AddItem TMPS.Item(I)
    Next I
    
    'Select the last item in the order summary list
    lstOrder.ListIndex = lstOrder.ListCount - 1
    
    'Display the keypad
    cmdChange.Enabled = True
    fraKeypad.Visible = True
    
    DOUPDATE = True
    
End Sub

'Registers that the customer has given the cashier an amount of money
Private Sub cmdChange_Click()

    'Input the cash amount
    Dim S As Single
    S = Val(txtKeypad.Text)
    
    With ORDERS(CURRENTORDER)
        .CASH = Round(S, 2) 'Store the cash amount to the order info
        .CHANGE = Round(.CASH - .TOTAL, 2) 'Calculate the needed change
        
        'Check if the customer gave enough cash
        If .CHANGE < 0 Then
            Call MsgBox("Not enough cash! Need " & ToMoney(-.CHANGE), vbCritical, "Not enough cash")
            Call cmdClearKeypad_Click
            Exit Sub
        End If
        
        'If change has already been taken, therefore this is just an adjustment...
        If CHANGETAKEN Then
            'Remove the last two items (cash and change) in the order summary listbox if change has already been taken
            lstOrder.RemoveItem (lstOrder.ListCount - 1)
            lstOrder.RemoveItem (lstOrder.ListCount - 1)
        End If
        
        'Re-add the new cash and change amounts
        lstOrder.AddItem "CASH:                    " & ToMoney(.CASH, 10)
        lstOrder.AddItem "CHANGE:                  " & ToMoney(.CHANGE, 10)
        lstOrder.ListIndex = lstOrder.ListCount - 1 'Select the last item in the list
        
    End With

    'Clear the keypad textbox
    txtKeypad.Text = ""
    CHANGETAKEN = True 'Set the changetaken variable
    
End Sub

'Clears the pizza frame
Private Sub cmdClear_Click()
    Dim I As Integer
    For I = 0 To 7
        'Clear the existing toppings
        chkTop(I).Value = Unchecked
        chkTop2(I).Value = Unchecked
    Next I
    optSize(0).Value = True
    optFull.Value = True
    optFull_Click
End Sub

'Clears the order summary and the keypad
Private Sub cmdClearCalc_Click()
    lstOrder.Clear
    cmdChange.Enabled = False
    fraKeypad.Visible = False
    txtKeypad.Text = ""
    CHANGETAKEN = False
End Sub

'Clears the keypad textbox
Private Sub cmdClearKeypad_Click()
    txtKeypad.Text = ""
End Sub

'Clears the slice frame
Private Sub cmdClearSlice_Click()
    Dim I As Integer
    For I = 0 To 7
        'Clear the existing toppings
        chkSliceTop(I).Value = Unchecked
    Next I
End Sub

'Deletes the selected order
Private Sub cmdDelOrder_Click()

    Dim TMPINDEX As Integer 'Temporary selection index
    
    'Only delete if something is selected and there's more than one item in the listbox
    If lstQueue.ListIndex >= 0 And lstQueue.ListCount > 1 Then
    
        'Delete the order from the orders array
        Call DelOrder(ORDERS, lstQueue.ListIndex)
        
        'Select a new order in the queue list
        TMPINDEX = lstQueue.ListIndex
        lstQueue.RemoveItem (lstQueue.ListIndex)
        lstQueue.ListIndex = IIf(TMPINDEX >= lstQueue.ListCount, lstQueue.ListCount - 1, TMPINDEX)
        
        'Calculate the new currentorder and queuecount values
        CURRENTORDER = lstQueue.ListIndex
        QUEUECOUNT = lstQueue.ListCount - 1
        
        DOUPDATE = True 'Changes have been made, so save the file on form unload
    End If
End Sub

'Deletes a pizza from the current order
Private Sub cmdDelPizza_Click()

    Dim TMPINDEX As Integer 'Temporary selection index
    
    'Only delete a pizza if there's more than one pizza
    If UBound(ORDERS(CURRENTORDER).PIZZADATA) > 0 Then
        
        'Delete the pizza from the current order
        Call DelPizza(ORDERS(CURRENTORDER), cboPizza.ListIndex)
        
        'Remove the pizza from its respective combo box
        TMPINDEX = cboPizza.ListIndex
        cboPizza.RemoveItem (cboPizza.ListCount - 1) 'Remove the last pizza in the list
        If TMPINDEX = cboPizza.ListCount Then TMPINDEX = TMPINDEX - 1
        
        'Update the form with the newly selected pizza
        cboPizza.ListIndex = TMPINDEX
        cboPizza_Click
        
    ElseIf UBound(ORDERS(CURRENTORDER).PIZZADATA) = 0 Then
    
        'If it's the last pizza, just uncheck the Pizza box
        chkNoPizza.Value = Unchecked
        chkNoPizza_Click
        
    End If
End Sub

'Deletes a pop from the pop value variable
Private Sub cmdDelPop_Click(Index As Integer)
    With ORDERS(CURRENTORDER)
        .POP(Index) = .POP(Index) - 1
        If .POP(Index) < 0 Then .POP(Index) = 0
        txtPop(Index).Text = .POP(Index)
    End With
End Sub

'Deletes a slice from the current ORDERS array
Private Sub cmdDelSlice_Click()

    Dim TMPINDEX As Integer 'Temporary selection index
    
    'Only delete a slice if there's more than one slice
    If UBound(ORDERS(CURRENTORDER).SLICES) > 0 Then
        
        'Delete the slice from the current order
        Call DelSlice(ORDERS(CURRENTORDER), cboSlices.ListIndex)
        
        'Remove the slice from its combo box
        TMPINDEX = cboSlices.ListIndex
        cboSlices.RemoveItem (cboSlices.ListCount - 1) 'Remove the last slice in the list
        If TMPINDEX = cboSlices.ListCount Then TMPINDEX = TMPINDEX - 1
        
        'Update the form
        cboSlices.ListIndex = TMPINDEX
        cboSlices_Click
        
    ElseIf UBound(ORDERS(CURRENTORDER).SLICES) = 0 Then
    
        'If it's the last slice, just uncheck the Slice box
        chkNoSlice.Value = Unchecked
        chkNoSlice_Click
        
    End If
End Sub

'Adds a number from the keypad onto the keypad text box
Private Sub cmdKeyPad_Click(Index As Integer)

    'If the index is less than ten, then its a number button from 0-9
    If Index < 10 Then
        txtKeypad.Text = txtKeypad.Text & Index
        
    'If the index equals ten, then it's the decimal point
    Else
        txtKeypad.Text = txtKeypad.Text & "."
    End If
    
End Sub

'Makes a new order
Private Sub cmdNewOrder_Click()

    Dim I As Integer 'For loop variable
    
    'Reset the form
    lstOrder.Clear
    cmdChange.Enabled = False
    fraKeypad.Visible = False
    txtKeypad.Text = ""
    CHANGETAKEN = False
    cboPizza.Clear
    cboSlices.Clear
        
    'Check if we're actually adding a new order, and not just using the sub for something else that just clears the form
    If ADDNEW Then
    
        'Adjust the queuecount and currentorder variables
        QUEUECOUNT = QUEUECOUNT + 1
        CURRENTORDER = QUEUECOUNT
        
        'Add an order to the ORDERS array
        Call AddOrder(ORDERS)
        
        'Add the order to the queue list
        lstQueue.AddItem (QUEUECOUNT + 1) & ") " & ORDERS(CURRENTORDER).ID
        LOADING = True
        lstQueue.ListIndex = lstQueue.ListCount - 1 'Select the latest order
        LOADING = False
        
        'Reset the pizza and slice objects
        cboPizza.AddItem "1"
        cboPizza.ListIndex = 0
        chkNoPizza.Value = Checked
        cboSlices.AddItem "1"
        cboSlices.ListIndex = 0
        chkNoSlice.Value = Unchecked
        
        'Reset the pop
        For I = 0 To 3
            txtPop(I).Text = "0"
        Next I
    
    End If
    
End Sub

'Creates a special pizza arrangement
Private Sub cmdSpecial_Click(Index As Integer)
    Dim I As Integer
    For I = 0 To 7
        'Load the special toppings
        chkTop(I).Value = SPECIALTOPS(Index, I)
        chkTop2(I).Value = 0
        ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).TOPS(I) = SPECIALTOPS(Index, I)
        ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).TOPS2(I) = 0
    Next I
    
    'Set it to full mode, can't be half-half
    optFull.Value = True
    optFull_Click
End Sub

'Creates a special slice arrangement
Private Sub cmdSpecialSlice_Click(Index As Integer)
    Dim I As Integer
    For I = 0 To 7
        'Load the specialty toppings
        chkSliceTop(I).Value = SPECIALTOPS(Index, I)
        ORDERS(CURRENTORDER).SLICES(cboSlices.ListIndex).TOPS(I) = SPECIALTOPS(Index, I)
    Next I
End Sub

'Resets all of the common variables
Public Sub Reload()

    Dim I As Integer, J As Integer 'For loop variables

    'Reset the order queue and current order variables
    CURRENTORDER = 0
    QUEUECOUNT = 0
    
    'Reset various names
    For I = 0 To 3
        'Pizza sizes
        optSize(I).Caption = PIZZASIZES(I)
    Next I
    
    For I = 0 To 7
        'Toppings
        chkTop(I).Caption = TOPNAMES(I)
        chkTop2(I).Caption = TOPNAMES(I)
        chkSliceTop(I).Caption = TOPNAMES(I)
    Next I
    
    For I = 0 To 3
        'Pop names
        cmdAddPop(I).Caption = POPNAMES(I)
    Next I
    
    For I = 0 To 2
       'Special names
       cmdSpecial(I).Caption = SPECIALNAMES(I)
       cmdSpecialSlice(I).Caption = SPECIALNAMES(I)
    Next I
    
    'Select the first pizza by default
    cboPizza.ListIndex = 0
    cboPizza_Click
    
    'Deselect the slices
    chkNoSlice.Value = Unchecked
    chkNoSlice_Click
    cboSlices.ListIndex = 0
    optFull_Click
    
    'Load previous orders
    If FileExists(App.PATH & "\Data\store" & (CURRENTSTORE + 1) & ".txt") Then
    
        Call ReadFile(CURRENTSTORE, ORDERS)
        
        'Add the previousorders to the queue
        For I = 0 To UBound(ORDERS)
            lstQueue.AddItem (I + 1) & ") " & ORDERS(I).ID
        Next I
    
        Call AddOrder(ORDERS) 'Only have to do this if we load previous orders, otherwise Main() takes care of it
    End If
    
    'Create a new order
    QUEUECOUNT = UBound(ORDERS)
    CURRENTORDER = QUEUECOUNT
    lstQueue.AddItem (QUEUECOUNT + 1) & ") " & ORDERS(CURRENTORDER).ID

    lstQueue.ListIndex = lstQueue.ListCount - 1
    
    ADDNEW = True
    
    'Add the date to the form's caption
    Me.Caption = Me.Caption & " - " & Format(Now, "dddd, MMMM dd, yyyy")
    
    'Add the store number to the form's caption
    Me.Caption = Me.Caption & " - Store #" & (CURRENTSTORE + 1)
    
    DOUPDATE = False
    
    Exit Sub 'Avoid the BadStore block
    
BadStore:
    Call MsgBox("Invalid store number! Now closing program.", vbCritical, "Invalid store")
    Unload Me
    End
    
End Sub

'Called when the form is closed
Private Sub Form_Unload(Cancel As Integer)
    'Write everything to the updated data file
    If DOUPDATE Then Call UpdateFile(CURRENTSTORE, ORDERS)
End Sub

'Updates the form to whatever order was clicked in the queue
Private Sub lstQueue_Click()
    Dim I As Integer
    
    If LOADING Then Exit Sub
    
    'Load the selected order from the queue
    CURRENTORDER = lstQueue.ListIndex
    ADDNEW = False
    cmdNewOrder_Click
    ADDNEW = True
    
    With ORDERS(CURRENTORDER)
    
        'Update the pizza frame
        If .DOPIZZA Then
            For I = 0 To UBound(.PIZZADATA)
                cboPizza.AddItem (I + 1)
            Next I
        Else
            cboPizza.AddItem "1"
        End If
        cboPizza.ListIndex = 0
        chkNoPizza.Value = -CInt(.DOPIZZA)
        chkNoPizza_Click
        
        'Update the slices frame
        If .DOSLICES Then
            For I = 0 To UBound(.SLICES)
                cboSlices.AddItem (I + 1)
            Next I
        Else
            cboSlices.AddItem "1"
        End If
        cboSlices.ListIndex = 0
        chkNoSlice.Value = -CInt(.DOSLICES)
        chkNoSlice_Click
        
        'Update the pop frame
        For I = 0 To 3
            txtPop(I).Text = .POP(I)
        Next I
    End With
    
End Sub

'Exits the program
Private Sub mnuExit_Click()
    Dim MSGRES As Integer
    MSGRES = MsgBox("Are you sure you want to exit?", vbCritical + vbOKCancel, "Exit")
    If MSGRES = vbCancel Then Exit Sub 'Exit this sub if they click cancel
    End
End Sub

'Logs out of the cashier form
Private Sub mnuLogout_Click()
    Dim MSGRES As Integer
    MSGRES = MsgBox("Are you sure you want to log out?", vbInformation + vbOKCancel, "Logout")
    If MSGRES = vbCancel Then Exit Sub 'Exit this sub if they click canel
    
    'Close this form, show the login screen, and reload the main.txt file
    Unload Me
    frmIntro.Show
    Call Main
    
End Sub

'Called when the Full option button is chosen
Private Sub optFull_Click()
    Dim I As Integer 'For loop variable
    
    'If optFull is currently chosen...
    If optFull Then
    
        'Disable the half-half objects
        fraHalf.Enabled = False
        For I = 0 To 7
            chkTop2(I).Enabled = False
        Next I
        
        'Update the current order
        ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).FULLMODE = True
        
        'Set the Full frame to say Full
        fraFull.Caption = "Full"
    End If
End Sub

'Called when the Half option button is chosen
Private Sub optHalf_Click()
    Dim I As Integer 'For loop variable
    
    'If optHalf is current chosen...
    If optHalf Then
    
        'Enable the half-half objects
        fraHalf.Enabled = True
        For I = 0 To 7
            chkTop2(I).Enabled = True
        Next I
        
        'Update the current order
        ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).FULLMODE = False
        
        'Set the Full frame to say Half
        fraFull.Caption = "Half"
    End If
End Sub

'Updates the current order when a new size of pizza is clicked
Private Sub optSize_Click(Index As Integer)
    ORDERS(CURRENTORDER).PIZZADATA(cboPizza.ListIndex).SIZE = Index
End Sub

'Handles the changes that need to be made when a value is entered manually into the pop text boxes
Private Sub txtPop_Change(Index As Integer)
    Dim TMPSTART As Integer
    With txtPop(Index)
        TMPSTART = .SelStart
        .Text = Val(.Text)
        ORDERS(CURRENTORDER).POP(Index) = Val(.Text)
        .SelStart = TMPSTART
    End With
End Sub
