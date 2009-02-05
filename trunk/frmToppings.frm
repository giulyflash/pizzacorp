VERSION 5.00
Begin VB.Form frmToppings 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toppings"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmToppings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraReturn 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   3975
      Begin VB.OptionButton optReturn 
         BackColor       =   &H00000000&
         Caption         =   "Less than"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optReturn 
         BackColor       =   &H00000000&
         Caption         =   "Exact"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   18
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optReturn 
         BackColor       =   &H00000000&
         Caption         =   "At least"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.OptionButton optChoice 
      BackColor       =   &H00000000&
      Caption         =   "By kind of toppings"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1695
   End
   Begin VB.OptionButton optChoice 
      BackColor       =   &H00000000&
      Caption         =   "By # of toppings"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Go"
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
      TabIndex        =   13
      Top             =   4080
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
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame fraTopping 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   10
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkTop 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame fraTopping 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.ComboBox cboNumToppings 
         Height          =   315
         ItemData        =   "frmToppings.frx":08CA
         Left            =   1920
         List            =   "frmToppings.frx":0901
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNumToppings 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Number of toppings: "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmToppings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The Toppings selection form
'Allows a user to pick out various topping arrangements and create a report based on it


'Closes the form
Private Sub cmdClose_Click()
    Unload Me
End Sub

'Performs the report on pizzas with whatever arrangement of toppings
Private Sub cmdGo_Click()
    Dim I As Integer, J As Integer, K As Integer, L As Integer 'Four for-loops
    Dim CT As Integer 'Temp toppings count
    Dim TOPCT(0 To 9, 0 To 3) As Integer 'Topping count per size, per store
    Dim SLICECT(0 To 9) As Integer 'Slice count per store
    Dim TMPCOMPARE As Boolean 'Temporary comparison variable
    
    'There are two option buttons, one for each frame; see which one is selected
    If optChoice(0).Value = True Then
        
        'By number of toppings
        For I = 0 To 9
            
            If STOREDATA(I).EXISTS Then
                
                With STOREDATA(I)
                    For J = 0 To UBound(.ORDERS)
                    
                        'Pizzas
                        For K = 0 To UBound(.ORDERS(J).PIZZADATA)
                            
                            CT = 0
                            For L = 0 To 7
                                CT = CT + .ORDERS(J).PIZZADATA(K).TOPS(L)
                                If Not .ORDERS(J).PIZZADATA(K).FULLMODE Then CT = CT + .ORDERS(J).PIZZADATA(K).TOPS2(L)
                            Next L
                            
                            'Set the comparison
                            If optReturn(0) Then
                                'At Least
                                TMPCOMPARE = (CT >= cboNumToppings.ListIndex)
                            ElseIf optReturn(1) Then
                                'Less than
                                TMPCOMPARE = (CT < cboNumToppings.ListIndex)
                            Else
                                'Exact
                                TMPCOMPARE = (CT = cboNumToppings.ListIndex)
                            End If
                            
                            If TMPCOMPARE Then
                                TOPCT(I, .ORDERS(J).PIZZADATA(K).SIZE) = TOPCT(I, .ORDERS(J).PIZZADATA(K).SIZE) + 1
                            End If
                        
                        Next K
                        
                        
                        'Slices
                        For K = 0 To UBound(.ORDERS(J).SLICES)
                            
                            CT = 0
                            For L = 0 To 7
                                CT = CT + .ORDERS(J).SLICES(K).TOPS(L)
                            Next L
                            
                            'Set the comparison
                            If optReturn(0) Then
                                'At Least
                                TMPCOMPARE = (CT >= cboNumToppings.ListIndex)
                            ElseIf optReturn(1) Then
                                'Less than
                                TMPCOMPARE = (CT < cboNumToppings.ListIndex)
                            Else
                                'Exact
                                TMPCOMPARE = (CT = cboNumToppings.ListIndex)
                            End If
                            
                            If TMPCOMPARE Then
                                SLICECT(I) = SLICECT(I) + 1
                            End If
                        
                        Next K
                        
                    Next J
                
                End With
                
            End If
            
        Next I
        
        'Display the results in the head office form listbox
        With frmHeadOffice.lstMain
        
            'Display detailed results if necessary (by store)
            If frmHeadOffice.chkDetail Then
            
                For J = 0 To 9
                    If STOREDATA(J).EXISTS Then
                        .AddItem "  STORE #" & (J + 1)
                        For I = 0 To 3
                            .AddItem "    " & PIZZASIZES(I) & " " & cboNumToppings.ListIndex & "-item pizzas: " & TOPCT(J, I)
                        Next I
                        .AddItem "    " & cboNumToppings.ListIndex & "-item slices: " & SLICECT(J)
                    End If
                Next J
                
                .AddItem ""
                
            End If
            
            'Display totals
            For I = 0 To 3
                CT = 0
                For J = 0 To 9
                    CT = CT + TOPCT(J, I)
                Next J
                .AddItem "  " & PIZZASIZES(I) & " " & cboNumToppings.ListIndex & "-item pizzas: " & CT
            Next I
                
            CT = 0
            For J = 0 To 9
                CT = CT + SLICECT(J)
            Next J
            .AddItem "  " & cboNumToppings.ListIndex & "-item slices: " & CT
            .AddItem ""
            
        End With
        
    Else
    
        'By kind of toppings
        For I = 0 To 9
        
            If STOREDATA(I).EXISTS Then
            
                With STOREDATA(I)
                    For J = 0 To UBound(.ORDERS)
                    
                        'Pizzas
                        For K = 0 To UBound(.ORDERS(J).PIZZADATA)

                            CT = 0
                            For L = 0 To 7
                            
                                'Set the comparison
                                If optReturn(0) Then
                                    'At Least
                                    TMPCOMPARE = ((.ORDERS(J).PIZZADATA(K).TOPS(L) = 0) And (chkTop(L).Value = 1))
                                ElseIf optReturn(2) Then
                                    'Exact
                                    TMPCOMPARE = Not (.ORDERS(J).PIZZADATA(K).TOPS(L) = chkTop(L).Value)
                                End If
                            
                                If TMPCOMPARE Then
                                    CT = 1
                                End If
                            Next L
                            
                            If CT = 1 And Not .ORDERS(J).PIZZADATA(K).FULLMODE Then
                                'Check the other half
                                CT = 0
                                For L = 0 To 7
                                
                                    'Set the comparison
                                    If optReturn(0) Then
                                        'At Least
                                        TMPCOMPARE = ((.ORDERS(J).PIZZADATA(K).TOPS2(L) = 0) And (chkTop(L).Value = 1))
                                    ElseIf optReturn(2) Then
                                        'Exact
                                        TMPCOMPARE = Not (.ORDERS(J).PIZZADATA(K).TOPS2(L) = chkTop(L).Value)
                                    End If
                                    
                                    If TMPCOMPARE Then
                                        CT = 1
                                    End If
                                Next L
                            End If
                            
                            If CT = 0 Then
                                TOPCT(I, .ORDERS(J).PIZZADATA(K).SIZE) = TOPCT(I, .ORDERS(J).PIZZADATA(K).SIZE) + 1
                            End If
                        
                        Next K
                            
                        'Slices
                        For K = 0 To UBound(.ORDERS(J).SLICES)
                        
                            CT = 0
                            For L = 0 To 7
                            
                                    'Set the comparison
                                    If optReturn(0) Then
                                        'At Least
                                        TMPCOMPARE = ((.ORDERS(J).SLICES(K).TOPS(L) = 0) And (chkTop(L).Value = 1))
                                    ElseIf optReturn(2) Then
                                        'Exact
                                        TMPCOMPARE = Not (.ORDERS(J).SLICES(K).TOPS(L) = chkTop(L).Value)
                                    End If
                                
                                If TMPCOMPARE Then
                                    CT = 1
                                End If
                            Next L
                            
                            If CT = 0 Then
                                SLICECT(I) = SLICECT(I) + 1
                            End If
                            
                        Next K
                        
                    Next J
                    
                End With
                
            End If
            
        Next I
        
        'Display results
        With frmHeadOffice.lstMain
        
            'Show the toppings that were searched
            .AddItem "  Toppings searched:"
            For I = 0 To 7
                If chkTop(I) Then .AddItem "    " & TOPNAMES(I)
            Next I
            .AddItem ""
            
            'Show detailed results if necessary
            If frmHeadOffice.chkDetail Then
                
                For J = 0 To 9
                    If STOREDATA(J).EXISTS Then
                        .AddItem "  STORE #" & (J + 1)
                        For I = 0 To 3
                            .AddItem "    " & PIZZASIZES(I) & " pizzas: " & TOPCT(J, I)
                        Next I
                        .AddItem "    Slices: " & SLICECT(J)
                    End If
                Next J
                .AddItem ""
            
            End If
            
            'Display totals
            For I = 0 To 3
                CT = 0
                For J = 0 To 9
                    CT = CT + TOPCT(J, I)
                Next J
                .AddItem "  " & PIZZASIZES(I) & " pizzas: " & CT
            Next I
                
            CT = 0
            For J = 0 To 9
                CT = CT + SLICECT(J)
            Next J
            .AddItem "  Slices: " & CT
            .AddItem ""
            
        End With
                        
    End If
    
    Call frmHeadOffice.GoToEnd(frmHeadOffice.lstMain)
    frmHeadOffice.SetFocus
    
End Sub

'Updates the topping objects with the topping names
Private Sub Form_Activate()
    Dim I As Integer
    For I = 0 To 7
        chkTop(I).Caption = TOPNAMES(I)
    Next I
End Sub

'Updates the form
Private Sub Form_Load()
    cboNumToppings.ListIndex = 0
    optChoice_Click (0)
End Sub

'Updates object .Enabled properties on the form when a different search choice is selected
Private Sub optChoice_Click(Index As Integer)
    Dim I As Integer 'For loop variable
    
    'Update the frames' .Enabled properties to correspond to their respective option buttons
    fraTopping(0).Enabled = optChoice(0).Value
    fraTopping(1).Enabled = optChoice(1).Value
    
    'Update the objects on each frame
    lblNumToppings.Enabled = optChoice(0).Value
    cboNumToppings.Enabled = optChoice(0).Value

    'We can't have the "less than" option for 'By Topping', so let's disable it if necessary
    If optReturn(1) Then optReturn(2).Value = True
    optReturn(1).Enabled = optChoice(0).Value
    
    'Update the topping checkboxes
    For I = 0 To 7
        chkTop(I).Enabled = optChoice(1).Value
    Next I
    
End Sub
