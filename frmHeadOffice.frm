VERSION 5.00
Begin VB.Form frmHeadOffice 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pizza R Us - Head Office"
   ClientHeight    =   6960
   ClientLeft      =   3045
   ClientTop       =   3105
   ClientWidth     =   7920
   Icon            =   "frmHeadOffice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7920
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   7635
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.CommandButton cmdSaveReport 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save Report..."
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
      TabIndex        =   14
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CheckBox chkDetail 
      BackColor       =   &H00000000&
      Caption         =   "Detail"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   4800
      Width           =   855
   End
   Begin VB.ListBox lstMain 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7695
   End
   Begin VB.CommandButton cmdOrders 
      BackColor       =   &H00E0E0E0&
      Caption         =   "View Orders..."
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrices 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Price Chart..."
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdDataFile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Load Data Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdGraph 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdPossCombos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Combos/ Specials"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSizeRate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pizza Size Rate"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPopSales 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pop Sales"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdToppingSales 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Toppings Sales"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPizzaFreq 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pizza Frequency..."
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalesRate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sales Rate"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label lblGraph 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6720
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   4680
      X2              =   4680
      Y1              =   4680
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   7920
      Y1              =   4680
      Y2              =   4680
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
Attribute VB_Name = "frmHeadOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Head Office form
'Performs various reports on the current data, as well as a sales rate graph


Option Explicit

'Form variables
    'Revenue, Sales and Tax amounts should be accessible by any procedure in the form, so those variables are up here
    Dim TOTALREV As Single, SREV(0 To 9) As Single 'Total revenue and revenue by store (revenue = sales + tax)
    Dim TOTALTAX As Single, STAX(0 To 9) As Single 'Total tax and tax by store
    Dim TOTALSALES As Single, SSALES(0 To 9) As Single 'Total sales and sales by store
    
'Clears the main listbox
Private Sub cmdClear_Click()
    lstMain.Clear
End Sub

'Loads the data files from the stores
Private Sub cmdDataFile_Click()
    Dim I As Integer
    
    'Clear the list and load the files
    lstMain.Clear
    ReadFiles
    
    'Display what stores were loaded
    For I = 0 To 9
        If STOREDATA(I).EXISTS Then
            lstMain.AddItem "Loaded data for Store #" & (I + 1)
        End If
    Next I
    lstMain.AddItem "" 'Add a blank line
    
    Call GoToEnd(lstMain) 'Select the last item in the list
End Sub

'Exits the program
Private Sub cmdExit_Click()
    Dim MSGRES As Integer
    
    'Ask if the user would like to save a report before exiting
    MSGRES = MsgBox("Do you want to save a report?", vbYesNoCancel, "Save report?")
    
    If MSGRES = vbYes Then Call cmdSaveReport_Click 'If Yes, then save the report
    If MSGRES = vbCancel Then Exit Sub 'If they click Cancel, just exit the sub altogether
    
    'End the program
    Unload Me
    End
End Sub

'Graphs the sales rates of each store
Private Sub cmdGraph_Click()

    Dim I As Integer 'For loop variable
    Dim VALUES(0 To 9) As Single 'The ten values to graph
    Dim MAXVALUE As Single 'The largest value (for drawing the graph at a nice scale that fits the whole box)
    Dim GX1 As Integer, GX2 As Integer, GY1 As Integer, GY2 As Integer 'Bounds of the graph
    Dim BX1 As Integer, BX2 As Integer, BY1 As Integer, BY2 As Integer 'Bounds of the current bar
    
    If cmdGraph.Caption = "Graph" Then
        'Toggle the graph visibility
        cmdGraph.Caption = "Hide"
        picGraph.Visible = True
        
        'Clear the picture
        picGraph.Cls
        
        'Draw a title on the graph
        picGraph.FontSize = 14
        picGraph.FontBold = True
        picGraph.ForeColor = vbBlue
        
            'Get the width of the title in pixels by putting it into a variable-width lable first
            lblGraph.Caption = "Sales Rates per store"
            picGraph.CurrentX = 0.8 * ((picGraph.Width / 2) - (lblGraph.Width / 2))
            picGraph.CurrentY = 150
            picGraph.Print lblGraph.Caption
            
        'Draw a bar for each store
        
            'Set the bounds of the graph
            GX1 = 150
            GX2 = picGraph.Width - 150
            GY1 = 600
            GY2 = picGraph.Height - 450
            picGraph.Line (GX1, GY1)-(GX2, GY2), vbBlack, B
            
            'Move GY1 down to make some whitespace for the number values
            GY1 = 900
            
            'Get the ten values
            cmdSalesRate_Click
            For I = 0 To 9
                VALUES(I) = SSALES(I)
            Next I
            
            'Get the max value
            For I = 0 To 9
                If VALUES(I) > MAXVALUE Then MAXVALUE = VALUES(I)
            Next I
            
            'Check if the MAXVALUE is 0, which will cause an error
            If MAXVALUE = 0 Then MAXVALUE = 1 'Set it to 1 to avoid the error
            
            'Graph the ten bars and amounts
            For I = 0 To 9
                
                'Set the bounds of the bar
                BX1 = GX1 + ((GX2 - GX1) / 10 * I) + 60
                BX2 = GX1 + ((GX2 - GX1) / 10 * (I + 1)) - 60
                BY2 = GY2 - 30
                BY1 = BY2 - ((VALUES(I) / MAXVALUE) * (GY2 - GY1 - 60))
                picGraph.Line (BX1, BY1)-(BX2, BY2), vbRed, BF
                
                'Draw the store number
                picGraph.FontSize = 10
                picGraph.FontBold = True
                picGraph.ForeColor = vbBlack
                picGraph.CurrentX = BX1 + ((BX2 - BX1) / 2) - 120
                picGraph.CurrentY = GY2 + 60
                picGraph.Print (I + 1)
                
                'Draw the amounts
                picGraph.FontSize = 8
                picGraph.ForeColor = vbRed
                picGraph.FontBold = False
                picGraph.CurrentX = BX1
                picGraph.CurrentY = BY1 - 240
                picGraph.Print ToMoney(VALUES(I))
                
            Next I
    Else
        'If they click it when it says Hide, then hide the graph
        cmdGraph.Caption = "Graph"
        picGraph.Visible = False
    End If
    
End Sub

'Display the View Orders form
Private Sub cmdOrders_Click()
    frmViewOrder.Show
End Sub

'Display the Toppings selection form
Private Sub cmdPizzaFreq_Click()
    frmToppings.Show
End Sub

'Displays the pop sales rates
Private Sub cmdPopSales_Click()
    Dim I As Integer, J As Integer, K As Integer, CT As Integer 'For loop and counter variables
    Dim POPCT(0 To 9, 0 To 3) As Integer 'The pop amounts for each store, and for each kind of pop
    
    'Add up the pop amounts for each store and kind of pop
    For I = 0 To 9
        If STOREDATA(I).EXISTS Then
            For J = 0 To UBound(STOREDATA(I).ORDERS)
                For K = 0 To 3
                    POPCT(I, K) = POPCT(I, K) + STOREDATA(I).ORDERS(J).POP(K)
                Next K
            Next J
        End If
    Next I
    
    'Display detailed results (from each store)
    If chkDetail Then
        For I = 0 To 9
            If STOREDATA(I).EXISTS Then
                lstMain.AddItem "  STORE #" & (I + 1)
                For J = 0 To 3
                    lstMain.AddItem "    " & POPNAMES(J) & ": " & POPCT(I, J) & " cans"
                Next J
            End If
        Next I
        lstMain.AddItem ""
    End If
    
    'Display the totals
    For I = 0 To 3 'For each kind
        CT = 0
        For J = 0 To 9 'For each store, calculate a sum
            CT = CT + POPCT(J, I)
        Next J
        lstMain.AddItem "  Total " & POPNAMES(I) & ": " & CT & " cans"
    Next I
    
    'Add a blank line and select the last line
    lstMain.AddItem ""
    Call GoToEnd(lstMain)
    
End Sub

'Displays the amounts combos (Specialty and slice/pop) for each store
Private Sub cmdPossCombos_Click()

    Dim I As Integer, J As Integer, K As Integer, CT As Integer 'For loop and counter variables
    Dim SPECIALCT(0 To 9, 0 To 2) As Integer 'Amount of special combos per special type, per store
    Dim SLICECT(0 To 9) As Integer 'Amount of slice/pop combos per store
    Dim TMPORDER As Order, TMPC As New StringCollection 'Temporary order and string collection
    
    'Parse the orders to get the specialcombo and slicecombo numbers
    For I = 0 To 9
        If STOREDATA(I).EXISTS Then
            For J = 0 To UBound(STOREDATA(I).ORDERS)
            
                'Make a copy of the order before parsing so that the original order isn't affected
                Call CopyOrder(STOREDATA(I).ORDERS(J), TMPORDER)
                Call ParseOrder(TMPORDER, TMPC)
                
                'Add to the slice and special counters
                SLICECT(I) = SLICECT(I) + TMPORDER.SLICECOMBOS
                For K = 0 To 2
                    SPECIALCT(I, K) = SPECIALCT(I, K) + TMPORDER.SPECIALCOMBOS(K)
                Next K
            
            Next J
        End If
    Next I
    
    'Display the results
    If chkDetail Then
        For I = 0 To 9
            If STOREDATA(I).EXISTS Then
                lstMain.AddItem "  STORE #" & (I + 1)
                For J = 0 To 2
                    lstMain.AddItem "    " & SPECIALNAMES(J) & " Pizzas: " & SPECIALCT(I, J)
                Next J
                lstMain.AddItem "    Slice-pop combos: " & SLICECT(I)
            End If
        Next I
        lstMain.AddItem ""
    End If
    
    'Display totals
    For I = 0 To 2
        CT = 0
        For J = 0 To 9
            CT = CT + SPECIALCT(J, I)
        Next J
        lstMain.AddItem "  Total " & SPECIALNAMES(I) & " Pizzas: " & CT
    Next I
    
    CT = 0
    For I = 0 To 9
        CT = CT + SLICECT(I)
    Next I
    lstMain.AddItem "  Total slice-pop combos: " & CT
    
    'Add a blank line and select the last line
    lstMain.AddItem ""
    Call GoToEnd(lstMain)
    
End Sub

'Display the price chart form
Private Sub cmdPrices_Click()
    frmPriceChart.Show
End Sub

'Calculates and displays various useful sale rates
Private Sub cmdSalesRate_Click()

    Dim I As Integer, J As Integer 'For loop variables
    
    'Reset the total revenue, tax and sales amounts
    TOTALREV = 0
    TOTALTAX = 0
    TOTALSALES = 0
    
    'Calculate and display total revenue
    For I = 0 To 9
        If STOREDATA(I).EXISTS Then
            
            'Reset the per-store sales rates
            SREV(I) = 0
            STAX(I) = 0
            SSALES(I) = 0
            
            'Calculate the per-store sales rates
            For J = 0 To UBound(STOREDATA(I).ORDERS)
                SREV(I) = SREV(I) + STOREDATA(I).ORDERS(J).TOTAL
                STAX(I) = STAX(I) + STOREDATA(I).ORDERS(J).TAX
                SSALES(I) = SSALES(I) + STOREDATA(I).ORDERS(J).PRICE
            Next J
            
            If chkDetail Then
                'Display the sales rates per store
                lstMain.AddItem "  STORE #" & (I + 1)
                lstMain.AddItem "    Total Revenue: " & ToMoney(SREV(I))
                lstMain.AddItem "    Total Tax: " & ToMoney(STAX(I))
                lstMain.AddItem "    Total Sales: " & ToMoney(SSALES(I))
                lstMain.AddItem ""
            End If
            
            'Add to the total sales variables
            TOTALREV = TOTALREV + SREV(I)
            TOTALTAX = TOTALTAX + STAX(I)
            TOTALSALES = TOTALSALES + SSALES(I)
            
        End If
    Next I
    
    'Display the total sales
    lstMain.AddItem "  TOTAL REVENUE: " & ToMoney(TOTALREV)
    lstMain.AddItem "  TOTAL TAX: " & ToMoney(TOTALTAX)
    lstMain.AddItem "  TOTAL SALES: " & ToMoney(TOTALSALES)
    
    'Add a blank line and select the last line
    lstMain.AddItem ""
    Call GoToEnd(lstMain)
    
End Sub

'Writes the contents of lstMain to a text file
Private Sub cmdSaveReport_Click()

    Dim FILENUM As Integer, PATH As String 'File number and path of the file
    Dim I As Integer 'For loop variable
    
    'Set the file path and file number
    PATH = App.PATH & "\Reports\" & Replace(Replace(Now, ":", "-"), "/", "-") & ".txt"
    FILENUM = FreeFile
    
    'Open the file for writing (Creates a new file automatically if it doesn't exist yet, which it shouldn't because the filename is set to the date and time that it's created at)
    Open PATH For Output As #FILENUM
    
        'Print each line of lstMain to the text file
        For I = 0 To lstMain.ListCount - 1
            Print #FILENUM, lstMain.List(I)
        Next I
        
    'Close the file
    Close #FILENUM
    
    'Tell the user that the file was saved and where it was saves
    Call MsgBox("Saved to """ & PATH & """", vbInformation, "Saved report")
    
End Sub

'Displays the pizza size rates per store
Private Sub cmdSizeRate_Click()

    Dim I As Integer, J As Integer, K As Integer 'For loop variables
    Dim SIZECT(0 To 9, 0 To 3) As Integer 'Amount of pizzas ordered per size, per store
    Dim SLICECT(0 To 9) As Integer 'Amount of slices ordered per store
    Dim CT As Integer 'Temporary counter variable
    
    For I = 0 To 9 'For each store
    
        If STOREDATA(I).EXISTS Then
            With STOREDATA(I)
                For J = 0 To UBound(.ORDERS)
                
                    'Pizza
                    If .ORDERS(J).DOPIZZA Then
                        For K = 0 To UBound(.ORDERS(J).PIZZADATA)
                            SIZECT(I, .ORDERS(J).PIZZADATA(K).SIZE) = SIZECT(I, .ORDERS(J).PIZZADATA(K).SIZE) + 1
                        Next K
                    End If
                    
                    'Slices
                    If .ORDERS(J).DOSLICES Then
                        For K = 0 To UBound(.ORDERS(J).SLICES)
                            SLICECT(I) = SLICECT(I) + 1
                        Next K
                    End If
                    
                Next J
            End With
        End If
    Next I
    
    'Display the data per store
    If chkDetail Then
    
        For I = 0 To 9
            If STOREDATA(I).EXISTS Then
                lstMain.AddItem "  STORE #" & (I + 1)
                For J = 0 To 3
                    lstMain.AddItem "    " & PIZZASIZES(J) & " pizzas: " & SIZECT(I, J)
                Next J
                lstMain.AddItem "    Slices: " & SLICECT(I)
            End If
        Next I
        lstMain.AddItem ""
    End If
    
    'Calculate and display the totals
    For I = 0 To 3
        CT = 0
        For J = 0 To 9
            CT = CT + SIZECT(J, I)
        Next J
        lstMain.AddItem "  Total " & PIZZASIZES(I) & " pizzas: " & CT
    Next I
    
    CT = 0
    For I = 0 To 9
        CT = CT + SLICECT(I)
    Next I
        
    lstMain.AddItem "  Total Slices: " & CT
    
    'Add a blank line and select the last line
    lstMain.AddItem ""
    Call GoToEnd(lstMain)
    
End Sub

'Selects the last line in a listbox
Public Sub GoToEnd(LST As ListBox)
    LST.ListIndex = LST.ListCount - 1
End Sub

'Displays topping rates per store
Private Sub cmdToppingSales_Click()
    Dim I As Integer, J As Integer, K As Integer, L As Integer 'For loop variables
    Dim CT As Integer 'Counter variable
    Dim TOPS(0 To 9, 0 To 7) As Integer 'Topping amounts per type, per store
    Dim HALFTOPS(0 To 9, 0 To 7) As Integer 'Halftopping amounts per type, per store
    Dim SLICETOPS(0 To 9, 0 To 7) As Integer 'Slice topping amounts per type, per store
    
    'For each store
    For I = 0 To 9
        
        If STOREDATA(I).EXISTS Then
            With STOREDATA(I)
                For J = 0 To UBound(.ORDERS)
                    
                    'Pizza
                    If .ORDERS(J).DOPIZZA Then
                        For K = 0 To UBound(.ORDERS(J).PIZZADATA)
                            For L = 0 To 7
                            
                                'Full toppings
                                If .ORDERS(J).PIZZADATA(K).TOPS(L) = 1 Then
                                    TOPS(I, L) = TOPS(I, L) + 1
                                End If
                                
                                If .ORDERS(J).PIZZADATA(K).FULLMODE = False And .ORDERS(J).PIZZADATA(K).TOPS2(L) = 1 Then
                                    HALFTOPS(I, L) = HALFTOPS(I, L) + 1
                                End If
                            Next L
                        Next K
                    End If
                    
                    'Slices
                    If .ORDERS(J).DOSLICES Then
                        For K = 0 To UBound(.ORDERS(J).SLICES)
                            For L = 0 To 7
                                If .ORDERS(J).SLICES(K).TOPS(L) = 1 Then
                                    SLICETOPS(I, L) = SLICETOPS(I, L) + 1
                                End If
                            Next L
                        Next K
                    End If
                    
                Next J
            End With
        End If
    Next I
    
    'Display the data per store
    If chkDetail Then
        For I = 0 To 9
            If STOREDATA(I).EXISTS Then
            
                lstMain.AddItem "  STORE #" & (I + 1)
                
                'Full toppings
                lstMain.AddItem "    Full toppings"
                For J = 0 To 7
                    lstMain.AddItem "      " & TOPNAMES(J) & ": " & TOPS(I, J) & " units"
                Next J
                
                'Half-toppings
                lstMain.AddItem "    Half-toppings"
                For J = 0 To 7
                    lstMain.AddItem "      " & TOPNAMES(J) & ": " & HALFTOPS(I, J) & " units"
                Next J
        
                'Slice toppings
                lstMain.AddItem "    Slice toppings"
                For J = 0 To 7
                    lstMain.AddItem "      " & TOPNAMES(J) & ": " & SLICETOPS(I, J) & " units"
                Next J
                
                lstMain.AddItem ""
            
            End If
        Next I
        
    End If
    
    'Display the totals
    For I = 0 To 7
        CT = 0
        For J = 0 To 9
            CT = CT + TOPS(J, I) + HALFTOPS(J, I) + SLICETOPS(J, I)
        Next J
        lstMain.AddItem "  Total " & TOPNAMES(I) & ": " & CT & " units"
    Next I
    
    'Add a blank line and select the last line
    lstMain.AddItem ""
    Call GoToEnd(lstMain)
    
End Sub

'Called when the form is loaded
Private Sub Form_Load()
    cmdDataFile_Click 'Load the data files
End Sub

'Called when the form is unloaded
Private Sub Form_Unload(Cancel As Integer)
    'Close all other open forms
    Unload frmPriceChart
    Unload frmToppings
    Unload frmViewOrder
End Sub

'Exits the program
Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

'Logs out of the head office form
Private Sub mnuLogout_Click()
    Dim MSGRES As Integer
    
    'Confirm that the user wants to log out
    MSGRES = MsgBox("Are you sure you want to log out?", vbOKCancel, "Logout")
    If MSGRES = vbCancel Then Exit Sub 'Exit the sub if they click No
    
    'Ask the user if they would like to save a report
    MSGRES = MsgBox("Do you want to save a report?", vbYesNoCancel, "Save report?")
    If MSGRES = vbYes Then Call cmdSaveReport_Click
    If MSGRES = vbCancel Then Exit Sub
    
    'Close this form and load the intro form
    Unload Me
    frmIntro.Show
    Call Main 'Reload the main variables from main.txt, for use in the cashier program
    
End Sub
