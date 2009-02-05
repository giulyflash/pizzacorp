Attribute VB_Name = "modStuff"
'The central module that houses many useful subroutines


Option Explicit

'Global Variables
    'Various strings
    Public PIZZASIZES(0 To 3) As String
    Public POPNAMES(0 To 3) As String
    Public TOPNAMES(0 To 7) As String
    Public SPECIALTOPS(0 To 2, 0 To 7) As Integer
    Public SPECIALNAMES(0 To 2) As String

    'Various prices
    Public POPCOST As Single                   '$1.00 for a can of pop
    Public SLICECOST As Single                 '$2.00 for a slice of pizza
    Public SLICECOMBO As Single                '$0.50 discount for a pop-slice combo
    Public SPECIALCOST(0 To 3) As Single       'Prices for the four specialty pizza sizes
    Public PIZZACOST(0 To 3) As Single         'Prices for the four sizes of pizza
    Public TOPCOST(0 To 3) As Single           'Prices for the four sizes of toppings
    Public HALFTOPCOST(0 To 3) As Single       'Prices for the four sizes of half-toppings
    Public TAXAMT As Single
    
'Global types
Public Type Pizza
    FULLMODE As Boolean         'TRUE=full pizza, FALSE=half-half pizza
    TOPS(0 To 7) As Integer     'Toppings for the first half
    TOPS2(0 To 7) As Integer    'Toppings for the second half
    SIZE As Integer             '0=Small, 1=Medium, 2=Large, 3=XLarge
End Type

Public Type Slice
    TOPS(0 To 7) As Integer     'Toppings for the slice
End Type

Public Type Order
    PIZZADATA() As Pizza        'Array of pizzas
    DOPIZZA As Boolean          'TRUE=pizza was ordered, FALSE=pizza was not ordered
    SLICES() As Slice           'Array of slices
    DOSLICES As Boolean         'Ditto for DOPIZZA
    POP(0 To 3) As Integer      'Amount of pop ordered (4 different kinds)
    ID As String                'Date and time of the order
    PRICE As Single             'Total price of the order (before tax)
    TAX As Single               'Amount of tax calculated from price
    TOTAL As Single             'Total (price + tax)
    CASH As Single              'Cash given to cashier
    CHANGE As Single            'Change given back to customer
    SPECIALCOMBOS(0 To 2) As Integer 'Used for head office analysis
    SLICECOMBOS As Integer           'Used for head office analysis
End Type

Public Type Store
    ORDERS() As Order  'Each store has orders
    EXISTS As Boolean 'Each store also exists or doesn't exist
End Type

'Global Variables
Public ORDERS() As Order 'Orders for the current store in the cashier program
Public STOREDATA(0 To 9) As Store 'Data for each store in the head office program
Public CURRENTSTORE As Integer 'The current store if in cashier mode

'Main sub that initializes variables
'call Main()
Public Sub Main()

    Dim I As Integer 'For loop variable
    
    'Load the names and prices from the main data file
    LoadNames
    
    'Reset the orders arrays
    ReDim ORDERS(0 To 0)
    ReDim ORDERS(0).PIZZADATA(0 To 0)
    ReDim ORDERS(0).SLICES(0 To 0)
    ORDERS(0).ID = Now
    ORDERS(0).PIZZADATA(0).FULLMODE = True
    ORDERS(0).DOPIZZA = True
    
    'Reset the store data arrays
    For I = 0 To 9
        ReDim STOREDATA(I).ORDERS(0 To 0)
        ReDim STOREDATA(I).ORDERS(0).PIZZADATA(0 To 0)
        ReDim STOREDATA(I).ORDERS(0).SLICES(0 To 0)
    Next I
    
End Sub

'Loads various names (pizza sizes, topping names, etc.) from the main.txt file
'call LoadNames()
Public Sub LoadNames()

    'If there are any errors, go to the BadFile handler, because the file either doesn't exist or is corrupt
    On Error GoTo BadFile
    
    'Read everything from a file
    Dim PATH As String, FILENUM As Integer, I As Integer, J As Integer, S As String
    
    FILENUM = FreeFile                 'Grab an unused file number
    PATH = App.PATH & "\Data\main.txt" 'Set the path of the file
    Open PATH For Input As #FILENUM    'Open the file
    
        'Pizza sizes
        For I = 0 To 3
            Line Input #FILENUM, PIZZASIZES(I)
        Next I
        
        'Pop names
        For I = 0 To 3
            Line Input #FILENUM, POPNAMES(I)
        Next I
        
        'Topping names
        For I = 0 To 7
            Line Input #FILENUM, TOPNAMES(I)
        Next I
        
        'Specialty pizza data
        For I = 0 To 2
            Line Input #FILENUM, SPECIALNAMES(I)
            
            'Specialty toppings
            Line Input #FILENUM, S
            For J = 0 To 7
                SPECIALTOPS(I, J) = Val(Mid(S, J + 1, 1))
            Next J
        Next I
        
        'Pizza costs
        Line Input #FILENUM, S
        For I = 0 To 3
            PIZZACOST(I) = Val(Split(S, " ")(I))
        Next I
        
        'Topping costs
        Line Input #FILENUM, S
        For I = 0 To 3
            TOPCOST(I) = Val(Split(S, " ")(I))
        Next I
        
        'Half topping costs
        Line Input #FILENUM, S
        For I = 0 To 3
            HALFTOPCOST(I) = Val(Split(S, " ")(I))
        Next I
        
        'Specialty pizza costs
        Line Input #FILENUM, S
        For I = 0 To 3
            SPECIALCOST(I) = Val(Split(S, " ")(I))
        Next I
        
        'Pop, slice, combo discount and tax costs
        Line Input #FILENUM, S
        POPCOST = Val(S)
        Line Input #FILENUM, S
        SLICECOST = Val(S)
        Line Input #FILENUM, S
        SLICECOMBO = Val(S)
        Line Input #FILENUM, S
        TAXAMT = Val(S)
    
    Close #FILENUM
    
    Exit Sub 'Avoid the BadFile handler

'Display an error message if there was an error loading the main data file
BadFile:

    Call MsgBox("There was an error loading the main.txt data file!" & vbNewLine & "Make sure that main.txt exists in the Data directory.", vbCritical, "Error loading database")
    End
    
End Sub

'Adds another element to an Order array
'call AddOrder(order array)
Public Sub AddOrder(O() As Order)

    'Redimension the array
    ReDim Preserve O(0 To UBound(O) + 1)
    
    'Reset the variables inside the newest element
    O(UBound(O)).ID = Now 'Give it an ID
    
    ReDim O(UBound(O)).PIZZADATA(0 To 0) 'Reset its pizza data array
    O(UBound(O)).PIZZADATA(0).FULLMODE = True
    O(UBound(O)).DOPIZZA = True
    
    ReDim O(UBound(O)).SLICES(0 To 0) 'Reset its slices data array
End Sub

'Deletes an element from an Order array
'call DelOrder(order array, index of order to remove)
Public Sub DelOrder(O() As Order, Index As Integer)
    'Bump everything down
    Dim I As Integer, J As Integer, K As Integer 'For loop variables
    If Index < UBound(O) Then 'Only have to bump it down if it's not the last element being deleted
        For I = Index To UBound(O) - 1
            Call CopyOrder(O(I + 1), O(I))
        Next I
    End If
    
    'Redimension the array to have one less element
    ReDim Preserve O(0 To UBound(O) - 1)
End Sub

'Copies the contents of one Order to another
'call CopyOrder(input order, output order)
Public Sub CopyOrder(O1 As Order, ByRef O2 As Order)
    Dim J As Integer, K As Integer 'For loop variables
    
    'Copy the money amounts and ID
    O2.CASH = O1.CASH
    O2.CHANGE = O1.CHANGE
    O2.DOPIZZA = O1.DOPIZZA
    O2.DOSLICES = O1.DOSLICES
    O2.ID = O1.ID
    O2.PRICE = O1.PRICE
    O2.TAX = O1.TAX
    O2.TOTAL = O1.TOTAL
            
    'Copy the pop data
    For J = 0 To 3
        O2.POP(J) = O1.POP(J)
    Next J
            
    'Copy the pizza data
    ReDim O2.PIZZADATA(0 To UBound(O1.PIZZADATA)) 'Redim the pizza array to match
    For J = 0 To UBound(O2.PIZZADATA)
        O2.PIZZADATA(J).FULLMODE = O1.PIZZADATA(J).FULLMODE
        O2.PIZZADATA(J).FULLMODE = O1.PIZZADATA(J).FULLMODE
        O2.PIZZADATA(J).SIZE = O1.PIZZADATA(J).SIZE
        For K = 0 To 7
            O2.PIZZADATA(J).TOPS(K) = O1.PIZZADATA(J).TOPS(K)
            O2.PIZZADATA(J).TOPS2(K) = O1.PIZZADATA(J).TOPS2(K)
        Next K
    Next J
            
    'Copy the slices data
    ReDim O2.SLICES(0 To UBound(O1.SLICES))
    For J = 0 To UBound(O2.SLICES)
        For K = 0 To 7
            O2.SLICES(J).TOPS(K) = O1.SLICES(J).TOPS(K)
        Next K
    Next J
    
End Sub

'Adds another element to a Pizza array
'call AddPizza(order element)
Public Sub AddPizza(O As Order)
    ReDim Preserve O.PIZZADATA(0 To UBound(O.PIZZADATA) + 1)
    O.PIZZADATA(UBound(O.PIZZADATA)).FULLMODE = True
End Sub

'Deletes an element from a Pizza array
'call DelPizza(order element, index to remove)
Public Sub DelPizza(O As Order, Index As Integer)
    'Bump everything down
    Dim I As Integer, J As Integer
    If Index < UBound(O.PIZZADATA) Then
        For I = Index To UBound(O.PIZZADATA) - 1
            O.PIZZADATA(I).FULLMODE = O.PIZZADATA(I + 1).FULLMODE
            O.PIZZADATA(I).SIZE = O.PIZZADATA(I + 1).SIZE
            For J = 0 To 7
                O.PIZZADATA(I).TOPS(J) = O.PIZZADATA(I + 1).TOPS(J)
                O.PIZZADATA(I).TOPS2(J) = O.PIZZADATA(I + 1).TOPS2(J)
            Next J
        Next I
    End If
    ReDim Preserve O.PIZZADATA(0 To UBound(O.PIZZADATA) - 1)
End Sub

'Adds another element to a Slice array
'call AddSlice(order element)
Public Sub AddSlice(O As Order)
    ReDim Preserve O.SLICES(0 To UBound(O.SLICES) + 1)
End Sub

'Deletes an element from a Slice array
'call DelSlice(order element, index to remove)
Public Sub DelSlice(O As Order, Index As Integer)
    'Bump everything down
    Dim I As Integer, J As Integer
    If Index < UBound(O.SLICES) Then
        For I = Index To UBound(O.SLICES) - 1
            For J = 0 To 7
                O.SLICES(I).TOPS(J) = O.SLICES(I + 1).TOPS(J)
            Next J
        Next I
    End If
    ReDim Preserve O.SLICES(0 To UBound(O.SLICES) - 1)
End Sub

'Returns a string dollar representation of a Single money amount
'variable = ToMoney("123456.78") 'Will return "$123,456.78"
Public Function ToMoney(Amt As Single, Optional Pad As Integer) As String
    ToMoney = Right(String(Pad, " ") & Format(Amt, "$###,###,##0.00"), IIf(Pad > Len(Format(Amt, "$###,###,##0.00")), Pad, Len(Format(Amt, "$###,###,##0.00"))))
End Function

'Parses an order into a receipt, stores results into collection C
'call ParseOrder(order name, output string collection)
Public Sub ParseOrder(O As Order, ByRef C As StringCollection)

    Dim I As Integer, J As Integer, K As Integer 'For loop variables
    Dim FLAG As Integer 'Needed for determining if a pizza is a specialty
    Dim POPCOMBOS As Single, NUMPOPS As Single, NUMSLICES As Single 'Used to determine the number of slice/pop combos
    Dim SPECTOPCT As Integer 'Temporary counter for the number of toppings on a specialty pizza
    
    On Error GoTo BadMark 'This happens when Mark messes around
    
    'With O makes it easier to access its properties
    With O
        
        'Reset the price and combo counts
        .PRICE = 0
        .SLICECOMBOS = 0
        For I = 0 To 2
            .SPECIALCOMBOS(I) = 0
        Next I
        
        'Parse the PIZZA
        If .DOPIZZA Then
        
            'Calculate the pizza costs
            FLAG = 0
                For J = 0 To UBound(.PIZZADATA) 'For each ordered pizza
                    
                    'Check for specials
                    For I = 0 To 2 'For each available special
                    
                        FLAG = 0
                        
                        'Only compare the toppings if it's a full pizza,
                        'because half-half pizzas can't be specials
                        If .PIZZADATA(J).FULLMODE Then
                        
                            'See if the toppings for that pizza match the ones for the special
                            
                            SPECTOPCT = 0
                            
                            For K = 0 To 7 'For each topping
                            
                                'Compare each topping
                                If Not .PIZZADATA(J).TOPS(K) = SPECIALTOPS(I, K) Then
                                
                                    'If it finds even one topping that doesn't match, set a flag
                                    FLAG = 1
                                    Exit For
                                    
                                End If
                                
                                If SPECIALTOPS(I, K) = 1 Then SPECTOPCT = SPECTOPCT + 1
                            Next K
                        Else
                            FLAG = 1 'Invalid special if it's half-half
                        End If
                        
                        If FLAG = 0 Then
                            'If it finds a special, process it and exit the loop
                            .PRICE = .PRICE + SPECIALCOST(.PIZZADATA(J).SIZE) + (SPECTOPCT * TOPCOST(.PIZZADATA(J).SIZE))
                            .SPECIALCOMBOS(I) = .SPECIALCOMBOS(I) + 1
                            
                            C.Add Left("SPECIAL " & UCase(PIZZASIZES(.PIZZADATA(J).SIZE)) & " " & UCase(SPECIALNAMES(I)) & String(25, " "), 25) & ToMoney(SPECIALCOST(.PIZZADATA(J).SIZE) + (4 * TOPCOST(.PIZZADATA(J).SIZE)), 10)
                            
                            Exit For
                        End If
                        
                    Next I
                    
                        If FLAG = 1 Then
                        
                            'It's just a regular pizza
                            .PRICE = .PRICE + PIZZACOST(.PIZZADATA(J).SIZE)
                            C.Add Left(UCase(PIZZASIZES(.PIZZADATA(J).SIZE)) & IIf(.PIZZADATA(J).FULLMODE = True, " FULL PIZZA", " HALF-HALF PIZZA") & String(25, " "), 25) & ToMoney(PIZZACOST(.PIZZADATA(J).SIZE), 10)
                            
                            'If it's a full pizza, add up the price for its toppings
                            If .PIZZADATA(J).FULLMODE Then
                                For K = 0 To 7
                                    If .PIZZADATA(J).TOPS(K) = 1 Then
                                        .PRICE = .PRICE + TOPCOST(.PIZZADATA(J).SIZE)
                                        C.Add Left("  " & UCase(TOPNAMES(K)) & String(25, " "), 25) & ToMoney(TOPCOST(.PIZZADATA(J).SIZE), 10)
                                    End If
                                Next K
                            Else
                                'Divide topping costs by two for half-half pizzas
                                'The first half
                                C.Add "  HALF 1:"
                                For K = 0 To 7
                                    If .PIZZADATA(J).TOPS(K) = 1 Then
                                        .PRICE = .PRICE + HALFTOPCOST(.PIZZADATA(J).SIZE)
                                        C.Add Left("    " & UCase(TOPNAMES(K)) & String(25, " "), 25) & ToMoney(HALFTOPCOST(.PIZZADATA(J).SIZE), 10)
                                    End If
                                Next K
                                
                                'The second half
                                C.Add "  HALF 2:"
                                For K = 0 To 7
                                    If .PIZZADATA(J).TOPS2(K) = 1 Then
                                        .PRICE = .PRICE + HALFTOPCOST(.PIZZADATA(J).SIZE)
                                        C.Add Left("    " & UCase(TOPNAMES(K)) & String(25, " "), 25) & ToMoney(HALFTOPCOST(.PIZZADATA(J).SIZE), 10)
                                    End If
                                Next K
                                
                            End If
                            
                        End If
                Next J
        End If
        
        'SLICES and POP
        NUMSLICES = 0
        NUMPOPS = 0
        
        'Add up the number of slices and add it to the price
        If .DOSLICES Then
            .PRICE = .PRICE + ((UBound(.SLICES) + 1) * SLICECOST)
            NUMSLICES = UBound(.SLICES) + 1
        End If
        
        'Add up the number of pops and add it to the price
        For I = 0 To 3
            .PRICE = .PRICE + (POPCOST * .POP(I))
            NUMPOPS = NUMPOPS + .POP(I)
        Next I
        
        'Display the slices and their toppings
        If .DOSLICES Then
            C.Add "SLICE   Qty: " & NUMSLICES
            For I = 0 To UBound(.SLICES)
                C.Add "  SLICE"
                For J = 0 To 7
                    If .SLICES(I).TOPS(J) = 1 Then C.Add "    " & UCase(TOPNAMES(J))
                Next J
            Next I
        End If

        'Display the pops
        If NUMPOPS > 0 Then C.Add "POP   Qty: " & NUMPOPS
        For I = 0 To 3
            If .POP(I) > 0 Then C.Add "  " & UCase(POPNAMES(I)) & "   Qty: " & .POP(I)
        Next I
        
        'Combos; subtract slice-pop combo discounts from the price
        POPCOMBOS = IIf(NUMSLICES - NUMPOPS < 0, NUMSLICES, NUMPOPS)
        .SLICECOMBOS = POPCOMBOS 'Store the combo number to the SLICECOMBOS variable
        
        'If there are slice/pop combos...
        If POPCOMBOS > 0 Then
            .PRICE = .PRICE - (POPCOMBOS * SLICECOMBO) 'Make the discount
            
            'Display the combos and extra slices/pop
            C.Add ""
            C.Add Left("SLICE-POP COMBO   Qty: " & POPCOMBOS & String(25, " "), 25) & ToMoney(POPCOMBOS * (SLICECOST + POPCOST - SLICECOMBO), 10)
            If (NUMSLICES - POPCOMBOS) > 0 Then C.Add Left("EXTRA SLICES   Qty: " & (NUMSLICES - POPCOMBOS) & String(25, " "), 25) & ToMoney((NUMSLICES - POPCOMBOS) * SLICECOST, 10)
            If (NUMPOPS - POPCOMBOS) > 0 Then C.Add Left("EXTRA POP   Qty: " & (NUMPOPS - POPCOMBOS) & String(25, " "), 25) & ToMoney((NUMPOPS - POPCOMBOS) * POPCOST, 10)
        Else
            'If there's no combos, just display the slices and pop totals
            C.Add ""
            If NUMSLICES > 0 Then C.Add "SLICE TOTAL              " & ToMoney(NUMSLICES * SLICECOST, 10)
            If NUMPOPS > 0 Then C.Add "POP TOTAL                " & ToMoney(NUMPOPS * POPCOST, 10)
        End If
        
        'Calculate and display the price
        .PRICE = Round(.PRICE, 2) 'Round everything to 2 decimal places to avoid calculation errors later
        .TAX = Round(.PRICE * TAXAMT, 2)
        .TOTAL = .PRICE + .TAX
        C.Add ""
        C.Add "SUBTOTAL                 " & ToMoney(.PRICE, 10)
        C.Add "TAX " & (TAXAMT * 100) & "%                  " & ToMoney(.TAX, 10)
        C.Add "TOTAL                    " & ToMoney(.TOTAL, 10)
    
    End With
    
    Exit Sub
    
BadMark:

    Call MsgBox("Mark, stop messing around...", vbCritical, "Overflow error")
    
End Sub

'Writes the given O() array to the specified store data file
'call UpdateFile(store number, orders array)
Public Sub UpdateFile(STORENUM As Integer, O() As Order)

    Dim PATH As String, FILENUM As Integer 'Pathname and filenumber variables
    Dim I As Integer, J As Integer, K As Integer 'For loop variables
    Dim BUILD As String, S As String 'Temporary output strings
    
    'Set the path and file number
    PATH = App.PATH & "\Data\store" & (STORENUM + 1) & ".txt"
    FILENUM = FreeFile
    
    'Open the file for writing
    Open PATH For Output As #FILENUM
    
        'Write each order to the file
        For I = 0 To UBound(O)
            With O(I)
            
                'We have to divide everything with a single space so that things remain constant when reading the file back in
                
                'Skip the order if there's no price (it isn't even a real order then)
                If .PRICE = 0 Then GoTo SkipOrder
                
                'First line: ID (date and time), cash given by customer, change given to customer, price, tax, and total
                Print #FILENUM, Replace(.ID, " ", "_") & _
                    " " & .CASH & " " & .CHANGE & " " & .PRICE & _
                    " " & .TAX & " " & .TOTAL
                    
                'Second line: True/False (if pizza was ordered or not), number of pizzas ordered,
                '             True/False (if slices were ordered or not), number of slices,
                '             amount of the four kinds of pop
                Print #FILENUM, .DOPIZZA & " " & UBound(.PIZZADATA) & _
                    " " & .DOSLICES & " " & UBound(.SLICES) & " " & _
                    .POP(0) & " " & .POP(1) & " " & .POP(2) & " " & .POP(3)
                
                'Write the pizza data
                BUILD = ""
                If .DOPIZZA Then
                    For J = 0 To UBound(.PIZZADATA)
                        BUILD = BUILD & -CInt(.PIZZADATA(J).FULLMODE) 'Negate it to remove the negative sign when true
                        BUILD = BUILD & .PIZZADATA(J).SIZE
                        For K = 0 To 7 'Writing the toppings interleaved (1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2, 1, 2)
                            BUILD = BUILD & .PIZZADATA(J).TOPS(K)
                            BUILD = BUILD & .PIZZADATA(J).TOPS2(K)
                        Next K
                    Next J
                End If
                Print #FILENUM, BUILD
                
                'Write the slices data
                BUILD = ""
                If .DOSLICES Then
                    For J = 0 To UBound(.SLICES)
                        For K = 0 To 7
                            BUILD = BUILD & .SLICES(J).TOPS(K)
                        Next K
                    Next J
                End If
                Print #FILENUM, BUILD
                
SkipOrder:
                
            End With
        Next I
        
    'Close the file
    Close #FILENUM
    
End Sub

'Reads in the orders from a single store
'call ReadFile(number of the store, orders array to output to)
Public Sub ReadFile(STORENUM As Integer, ByRef O() As Order)
    
    Dim J As Integer, K As Integer, L As Integer 'For loop variables
    Dim CT As Integer 'Counts the lines in the file
    Dim FILENUM As Integer, PATH As String 'Path and filenumber of the file
    Dim FILELINES() As String, NUMORDERS As Integer 'The lines of text of the file, and the calculated number of orders
    Dim CURLINE1 As Variant, CURLINE2 As Variant, CURLINE3 As String, CURLINE4 As String 'The four lines of text for each order (temporary)
    ReDim FILELINES(0 To 0) 'Filelines is a variable-length array, so we need to redimension it first
    
        'Set the path and filenumber
        PATH = App.PATH & "\Data\store" & (STORENUM + 1) & ".txt"
        
        'Make sure the file ecists before trying to open it
        If FileExists(PATH) Then
            
            FILENUM = FreeFile
            
            'Open the file
            Open PATH For Input As #FILENUM
            
                'Redimension the file lines again, and reset the counter
                ReDim FILELINES(0 To 0)
                CT = 0
                
                'Read in the lines of text from the file
                Do While Not EOF(FILENUM)
                    ReDim Preserve FILELINES(0 To CT)
                    Line Input #FILENUM, FILELINES(CT)
                    CT = CT + 1
                Loop
                CT = CT - 1 'Subtracting 1 will make CT divisible by 4
                
                'There aren't any lines of text in the file
                If CT = -1 Then
                    
                    'Delete the file because it shouldn't exist in the first place
                    Close #FILENUM
                    Kill PATH
                    
                    'No point in continuing the sub, so exit it
                    Exit Sub
                End If
                
                'Parse the data
                
                'Calculate the number of orders, and redimension the output orders array
                NUMORDERS = (CT + 1) / 4
                ReDim O(0 To (NUMORDERS - 1))
                
                For J = 0 To NUMORDERS - 1
                
                    'Split the first two lines of text for the order into arrays
                    CURLINE1 = Split(FILELINES(J * 4), " ")
                    CURLINE2 = Split(FILELINES(J * 4 + 1), " ")
                    
                    'Load the second two lines
                    CURLINE3 = FILELINES(J * 4 + 2)
                    CURLINE4 = FILELINES(J * 4 + 3)
                    
                    With O(J)
                        
                        'The first line has the ID, cash, change, price, tax and total,
                        'all separated by single spaces
                        .ID = Replace(CURLINE1(0), "_", " ")
                        .CASH = Val(CURLINE1(1))
                        .CHANGE = Val(CURLINE1(2))
                        .PRICE = Val(CURLINE1(3))
                        .TAX = Val(CURLINE1(4))
                        .TOTAL = Val(CURLINE1(5))
                        
                        'The second line has TRUE/FALSE for pizza and slices,
                        'the number of pizzas and slices, and the amount of pops
                        'Parse the pizza/slice headers
                        .DOPIZZA = CBool(CURLINE2(0))
                        ReDim .PIZZADATA(0 To CInt(CURLINE2(1)))
                        .DOSLICES = CBool(CURLINE2(2))
                        ReDim .SLICES(0 To CInt(CURLINE2(3)))
                        
                        'Read in the amount of pops
                        For K = 0 To 3
                            .POP(K) = CInt(CURLINE2(4 + K))
                        Next K
                        
                        'Parse the pizzadata if it exists
                        If .DOPIZZA Then
                            For K = 0 To UBound(.PIZZADATA)
                                .PIZZADATA(K).FULLMODE = CBool(-CInt(Mid(CURLINE3, 1 + (K * 18), 1)))
                                .PIZZADATA(K).SIZE = CInt(Mid(CURLINE3, 2 + (K * 18), 1))
                                
                                For L = 0 To 7
                                    .PIZZADATA(K).TOPS(L) = CInt(Mid(CURLINE3, 3 + (K * 18) + (L * 2), 1))
                                    .PIZZADATA(K).TOPS2(L) = CInt(Mid(CURLINE3, 4 + (K * 18) + (L * 2), 1))
                                Next L
                            Next K
                        End If
                        
                        'Parse the slice data if it exists
                        If .DOSLICES Then
                            For K = 0 To UBound(.SLICES)
                                For L = 0 To 7
                                    .SLICES(K).TOPS(L) = CInt(Mid(CURLINE4, 1 + (K * 8), 1))
                                Next L
                            Next K
                        End If
                        
                    End With
                    
                Next J
                
            Close #FILENUM
            
        End If
        
End Sub

'Reads in the data files from all of the stores, puts it into the STOREDATA() array
'call ReadFiles()
Public Sub ReadFiles()

    Dim I As Integer, J As Integer, K As Integer, L As Integer 'For-loop variables
    Dim CT As Integer 'Loop counter variable
    Dim FILENUM As Integer, PATH As String 'The pathname and filenumber of the file
    Dim FILELINES() As String, NUMORDERS As Integer
    Dim CURLINE1 As Variant, CURLINE2 As Variant, CURLINE3 As String, CURLINE4 As String
    ReDim FILELINES(0 To 0)
    
    'Read each of the ten store files
    For I = 0 To 9
    
        'Set the pathname of the file
        PATH = App.PATH & "\Data\store" & (I + 1) & ".txt"
        
        'See if the file exists
        If FileExists(PATH) Then
        
            'If it exists, read the file
            STOREDATA(I).EXISTS = True
            Call ReadFile(I, STOREDATA(I).ORDERS)
            
        Else
        
            'The store file doesn't exist yet
            STOREDATA(I).EXISTS = False
            
        End If
        
    Next I
    
End Sub

'Returns true or false, whether or not a file of PATH exists
'variable = FileExists(path of file)
Public Function FileExists(PATH As String) As Boolean
    Dim FILENUM As Integer 'The file number
    
    'If there's an error, the file doesn't exist or it's already open
    On Error GoTo BadFile
    
    'Try opening and closing the file
    FILENUM = FreeFile
    Open PATH For Input As #FILENUM
    Close #FILENUM
    
    'If the file was successfully opened, the file exists
    FileExists = True
    Exit Function 'Exit the function to avoid the BadFile handler
    
BadFile:
    FileExists = False

End Function

'Writes the current set of names of things to the main.txt data file
'call UpdateNames
Public Sub UpdateNames()

    Dim I As Integer, J As Integer, S As String 'For loops and temporary string variables
    Dim FILENUM As Integer, PATH As String
    
    'Set the file number and path for writing
    FILENUM = FreeFile
    PATH = App.PATH & "\Data\main.txt"
    
    'Open the file for writing
    Open PATH For Output As #FILENUM
    
        'Size names
        For I = 0 To 3
            Print #FILENUM, PIZZASIZES(I)
        Next I
        
        'Pop names
        For I = 0 To 3
            Print #FILENUM, POPNAMES(I)
        Next I
        
        'Topping names
        For I = 0 To 7
            Print #FILENUM, TOPNAMES(I)
        Next I
        
        'Specialty pizzas
        For I = 0 To 2
            Print #FILENUM, SPECIALNAMES(I)
            
            'Build the toppings string
            S = ""
            For J = 0 To 7
                S = S & SPECIALTOPS(I, J)
            Next J
            Print #FILENUM, S
        Next I
        
        'Pizza costs
        S = ""
        For I = 0 To 3
            S = S & " " & PIZZACOST(I)
        Next I
        Print #FILENUM, Mid(S, 2) 'Start after the first space
        
        'Topping costs
        S = ""
        For I = 0 To 3
            S = S & " " & TOPCOST(I)
        Next I
        Print #FILENUM, Mid(S, 2)
        
        'Half-topping costs
        S = ""
        For I = 0 To 3
            S = S & " " & HALFTOPCOST(I)
        Next I
        Print #FILENUM, Mid(S, 2)
        
        'Specialty pizza costs
        S = ""
        For I = 0 To 3
            S = S & " " & SPECIALCOST(I)
        Next I
        Print #FILENUM, Mid(S, 2)
        
        'Write the last four prices
        Print #FILENUM, CStr(POPCOST)
        Print #FILENUM, CStr(SLICECOST)
        Print #FILENUM, CStr(SLICECOMBO)
        Print #FILENUM, CStr(TAXAMT)
        
    'Close the file
    Close #FILENUM
    
End Sub
