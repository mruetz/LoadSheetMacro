Attribute VB_Name = "Module3"

Public bothSideThreshold As Integer

Public FIT_FACTOR As Integer ' The col with greatest stack (stack=0 if impossible)
        ' with existing partial behind + (this many) point PER LAYER of full stack
        
Public MIN_BW_FACTOR As Integer ' The col with a BW at the end LOSES points = (this)*5-(layers in partial)
        ' if a stack with this partial is impossible
     
Public EVEN_FACTOR As Integer ' The col that is the shortest + (this many) point for every space
        ' it is shorter by
        
Sub AutoBuild()

    Application.DisplayAlerts = False
    Worksheets("26 Pallets").Activate

    If Not rangeIsEmpty("X5", "X7") Then

        If Not IsEmpty(Range("X5")) And IsNumeric(Range("X5").Value) Then
            FIT_FACTOR = CInt(Range("X5").Value)
            MsgBox ("FIT: " & FIT_FACTOR)
        Else
            FIT_FACTOR = 1
        End If
        
        If Not IsEmpty(Range("X6")) And IsNumeric(Range("X6").Value) Then
            MIN_BW_FACTOR = CInt(Range("X6").Value)
            MsgBox ("MIN_BW: " & MIN_BW_FACTOR)
        Else
            MIN_BW_FACTOR = 1
        End If
        
        If Not IsEmpty(Range("X7")) And IsNumeric(Range("X7").Value) Then
            EVEN_FACTOR = CInt(Range("X7").Value)
            MsgBox ("EVEN: " & EVEN_FACTOR)
        Else
            EVEN_FACTOR = 1
        End If
        
        success = AutoBuildSheet(FIT_FACTOR, MIN_BW_FACTOR, EVEN_FACTOR)
    
    Else
        MsgBox ("RUNNING 3 DIFFERENT FACTOR SETS!")
        
        cleared = clearTrailer()
        If cleared Then
            ' Reccursively call AutoBuild with the first set of factors
            success = AutoBuildSheet(1, 3, 5)
            ' Save this file (CHANGE FOLDER AS NEEDED) and call again with
            MsgBox ("FIRST TRIAL FINISHED!!!")
            
             ActiveWorkbook.SaveAs "LoadSheet_" & Range("H5").Value & "(T1).xlsm"
            ' Use the clearTrailer to empty the cells for the next trial
            cleared = clearTrailer()
        Else
            MsgBox ("ERROR: ISSUE CLEARING TRAILER FOR 1st TRIAL!!")
        End If
        
        If cleared Then
            ' Repeat this process for trials 2 and 3
            success = AutoBuildSheet(5, 1, 3)
            ActiveWorkbook.SaveAs "LoadSheet_" & Range("H5").Value & "(T2).xlsm"
            cleared = clearTrailer()
        Else
            MsgBox ("ERROR: ISSUE CLEARING TRAILER FOR 2nd TRIAL!!")
        End If
        
        If cleared Then
            success = AutoBuildSheet(3, 5, 1)
            ActiveWorksheet.SaveAs "LoadSheet_" & Range("H5").Value & "(T3).xlsm"
            cleared = clearTrailer()
        Else
            MsgBox ("ERROR: ISSUE CLEARING TRAILER FOR 3rd TRIAL!!")
        End If
        
        
        
        If cleared Then
            MsgBox ("All 3 versions were successfully built!")
            ActiveWorksheet.SaveAs CurDir & Application.PathSeparator & "Load Sheet Auto.xlsm"
            Application.Run "Module1.Clear"
        End If
            
        
    End If

End Sub



Function AutoBuildSheet(Optional cur_FIT_FACTOR As Integer, Optional cur_MIN_BW_FACTOR As Integer, Optional cur_EVEN_FACTOR As Integer)

    FIT_FACTOR = cur_FIT_FACTOR
    
    MIN_BW_FACTOR = cur_MIN_BW_FACTOR
    
    EVEN_FACTOR = cur_EVEN_FACTOR

    Worksheets("26 Pallets").Activate
    
    ' This variable is the minumum number of layers for deliveries
    ' that will be distributed across BOTH columns
    ' (If the delivery has at least 6 full skids (6 skids * 7 layers/skid = 42 layers), it will be put in both cols)
    bothSideThreshold = 42
    
    Dim lastOne
    lastOne = False
    
    Dim secondLast
    seconfLast = False
    
    ' This represents the row for the store # for the current drop being placed in the list
    Dim currentDropRow As Integer
    currentDropRow = 6
    
    Dim currentDropCell
    currentDropCell = "J" & CStr(currentDropRow)
    
    Dim success
    success = 0
    
    ' First, starting in Cell J6 (first drop) go down to the last item in the list
    Range(currentDropCell).Select
    
    ' If the top cell is empty, there is no data in the dump. Display error.
    
    Dim dict
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Setting the trailer area font size to 16 for the start
    With ActiveSheet.Range("B4:C16")
        .Font.Size = 16
    End With
    
    ' Delete all comments in the trailer area
    For Each cmt In ActiveSheet.Comments
        cmt.Delete
    Next cmt
    
    
    If IsEmpty(ActiveCell) Then
        MsgBox ("Error: Please paste route schedule into dump area")
    Else
        ' This loop finds and selects the LAST delivery to be placed FIRST in the truck
        ' iterate down past the last item on the list
        While (Not IsEmpty(ActiveCell))
            ' active cell becomes NEXT item down in dump list
            currentDropRow = currentDropRow + 1
            currentDropCell = "J" & CStr(currentDropRow)
            Range(currentDropCell).Select
        Wend
        
        ' This sets active cell to the LAST store # cell on the list (bottom)
        currentDropRow = currentDropRow - 1
        currentDropCell = "J" & CStr(currentDropRow)
        Range(currentDropCell).Select
    End If
    
    ' Now the drops are placed on the trailer side, starting with the last delivery
    ' and moving up the list. Once the first delivery is placed, loop STOPS.
    
    ' deliveriesPlaced keeps track of how many have already been placed in the truck section
    deliveriesPlaced = 0
    
    
    While (currentDropRow > 5)
        
        ' This variable tracks the number for the delivery being placed in this iteration
        curStoreNum = ActiveCell.Value
        
        ' If this is one of the first two deliveries, it will favor placement in column B
        If deliveriesPlaced = 0 Or deliveriesPlaced = 1 Then
            firstTwo = True
        Else
            firstTwo = False
        End If

        layerCountCell = "K" & CStr(currentDropRow)   ' Gets the cells with layer count for current deliv
        nextLayerCountCell = "K" & CStr(currentDropRow - 1) 'layer count for NEXT deliv
        nextNextLayerCountCell = "K" & CStr(currentDropRow - 2) ' layer count cell for 2 deliveries ahead
        kegCountCell = "L" & CStr(currentDropRow) ' Keg count for current deliv
        numKegs = 0 ' Reset keg count to zero for this deliv
        
        If Not IsEmpty(Range(kegCountCell)) And IsNumeric(Range(kegCountCell).Value) Then
            MsgBox ("This delivery has " & Range(kegCountCell).Value & " Kegs!!!")
            numKegs = CInt(Range(kegCountCell).Value)
        End If
        
        layers = Range(layerCountCell).Value
        
        futurePartialLayers = 0
        
        ' This gets the partial layers for the NEXT delivery, to plan for a potential triple stack.
        If IsNumeric(Range(nextLayerCountCell).Value) And Not IsEmpty(Range(nextLayerCountCell)) Then
        
            futurePartialLayers = CInt(Range(nextLayerCountCell).Value) Mod 7
        
        
        ' If next delivery does NOT have layers, check to see if current deliv. is the LAST one with layers given
            ' If it's the last delivery, bothSidePlace will be used to keep weight more balanced
        ElseIf ActiveCell.Row = 6 Or rangeIsEmpty("K6", nextLayerCountCell) = True Then
            MsgBox ("THIS IS THE LAST DELIVERY TO BE PLACED")
            lastOne = True
            
        ElseIf rangeIsEmpty("K6", nextNextLayerCountCell) = True Then
            MsgBox ("SECOND LAST")
            secondLast = True
            
        End If
        
        
        
        
        If Not curStoreNum = "" And Not lastOne Then
            
            ' The smartPlace fn is called to place each individual delivery's load, one-by-one
            success = smartPlace(layers, curStoreNum, futurePartialLayers, False, firstTwo, numKegs)
        
            If success <> 1 Then
                MsgBox ("There was an issue placing the delivery: " & curStoreNum)
                Stop
            End If
            
        ElseIf lastOne Then ' If placing the LAST DELIVERY TO BE PLACED
            
            ' smartPlace is called again, this time with lastOne (4th arg) set to TRUE
            success = smartPlace(layers, curStoreNum, futurePartialLayers, True, firstTwo, numKegs) ' lastOne is now TRUE
            
            If success <> 1 Then
                MsgBox ("There was an issue placing the delivery: " & curStoreNum)
                Stop
            End If
            
        End If
        
        ' Iterates to the next delivery before the loop repeats...
        currentDropRow = currentDropRow - 1
        currentDropCell = "J" & CStr(currentDropRow)
        Range(currentDropCell).Select
        
        deliveriesPlaced = deliveriesPlaced + 1
    Wend
    
    AutoBuildSheet = success
        
End Function



' The big 'n ugly smartPlace controlls the placement of an individual delivery in the truck section.
' There are 3 types of placement for whole deliveries:
'       1. left or right side (column B or C) with partial stacking behind if possible
'       2. Both side placement, where the left and right side are used, with partial stacking with the best possible
'           partial on either side of the truck
'       3. Staggered placement, where the partial is stacked on one side, and the full skids on the other,
'           which makes room for a triple stack for the next placement if the resultant stack is <= 7 Layers.

'   ARGS:   layers - (int) the total number of layers for the store being placed
'           curStoreNum - (int) the number for the store being placed
'           futurePartialLayers (int) - number of partial layers for the NEXT delivery to be placed, to decide whether
'               offset partial placement should be used to create a triple stack with the next delivery.
'           lastOne (bool) - TRUE if this is the last delivery to be placed, FALSE otherwise
'           firstTwo (bool) - TRUE if this is the one of the first two deliveries to be placed, FALSE otherwise

'   RET:    1 - (int) represents successful placement (no errors)
'           0 - (int) represents fail case
Function smartPlace(layers, curStoreNum, futurePartialLayers, lastOne, firstTwo, numKegs)

    Dim proposedCell
    
    ' Get the number of full skids, and partial layers by division by 7 with the layer count
    
    If IsNumeric(layers) Then ' Check that a NUMBER is in the Layers cell for this deliv.
        numFullSkids = Application.WorksheetFunction.RoundDown(layers / 7, 0)
        partialLayers = layers Mod 7
    Else
        smartPlace = 1 ' If not a number, this function will return and do nothing
    End If
    
    fullSkidStartCell = ""
    
    
    ' chooseColumn function is used to pick the optimal column and return the top-most
    ' placeable cell address in the column (including partials if a stack is possible)
    proposedCell = chooseColumn(partialLayers, numFullSkids, firstTwo)
    Range(proposedCell).Select

    MsgBox ("Load will start from " & proposedCell)
    
    
    ' This variable tracks whether the partial skid has been placed or not
    Dim partialPlaced
    partialPlaced = False
    
    ' If this is the last delivery or at least 6 pallets, distribute accross both sides of trailer for even columns
    If lastOne = True Or layers >= bothSideThreshold Then
    
        MsgBox ("BOTH-SIDE  FOR " & curStoreNum)
        
        smartPlace = bothSidePlace(proposedCell, curStoreNum, partialLayers, numFullSkids, numKegs)
        
        ' The partial would have been placed by the bothSidePlace function.
        partialPlaced = True
        
        
    ' If NOT the last deliv, and partial skid is returned from chooseColumn, a stack is possible! Start with that...
    ElseIf ActiveCell.Value Like "*-*" And partialLayers > 0 Then
    
        ' Gets the actual layer count from the string (ex. 4545-3L would return 3) in the partial cell
        layersBehind = CInt(getPartialLayersFromCell(proposedCell))
        
        ' If a triple stack is possible with the delivery after this one, so use OFFSET placement
            ' Must add up to < 8 Layers, maximum num of different partials on a stack is 3
        MsgBox ("TRIPLE??" & layersBehind & "  " & partialLayers & "  " & futurePartialLayers)
            
        If layersBehind + partialLayers + futurePartialLayers < 8 And layersBehind > 0 And futurePartialLayers > 0 Then
            
            MsgBox ("VALID PARTIAL FOUND!")
            MsgBox ("OFFSET PARTIAL PLACE")
            
            ' Use offset partial placement to save room for the third delivery in the stack (partials on one side, fulls on other)
                'This stacks the partial on one side, and returns the first cell on OTHER side where the FULL skids can be placed
            fullSkidStartCell = offsetPartialPlace(proposedCell, curStoreNum, partialLayers, numFullSkids, True)
            
            If fullSkidStartCell <> "err" Then
                partialPlaced = True ' done by offsetPartialPlace
                proposedCell = fullSkidStartCell ' cell to start placing fulls
                Range(proposedCell).Select
            End If
            
        ' ElseIf futurePartialLayers = 0 Or (futurePartialLayers + layersBehind > 4 And futurePartialLayers + layersBehind < 5)
            
        ' If only a 2-stack is possible, just place the whole delivery on the side where partial can be stacked
        ElseIf layersBehind + partialLayers < 8 And layersBehind > 0 Then
            
            MsgBox ("VALID PARTIAL FOUND!")
            MsgBox ("ONE-SIDE PLACE")
            
            success = stackPartials(curStoreNum, partialLayers, numFullSkids, proposedCell, False, numKegs)
            
            If success = 1 Then
                partialPlaced = True
            End If
            
        End If
        
        ' Now that the partial has been placed, iterate to the next cell down to start placing fulls...
            ' UNLESS offsetPartialPlace was used (next free cell is given already, and will be empty)
        If ActiveCell.Column = 2 And Not IsEmpty(ActiveCell) Then
            proposedCell = "B" & ActiveCell.Row + 1
        ElseIf Not IsEmpty(ActiveCell) Then
            proposedCell = "C" & ActiveCell.Row + 1
        End If
        
    ' Finally, If an empty cell was given then the partial can't be stacked, so put a BW under the one behind if < 5 Layers
    ' This block does not place the partial, it will be put in after the full skids later.
    ElseIf IsEmpty(ActiveCell) Then
            
        If proposedCell Like "*B*" Then
            cellBehind = "B" & ActiveCell.Row - 1
        Else
            cellBehind = "C" & ActiveCell.Row - 1
        End If
        
        layersBehind = getPartialLayersFromCell(cellBehind)
        
        ' If a BW should be placed, and one is not already there (already placed in bothSidePlace) then place BW under
        If CInt(layersBehind) > 0 And CInt(layersBehind) < 5 And InStr(Range(cellBehind).Value, "BW") = 0 Then
            Range(cellBehind).Value = Range(cellBehind).Value & Chr(10) & "BW"
        End If
    End If
    
    ' If the partial was stacked, it will now select the next free space in front
    Range(proposedCell).Select
    
    ' If there are NO full skids, only partial, and it wasn't stacked:
    If numFullSkids = 0 And partialPlaced = False And partialLayers > 0 Then
    
        ' Place the partial in proposedCell
        ActiveCell.Value = curStoreNum & "-" & CStr(partialLayers) & "L"
        
        ' A comment is added to this cell to indicate that it should NOT be stacked on
        If Not ActiveCell.Comment Is Nothing Then
            ActiveCell.Comment.Delete
        End If
        ActiveCell.AddComment "DO NOT STACK ON: " & CStr(curStoreNum)
        
        partialPlaced = True
        
    ElseIf lastOne = False And layers < bothSideThreshold And numFullSkids > 0 Then
        ' If there ARE full skids, the first one is placed here
        ActiveCell.Value = curStoreNum
    End If
    
    ' Converts the col NUMBER into respective LETTER
    If ActiveCell.Column = 2 Then
        curCol = "B"
    End If
    If ActiveCell.Column = 3 Then
        curCol = "C"
    End If
        
    curRow = 0
    
    ' Place THE REST of the full skids for this delivery in a vertical row
    i = 1 ' Keeps track of fulls placed so far. One skid was just placed...
    
    ' Places more skids WHILE the selected store is still the same
    ' AND while there are still skids to place
    ' AND if it isn't the last delivery (done by bothSidePlace instead)
    ' AND if it isn't more than 5 full skids (done by bothSidePlace instead)
    ' AND if placement does not overflow the truck
    Do While ActiveCell.Value = curStoreNum And i < numFullSkids And lastOne = False And layers < bothSideThreshold And curRow < 17
        
        curRow = ActiveCell.Row
    
        proposedCell = curCol & CStr(curRow + 1)
        
        ' If this skid doesn't go past the end of the trailer
        If curRow + 1 < 17 Then
        
            Range(proposedCell).Select
            
            ' Then the skid is placed in the cell
            ActiveCell.Value = curStoreNum
            i = i + 1
        Else
            ' The trailer will overflow; print error.
            MsgBox ("(1)ERROR: This route will likely not fit on the trailer." & Chr(10) & "It is recommended that you count this run on the floor.")
            smartPlace = 0
            Exit Do
        End If
    Loop
    
    curRow = ActiveCell.Row + 1
    
    
    If partialPlaced = False And partialLayers > 0 And lastOne = False Then
        ' No stack behind, so simply place partial after the full skids
        proposedCell = curCol & CStr(curRow)
        
        If curRow < 17 Then
            Range(proposedCell).Select
            ActiveCell.Value = curStoreNum & "-" & CStr(partialLayers) & "L"
        Else
            ' The trailer will overflow; print error.
            MsgBox ("(1)ERROR: This route will likely not fit on the trailer." & Chr(10) & "It is recommended that you count this run on the floor.")
            smartPlace = 0
        End If
        
    End If
    
    partialPlaced = False ' reset to false for next delivery
    
    smartPlace = 1
End Function





' Takes the number of skids and leftover layers of one delivery, and
' Selects the best Column to place it in.
' If a cell with an existing partial is returned, the partial can be stacked here.

'   ARGS:   partialLayers (int) - number of partial layers in this delivery (>= 0)
'           numFullSkids (int) - number of full skids to be placed for this delivery (>= 0)

'   RET: Returns the cell ID of the cell where the first skid (partial OR full) can be placed
Function chooseColumn(partialLayers, numFullSkids, firstTwo)
    
    currentPlaceRow = 4
    currentPlaceCol = "B"
    proposedCell = currentPlaceCol & CStr(currentPlaceRow)

    colB_points = 0
    colC_points = 0
    
    colB_len = 0
    colC_len = 0
    
    ' Starting from top left pos, and moving down the trailer to look for best placement
    Range(proposedCell).Select
    
    
    colB_len = getColLength("B")
    colC_len = getColLength("C")
    
    ' First allocate points for length, so the col with less pallets has more points.
    colB_points = EVEN_FACTOR * (13 - colB_len)
    colC_points = EVEN_FACTOR * (13 - colC_len)
    
    ' If this is the first or secong deliv to be placed...
    If firstTwo And (ActiveCell.Row + numFullSkids) <= 9 Then
        ' Adds a lot of points to column B for the first two deliveries placed
        colB_points = colB_points + numFullSkids
    End If
    
    MsgBox ("B is " & CStr(colB_len) & " long.  C is " & CStr(colC_len) & " long")
    
    ' Now check the potential stacks that can be made at the end of each col, so the
    ' col with the stack closest to max size (7) is favored.
    
    ' FIRST THIS IS DONE FOR COLUMN B
    colB_startCell = "B" & 4 + colB_len
    colB_behindStartCell = "B" & 3 + colB_len
    Range(colB_behindStartCell).Select
    
    If ActiveCell.Value Like "*-*" Then
    
        MsgBox ("DETECTED A PARTIAL IN CELL " & colB_behindStartCell)
    
        ' Isolating the tier count for the partial behind
        partialLayersBehind = getPartialLayersFromCell(ActiveCell.Address)
        
        ' If it is a legal stack, award more points for bigger stack
        If partialLayersBehind + partialLayers < 8 And partialLayers > 0 Then
            colB_points = colB_points + FIT_FACTOR * (partialLayers + partialLayersBehind)
            MsgBox ("LEGAL STACK in col B")
            colB_startCell = colB_behindStartCell
            
        ' If a stack cannot be made, and partial behind needs BWs under it
        ElseIf partialLayersBehind < 5 Then
        ' Then take away points for this col(don't want BWs)
            colB_points = colB_points - (MIN_BW_FACTOR * (5 - partialLayersBehind))
        End If
    Else
        MsgBox ("No partial in col B")
    End If
    
    ' NEXT IT IS DONE FOR COLUMN C
    colC_startCell = "C" & 4 + colC_len
    colC_behindStartCell = "C" & 3 + colC_len
    Range(colC_behindStartCell).Select
    
    If ActiveCell.Value Like "*-*" Then
    
        MsgBox ("DETECTED A PARTIAL IN CELL " & colC_behindStartCell)
    
        ' Isolating the tier count for the partial behind
        partialLayersBehind = getPartialLayersFromCell(ActiveCell.Address)
        
        ' If it is a legal stack, award more points for bigger stack
        If partialLayersBehind + partialLayers < 8 And partialLayers > 0 Then
            colC_points = colC_points + FIT_FACTOR * (partialLayers + partialLayersBehind)
            MsgBox ("LEGAL STACK in col C")
            colC_startCell = colC_behindStartCell
            
        ' If a stack cannot be made, and partial behind needs BWs under it
        ElseIf partialLayersBehind < 5 Then
        ' Then take away points for this col(don't want BWs)
            colC_points = colC_points - (MIN_BW_FACTOR * (5 - partialLayersBehind))
        End If
        
    Else
        MsgBox ("No partial in col C")
    End If
    
    ' Returns the Cell address of the EITHER the FIRST FREE SPACE in the optimal column
    ' OR the PARTIAL SKID at the end of the chosen column if a stack is possible.
    
    If colB_points > colC_points Then
        proposedCell = colB_startCell
        MsgBox ("VERDICT: Col B")
    ElseIf colB_points < colC_points Then
        proposedCell = colC_startCell
        MsgBox ("VERDICT: Col C")
    Else
        MsgBox ("Column B and C are EQUALLY favorable for next placement." & Chr(10) & " DISTRIBUTING EVENLY!")
        proposedCell = colC_startCell
    End If
    
    chooseColumn = proposedCell

End Function



' This function returns the length of the column (occupied cells) given as arg.
    ' ARG: col - Either  'B' or 'C' (/'b' or 'c') representing the
    ' trailer column to return the length of
    
    ' RET: (int) the 'length' of the column (in rows) that is signified by "col"
Function getColLength(col)
    ' Capitalize the Column idicator (maybe useless)
    If col = "b" Then
        col = "B"
    ElseIf col = "c" Then
        col = "C"
    End If
    
    colLen = 0
    
    curReadRow = 4
    curReadCol = CStr(col)
    curReadCell = curReadCol & CStr(curReadRow)
    
    Range(curReadCell).Select

    Do While (Not IsEmpty(ActiveCell))
    
        colLen = colLen + 1
        
        curReadRow = curReadRow + 1
        curReadCell = curReadCol & CStr(curReadRow)
        Range(curReadCell).Select
        
    Loop
    
    ' The length of Column "col" is returned by this function.
    getColLength = colLen

End Function


' This function distributes a delivery accross both sides of the truck so that the
' columns have equal length or 1 skid away from equal. Places partial AND full skids.
'   ARGS:   firstCell - The cell address for the first skid that will be placed
'           storeNum - The store number of the delivery being placed
'           partialLayers (int) - number of partial layers in this delivery (>= 0)
'           numFullSkids (int) - number of full skids to be placed for this delivery (>= 0)

'   RET:    1 - operation was successful
'           0 - operation failed
'           -1 - look outside, pigs might be flying
Function bothSidePlace(firstCell, storeNum, partialLayers, numFullSkids, numKegs)

    Range(firstCell).Select
    If ActiveCell.Row > 5 Then
        checkRow = ActiveCell.Row - 2 ' ActiveCell.Row
    Else
        checkRow = 4
    End If
    
    partialPlaced = False
    numPlaced = 0
    
    Do While (numPlaced < numFullSkids And checkRow < 17)
    
        ' If NOT the last skid, place them in B, C, B, so on..
        ' Last skid should go in Col C if the columns can't be equal length.
        B_cell = Range("B" & checkRow)
        C_cell = Range("C" & checkRow)
        
        B_cell_InFront = Range("B" & checkRow + 1)
        C_cell_InFront = Range("C" & checkRow + 1)
        
        B_cellPartialLayers = getPartialLayersFromCell("B" & checkRow)
        C_cellPartialLayers = getPartialLayersFromCell("C" & checkRow)
        
        MsgBox ("BSP: partials - " & partialLayers & "  B's - " & B_cellPartialLayers & "  C's - " & C_cellPartialLayers)
        
        ' This section places the partial layers for this delivery if it has partials
        If partialLayers > 0 Then
        
            If (B_cellPartialLayers > 0 And C_cellPartialLayers > 0) And partialPlaced = False Then
            ' If both col's have a partial at the end, stack on the largest as long as its <= 7 layer when stacked
                If B_cellPartialLayers + partialLayers < 8 And C_cellPartialLayers + partialLayers < 8 Then
                
                    If B_cellPartialLayers > C_cellPartialLayers And IsEmpty(B_cell_InFront) Then
                        ' STACK PARTIAL ON B CELL
                        MsgBox ("Col B stack would be bigger; Stacking here")
                        success = stackPartials(storeNum, partialLayers, numFullSkids, "C" & checkRow, False, numKegs)
                        
                    ElseIf IsEmpty(C_cell_InFront) Then
                        ' STACK ON C CELL
                        MsgBox ("Col C stack would be bigger; Stacking here")
                        success = stackPartials(storeNum, partialLayers, numFullSkids, "C" & checkRow, False, numKegs)
                    End If
                    
                    If success = 1 Then
                        partialPlaced = True
                    End If
                        
                End If
                
            ElseIf B_cellPartialLayers > 0 And B_cellPartialLayers + partialLayers < 8 And IsEmpty(B_cell_InFront) And partialPlaced = False Then
            ' There is a valid stack with the cell in col B
                MsgBox ("Only B has a valid partial!")
                
                success = stackPartials(storeNum, partialLayers, numFullSkids, "B" & checkRow, False, numKegs)
                
                If success = 1 Then
                    partialPlaced = True
                End If
                
            ElseIf C_cellPartialLayers > 0 And C_cellPartialLayers + partialLayers < 8 And IsEmpty(C_cell_InFront) And partialPlaced = False Then
            ' There is a valid stack with the cell in col C
                MsgBox ("Only C has a valid partial!")
                
                success = stackPartials(storeNum, partialLayers, numFullSkids, "C" & checkRow, False, numKegs)
                
                If success = 1 Then
                    partialPlaced = True
                End If

            End If
            
        Else
            ' If there is no partial for this delivery, set it to already placed.
            partialPlaced = True
        End If

        Dim cellBehind As String
    
        If IsEmpty(B_cell) And IsEmpty(C_cell) Then
            MsgBox ("BOTH C and B cells are free. Placing in C")
            Range("C" & checkRow).Select
            ActiveCell.Value = storeNum
            numPlaced = numPlaced + 1
            
            ' The cell behind the one that was just placed is selected to check if it needs a BW
            cellBehind = "C" & CStr(checkRow - 1)
            
        ElseIf IsEmpty(B_cell) Then
            MsgBox ("ONLY B cell is free. Placing in B")
            Range("B" & checkRow).Select
            ActiveCell.Value = storeNum
            numPlaced = numPlaced + 1
            
            cellBehind = "B" & CStr(checkRow - 1)
            
        ElseIf IsEmpty(C_cell) Then
            MsgBox ("ONLY C cell is free. Placing in C")
            Range("C" & checkRow).Select
            ActiveCell.Value = storeNum
            numPlaced = numPlaced + 1
            
            cellBehind = "C" & CStr(checkRow - 1)
            
        Else
            checkRow = checkRow + 1
        End If
        
        If cellBehind <> "" Then ' If a full has been placed, check behind for a partial...
            cellBehindLayers = getPartialLayersFromCell(cellBehind)
        
            ' If the cell behind needs a BW because it's <5 Layers, stack cannot be made and a BW is not already placed...
            If cellBehindLayers < 5 And cellBehindLayers > 0 And (partialLayers + cellBehindLayers > 7 Or partialPlaced) And Not InStr("BW", Range(cellBehind).Value) Then
            
                ' Place a BW under the partial behind
                Range(cellBehind).Select
                MsgBox ("BSP doing BW!")
                ActiveCell.Value = ActiveCell.Value & Chr(10) & "BW"
                
            End If
            
        End If
        
        cellBehind = ""
        
    Loop
    
    ' IF ALL full skids have been placed, but partial has not, place partial in next available spot
    Do While Not partialPlaced And partialLayers > 0 And checkRow < 17
        
        B_cell = Range("B" & checkRow)
        C_cell = Range("C" & checkRow)
        
        If IsEmpty(B_cell) And IsEmpty(C_cell) Then
            MsgBox ("BOTH C and B cells are free. Placing PARTIAL in C")
            success = stackPartials(storeNum, partialLayers, numFullSkids, "C" & checkRow, False, numKegs)
            partialPlaced = True
                
        ElseIf IsEmpty(B_cell) Then
            MsgBox ("ONLY B cell is free. Placing PARTIAL in B")
            success = stackPartials(storeNum, partialLayers, numFullSkids, "B" & checkRow, False, numKegs)
            partialPlaced = True
                
        ElseIf IsEmpty(C_cell) Then
            MsgBox ("ONLY C cell is free. Placing PARTIAL in C")
            success = stackPartials(storeNum, partialLayers, numFullSkids, "C" & checkRow, False, numKegs)
            partialPlaced = True
                
        Else
            checkRow = checkRow + 1
        End If
    Loop
    
    
    If checkRow > 16 Then
        MsgBox ("(1)ERROR: This route will likely not fit on the trailer." & Chr(10) & "It is recommended that you count this run on the floor.")
    End If

    bothSidePlace = 1

End Function

' This function does a placement where the fulls are on a different side than the partial for the given delivery
' USEFUL FOR setting up a triple stack if one can be made (check this first with placement foresight)
'   ARGS:   firstCell (str) - string-form address of the first cell offset placement begins (should be a partial)
'           storeNum (str) - store number for the delivery being placed
'           partialLayers (int) - number of partial layers in this delivery to stack on firstCell
'           numFullSkids (int) - number of full skids for this delivery, to palce on opposite side of partial.

'   RET:    (str) - string-form address of the starting cell for the full goods to be placed in one col (done in smartPlace)
'           0 - The operation failed
Function offsetPartialPlace(firstCell, storeNum, partialLayers, numFullSkids, tripleStack)

    Range(firstCell).Select
    
    MsgBox ("offsetPartialPlace is doing this one!")
    
    success = stackPartials(storeNum, partialLayers, numFullSkids, firstCell, True, numKegs)

    If Not success = 1 Then
        MsgBox ("ERROR: A problem occured while stacking partial skids for offset-placement!")
        offsetPartialPlace = 0
    Else
        ' Since the partial has been placed, set the return address to the next free cell in the OTHER column.
        
        checkRow = 4
        
        ' Selects the OTHER col (without end partial) to start placing fulls
        If ActiveCell.Column = 2 Then
            Range("C" & checkRow).Select
            checkCol = "C"
        ElseIf ActiveCell.Column = 3 Then
            Range("B" & checkRow).Select
            checkCol = "B"
        Else
            MsgBox ("ERROR: Something went seriously wrong in offset partial placement!" & Chr(10) & "Call an engineer.")
            offsetPartialPlace = 0
        End If

        ' Loop scans down the column until a free cell is found.
        Do While Not IsEmpty(ActiveCell) And checkRow < 17
            ' If partial is in this column < 5L w/o a BW, it places one under
            If ActiveCell.Value Like "*-*" And getPartialLayersFromCell(ActiveCell.Address) < 5 Then
                ActiveCell.Value = ActiveCell.Value & Chr(10) & "BW"
            End If
                
            ' iterate down to next cell in this column
            checkRow = checkRow + 1
            Range(checkCol & checkRow).Select
        Loop
        
        If checkRow <> 17 Then
            offsetPartialPlace = CStr(checkCol & checkRow)
        Else
            offsetPartialPlace = "err" ' to indicate error
            MsgBox ("(2)ERROR: This route will not fully fit in the truck." & Chr(10) & "It is recommended that you count this run on the floor.")
        End If
        
    End If

End Function


' This function parses the text in a cell, and if there is a partial (indicated by presence of "-")
' then it returns this integer value. For stacks of two or more partials if..

' SUMS THE STACKED PARTIALS AND RETURNS THE PRE_EXISTING STACK'S HEIGHT

'   ARG:    cellAddress - address of the cell containing a partial, for which the layer count is extracted

'   RET:    Integer number (1-6) representing the layers of the partial in cellAddress
Function getPartialLayersFromCell(cellAddress)

    Range(cellAddress).Select

    If ActiveCell.Value Like "*-*" Then

        split1 = Split(ActiveCell.Value, "-")
        
        
        numberOfPartials = UBound(split1) - LBound(split1) + 1
        
        For i = 0 To numberOfPartials - 1
            testStr = testStr & " , " & split1(i)
        Next
        indexToParse = 0
        totalLayers = 0
                
        Do While indexToParse < numberOfPartials
            If split1(indexToParse) Like "*L*" Then ' This array item is a layer count, not store number
                
                partialLayersBehind = Split(split1(indexToParse), "L")
                totalLayers = totalLayers + CInt(partialLayersBehind(0))
            End If
            indexToParse = indexToParse + 1
        Loop
        
        MsgBox ("The total partial layers in this cell are " & CStr(totalLayers))
        
        getPartialLayersFromCell = CInt(totalLayers)

    Else
        getPartialLayersFromCell = 0

    End If
    
End Function

' This function stacks a new partial skid on-top or under an existing partial and formats the cell properly.
    ' NOTE: This function will not stack more than 3 deliveries
    
'   ARGS:   storeNum_newPartial - Store number of the new partial being stacked onto an existing one
'           partialLayers - number of layers for the newPartial
'           numFullSkids - number of full skids in this delivery, to determine if the new partial can go on bottom
'           partialCell - cell address of the existing partial that is already placed

'   RET:    1 - operation was successful
'           0 - operation failed
'           -1 - They're coming... RUN!
Function stackPartials(storeNum_newPartial, partialLayers, numFullSkids, partialCell, tripleStack, numKegs)

    partialLayersBehind = CInt(getPartialLayersFromCell(partialCell))
    partialLayers = CInt(partialLayers)
    numFullSkids = CInt(numFullSkids)
    
    MsgBox ("STACKING PARTIAL SIZE : " & partialLayers & " ONTO EXISTING : " & partialLayersBehind)

    ' If Active Cell has a comment that says DO NOT STACK then it should not be stacked on top of, so stack under if <= 7Layers
    ' The number of "-" are counted in the cell to determine the number of partials so far. If more than 2 already, won't stack!
    If (partialLayers <= partialLayersBehind + 1 Or numFullSkids = 0) And ActiveCell.Comment Is Nothing And Len(ActiveCell.Value) < 19 Then
        ' Places the NEW partial ON TOP of the partial behind (new one is lighter)
        ' KEGS can also be placed with this partial.
        If numKegs > 0 Then
            ActiveCell.Value = (CStr(storeNum_newPartial) & "-" & partialLayers & "L" & "+" & numKegs & "KEGS" & Chr(10) & ActiveCell.Value)
        Else
            ActiveCell.Value = (CStr(storeNum_newPartial) & "-" & partialLayers & "L" & Chr(10) & ActiveCell.Value)
        End If
        
        partialPlaced = True
                
    ElseIf partialLayersBehind > 0 And Len(ActiveCell.Value) < 19 Then
        'Places the NEW partial BELOW the partial behind
        If numKegs > 0 Then
            ActiveCell.Value = (ActiveCell.Value & Chr(10) & CStr(storeNum_newPartial) & "-" & partialLayers & "L" & "+" & numKegs & "KEGS")
        Else
            ActiveCell.Value = (ActiveCell.Value & Chr(10) & CStr(storeNum_newPartial) & "-" & partialLayers & "L")
        End If
        
        partialPlaced = True
        
    End If
    
    ' There is no existing partial in the given cell, put it there alone
    If Not partialPlaced And partialLayers > 0 And IsEmpty(ActiveCell) Then
        
        ActiveCell.Value = CStr(storeNum_newPartial) & "-" & partialLayers & "L"
    
        If numFullSkids = 0 Then
            If Not ActiveCell.Comment Is Nothing Then
                ActiveCell.Comment.Delete
            End If
            ActiveCell.AddComment "DO NOT STACK ON: " & CStr(curStoreNum)
        End If
        
        partialPlaced = True
        
    End If
    
    ' If the STACK still has < 5 Layers add BWs and change font size to fit 3 rows
    If getPartialLayersFromCell(ActiveCell.Address) < 5 And partialLayersBehind > 0 And Not tripleStack Then
        ActiveCell.Value = ActiveCell.Value & Chr(10) & "BW"
        ActiveCell.Font.Size = 12
    ElseIf tripleStack And partialLayersBehind > 0 Then
        ActiveCell.Font.Size = 12
    End If
        
    
    ' Set return value to 1 if the stack was successful.
    If partialPlaced Then
        stackPartials = 1
    Else
        MsgBox ("ERROR: something went wrong when stacking partial skids!")
        stackPartials = 0
    End If
    
End Function

Function rangeIsEmpty(fromCell, toCell)

    If WorksheetFunction.CountA(Range(fromCell & ":" & toCell)) = 0 Then
        rangeIsEmpty = True
    Else
        rangeIsEmpty = False
    End If
End Function


Function clearTrailer()

    With ActiveSheet.Range("B4:C16")
        .Font.Size = 16
        .Value = ""
    End With
    
    If rangeIsEmpty("B4", "C16") Then
        clearTrailer = True
    Else
        clearTrailer = False
    End If
    
End Function

