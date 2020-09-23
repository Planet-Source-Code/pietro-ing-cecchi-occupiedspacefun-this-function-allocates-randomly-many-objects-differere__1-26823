Attribute VB_Name = "OccupiedSpaceRoutines"
Public Enum ActionCommandEnum
   ClearContainerRegion = 1
   AddObjectsToContainer = 2
   ShowCells = 4
   HideCells = 8
End Enum
Public Enum ReturnErrors
   NoERRORS = 0
   ErrorInCLEAR = 1
   ErrorInADD = 2
   ErrorInSHOWCELLS = 4
   ErrorInHIDECELLS = 8
End Enum
Public Enum BackUpRestoreEnum
   BackUp = 1
   Restore = 2
End Enum

Public Function OccupiedSpaceFun(ByVal ContainerOBJ As Object, ByVal PopulatingOBJ As Object, ByVal ActionCommand As ActionCommandEnum, Optional ByRef IOvariable As Integer = 1) As ReturnErrors
'IOvariable, optional, is used as follows:
'- during CLEAR, to output the calculated MaxAllocableObjectsNumber
'  that can be used by the application to initially set the UpDown1.Max value
'- during ADD objects: A) to supply the function with the number of objects to add to
'  the container B) to return to the application the real number of objects added to container
   
'EXAMPLE OF USAGE:
'   create a form and pull down a certain number of indexed objects (labels, images or whatsoever)
'   pull down also a container (where the objects will be placed), i.e. a picture box, or you can use the form as container
'   size at your pleasure both the container and each object
'   create a module and place there the code (the enums and the function) of this module
'   then use these calls (see enums for action commands and errors):
'   CALLS:
'      1) to clear and show cells (MaxAllocObjsNumber is returned). Check ret for error (no error: ret=0)
'         ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + ShowCells, MaxAllocObjsNumber)
'         remember that MaxAllocObjsNumber contains the max number of allocable abject
'      2) to clear and hide cells (MaxAllocObjsNumber is returned). Check ret for error (no error: ret=0)
'         ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + HideCells, MaxAllocObjsNumber)
'      3) to add objects to container (as input allocatenumber=number of objects to allocate,
'                                      instead as returned value allocatenumber=number of objects really allocated,
'                                      if not the same ret will be >0)
'         ret = OccupiedSpaceFun(ChoosenContainer, Label1, AddObjectsToContainer, allocatenumber)
'         remember to check both ret and allocatenumber, after the call
'   NOTE:
'      multiple commands are accepted. Anyway the following combinations
'      give unvoreseen results:
'      clear + show + add  adds all! because clear returns MaxAllocObjsNumber that is considered the number of objects
'                          to be added to container (to add a defined number use a separated dedicated call)
'      clear + hide + show  result is an hide, because show is processed before hide (first shows, then hides cells)
'   NOTE 1:
'      each action command may generate an error return (see error enum) in ret
'      the first error verified exits the function with the relative error code
'end of example of usage


   Static nx As Integer
   Static ny As Integer
   Static CellWidth As Double
   Static CellHeight As Double
   Static MaxAllocableObjectsNumber As Integer
   Static MaxCellsNumber As Integer
   Static AverageCellsPerObject As Integer  '1,4,9,16,etc (n^2)
   Static AllocableObjectsVector() As Integer
   Static AllocableCellsVector() As Integer
   Static CoordinateCellX() As Double
   Static CoordinateCellY() As Double
   
   'start returning no errors, then collect errors, if any,
   'from all action commands sections
   OccupiedSpaceFun = NoERRORS
   
'multiple commands allowed:
'CLEAR:
   If (ActionCommand And ClearContainerRegion) = ClearContainerRegion Then
      '- make objects invisible
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         PopulatingOBJ(a).Visible = False
      Next
      '- assure objects are in their container
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         Set PopulatingOBJ(a).Container = ContainerOBJ
      Next
      '- find CellWidth, Cellheight
      '  as minimum width and height of populating objects
      Dim MinWidthOfOBJs As Double, MinHeightOfOBJs As Double
      MinWidthOfOBJs = ContainerOBJ.Width
      MinHeightOfOBJs = ContainerOBJ.Height
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         If MinWidthOfOBJs > PopulatingOBJ(a).Width Then MinWidthOfOBJs = PopulatingOBJ(a).Width
         If MinHeightOfOBJs > PopulatingOBJ(a).Height Then MinHeightOfOBJs = PopulatingOBJ(a).Height
      Next
      '- normalize cell dimensions to pixel
      CellWidth = Fix((MinWidthOfOBJs + Screen.TwipsPerPixelX) / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
      CellHeight = Fix((MinHeightOfOBJs + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY
      'find Max Width and Height of populating objects
      Dim MaxWidthOfOBJs As Double, MaxHeightOfOBJs As Double
      MaxWidthOfOBJs = 0
      MaxHeightOfOBJs = 0
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         If MaxWidthOfOBJs < PopulatingOBJ(a).Width Then MaxWidthOfOBJs = PopulatingOBJ(a).Width
         If MaxHeightOfOBJs < PopulatingOBJ(a).Height Then MaxHeightOfOBJs = PopulatingOBJ(a).Height
      Next
      'are objects all equally sized?
      allequal = ((MinWidthOfOBJs = MaxWidthOfOBJs) And (MinHeightOfOBJs = MaxHeightOfOBJs))
      
'not used:
      '- compute CellSurface
      CellSurface = CellWidth * CellHeight
      
      '- find total surface of objects, in cell units
      Dim TotalSurfaceCells As Long
      TotalSurfaceCells = 0
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         cw = PopulatingOBJ(a).Width / CellWidth
         If cw - Fix(cw) > 0 Then
            cw = Fix(cw) + 1
         Else
            cw = Fix(cw)
         End If
         ch = PopulatingOBJ(a).Height / CellHeight
         If ch - Fix(ch) > 0 Then
            ch = Fix(ch) + 1
         Else
            ch = Fix(ch)
         End If
         
         TotalSurfaceCells = TotalSurfaceCells + cw * ch
      Next
      
      factor = IIf(allequal, 1, 1.667) '=1 if allequal, else 1.667 (40%=(1-1/1.667)*100 space lost, is not a puzzle!)
      '- compute AverageCellsPerObject
      AverageCellsPerObject = factor * TotalSurfaceCells / PopulatingOBJ.Count
      'the 40% space lost (allequal=False) becomes even greater (50%) with this needed normalization
      If AverageCellsPerObject - Fix(AverageCellsPerObject) > 0 Then
         AverageCellsPerObject = Fix(AverageCellsPerObject) + 1
      Else
         AverageCellsPerObject = Fix(AverageCellsPerObject)
      End If
      'calculate MaxCellsNumber and estimate MaxAllocableObjectsIndex
      nx = Fix(ContainerOBJ.ScaleWidth / (CellWidth + Screen.TwipsPerPixelX))
      ny = Fix(ContainerOBJ.ScaleHeight / (CellHeight + Screen.TwipsPerPixelY))
      MaxCellsNumber = nx * ny
      MaxAllocableObjectsNumber = Fix(MaxCellsNumber / AverageCellsPerObject)
      'clamp to existent populating objects
      If MaxAllocableObjectsNumber > PopulatingOBJ.Count Then
         MaxAllocableObjectsNumber = PopulatingOBJ.Count
         MaxCellsNumber = MaxAllocableObjectsNumber * AverageCellsPerObject
         If MaxCellsNumber > nx Then
            ny = Fix(MaxCellsNumber / nx) + 1
         Else
            ny = 1
            nx = MaxCellsNumber
         End If
         'at the end, make 'occupied space' a rectangle
         MaxCellsNumber = nx * ny
      End If
      'output MaxAllocableObjectsNumber
      IOvariable = MaxAllocableObjectsNumber
      'go on
      If MaxAllocableObjectsNumber > 0 Then
         ReDim AllocableObjectsVector(0 To PopulatingOBJ.Count - 1)
         ReDim AllocableCellsVector(0 To MaxCellsNumber - 1) As Integer
         'initialize vectors
         For a = 0 To PopulatingOBJ.Count - 1
            DoEvents
            AllocableObjectsVector(a) = -1 'value that indicates not yet allocated object
         Next
         For a = 0 To MaxCellsNumber - 1
            DoEvents
            AllocableCellsVector(a) = -1   'value that indicates not yet allocated object
         Next
         'compute coordinates of all cells
         ReDim CoordinateCellX(0 To MaxCellsNumber - 1)
         ReDim CoordinateCellY(0 To MaxCellsNumber - 1)
         cellscount = 0
         For ay = 0 To ny - 1
            For ax = 0 To nx - 1
               DoEvents
               If cellscount > MaxCellsNumber - 1 Then Exit For: Exit For
               'assigns coordinates
               'also an offset for both x and y is added
               'to center the obj's in the container available area (second terms of additions)
               CoordinateCellX(cellscount) = ax * (CellWidth + Screen.TwipsPerPixelX) + (ContainerOBJ.ScaleWidth - nx * (CellWidth + Screen.TwipsPerPixelX)) \ 2
               CoordinateCellY(cellscount) = ay * (CellHeight + Screen.TwipsPerPixelY) + (ContainerOBJ.ScaleHeight - ny * (CellHeight + Screen.TwipsPerPixelY)) \ 2
               cellscount = cellscount + 1
            Next
         Next
      Else
         OccupiedSpaceFun = ErrorInCLEAR
         Exit Function
      End If
      'clear any previous cell pattern
      ContainerOBJ.Cls 'clears cells pattern
   End If
'SHOW CELLS
   If (ActionCommand And ShowCells) = ShowCells Then
      ReDim VisibleLabels(0 To PopulatingOBJ.Count - 1)
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         VisibleLabels(a) = PopulatingOBJ(a).Visible
         If PopulatingOBJ(a).Visible Then PopulatingOBJ(a).Visible = False
      Next
      ContainerOBJ.AutoRedraw = False
      forecolorsave = ContainerOBJ.ForeColor
      ContainerOBJ.ForeColor = vbWhite
      For aa = 0 To MaxCellsNumber - 1
         DoEvents
         X1 = CoordinateCellX(aa)     'left
         Y1 = CoordinateCellY(aa)     'top
         X2 = X1 + CellWidth          'right
         Y2 = Y1 + CellHeight         'bottom
         ContainerOBJ.Line (X1, Y1)-(X2, Y2), vbWhite, B
         ContainerOBJ.CurrentX = X1
         ContainerOBJ.CurrentY = Y1
         ContainerOBJ.Print aa
      Next
      ContainerOBJ.ForeColor = forecolorsave
      ContainerOBJ.AutoRedraw = True
      
      For a = 0 To PopulatingOBJ.Count - 1
         DoEvents
         If VisibleLabels(a) Then PopulatingOBJ(a).Visible = True
      Next
   End If
'HIDE CELLS
   If (ActionCommand And HideCells) = HideCells Then
      ContainerOBJ.Cls 'clears cells pattern
   End If
'ADD OBJECTS
   If (ActionCommand And AddObjectsToContainer) = AddObjectsToContainer Then
      MaxTrialsRandom = 3 * MaxAllocableObjectsNumber 'of finding randomly a free allocation
      Dim WidthCells, HeightCells  'width and height of object in cell units
      
      objnumber = IOvariable
      For a = 1 To objnumber
         trials = 0: trials2 = 0: trials3 = 0
again1:
         DoEvents
         Randomize Timer
         rndnumber = CInt(Fix(Rnd * PopulatingOBJ.Count)) 'gives a rnd integer between 0 and PopulatingOBJ.Count-1
         If AllocableObjectsVector(rndnumber) = -1 Then 'allocable objet
again2:
            DoEvents
            Randomize Timer
            rndnumber2 = CInt(Fix(Rnd * MaxCellsNumber)) 'gives a rnd integer between 0 and MaxCellsNumber-1
again22:
            DoEvents
            If AllocableCellsVector(rndnumber2) = -1 Then 'allocable cell
               fract = PopulatingOBJ(rndnumber).Width / CellWidth
               If fract - Fix(fract) > 0 Then fract = fract + 1
               WidthCells = Fix(fract)
               fract = PopulatingOBJ(rndnumber).Height / CellHeight
               If fract - Fix(fract) > 0 Then fract = fract + 1
               HeightCells = Fix(fract)
            
               'check allocability
               available = True: outofallocablespace = False
               fragmented = False
               For yy = 0 To HeightCells - 1 'columns
                  For xx = rndnumber2 To rndnumber2 + WidthCells - 1 'rows
                     DoEvents
                     If xx + yy * nx > MaxCellsNumber - 1 Then
                        outofallocablespace = True
                        Exit For
                        Exit For
                     End If
                     If AllocableCellsVector(xx + yy * nx) > -1 Then
                        available = False
                        Exit For
                        Exit For
                     End If
                     If Fix((xx + yy * nx) / nx) > Fix((rndnumber2 + yy * nx) / nx) Then
                        fragmented = True
                        Exit For
                        Exit For
                     End If
                  Next
               Next
                   
               If (Not available) Or outofallocablespace Or fragmented Then
                  trials3 = trials3 + 1
                  If trials3 > MaxTrialsRandom Then
                     
                     'explore intelligently the cell space for possible allocation
                     If aaamin < MaxCellsNumber - 1 Then
                        For aaa = aaamin To MaxCellsNumber - 1
                           If aaa = 0 Then Previous = 1
                           If (AllocableCellsVector(aaa) = -1) And Previous > 0 Then
                              rndnumber2 = aaa
                              Previous = AllocableCellsVector(aaa)
                              aaamin = aaa + 1
                              GoTo again22
                           End If
                           Previous = AllocableCellsVector(aaa)
                        Next
                     End If
                     aaamin = 0
                     'no possible allocation found
                     GoTo again1 'try with another object
                  End If
                  GoTo again2
               End If
            
               'can allocate object
               'mark occupied space
               For yy = 0 To HeightCells - 1 'columns
                  For xx = rndnumber2 To rndnumber2 + WidthCells - 1 'rows
                     DoEvents
                     AllocableCellsVector(xx + yy * nx) = rndnumber2
                  Next
               Next
               'mark object allocated
               AllocableObjectsVector(rndnumber) = rndnumber
               'return number of objects allocated till now
               IOvariable = a
               'locate object in its container
               PopulatingOBJ(rndnumber).Left = CoordinateCellX(rndnumber2)
               PopulatingOBJ(rndnumber).Top = CoordinateCellY(rndnumber2)
               'make object visible
               PopulatingOBJ(rndnumber).Visible = True
               'put on top
               PopulatingOBJ(rndnumber).ZOrder
            Else 'try another cell
               trials2 = trials2 + 1
               If trials2 > MaxTrialsRandom Then
                  'explore intelligently the cell space for possible allocation
                  If aaamin < MaxCellsNumber - 1 Then
                     For aaa = aaamin To MaxCellsNumber - 1
                        If aaa = 0 Then Previous = 1
                        If (AllocableCellsVector(aaa) = -1) And Previous > 0 Then
                           rndnumber2 = aaa
                           Previous = AllocableCellsVector(aaa)
                           aaamin = aaa + 1
                           GoTo again22
                        End If
                        Previous = AllocableCellsVector(aaa)
                     Next
                  End If
                  aaamin = 0
                  'no possible allocation found
                  GoTo again1 'try with another object
               End If
               GoTo again2
            End If
         Else 'try another object (rndnumber)
            trials = trials + 1
            If trials > MaxTrialsRandom Then
               IOvariable = a - 1 'return number of objects allocated
               OccupiedSpaceFun = ErrorInADD
               Exit For
            End If
            GoTo again1
         End If
      Next
   End If


End Function


Public Sub RunTimeSizeObjects(ByVal PopulatingOBJ As Object)
   'basic dimensions of labels, in pixel
   Const LBLwidthPixels = 50
   Const LBLheightPixels = 25
   
   'make all labels invisible
   For a = 0 To PopulatingOBJ.Count - 1
      DoEvents
      PopulatingOBJ(a).Visible = False
   Next
   
   'PopulatingOBJ dimensions are initialized to
   'random size dimensions from 1/2 to 2 times the
   'dimensions of the constants
   For a = 0 To PopulatingOBJ.Count - 1
      DoEvents
      With PopulatingOBJ(a)
         .Move .Left, .Top, (Fix((Rnd * 4) + 1) / 2) * LBLwidthPixels * Screen.TwipsPerPixelX, (Fix((Rnd * 4) + 1) / 2) * LBLheightPixels * Screen.TwipsPerPixelY
         .Caption = "(" & Format(a, "000") & ")"
         'colored labels
         BrightRandomColor PopulatingOBJ(a) 'both fore and back colors
      End With
   Next
End Sub


Private Sub BrightRandomColor(ByVal obj As Object)
'randomly assigns forecolor (black or white) and
'backcolor (bright colors) to obj
   
   Const colminR = 100
   Const colrndR = 255 - colminR
   Const colminG = 50
   Const colrndG = 255 - colminG
   Const colminB = 100
   Const colrndB = 255 - colminB
   
   With obj
      Randomize Timer
      
      'tricky way of obtaining bright random colors
      red = colminR + Rnd * colrndR
      green = colminG + Rnd * colrndG
      blue = colminB + Rnd * colrndB
      Max = 0
      If green > red Then Max = green Else Max = red
      If blue > Max Then Max = blue
      clamp = 255
      If red = Max Then red = clamp
      If green = Max Then green = clamp
      If blue = Max Then blue = clamp

      If red + green + blue < 2 * (colminR + colminG + colminB) Then
         .ForeColor = vbWhite
      Else
         .ForeColor = vbBlack
      End If
      
      .BackColor = RGB(red, green, blue)
   
   End With
End Sub

Public Sub DesignTimeNameObjects(ByVal PopulatingOBJ As Object)
   'restore design time properties of objects
   DesignTimeBackUpRestoreObjects PopulatingOBJ, Restore
   'name each object (label) with its index, color them randomly
   'make all objects invisible
   For a = 0 To PopulatingOBJ.Count - 1
      DoEvents
      PopulatingOBJ(a).Caption = "(" & Format(a, "000") & ")"
      PopulatingOBJ(a).Visible = False
      'colored labels
      BrightRandomColor PopulatingOBJ(a) 'both fore and back colors
   Next
End Sub

Public Sub DesignTimeBackUpRestoreObjects(ByVal PopulatingOBJ As Object, ByVal action As BackUpRestoreEnum)
   Static SavedObjectsWidth() As Double
   Static SavedObjectsHeight() As Double
   Static SavedObjectsCaption() As String
   Static SavedObjectsForeColor() As Long
   Static SavedObjectsBackColor() As Long
   'saves dimensions, caption and colors of design time objects
   'for further restore
   For a = 0 To PopulatingOBJ.Count - 1
      DoEvents
      Select Case action
         Case BackUp
            If a = 0 Then 'only at the very beginning, and once
               ReDim SavedObjectsWidth(0 To PopulatingOBJ.Count - 1)
               ReDim SavedObjectsHeight(0 To PopulatingOBJ.Count - 1)
               ReDim SavedObjectsCaption(0 To PopulatingOBJ.Count - 1)
               ReDim SavedObjectsForeColor(0 To PopulatingOBJ.Count - 1)
               ReDim SavedObjectsBackColor(0 To PopulatingOBJ.Count - 1)
            End If
            
            SavedObjectsWidth(a) = PopulatingOBJ(a).Width
            SavedObjectsHeight(a) = PopulatingOBJ(a).Height
            SavedObjectsCaption(a) = PopulatingOBJ(a).Caption
            SavedObjectsForeColor(a) = PopulatingOBJ(a).ForeColor
            SavedObjectsBackColor(a) = PopulatingOBJ(a).BackColor
         Case Restore
            'make object invisible
            PopulatingOBJ(a).Visible = False
            'restore
            PopulatingOBJ(a).Width = SavedObjectsWidth(a)
            PopulatingOBJ(a).Height = SavedObjectsHeight(a)
            PopulatingOBJ(a).Caption = SavedObjectsCaption(a)
            PopulatingOBJ(a).ForeColor = SavedObjectsForeColor(a)
            PopulatingOBJ(a).BackColor = SavedObjectsBackColor(a)
      End Select
   Next
End Sub

