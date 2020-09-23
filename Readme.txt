OccupiedSpaceFun, this function allocates randomly many objects (differerent size), no collision!   
submitted to www.planet-source-code.com/vb the 31st August 2K1, by Pietro Cecchi
Email: pietrocecchi@inwind.it

Category: math
Level: advanced

Title: OccupiedSpaceFun, this function allocates randomly many objects (differerent size), no collision!

Populates container (form or picture) with objects, using the 'non-occupied space' concept. 
Any added item, while populating the container, takes some cells from the 'unoccupied space': 
it is a very fast process. Accurate demo and function design. Don't miss it!
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
OccupiedSpaceFun demo features: 
- 2 switchable containers for the objects
- viewable 'non-occupied space' cells
- design/run time object sizes
 
On ideas proposed by Gunnar g68_se@yahoo.se, THIRD EMAIL (23rd August 2K1):

Thank you very much for your time and efforts Pietro !

  It works very fine, except for one small thing: 

  My labels are all in different sizes, so it would
have to be some kind of combination of no-collision
and non-occupied space if i was to be able to
implement it :)

Best regards,
  Gunnar.

SECOND EMAIL:

Pietro, one more thing:

What i really, originally was thinking, was to be able
to 

1. Press a button, and choose a number to allocate,
let's say 20 labels, then
2. Press the button again if one likes, and chose
another number to allocate, let's say 12.
..
..
3. Press (any) more times, until the MaxObjectIndex is
reached (or, when it's not possible to allocate more
without collision).

And then the question (from my earlier mail) would
come (yes/no/cancel).

The program would know from the first time, where to
find the "unoccupied space".

I guess this would take some kind of storing the
occupied spaces or something

if i should be able to use the code, it has to be done
with (invisible) objects placed at design-time (100),
not loaded :)

I wish you many votes at PSC, i know you got mine !

/Gunnar.

FIRST EMAIL WAS (which originated my post: How to allocate HUNDREDS rectangles (labels, images...) without overlapping, no API!):

Hello Pietro !

I was studying your excellent code-example of the
rectangle-Intersections.

I got inspired by it, and i have a problem in my
application, but i do not know how to code it, i was
thinking maybe you could help me.

* I have an array of labels  
* I also have a function to place the labels out
randomly over the screen

Private sub Place_Labels

  Dim lIndex As Long

  Randomize timer

  For lIndex = 0 To cmd.Count - 1
    cmd(lIndex).Move (Rnd * (ScaleWidth -
cmd(lIndex).Width)), (Rnd * (ScaleHeight -            
    cmd(lIndex).Height))
  Next

End sub

This procedure sometimes places the labels "on top of
eachother".

1)  I would like to add some code to prevent
collisions, so that every labels is placed on a
separate position.

Maybe your code could help me here, but i do not
really know how to do this,

i think i'd have to go through all of my labels, and
check if there is already
a label placed there, but i have no idea of how to
code this.

Can you help me ?

thanks in advance !

Best regards from sweden,

  Gunnar.
++++++++++++++++++++++++++++++++++++++++++++++++++

This program has been written in Visual Basic 6.

I hope yow like the effort, have fun :)



====================================================
'USAGE
USAGE notes
   Even if this is a very small application, some notes of usage are needed: 1) usage of function, 2) usage of demo program.
   
 1) USAGE OF FUNCTION:

1A) the function enums
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

1B) the function
     Public Function OccupiedSpaceFun(ByVal ContainerOBJ As Object, ByVal PopulatingOBJ As Object, ByVal ActionCommand As ActionCommandEnum, Optional ByRef IOvariable As Integer = 1) As ReturnErrors

1C) a very important, input/output argument: IOvariable (all others are input only, i.e. ByVal)
   IOvariable, optional, is used as follows:
   - during CLEAR, to output the calculated MaxAllocableObjectsNumber
     that can be used by the application to initially set the UpDown1.Max value
   - during ADD objects: A) to supply the function with the number of objects to add to
     the container B) to return to the application the real number of objects added to container
   
1D) EXAMPLE OF USAGE of function:
   create a form and pull down a certain number of indexed objects (labels, images or whatsoever)
   pull down also a container (where the objects will be placed), i.e. a picture box, or you can use the form as container
   size at your pleasure both the container and each object
   create a module and place there the code (the enums and the function) of this module
   then use these calls (see enums for action commands and errors):
CALLS:
      a) to clear and show cells (MaxAllocObjsNumber is returned). Check ret for error (no error: ret=0)
         ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + ShowCells, MaxAllocObjsNumber)
            remember that MaxAllocObjsNumber after the call contains the max number of allocable abject
      b) to clear and hide cells (MaxAllocObjsNumber is returned). Check ret for error (no error: ret=0)
         ret = OccupiedSpaceFun(ChoosenContainer, Label1, ClearContainerRegion + HideCells, MaxAllocObjsNumber)
            MaxAllocObjsNumber As above
      c) to add objects to container (as input allocatenumber=number of objects to allocate,
                                      instead as returned value allocatenumber=number of objects really allocated,
                                      if not the same ret will be >0)
         ret = OccupiedSpaceFun(ChoosenContainer, Label1, AddObjectsToContainer, allocatenumber)
            remember to check both ret and allocatenumber, after the call
NOTE:
      multiple commands are accepted. Anyway the following combinations
      give unvoreseen results, and should be avoided (particularly the combination of clear and add commands):
      clear + show + add  adds all! because clear returns MaxAllocObjsNumber that is considered the number of objects
                          to be added to container (to add a defined number use a separated dedicated call)
      clear + hide + show  result is an hide, because show is processed before hide (first shows, then hides cells)
   NOTE also:
      each action command may generate an error return (see error enum) in ret
      the first error verified exits the function with the relative error code
 end of example of usage of function

1E) Mechanism of allocation
In allocating an object, the OccupiedSpaceFun follows these criteria:
   a) it tries 3 by 'max allocable objects number' times, a random allocation, of a random object
   b) in case on unsucces, it explores the available elemetary cells sequentially, for a possible allocation (more cells are involved, while allocating an object)
   c) if this further trial results in an unsuccess, it tries to allocate a different available object, choosen randomly
   e) should all of the above fail, the function exits, returning an error  (that, once detected by the demo program, blinks the cyan LED for few seconds)


 2) USAGE OF DEMO PROGRAM:

2A) run/design time command button, using 'run time' random sizing of objects:
When started, the program redimensions randomly all available objects (labels). This is made only once, so, even pushing the 'clear' button, the labels will come up always with the same dimensions (of course, at each 'clear', they will be located somewhereelse, in the same 'unoccupied space'). Same dimensions means same 'maximum allocable objects', as can be seen from add command caption or from etched menu comment, and hence same 'unoccupied space' (same elementary cell size and same number of cells composing the 'unoccupied space'.
To have a different size of labels you must restart the program ('maximum allocable objects number' will vary, because at any restart the total surface of objects, in general, varies) .

2B) run/design time command button, using 'design time'  objects sizes:
In this case, the sizes assigned at design time are used. So, even if the objects will be located randomly, the elementary cell size will not vary, neither the number of them ('unoccupied space') nor  the max allocable objects number. Even restarting the program the elementary cell size and the number of cells will remain the same.
Only changing, at design time, the size of at least one object, the 'occupied space' configuration will change (cell size and number of cells + 'max allocable objects number').

2C) choose container command button
Two containers for the objects are provided, in this demo: The Form and a Picture. If the container is Picture (the grayed rectangle upon the command controls panel), as it happens at the start of the program, all allocated objects will be visible. Instaed, if the container is Form (the whole form, not a part of it), some of the allocated objects could not be visible, because covered by the command panel. To see them anyway, just click on the form area, and the command panel will disappear. Click again to let the command panel reappear.

2D) clear command button
This command clears the container from any previously allocated object. Intrinsically also resets the 'add button' and the 'updown button' to initial values, basing on the 'max allocable objects number', returned by the OccupiedSpaceFun function, at 'action clear' processing.

2E) add objects command button
This button, programmed by the updown button, at its right, will try to add a certain number of objects at once. Also all remaining objects (to program the button to all remaining objects, decrease the updown button till 0 is reached, at this time the add button will be disabled (no meaning of addin zero objects), CONTINUE decreasing pushing the updown button once more downward, and, because the wrap propertyin the updawn button, the add butto is now set to max remaining objects. This is much more easy to do than to explain, so please, try once, and you will learn it forever (the command panel has been, a little immodestly, optimized for maximum performance).

2F) blinking cyan LED (located between the 'add objects command button' and the updown button)
Should The 'add objects command button' fail of adding all the programmed objects, this LED will blink for a while.

2G) updown button
To program the number of objects to be added in the container, at any pushing of the 'add objects command button'

2H) show/hide cells button
Shows or hides the 'occupied space' in terms of elementary cells (each cell consisting of the minimum width and height found in the objects collection to allocate in the container)
 
2I) menu comment item
A visible but not enabled item of menu is, very unusually, used as a memo line to remind the 'maximum allocable objects number'

2J) menu about item
Credits.

2K) menu usage item
This page. The structure is very simple an may be used to create a sort of block notes, useful either during the design of an application, either to the use, as a reminder: an added space where to write something and retrieve it further!
 
      

