
https://www.vbforums.com/showthread.php?807263-VB6-LightWeight-COM-and-vbFriendly-BaseInterfaces

VB6 LightWeight COM and vbFriendly-BaseInterfaces

    Some stuff for the advanced VB-users among the community (or the curious) ...

    I was recently working on some things in this area (preparations for the
    C-Emitter of a new VB6-compiler, with regards to "C-style defined Classes") -
    and this Tutorial is more or less a "by-product".

    I've just brought parts of it into shape, since I think this stuff can be
    useful for the community even whilst working with the old compiler.

    To gain more IDE-safety (and keep some noise out of the Tutorial-Folders),
    I've decided to implement the Base-stuff in its own little Dll-Project:
    vbInterfaces.dll

    The sources for this Helper-Dll are contained in an appropriate Folder
    (vbFriendlyInterfaces\vbInterfaces-Dll\...) in this Tutorial-Zip here:
    vbFriendlyInterfaces.zip

    The Dll-Project currently contains vbFriendly (Callback-) Interfaces for:
    - IUnknown
    - IDispatch
    - IEnumVariant
    - IPicture

    Feel free to contribute stuff you think would be useful to include in the
    Dll-Project itself - although what it currently contains with regards to
    IUnknown and IDispatch, allows to develop your own vtMyInterface-stuff
    already "separately" (in a normal VB-StdExe-project for example).

    Before entering the Tutorial-Folder and start running the Examples, please make
    sure, that you compile the vbInterfaces.dll first from the above mentioned Folder.

    The above Zip contains currently a set of 10 Tutorial-Apps, all in their own Folders
    (numbered from 0 to 9, from "easy to more advanced") - and here is the
    Tutorial-FolderList:
    .. 0 - LightWeight COM without any Helpers
    .. 1 - LightWeight LateBound-Objects
    .. 2 - LightWeight EarlyBound-Objects
    .. 3 - LightWeight Object-Lists
    .. 4 - Enumerables per vbIEnumVariant
    .. 5 - MultiEnumerations per vbIEnumerable
    .. 6 - Performance of vbIDispatch
    .. 7 - Dynamic usage of vbIDispatch
    .. 8 - Simple SOAPDemo with vbIDispatch
    .. 9 - usage of vbIPictureDisp

    For the last two Tutorial-Demos above I will post separate CodeBank articles,
    since they are larger ones - and deserve a few Extra-comments.

    Maybe some explanations for NewComers to the topic, who want to learn what
    the terms "LightWeight COM", or "C-style Class-implementation" mean:

    First, there's a clear separation to be made between "a Class" and "an Object",
    since these terms mean two different things really, which we need to look at separately.

    - "a Class" is the "BluePrint", which lives in the static Memory of our running Apps or Dlls
    - "an Object" (aka "an Instance of a Class") lives as a dynamic Memory-allocation (which refers back to the "BluePrint").

    And VB-Objects (the ones we create as Instances from a VB-ClassModules "BluePrint" per New) are quite "large animals" -
    since they will take up roughly 116 Bytes per instance-allocation, even when they don't contain any Private Variable Definitions.

    A Lightweight COM-Object can be written in VB6 (later taking up only as few as 8Bytes per Instance),
    when we resort to *.bas-Modules (similar to the code-modules one would write in plain C).

    Here's some Code, how one would implement that (basically the same, as contained in Tutorial-Folder #0):

    Let's say we want to implement a lightweight COM-Class (MyClass), which has only a single
    Method (AddTwoLongs) in its Public Interface (IMyClass).

    We start with the "BluePrint", and the VB-Module which implements that "C-style" would contain only:
    Code:

    Private Type tMyCOMcompatibleVTable
      'Space for the 3 Function-Pointers of the IUnknown-Interface
      QueryInterface As Long
      AddRef         As Long
      Release        As Long
      'followed by Space for the single Function-Pointer of our concrete Method
      AddTwoLongs    As Long
    End Type

    Private mVTable As tMyCOMcompatibleVTable 'preallocated (static, non-Heap) Space for the VTable

    Public Function VTablePtr() As Long 'the only Public Function here (later called from modMyClassFactory)
      If mVTable.QueryInterface = 0 Then InitVTable 'initializes only, when not already done
      VTablePtr = VarPtr(mVTable) 'just hand out the Pointer to the statically defined mVTable-Variable
    End Function

    Private Sub InitVTable() 'this method will be called only once (and is thus not "performance-critical")
      mVTable.QueryInterface = FuncPtr(AddressOf modMyClassFactory.QueryInterface)
      mVTable.AddRef = FuncPtr(AddressOf modMyClassFactory.AddRef)
      mVTable.Release = FuncPtr(AddressOf modMyClassFactory.Release)
      
      mVTable.AddTwoLongs = FuncPtr(AddressOf modMyClassFactory.AddTwoLongs)
    End Sub

    I assume, the above is not that difficult to understand (most "static things" are easy this way) -
    what it ensures is, that it "gathers things in one static place" - in this case:
    "Function-Pointers in a certain Order" - this "List of Function-Pointers" remains (in its defined order)
    behind the static UDT-variable mVTable - and that was it already...

    What remains (perhaps a bit more difficult to understand to "make the leap") is,
    how the above code-definition will interact, when we now come to the "dynamic part"
    (the Objects and their instantiations from a BluePrint).

    To have the dynamic part more separated, let's use an additional module (modMyClassFactory):

    And as the choosen name (modMyClassFactory) suggests, this is the part which finally hands out
    the new Instances (similar to one of the 4 exported Functions, which any ActiveX-Dll needs to support,
    which is named 'DllGetClassFactory' for a reason).

    So let's show the ObjectCreation-Function in that *.bas Module first:
    Note, that UDT struct-definitions are only there for the compiler to "have info about needed space" -
    (I've marked these Length-Info parts in light orange below - and the dynamic allocation part in magenta)...
    Code:

    Private Type tMyObject 'the Object-Instances will occupy only 8Bytes (that's half the size of a Variant-Type)
      pVTable As Long
      RefCount As Long
    End Type
     
    'Factory Helper-Function to create a new Class-Instance (a new Object) of type IMyClass
    Public Function CreateInstance() As IMyClass '<- this Type is defined in a little TypeLib, contained in TutorialFolder #0
    Dim MyObj As tMyObject 'we use our UDT-based Object-Type in a Stack-Variable for more convenience
        MyObj.pVTable = modMyClassDef.VTablePtr 'whilst filling its members (as e.g. pVTable here)
        MyObj.RefCount = 1 '<- the obvious value, since we are about to create a "fresh instance"

    Dim pMem As Long
        pMem = CoTaskMemAlloc(LenB(MyObj)) 'allocate space for our little 8Byte large Object
        Assign ByVal pMem, MyObj, LenB(MyObj) 'copy-over the Data from our local MyObj-UDT-Variable
        Assign CreateInstance, pMem 'assign the new initialized Object-Reference to the Function-Result
    End Function

    What remains now, is to provide the Implementation-code for the 4 VTable-methods (which is contained in that same Module)
    Code:

    'IUnknown-Implementation
    Public Function QueryInterface(This As tMyObject, ByVal pReqIID As Long, ppObj As stdole.IUnknown) As Long '<- HResult
      QueryInterface = &H80004002 'E_NOINTERFACE, just for safety reasons ... but there will be no casts in our little Demo
    End Function

    Public Function AddRef(This As tMyObject) As Long
      This.RefCount = This.RefCount + 1
      AddRef = This.RefCount
    End Function

    Public Function Release(This As tMyObject) As Long
      This.RefCount = This.RefCount - 1
      Release = This.RefCount
      If This.RefCount = 0 Then CoTaskMemFree VarPtr(This) '<- here's the dynamic part again, when a Class-instance dies
    End Function

    'IMyClass-implementation (IMyClass only contains this single method)
    Public Function AddTwoLongs(This As tMyObject, ByVal L1 As Long, ByVal L2 As Long, Result As Long) As Long '<- HResult
      Result = L1 + L2 'note, that we set the Result ByRef-Parameter - not the Function-Result (which would be used for Error-Transport)
    End Function

    Finally (to have it complete) a Helper-Function and a few APIs, which are contained in another small *.bas Module
    Code:

    Declare Function CoTaskMemAlloc& Lib "ole32" (ByVal sz&)
    Declare Sub CoTaskMemFree Lib "ole32" (ByVal pMem&)
    Declare Sub Assign Lib "kernel32" Alias "RtlMoveMemory" (Dst As Any, Src As Any, Optional ByVal CB& = 4)
     
    Function FuncPtr(ByVal Addr As Long) As Long 'just a small Helper for the AddressOf KeyWord
      FuncPtr = Addr
    End Function

    So, what was (codewise) posted above, is complete - and how a bare-minimum-implementation
    for a lightweight "8-Byte large COM-object" could look like in VB6 (and not much different in C) -
    no need to copy it over into your own Modules because as said, this is all part of the first little
    Demo (in Tutorial-Folder #0, which also includes the needed TypeLib to run the thing).

    Happy studying and experimenting...

    Olaf 

    Last edited by Schmidt; Oct 18th, 2015 at 12:34 AM. 