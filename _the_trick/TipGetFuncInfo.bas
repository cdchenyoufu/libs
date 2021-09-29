Option Explicit

Private Declare Function EbGetExecutingProj Lib "vba6" ( _
                         ByRef cProject As IUnknown) As Long
Private Declare Function TipGetFunctionId Lib "vba6" ( _
                         ByVal cProj As IUnknown, _
                         ByVal bstrName As Long, _
                         ByVal bstrId As Long) As Long
Private Declare Function TipGetFuncInfo Lib "vba6" ( _
                         ByVal cProj As IUnknown, _
                         ByVal bstrId As Long, _
                         ByRef pFuncInfo As Long) As Long
Private Declare Function TipGetArgCount Lib "vba6" ( _
                         ByVal pFuncInfo As Long, _
                         ByRef lCount As Long) As Long
Private Declare Function TipGetArgType Lib "vba6" ( _
                         ByVal pFuncInfo As Long, _
                         ByVal lIndex As Long, _
                         ByRef lType As Long) As Long
Private Declare Sub TipReleaseFuncInfo Lib "vba6" ( _
                    ByVal pFuncInfo As Long)

Private Const INVOKE_STEP_INFO As Long = 0
Private Const INVOKE_STEP_OVER As Long = 1
Private Const INVOKE_CONTINUE As Long = 2 'WHILE STEPPING THROUGH
Private Const INVOKE_STOP1_STEP As Long = 3 '?? aborts step mode
Private Const INVOKE_STOP2_STEP As Long = 4 '?? aborts step mode
Private Const INVOKE_SAVE_AS_TEXT As Long = 5 ' save module as text file (save dialog)
Private Const INVOKE_INSERT_FILE As Long = 6
Private Const INVOKE_UNDO As Long = 7
Private Const INVOKE_REDO As Long = 8
'Private Const INVOKE_ As Long =9 '?? Nothing
Private Const INVOKE_WINDOW_SPLIT As Long = 10
Private Const INVOKE_CUT As Long = 11
Private Const INVOKE_COPY As Long = 12
Private Const INVOKE_PASTE As Long = 13
Private Const INVOKE_DELETE As Long = 14
Private Const INVOKE_FIND As Long = 15
Private Const INVOKE_FINDNEXT As Long = 16
Private Const INVOKE_FINDNEXTREVERSE As Long = 17 'up search if toggle is down by default, otherwise opposite
Private Const INVOKE_FIND_REPLACE As Long = 18
Private Const INVOKE_TOGGLE_BREAKPOINT As Long = 19
Private Const INVOKE_CLEAR_ALL_BREAKPOINTS As Long = 20
Private Const INVOKE_ADD_WATCH As Long = 21
Private Const INVOKE_EDIT_WATCH As Long = 22
Private Const INVOKE_QUICK_WATCH As Long = 23
'Private Const INVOKE_ As Long =24 '?? Nothing
Private Const INVOKE_VIEW_WATCH As Long = 25
Private Const INVOKE_VIEW_LOCALS As Long = 26
Private Const INVOKE_VIEW_IMMEDIATE As Long = 27
Private Const INVOKE_GOTO_DEFINITION As Long = 28 'object browser loaded with token to be defined
Private Const INVOKE_REFERENCES As Long = 29
Private Const INVOKE_CALL_STACK As Long = 30
'Private Const INVOKE_ As Long = 31 '?? CRASH
'Private Const INVOKE_ As Long =32 '?? Nothing
'Private Const INVOKE_ As Long =33'?? CRASH
Private Const INVOKE_INDENT As Long = 34
Private Const INVOKE_OUTDENT As Long = 35
Private Const INVOKE_LAST_POSITION As Long = 36
'Private Const INVOKE_ As Long =37'?? Nothing
Private Const INVOKE_TOGGLE_BOOKMARK As Long = 38
Private Const INVOKE_NEXT_BOOKMARK As Long = 39
Private Const INVOKE_PREVIOUS_BOOKMARK As Long = 40
Private Const INVOKE_CLEAR_ALL_BOOKMARKS As Long = 41
'Private Const INVOKE_ As Long =42 '?? Nothing
Private Const INVOKE_OBJECT_BROWSER As Long = 43 'object browser direct
Private Const INVOKE_STEP_OUT As Long = 44
Private Const INVOKE_LIST_PROPERTIES_METHODS As Long = 45
Private Const INVOKE_COMPLETE_WORD As Long = 46
Private Const INVOKE_LIST_CONSTANTS As Long = 47
Private Const INVOKE_BEEP As Long = 48 'Beep sound?
'Private Const INVOKE_ As Long =49'?? CRASH
Private Const INVOKE_DELETE_LINE As Long = 50 'delete line function (Ctrl+Y)
Private Const INVOKE_NEW_LINE As Long = 51
Private Const INVOKE_QUICK_INFO As Long = 52
Private Const INVOKE_COMMENT As Long = 53
Private Const INVOKE_UNCOMMENT As Long = 54
Private Const INVOKE_FINDNEXT_NO_DIALOG As Long = 55
Private Const INVOKE_FINDNEXTREVERSE_NO_DIALOG As Long = 56
Private Const INVOKE_PARAM_INFO As Long = 57
'Private Const INVOKE_ As Long = 58 '?? Nothing
'Private Const INVOKE_ As Long =59'?? Nothing
'Private Const INVOKE_ As Long =60'?? Nothing
Private Const INVOKE_SELECT_ALL As Long = 61

Private Declare Function apiEbInvokeItem Lib "vba6.dll" Alias "EbInvokeItem" (ByVal id As Long) As Long
                     
Sub Main()
    Dim cProj   As IUnknown
    Dim sFnId   As String
    Dim pFnInfo As Long
    Dim lArgs   As Long
    Dim lIndex  As Long
    Dim lType   As Long
    
    EbGetExecutingProj cProj
    TipGetFunctionId cProj, StrPtr("foo"), VarPtr(sFnId)
    TipGetFuncInfo cProj, StrPtr(sFnId), pFnInfo
    TipGetArgCount pFnInfo, lArgs
    
    For lIndex = 0 To lArgs
        ' // 0 - return type
        TipGetArgType pFnInfo, lIndex, lType
        Debug.Print lType
    Next
    
    TipReleaseFuncInfo pFnInfo
    
End Sub


Private Function TypeString(ByVal lType As Long) As String
    lType = lType And Not vbArray       ' We don't report if it's an array.
    Select Case lType
    Case vbInteger:     TypeString = "Integer"
    Case vbLong:        TypeString = "Long"
    Case vbSingle:      TypeString = "Single"
    Case vbDouble:      TypeString = "Double"
    Case vbCurrency:    TypeString = "Currency"
    Case vbDate:        TypeString = "Date"
    Case vbString:      TypeString = "String"
    Case vbObject:      TypeString = "Object"
    Case vbBoolean:     TypeString = "Boolean"
    Case vbVariant:     TypeString = "Variant"
    Case vbByte:        TypeString = "Byte"
    Case 24:            TypeString = "Void" ' This is the type value for SUB procedure's return value.
    Case Else:          TypeString = "Unknown"
    End Select
End Function

Public Function foo( _
                ByVal lArg1 As Long, _
                ByVal vArg2 As Variant, _
                ByVal sArg3 As String) As Boolean
                
End Function