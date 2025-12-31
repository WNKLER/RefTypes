Attribute VB_Name = "Module20"
'MIT License
'https://github.com/WNKLER/RefTypes
Option Private Module
Option Explicit

 #If VBA7 <> 1 Then
    Private Enum LONG_PTR: [_LONG_PTR]: End Enum
     Public Enum LongPtr:   [_LongPtr]: End Enum '// Must be Public for Enum-typed Public Property
    Private Const NullPtr = [_LongPtr]
 #Else
    Private Const NullPtr As LongPtr = 0
 #End If
    
Private Enum Context
 #If Win64 = 1 Then
    [_Win32] '// 0 on x64; undefined on x32.
 #End If
    [_Win64] '// 0 when [_Win32] is undefined; otherwise, [_Win32] + 1.
    [_PtrSize] = 4& + ([_Win64] * 4&)
End Enum

Private Const Win64 As Integer = [_Win64]
Private Const cLongPtr As Long = [_PtrSize]
Private Const wHalfPtr As Long = cLongPtr \ 4&

'// Implicit typing allows for (effectively) LongPtr-typed constants
Private Const oLongPtr = NullPtr + cLongPtr
Private Const oNativeCallBack = (NullPtr + 22) + (Win64 * 33)
Private Const oProcDscInfoPtr = (Win64 * oLongPtr) - (Not -Win64)
Private Const o8h = NullPtr + 8

Private Type HalfPtr
    Bytes As String * wHalfPtr
End Type

Private Type StackMemory
    Bytes(-1& To 0&) As HalfPtr
End Type

Private Type BindArgs
    This As LongPtr: pCalled As LongPtr: pActual As LongPtr
End Type

'// NOTE: This `redbinding` technique only works for VBA and p-code executables.
'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Internals] ////////////////////////////////////////////////////////////////////////////////

Private Enum ProcIndex                '// Provides identifiers for ImportTable indices
: [_GetP]:    [_PutP]:    [_MovP]     '// by matching member declaration-order to
: [_Get1]:    [_Put1]:    [_Mov1]     '// `LayoutImportTable()` return-order.
: [_Get2]:    [_Put2]:    [_Mov2]     '//
: [_Get4]:    [_Put4]:    [_Mov4]     '// Syntactic Sugar. Could just use literals.
: [_Get8]:    [_Put8]:    [_Mov8]
: [_GetB]:    [_PutB]:    [_MovB]

: [_GetPtr]:  [_LetPtr]:  [_CopyPtr]
: [_GetByte]: [_LetByte]: [_CopyByte]
: [_GetInt]:  [_LetInt]:  [_CopyInt]
: [_GetLng]:  [_LetLng]:  [_CopyLng]
: [_GetCur]:  [_LetCur]:  [_CopyCur]
: [_GetBool]: [_LetBool]: [_CopyBool]
End Enum

Private BindArgs As BindArgs

'// Never intended to be run.
Private Sub LayoutImportTable(ByRef A As LongPtr, ByRef AA As LongPtr)
    Select Case True
        Case True, False                       '// The presence of a procedure call anywhere in a Module's
        Case Else: Exit Sub                    '// code adds that procedure to the Module's ImportTable.
            Call GetP:   Call PutP:  Call MovP '// ImportTable entries are added in return-order (roughly).
            Call Get1:   Call Put1:  Call Mov1 '//
            Call Get2:   Call Put2:  Call Mov2 '// The only reason we need the ImportTable at all is because
            Call Get4:   Call Put4:  Call Mov4 '// we can't use `AddressOf` on Property Let/Set procedures.
            Call Get8:   Call Put8:  Call Mov8 '// Sacrificing the luxury of Property-based accessors
            Call GetB:   Call PutB:  Call MovB '// would greatly simplify this project.
        
            RefPtr(AA) = RefPtr(AA): Call CopyPtr(AA, A)
            RefByte(A) = RefByte(A): Call CopyByte(A, A)
            RefInt(AA) = RefInt(AA): Call CopyInt(AA, A)
            RefLng(AA) = RefLng(AA): Call CopyLng(AA, A)
            RefCur(AA) = RefCur(AA): Call CopyCur(AA, A)
            RefBool(A) = RefBool(A): Call CopyBool(A, A)
    End Select
End Sub

'// This does the thing
Private Property Let SetBind(ByVal Index_Called As ProcIndex, ByVal Index_Actual As ProcIndex)
    Dim ppCalled As LongPtr, ppActual As LongPtr
    ppCalled = GetBind(Index_Called)
    ppActual = GetBind(Index_Actual)
    MovPtr ppCalled, ppActual
End Property

'// [Helpers] //////////////////////////////////
Private Function GetBind(ByVal index As ProcIndex) As LongPtr
    GetBind = ImportTable + oLongPtr * index
    GetPtr GetBind, GetBind
    GetBind = GetBind + oProcDscInfoPtr
End Function

Private Function ImportTable() As LongPtr
  Const oImportTable = oLongPtr * (13 - Win64)
    ImportTable = EpiModule + oImportTable
    GetPtr ImportTable, ImportTable
End Function

Private Function EpiModule() As LongPtr
    EpiModule = UnWrapCallBack(AddressOf UnWrapCallBack)
    GetPtr EpiModule, EpiModule
    GetPtr EpiModule, EpiModule
End Function

Private Function UnWrapCallBack(ByVal AddressOf_Proc As LongPtr) As LongPtr
    GetPtr AddressOf_Proc, AddressOf_Proc + oNativeCallBack
 #If Win64 Then
    GetPtr AddressOf_Proc, AddressOf_Proc - oLongPtr
 #End If
    UnWrapCallBack = AddressOf_Proc + oProcDscInfoPtr
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Init] (Checkless. Runs only once, automatically.) /////////////////////////////////////////

Private Sub GetPtr(Optional ByRef Target As LongPtr, Optional ByVal Source As LongPtr)
    Rebind AddressOf GetPtr, AddressOf MovP
    GetPtr Target, Source
End Sub
Private Sub PutPtr(Optional ByVal Target As LongPtr, Optional ByRef Source As LongPtr)
    Rebind AddressOf PutPtr, AddressOf MovP
    PutPtr Target, Source
End Sub
Private Sub MovPtr(Optional ByVal Target As LongPtr, Optional ByVal Source As LongPtr)
    Rebind AddressOf MovPtr, AddressOf MovP
    MovPtr Target, Source
End Sub

Private Sub Rebind(ByVal AddressOf_Called As LongPtr, ByVal AddressOf_Actual As LongPtr)
    Bind NullPtr, AddressOf_Called, AddressOf_Actual
End Sub

Private Sub Bind(Optional ByVal Args As LongPtr, Optional ByRef Called As LongPtr, Optional ByRef Actual As LongPtr)
    With BindArgs               '// HighPtr(Here.Bytes)
        Dim Here As StackMemory '// With-block Accessor
        HighPtr(Here.Bytes) = VarPtr(Args) '// Set With-block address to VarPtr(Args)
        
        .pCalled = Called + oNativeCallBack
        .pActual = Actual + oNativeCallBack
     #If Win64 Then
        .pCalled = Called - oLongPtr
        .pActual = Actual - oLongPtr
     #End If
        .pCalled = Called + oProcDscInfoPtr
        .pActual = Actual + oProcDscInfoPtr
    End With
    
    Called = Actual
End Sub

Private Property Let HighPtr(ByRef HalfPtr() As LONG_PTR, ByVal Address As LongPtr)
    HalfPtr(0&) = Address
End Property

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Internal Accessors] (The actual code that runs when an exposed accessor is called.) ///////

'// [Pointer] ////////////////////////////////////
Private Function GetP(Optional ByRef Target As LongPtr) As LongPtr
    GetP = Target
End Function
     Private Sub PutP(Optional ByRef Target As LongPtr, Optional ByVal Source As LongPtr)
    Target = Source
End Sub
     Private Sub MovP(Optional ByRef Target As LongPtr, Optional ByRef Source As LongPtr)
    Target = Source
End Sub

'// [One Byte] /////////////////////////////////
Private Function Get1(Optional ByRef Target As Byte) As Byte
    Get1 = Target
End Function
     Private Sub Put1(Optional ByRef Target As Byte, Optional ByVal Source As Byte)
    Target = Source
End Sub
     Private Sub Mov1(Optional ByRef Target As Byte, Optional ByRef Source As Byte)
    Target = Source
End Sub

'// [Two Bytes] ////////////////////////////////
Private Function Get2(Optional ByRef Target As Integer) As Integer
    Get2 = Target
End Function
     Private Sub Put2(Optional ByRef Target As Integer, Optional ByVal Source As Integer)
    Target = Source
End Sub
     Private Sub Mov2(Optional ByRef Target As Integer, Optional ByRef Source As Integer)
    Target = Source
End Sub

'// [Four Bytes] ///////////////////////////////
Private Function Get4(Optional ByRef Target As Long) As Long
    Get4 = Target
End Function
     Private Sub Put4(Optional ByRef Target As Long, Optional ByVal Source As Long)
    Target = Source
End Sub
     Private Sub Mov4(Optional ByRef Target As Long, Optional ByRef Source As Long)
    Target = Source
End Sub

'// [Eight Bytes] //////////////////////////////
Private Function Get8(Optional ByRef Target As Currency) As Currency
    Get8 = Target
End Function
     Private Sub Put8(Optional ByRef Target As Currency, Optional ByVal Source As Currency)
    Target = Source
End Sub
     Private Sub Mov8(Optional ByRef Target As Currency, Optional ByRef Source As Currency)
    Target = Source
End Sub

'// [Boolean] //////////////////////////////////
Private Function GetB(Optional ByRef Target As Boolean) As Boolean
    GetB = Target
End Function
     Private Sub PutB(Optional ByRef Target As Boolean, Optional ByVal Source As Boolean)
    Target = Source
End Sub
     Private Sub MovB(Optional ByRef Target As Boolean, Optional ByRef Source As Boolean)
    Target = Source
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Exposed Accessors] ////////////////////////////////////////////////////////////////////////
'// --- These only run only once. //////////////////////////////////////////////////////////////
'// ------ (On first-call): ////////////////////////////////////////////////////////////////////
'// ----------- 1. Rebind Self /////////////////////////////////////////////////////////////////
'// ----------- 2. Invoke Self as-called ///////////////////////////////////////////////////////
'// ---------- [3. Return Result] //////////////////////////////////////////////////////////////

'// [Pointer] //////////////////////////////////
Public Property Get RefPtr(ByVal Target As LongPtr) As LongPtr
    SetBind([_GetPtr]) = [_GetP]
    RefPtr = RefPtr(Target)
End Property
Public Property Let RefPtr(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_LetPtr]) = [_PutP]
    RefPtr(Target) = Source
End Property

Public Sub CopyPtr(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyPtr]) = [_MovP]
    CopyPtr Target, Source
End Sub

'// [Byte] /////////////////////////////////////
Public Property Get RefByte(ByVal Target As LongPtr) As Byte
    SetBind([_GetByte]) = [_Get1]
    RefByte = RefByte(Target)
End Property
Public Property Let RefByte(ByVal Target As LongPtr, ByVal Source As Byte)
    SetBind([_LetByte]) = [_Put1]
    RefByte(Target) = Source
End Property

Public Sub CopyByte(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyByte]) = [_Mov1]
    CopyByte Target, Source
End Sub

'// [Integer] //////////////////////////////////
Public Property Get RefInt(ByVal Target As LongPtr) As Integer
    SetBind([_GetInt]) = [_Get2]
    RefInt = RefInt(Target)
End Property
Public Property Let RefInt(ByVal Target As LongPtr, ByVal Source As Integer)
    SetBind([_LetInt]) = [_Put2]
    RefInt(Target) = Source
End Property

Public Sub CopyInt(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyInt]) = [_Mov2]
    CopyInt Target, Source
End Sub

'// [Long] /////////////////////////////////////
Public Property Get RefLng(ByVal Target As LongPtr) As Long
    SetBind([_GetLng]) = [_Get4]
    RefLng = RefLng(Target)
End Property
Public Property Let RefLng(ByVal Target As LongPtr, ByVal Source As Long)
    SetBind([_LetLng]) = [_Put4]
    RefLng(Target) = Source
End Property

Public Sub CopyLng(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyLng]) = [_Mov4]
    CopyLng Target, Source
End Sub

'// [Currency] /////////////////////////////////
Public Property Get RefCur(ByVal Target As LongPtr) As Currency
    SetBind([_GetCur]) = [_Get8]
    RefCur = RefCur(Target)
End Property
Public Property Let RefCur(ByVal Target As LongPtr, ByVal Source As Currency)
    SetBind([_LetCur]) = [_Put8]
    RefCur(Target) = Source
End Property

Public Sub CopyCur(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyCur]) = [_Mov8]
    CopyCur Target, Source
End Sub

'// [Boolean] //////////////////////////////////
Public Property Get RefBool(ByVal Target As LongPtr) As Boolean
    SetBind([_GetBool]) = [_GetB]
    RefBool = RefBool(Target)
End Property
Public Property Let RefBool(ByVal Target As LongPtr, ByVal Source As Boolean)
    SetBind([_LetBool]) = [_PutB]
    RefBool(Target) = Source
End Property

Public Sub CopyBool(ByVal Target As LongPtr, ByVal Source As LongPtr)
    SetBind([_CopyBool]) = [_MovB]
    CopyBool Target, Source
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Exposed Utilities] (Assorted) /////////////////////////////////////////////////////////////

'// [tagVARIANT._Val] //////////////////////////
Public Property Get VarVal(ByRef VarVar As Variant) As LongPtr
    VarVal = RefPtr(VarPtr(VarVar) + o8h)
End Property
Public Property Let VarVal(ByRef VarVar As Variant, ByVal Val As LongPtr)
    RefPtr(VarPtr(VarVar) + o8h) = Val
End Property

'// `AddressOf` operator only accepts Sub/Function/Property_Get identifiers.
'//  Property_Let/Property_Set (propput[ref]) identifier operands are invalid.
Public Function RebindNonPut(ByVal AddressOf_Called As LongPtr, ByVal AddressOf_Actual As LongPtr) As LongPtr
    GetPtr AddressOf_Called, AddressOf_Called + oNativeCallBack
    GetPtr AddressOf_Actual, AddressOf_Actual + oNativeCallBack
 #If Win64 Then
    GetPtr AddressOf_Called, AddressOf_Called - oLongPtr
    GetPtr AddressOf_Actual, AddressOf_Actual - oLongPtr
 #End If
    AddressOf_Called = AddressOf_Called + oProcDscInfoPtr
    AddressOf_Actual = AddressOf_Actual + oProcDscInfoPtr
    
    GetPtr RebindNonPut, AddressOf_Called '// Returns the overwritten address
    MovPtr AddressOf_Called, AddressOf_Actual
End Function

