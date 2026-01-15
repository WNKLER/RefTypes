Attribute VB_Name = "RefTypes"
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

Private Const NullVar As Variant = Empty

Private Type HalfPtr
    Bytes As String * wHalfPtr
End Type

Private Type StackMemory
    Bytes(-1& To 0&) As HalfPtr
End Type

Private Type RebindArgs
    This As LongPtr: pCalled As LongPtr: pActual As LongPtr
End Type

Private Type tagVARIANT
    vt    As Integer
    wReserved1_2_3 As String * 3&
    val   As LongPtr
    valEx As LongPtr
End Type

'// NOTE: This `redbinding` technique only works for VBA and p-code executables.
'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Internals] ////////////////////////////////////////////////////////////////////////////////

Private Enum ImportTableIndex: [_SetBind]
: [_GetP]:    [_PutP]:    [_MovP] '// Provides identifiers for ImportTable indices
: [_Get1]:    [_Put1]:    [_Mov1] '// by matching the call-order in SetBind().
: [_Get2]:    [_Put2]:    [_Mov2] '//
: [_Get4]:    [_Put4]:    [_Mov4] '// Syntactic Sugar. Could just use literals.
: [_Get8]:    [_Put8]:    [_Mov8]
: [_LetV]:    [_SetV]:    [_MovV]
: [_GetU]:    [_PutU]:    [_MovU]

: [_GetPtr]:  [_LetPtr]:  [_CopyPtr]
: [_GetByte]: [_LetByte]: [_Copy1]
: [_GetInt]:  [_LetInt]:  [_Copy2]
: [_GetLng]:  [_LetLng]:  [_Copy4]
: [_GetCur]:  [_LetCur]:  [_Copy8]
: [_GetVar]:  [_LetVar]:  [_CopyVar]
: [_SetVar]

: [_GetVT]:   [_LetVT]:   [_CopyVT]
: [_GetVal]:  [_LetVal]:  [_CopyVal]
End Enum

Private IsBuilding As Boolean
Private RebindArgs As RebindArgs

'// [Helpers] //////////////////////////////////
Private Property Let SetBind(ByVal Index_Called As ImportTableIndex, Optional ByRef Target As Variant, Optional ByRef Source As Variant, ByVal Index_Actual As ImportTableIndex)
    Dim U As tagVARIANT, H() As HalfPtr

  Const A = NullPtr, AA = A, AAA = A, AAAA = A, L As Long = [_SetBind]
  Const V = NullVar, VV = V, VVV = V, VVVV = V, VVVVV = V

 Static IsBuilt As Boolean
 
    If IsBuilt Then '// Most of this is for compatibility with "Compile On Demand" / "Background Compile"
    ElseIf IsBuilding Then Exit Property
    Else:  IsBuilding = True: SetBind(L) = L

        Call GetP:    Call PutP:    Call MovP
        Call Get1:    Call Put1:    Call Mov1
        Call Get2:    Call Put2:    Call Mov2
        Call Get4:    Call Put4:    Call Mov4
        Call Get8:    Call Put8:    Call Mov8
        Call LetV:    Call SetV:    Call MovV(U, U)
        Call GetU(U): Call PutU(U): Call MovU(U, U)

        RefPtr(AAA) = RefPtr(AAAA): Call CopyPtr(A, A)
        RefByte(AA) = RefByte(AAA): Call Copy1(AAA, A)
        RefInt(AAA) = RefInt(AAAA): Call Copy2(AAA, A)
        RefLng(AAA) = RefLng(AAAA): Call Copy4(AAA, A)
        RefCur(AAA) = RefCur(AAAA): Call Copy8(AAA, A)
        RefVar(AAA) = RefVar(AAAA): Call CopyVar(A, A)
        Set RefVar(A) = Nothing

        VarVT(VVVV) = VarVT(VVVVV): Call CopyVT(VV, V)
        VarVal(VVV) = VarVal(VVVV): Call CopyVal(V, V)

        Call VarPtr(AA): HighPtr(H) = A
        Call Rebind: Call EnsureBindPtr
        Call UnWrapCallBack: Call EpiModule: Call ImportTable
        Call GetBind(L)

        IsBuilt = True: IsBuilding = False
    End If
    
    EnsureBindPtr
    CopyPtr GetBind(Index_Called), GetBind(Index_Actual)

 #If Win64 Then '// Can't write to ByRef VT_I8 Variant, so make it VT_CY.
    Select Case Index_Called: Case [_GetPtr], [_GetVal]
        With U: Dim Here As StackMemory
            HighPtr(Here.Bytes) = VarPtr(Target): .vt = &H4006
        End With
    End Select
 #End If

    Select Case Index_Called
        Case [_GetPtr]:  Target = RefPtr(Source):  Case [_LetPtr]:   RefPtr(Target) = Source: Case [_CopyPtr]: CopyPtr Target, Source
        Case [_GetByte]: Target = RefByte(Source): Case [_LetByte]: RefByte(Target) = Source: Case [_Copy1]:     Copy1 Target, Source
        Case [_GetInt]:  Target = RefInt(Source):  Case [_LetInt]:   RefInt(Target) = Source: Case [_Copy2]:     Copy2 Target, Source
        Case [_GetLng]:  Target = RefLng(Source):  Case [_LetLng]:   RefLng(Target) = Source: Case [_Copy4]:     Copy4 Target, Source
        Case [_GetCur]:  Target = RefCur(Source):  Case [_LetCur]:   RefCur(Target) = Source: Case [_Copy8]:     Copy8 Target, Source
        Case [_GetVar]:  Target = RefVar(Source):  Case [_LetVar]:   RefVar(Target) = Source: Case [_CopyVar]: CopyVar Target, Source
        Case [_SetVar]: Set RefVar(Target) = Source

        Case [_GetVT]:   Target = VarVT(Source):   Case [_LetVT]:     VarVT(Target) = Source: Case [_CopyVT]:   CopyVT Target, Source
        Case [_GetVal]:  Target = VarVal(Source):  Case [_LetVal]:   VarVal(Target) = Source: Case [_CopyVal]: CopyVal Target, Source
    End Select
End Property

Private Function GetBind(ByVal Index As ImportTableIndex) As LongPtr
    If IsBuilding Then Exit Function
    GetBind = RefPtr(ImportTable + oLongPtr * Index) + oProcDscInfoPtr
End Function

Private Function ImportTable() As LongPtr
  Const oImportTable = oLongPtr * (13 - Win64)
    If IsBuilding Then Exit Function
    ImportTable = RefPtr(EpiModule + oImportTable)
End Function

Private Function EpiModule() As LongPtr
    If IsBuilding Then Exit Function
    EpiModule = RefPtr(RefPtr(UnWrapCallBack(AddressOf CopyPtr)))
End Function

Private Function UnWrapCallBack(Optional ByVal AddressOf_Proc As LongPtr) As LongPtr
    If IsBuilding Then Exit Function
    AddressOf_Proc = RefPtr(AddressOf_Proc + oNativeCallBack)
 #If Win64 Then
    AddressOf_Proc = RefPtr(AddressOf_Proc - oLongPtr)
 #End If
    UnWrapCallBack = AddressOf_Proc + oProcDscInfoPtr
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Init] /////////////////////////////////////////////////////////////////////////////////////

Private Sub EnsureBindPtr()
    If IsBuilding Then Exit Sub
    Rebind NullPtr, AddressOf RefPtr, AddressOf GetP
    Rebind NullPtr, AddressOf CopyPtr, AddressOf MovP
End Sub

Private Sub Rebind(Optional ByVal Args As LongPtr, Optional ByRef Called As LongPtr, Optional ByRef Actual As LongPtr)
    If IsBuilding Then Exit Sub
    With RebindArgs             '// HighPtr(Here.Bytes)
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
    If IsBuilding Then Exit Property
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

'// [tagVARIANT] ///////////////////////////////
Private Sub LetV(Optional ByRef Target As Variant, Optional ByRef Source As Variant)
    Target = Source
End Sub
Private Sub SetV(Optional ByRef Target As Variant = Nothing, Optional ByRef Source As Variant = Nothing)
    Set Target = Source
End Sub
Private Sub MovV(ByRef Target As tagVARIANT, ByRef Source As tagVARIANT)
    Target = Source
End Sub

'// [tagVARIANT.val] ///////////////////////////
Private Function GetU(ByRef Target As tagVARIANT) As LongPtr
    GetU = Target.val
End Function
     Private Sub PutU(ByRef Target As tagVARIANT, Optional ByVal Source As LongPtr)
    Target.val = Source
End Sub
     Private Sub MovU(ByRef Target As tagVARIANT, ByRef Source As tagVARIANT)
    Target.val = Source.val
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
    If IsBuilding Then Exit Property Else SetBind([_GetPtr], RefPtr, (Target)) = [_GetP]
End Property
Public Property Let RefPtr(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Property Else SetBind([_LetPtr], (Target), (Source)) = [_PutP]
End Property

Public Sub CopyPtr(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_CopyPtr], (Target), (Source)) = [_MovP]
End Sub

'// [Byte] /////////////////////////////////////
Public Property Get RefByte(ByVal Target As LongPtr) As Byte
    If IsBuilding Then Exit Property Else SetBind([_GetByte], RefByte, (Target)) = [_Get1]
End Property
Public Property Let RefByte(ByVal Target As LongPtr, ByVal Source As Byte)
    If IsBuilding Then Exit Property Else SetBind([_LetByte], (Target), (Source)) = [_Put1]
End Property

Public Sub Copy1(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_Copy1], (Target), (Source)) = [_Mov1]
End Sub

'// [Integer] //////////////////////////////////
Public Property Get RefInt(ByVal Target As LongPtr) As Integer
    If IsBuilding Then Exit Property Else SetBind([_GetInt], RefInt, (Target)) = [_Get2]
End Property
Public Property Let RefInt(ByVal Target As LongPtr, ByVal Source As Integer)
    If IsBuilding Then Exit Property Else SetBind([_LetInt], (Target), (Source)) = [_Put2]
End Property

Public Sub Copy2(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_Copy2], (Target), (Source)) = [_Mov2]
End Sub

'// [Long] /////////////////////////////////////
Public Property Get RefLng(ByVal Target As LongPtr) As Long
    If IsBuilding Then Exit Property Else SetBind([_GetLng], RefLng, (Target)) = [_Get4]
End Property
Public Property Let RefLng(ByVal Target As LongPtr, ByVal Source As Long)
    If IsBuilding Then Exit Property Else SetBind([_LetLng], (Target), (Source)) = [_Put4]
End Property

Public Sub Copy4(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_Copy4], (Target), (Source)) = [_Mov4]
End Sub

'// [Currency] /////////////////////////////////
Public Property Get RefCur(ByVal Target As LongPtr) As Currency
    If IsBuilding Then Exit Property Else SetBind([_GetCur], RefCur, (Target)) = [_Get8]
End Property
Public Property Let RefCur(ByVal Target As LongPtr, ByVal Source As Currency)
    If IsBuilding Then Exit Property Else SetBind([_LetCur], (Target), (Source)) = [_Put8]
End Property

Public Sub Copy8(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_Copy8], (Target), (Source)) = [_Mov8]
End Sub

'// [Variant] //////////////////////////////////
Public Property Get RefVar(ByVal Target As LongPtr) As Variant
    If IsBuilding Then Exit Property Else SetBind([_GetVar], RefVar, (Target)) = [_MovV]
End Property
Public Property Let RefVar(ByVal Target As LongPtr, ByRef Source As Variant)
    If IsBuilding Then Exit Property Else SetBind([_LetVar], (Target), Source) = [_LetV]
End Property

Public Sub CopyVar(ByVal Target As LongPtr, ByVal Source As LongPtr)
    If IsBuilding Then Exit Sub Else SetBind([_CopyVar], (Target), (Source)) = [_MovV]
End Sub

Public Property Set RefVar(ByVal Target As LongPtr, ByRef Source As Variant)
    If IsBuilding Then Exit Property Else SetBind([_SetVar], (Target), Source) = [_SetV]
End Property

'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Exposed Utilities] (Assorted) /////////////////////////////////////////////////////////////

'// [tagVARIANT.vt] ////////////////////////////
Public Property Get VarVT(ByRef VarVar As Variant) As Integer
    If IsBuilding Then Exit Property Else SetBind([_GetVT], VarVT, VarVar) = [_Get2]
End Property
Public Property Let VarVT(ByRef VarVar As Variant, ByVal vt As Integer)
    If IsBuilding Then Exit Property Else SetBind([_LetVT], VarVar, (vt)) = [_Put2]
End Property

Public Sub CopyVT(ByRef Target As Variant, ByRef Source As Variant)
    If IsBuilding Then Exit Sub Else SetBind([_CopyVT], Target, Source) = [_Mov2]
End Sub

'// [tagVARIANT.val] ///////////////////////////
Public Property Get VarVal(ByRef VarVar As Variant) As LongPtr
    If IsBuilding Then Exit Property Else SetBind([_GetVal], VarVal, VarVar) = [_GetU]
End Property
Public Property Let VarVal(ByRef VarVar As Variant, ByVal val As LongPtr)
    If IsBuilding Then Exit Property Else SetBind([_LetVal], VarVar, (val)) = [_PutU]
End Property

Public Sub CopyVal(ByRef Target As Variant, ByRef Source As Variant)
    If IsBuilding Then Exit Sub Else SetBind([_CopyVal], Target, Source) = [_MovU]
End Sub


