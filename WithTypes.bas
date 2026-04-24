Attribute VB_Name = "WithTypes"
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
Private Const oLongPtrx2 = oLongPtr * 2
Private Const BasePtr = NullPtr, _
              NewBase = BasePtr + 0

Private Type HalfPtr
    Bytes As String * wHalfPtr
End Type

Private Enum Offsets
    [+0x0] = -1&
    [+0xLongPtr]
End Enum

Public Type WithBlock
    Offsets([+0x0] To [+0xLongPtr]) As HalfPtr
End Type

Private Type CallerContextArgs
    pContext As LongPtr: pFrameOffset As LongPtr: NewContext As LongPtr
End Type

'// With-block Interfaces for Intrinsic Types
Public Type IByte:         Byte As Byte:     End Type
Public Type IInteger:   Integer As Integer:  End Type
Public Type IBoolean:   Boolean As Boolean:  End Type
Public Type ILong:         Long As Long:     End Type
Public Type ISingle:     Single As Single:   End Type
Public Type ILongPtr:   LongPtr As LongPtr:  End Type
Public Type IString:     String As String:   End Type
Public Type IObject:     Object As Object:   End Type
Public Type ILongLong: LongLong As LongLong: End Type
Public Type IDouble:     Double As Double:   End Type
Public Type ICurrency: Currency As Currency: End Type
Public Type IDate:         Date As Date:     End Type
Public Type IVariant:   Variant As Variant:  End Type

Private CallerContextArgs As CallerContextArgs

Private Function test() As LongPtr
    With NewByte(VarPtr(test))
        Debug.Print VarPtr(.Byte) = VarPtr(test)
    End With
    
    With NewInteger(VarPtr(test))
        Debug.Print VarPtr(.Integer) = VarPtr(test)
    End With
    
    With NewBoolean(VarPtr(test))
        Debug.Print VarPtr(.Boolean) = VarPtr(test)
    End With
    
    With NewLong(VarPtr(test))
        Debug.Print VarPtr(.Long) = VarPtr(test)
    End With
    
    With NewSingle(VarPtr(test))
        Debug.Print VarPtr(.Single) = VarPtr(test)
    End With
    
    With NewLongPtr(VarPtr(test))
        Debug.Print VarPtr(.LongPtr) = VarPtr(test)
    End With
    
    With NewString(VarPtr(test))
        Debug.Print VarPtr(.String) = VarPtr(test)
    End With
    
    With NewObject(VarPtr(test))
        Debug.Print VarPtr(.Object) = VarPtr(test)
    End With
    
    With NewLongLong(VarPtr(test))
        Debug.Print VarPtr(.LongLong) = VarPtr(test)
    End With
    
    With NewDouble(VarPtr(test))
        Debug.Print VarPtr(.Double) = VarPtr(test)
    End With
    
    With NewCurrency(VarPtr(test))
        Debug.Print VarPtr(.Currency) = VarPtr(test)
    End With
    
    With NewDate(VarPtr(test))
        Debug.Print VarPtr(.Date) = VarPtr(test)
    End With
    
    With NewVariant(VarPtr(test))
        Debug.Print VarPtr(.Variant) = VarPtr(test)
    End With
    
End Function

'// [Examples] /////////////////////////////////////////////////////////////////////////////////
'// {CallerContext} is public, so you can implement this behavior for any type by matching this template
Public Function NewByte(ByVal This As LongPtr) As IByte
    CallerContext = This
End Function
Public Function NewInteger(ByVal This As LongPtr) As IInteger
    CallerContext = This
End Function
Public Function NewBoolean(ByVal This As LongPtr) As IBoolean
    CallerContext = This
End Function
Public Function NewLong(ByVal This As LongPtr) As ILong
    CallerContext = This
End Function
Public Function NewSingle(ByVal This As LongPtr) As ISingle
    CallerContext = This
End Function
Public Function NewLongPtr(ByVal This As LongPtr) As ILongPtr
    CallerContext = This
End Function
Public Function NewString(ByVal This As LongPtr) As IString
    CallerContext = This
End Function
Public Function NewObject(ByVal This As LongPtr) As IObject
    CallerContext = This
End Function
Public Function NewLongLong(ByVal This As LongPtr) As ILongLong
    CallerContext = This
End Function
Public Function NewDouble(ByVal This As LongPtr) As IDouble
    CallerContext = This
End Function
Public Function NewCurrency(ByVal This As LongPtr) As ICurrency
    CallerContext = This
End Function
Public Function NewDate(ByVal This As LongPtr) As IDate
    CallerContext = This
End Function
Public Function NewVariant(ByVal This As LongPtr) As IVariant
    CallerContext = This
End Function


'///////////////////////////////////////////////////////////////////////////////////////////////
'// [Worker] ///////////////////////////////////////////////////////////////////////////////////
'// 1. Gets the caller's base stack address. ({FrameBase})
'//     - The address of frame-offset 0x0 from the perspective of the calling precedure.
'//     - For a procedure defined in a standard module, frame-offset 0x0 holds a pointer to the module's Module-level variables.
'//     - For a procedure defined in a class module, frame-offset 0x0 holds the object pointer of the executing instance.
'//     - Parameters have a positive frame-offset and local variables have a negative frame-offset.
'//
'// 2. Get the {FOffset} argument of the caller's lblEX_FStI8 instruction
'//     - This frame-offset is where the caller writes the address of the callee's return value, and serves as the With-block's "context".
'//
'// 3. Compute the address of the caller's With-block context
'//     - {FrameBase} + {FrameOffset}
'//
'// 4. Write our own value ({NewContext}) to that address
'//
'// 5. Increment the caller's RSI to effectively jump over its instruction bytes for lblEX_FLdRf; lblEX_FStI8.
'//     - Prevents the caller from overwriting the address we just wrote

Public Property Let CallerContext(Optional ByRef Context As LongPtr, Optional ByRef FrameOffset As Long, ByVal NewContext As LongPtr)
'// [Constants] ////////////////////////////////////////////////////////////////////////////////
  'Const lblEX_ImpAdCallBasicCbFrame As Long = &HFFFF& And &H520
  'Const lblEX_ImpAdCallBasic        As Long = &HFFFF& And &H51F
  'Const lblBEX_LargeBos             As Long = &HFFFF& And &H267

  Const lblEX_FLdRf  As Long = &HFFFF& And &H29F, _
        lblEX_FLdRf2 As Long = lblEX_FLdRf * &H10000
  Const lblEX_FStI8  As Long = &HFFFF& And &H2BB
  
  Const cFLdRf As Long = 6&
  Const cFStI8 As Long = 6&

'// [Offsets] //////////////////////////////////////////////////////////////////////////////////
'// {o0_} - A base address; for defining offsets, symbolic only
'// {o1_} - A first-order offset; relative to some base address {o0_}
'// {o2_} - A second-order offset; relative to some first-order offset {o1_} ' _
     ...
  Const o0Err = NewBase
  Const o0EbThread = NewBase
  Const o0Exframe = NewBase
  Const o0FLdRf = NewBase
  
  Const o1EbThread = o0Err + oLongPtr * 6
  Const o1ExframeTOS = o0EbThread + oLongPtr * 2
  Const o1FrameBase = o0Exframe + oLongPtr * 5
  Const o1CallerRSI = o1FrameBase + oLongPtr * 7, _
        o2CallerRSI = o1CallerRSI - o1FrameBase

  Const o1CbFrame = o0FLdRf - 2
  Const o1FStI8 = o0FLdRf + cFLdRf, _
        o2FStI8 = o1FStI8 - o1CbFrame
  Const o1FStI8_FOffset = o1FStI8 + 2, _
        o2FStI8_FOffset = o1FStI8_FOffset - o1FStI8
  Const o1NextBos = o1FStI8 + cFStI8

'// [Procedure] ////////////////////////////////////////////////////////////////////////////////
    Dim FrameBase As LongPtr
    
    With CallerContextArgs
        Dim This As WithBlock
        ContextOf(This) = VarPtr(NewContext) - oLongPtrx2
        
        '// Walk down the callstack to the caller's Exframe
        .pContext = ObjPtr(Err) + o1EbThread
        .pContext = Context + o1ExframeTOS
        .pContext = Context '// CallerContext         (top of stack/this procedure)
        .pContext = Context '// New<Type>             (the callee)
        .pContext = Context '// <Caller_of_New<Type>> (the caller)
        
        '// Get caller's base stack address
        .pContext = .pContext + o1FrameBase
        FrameBase = Context
        
        '// Point this procedure's {Context} parameter at the caller's instruction pointer
        .pContext = .pContext + o2CallerRSI
        
        '// Point this procedure's {FrameOffset} parameter to where the caller's instruction pointer points
        .pFrameOffset = Context
        
        '// Sanity check the caller's bytecode
        If (FrameOffset And &HFFFF0000) = lblEX_FLdRf2 Then
            .pFrameOffset = .pFrameOffset + o2FStI8
        ElseIf (FrameOffset And &HFFFF&) = lblEX_FLdRf Then
            .pFrameOffset = .pFrameOffset + o1FStI8
        Else
            Exit Property
        End If
        
        If (FrameOffset And &HFFFF&) <> lblEX_FStI8 Then Exit Property
        
        '// Point this procedure's FrameOffset parameter at the FStI8_FOffset value
        .pFrameOffset = .pFrameOffset + o2FStI8_FOffset
        
        '// Advance the caller's instruction pointer
        Context = Context + o1NextBos
        
        '// Write the With-block context
        .pContext = FrameBase + FrameOffset
        Context = NewContext
    End With
End Property

Public Property Let ContextOf(ByRef Block As WithBlock, ByVal Address As LongPtr)
    WriteValueAtOffset Address, Block.Offsets, [+0xLongPtr]
End Property

Private Sub WriteValueAtOffset(ByVal Value As LongPtr, ByRef Offsets() As LONG_PTR, ByVal Offset As Offsets)
    Offsets(Offset) = Value
End Sub

