Attribute VB_Name = "WithTypes"
'MIT License
'https://github.com/WNKLER/RefTypes
Option Private Module
Option Explicit

 #If VBA7 <> 1 Then
    Private Enum LONG_PTR: [_LONG_PTR]: End Enum
     Private Enum LongPtr:  [_LongPtr]: End Enum '// Must be Private for Enum-typed Private Property
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
Private Const oLongPtrx2 = oLongPtr + oLongPtr

Private Type HalfPtr
    Bytes As String * wHalfPtr
End Type

Private Enum Offsets
    [+0x0] = -1&
    [+0xLongPtr]
End Enum

Private Type WithBlock
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

Private Function Test() As LongPtr
    With NewByte(VarPtr(Test))
        Debug.Print VarPtr(.Byte) = VarPtr(Test)
    End With
    
    With NewInteger(VarPtr(Test))
        Debug.Print VarPtr(.Integer) = VarPtr(Test)
    End With
    
    With NewBoolean(VarPtr(Test))
        Debug.Print VarPtr(.Boolean) = VarPtr(Test)
    End With
    
    With NewLong(VarPtr(Test))
        Debug.Print VarPtr(.Long) = VarPtr(Test)
    End With
    
    With NewSingle(VarPtr(Test))
        Debug.Print VarPtr(.Single) = VarPtr(Test)
    End With
    
    With NewLongPtr(VarPtr(Test))
        Debug.Print VarPtr(.LongPtr) = VarPtr(Test)
    End With
    
    With NewString(VarPtr(Test))
        Debug.Print VarPtr(.String) = VarPtr(Test)
    End With
    
    With NewObject(VarPtr(Test))
        Debug.Print VarPtr(.Object) = VarPtr(Test)
    End With
    
    With NewLongLong(VarPtr(Test))
        Debug.Print VarPtr(.LongLong) = VarPtr(Test)
    End With
    
    With NewDouble(VarPtr(Test))
        Debug.Print VarPtr(.Double) = VarPtr(Test)
    End With
    
    With NewCurrency(VarPtr(Test))
        Debug.Print VarPtr(.Currency) = VarPtr(Test)
    End With
    
    With NewDate(VarPtr(Test))
        Debug.Print VarPtr(.Date) = VarPtr(Test)
    End With
    
    With NewVariant(VarPtr(Test))
        Debug.Print VarPtr(.Variant) = VarPtr(Test)
    End With
    
End Function

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

Private Property Let CallerContext(Optional ByRef Context As LongPtr, Optional ByRef FrameOffset As Long, ByVal NewContext As LongPtr)
'// 1. Gets the caller's base stack address. (FrameBase)
'//     - The address of FrameOffset 0x0 from the perspective of the calling precedure.
'//     - For a procedure defined in a standard module, FrameOffset 0x0 holds a pointer to the module's Module-level variables.
'//     - For a procedure defined in a class module, FrameOffset 0x0 holds the object pointer of the executing instance.
'//     - Parameters have a positive FrameOffset and local variables have a negative FrameOffset.
'//
'// 2. Get the FrameOffset argument of the caller's lblEX_FStI8 instruction
'//     - This FrameOffset is where the caller writes the address of the callee's return value, and serves as the With-block's "context".
'//
'// 3. Compute the address of the caller's With-block context
'//     - FrameBase + FrameOffset
'//
'// 4. Write our own value (NewContext) to that address
'//
'// 5. Increment the caller's RSI to effectively jump over [lblEX_FLdRf; lblEX_FStI8]
'//     - Prevents the caller from overwriting the address we just wrote
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'  Const lblEX_ImpAdCallBasicCbFrame As Long = &HFFFF& And &H520
'  Const lblEX_ImpAdCallBasic        As Long = &HFFFF& And &H51F
  Const lblEX_FLdRf                 As Long = &HFFFF& And &H29F
  Const lblEX_FStI8                 As Long = &HFFFF& And &H2BB
'  Const lblBEX_LargeBos             As Long = &HFFFF& And &H267
  
  Const oEbThread = oLongPtr * 6
  Const oExframeTOS = oLongPtrx2
  Const oFrameBase = oLongPtr * 5
  Const oCallerRSI = NullPtr - (oLongPtr * 32)
  
  Const oCbFrame = NullPtr + 2
  Const oFLdRf = NullPtr + 0
  Const oFStI8 = NullPtr + 6
  Const oFOffset = NullPtr + 2
  Const oJumpSize = NullPtr + 12
  
    Dim FrameBase As LongPtr
    
    With CallerContextArgs
        Dim This As WithBlock
        ContextOf(This) = VarPtr(NewContext) - oLongPtrx2
        
        .pContext = ObjPtr(Err) + oEbThread
        .pContext = Context + oExframeTOS
        .pContext = Context '// CallerContext?
        .pContext = Context '// New<Type>?
        .pContext = Context '// <Caller_of_New<Type>>?
        
        .pContext = .pContext + oFrameBase
        FrameBase = Context
        
        .pContext = .pContext - oFrameBase
        .pContext = Context + oCallerRSI
        
        .pFrameOffset = Context
        
        If (FrameOffset And &HFFFF&) <> lblEX_FLdRf Then
            .pFrameOffset = .pFrameOffset + oCbFrame
            If (FrameOffset And &HFFFF&) <> lblEX_FLdRf Then Exit Property
        End If
        
        .pFrameOffset = .pFrameOffset + oFStI8
        If (FrameOffset And &HFFFF&) = lblEX_FStI8 Then Else Exit Property
            
        .pFrameOffset = .pFrameOffset + oFOffset
        Context = Context + oJumpSize
        
        .pContext = FrameBase + FrameOffset
        Context = NewContext
    End With
End Property

Private Property Let ContextOf(ByRef Block As WithBlock, ByVal Address As LongPtr)
    WriteValueAtOffset Address, Block.Offsets, [+0xLongPtr]
End Property

Private Sub WriteValueAtOffset(ByVal Value As LongPtr, ByRef Offsets() As LONG_PTR, ByVal Offset As Offsets)
    Offsets(Offset) = Value
End Sub

