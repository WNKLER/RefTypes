Attribute VB_Name = "RefTypes"
'================================================================================================================================'
' RefTypes                                                                                                                       '
'--------------------------------------------------------                                                                        '
' https://github.com/WNKLER/RefTypes                                                                                             '
'--------------------------------------------------------                                                                        '
' A VBA/VB6 Library for reading/writing intrinsic types at arbitrary memory addresses.                                           '
' Its defining feature is that this is achieved using truly native, built-in language features.                                  '
' It uses no API declarations and has no external dependencies.                                                                  '
'================================================================================================================================'
' MIT License                                                                                                                    '
'                                                                                                                                '
' Copyright (c) 2025 Benjamin Dovidio (WNKLER)                                                                                   '
'                                                                                                                                '
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated                   '
' documentation files (the "Software"), to deal in the Software without restriction, including without limitation                '
' the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,                   '
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:                         '
'                                                                                                                                '
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. '
'                                                                                                                                '
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO               '
' THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE                 '
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,            '
' TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.       '
'================================================================================================================================'
Option Private Module
Option Explicit

#If Win64 = 1 Then
    Private Const Win64 As Integer = 1
#Else
    Private Const Win64 As Integer = 0
    Public Type LongLong
        L0x0 As Long
        L0x4 As Long
    End Type
#End If

#If VBA7 = 0 Then
    Private Enum LONG_PTR: [_]: End Enum
     Public Enum LongPtr:  [_]: End Enum '// Must be Public for Enum-typed Public Property
#End If

Private Const LPTR_SIZE As Long = 4 + (Win64 * 4)
'*********************************************************************************'
' This section just sets up the array bounds according to the four `UserDefined`  '
' parameters. If you already know the correct bounds, you can just hardcode them. '

Private Const LPROXY_SIZE      As Long = 1         '// UserDefined (should be an integer divisor of <PTR_SIZE>)
Private Const LELEMENT_SIZE    As Long = LPTR_SIZE '// UserDefined (usually should be <PTR_SIZE>)
Private Const LSIZE_PROXIED    As Long = LELEMENT_SIZE - LPROXY_SIZE
Private Const LBLOCK_STEP_SIZE As Long = LELEMENT_SIZE * LSIZE_PROXIED

Private Enum BOUNDS_HELPER:
  [_BLOCK_ALLOCATION_SIZE] = 14 * LPTR_SIZE        '// UserDefined (the size of the region being proxied)
  [_0x0_ADDRESS_ALIGNMENT] = 0                     '// UserDefined (<BLOCK_ALLOCATION_ADDRESS> Mod <PTR_SIZE>)
  [_SIZE_OF_ALIGNED_BLOCK] = ([_0x0_ADDRESS_ALIGNMENT] - LPTR_SIZE) * ([_0x0_ADDRESS_ALIGNMENT] > 0) + [_BLOCK_ALLOCATION_SIZE]
  [_ELEMENT_SIZE_OF_PROXY] = ([_SIZE_OF_ALIGNED_BLOCK] + LBLOCK_STEP_SIZE - 1) \ LBLOCK_STEP_SIZE
  [_ELEMENT_SIZE_OF_BLOCK] = ([_ELEMENT_SIZE_OF_PROXY] * (LELEMENT_SIZE \ LPROXY_SIZE)) - [_ELEMENT_SIZE_OF_PROXY]
End Enum

Private Const LELEMENTS_LBOUND As Long = [_ELEMENT_SIZE_OF_PROXY] * -1
Private Const LELEMENTS_UBOUND As Long = [_ELEMENT_SIZE_OF_BLOCK] + (LELEMENTS_LBOUND < 0)
Private Const LBLOCK_OFFSET    As Long = [_ELEMENT_SIZE_OF_PROXY] * LELEMENT_SIZE '// Relative to VarPtr(<MemoryProxyVariable>)
Private Const LBLOCK_SIZE      As Long = [_ELEMENT_SIZE_OF_BLOCK] * LELEMENT_SIZE '
'*********************************************************************************'

Private Type ProxyElement               '// This could be anything. I have it as an array of bytes only for the
    ProxyAlloc(LPROXY_SIZE - 1) As Byte '// convenience of binding its size to a constant, and for maximum granularity.
End Type                                '// Although, be aware that this structure determines MemoryProxy alignment.

Private Type MemoryProxy                                           '// The declared type of `Elements()` can be any of
    Elements(LELEMENTS_LBOUND To LELEMENTS_UBOUND) As ProxyElement '// the following: Enum, UDT, or Alias (typedef)
End Type                                                           '// NOTE: A ProxyElement's Type must be smaller
                                                                   '// than the size of the Element it represents.
'******************************************************************'
' When passed to `InitByProxy()`, the `Initializer.Elements` array '
' provides access to fourteen, pointer-sized elements immmediately '
' following the `Initializer` variable's memory allocation.        '

Private Initializer   As MemoryProxy
' <Memory proxied by `Initializer`>
Private m_RefInt()    As Integer
Private m_RefLng()    As Long
Private m_RefSng()    As Single
Private m_RefDbl()    As Double
Private m_RefCur()    As Currency
Private m_RefDate()   As Date
Private m_RefStr()    As String
Private m_RefObj()    As Object
Private m_RefBool()   As Boolean
Private m_RefVar()    As Variant
Private m_RefUnk()    As IUnknown
'Private m_RefDec()    As Variant
Private m_RefByte()   As Byte
Private m_RefLngLng() As LongLong
Private m_RefLngPtr() As LongPtr
' <End of proxied memory block>
'*******************************************************************'
                                                                                               
'*************************************************************************************************'
' Inspired by Cristian Buse's `VBA-MemoryTools` <https://github.com/cristianbuse/VBA-MemoryTools> '
' Arbitrary memory access is achieved via a carefully constructed SAFEARRAY `Descriptor` struct.  '

Private m_cDims       As Integer
Private m_fFeatures   As Integer
Private m_cbElements  As Long
Private m_cLocks      As Long
Private m_pvData      As LongPtr
Private m_cElements   As Long
Private m_lLbound     As Long
'*************************************************************************************************'
Private IsInitialized As Boolean

Public Sub Initialize()
    If IsInitialized Then Exit Sub
    
    m_cDims = 1
    m_fFeatures = &H11 'FADF_FIXEDSIZE_AUTO
    m_cbElements = 0   'idk, might help prevent deallocation
    m_cLocks = 1
    m_pvData = 0
    m_cElements = 1
    m_lLbound = 0

    InitByProxy Initializer.Elements
    
    IsInitialized = True
End Sub
'*********************************************************************************'
' This is only possible because the compiler does not (or cannot?) discriminate   '
' between <Non-Intrinsic Array Argument> types passed to <Array Parameters> whose '
' <Declared Type> is an <Enum> or an <Alias> (a non-struct typdef).               '
' Such Array Parameters will accept any <UDT/Enum/Alias>-typed array argument.    '
'                                                                                 '
' Another key behavior is that (except for cDims, pvData, and Bounds) the array   '
' descriptor has no effect on indexing/reading/writing the array elements within  '
' the scope of the receiving procedure; indexing/reading/writing align with the   '
' declared type of the Array Parameter. (this behavior is not critical, but it    '
' greatly simplifies the implementation) NOTE: You cannot pass an element ByRef   '
' from inside the procedure. Doing so passes the address of its proxy.            '
'                                                                                 '
' Similarly, Array Parameters whose <Declared Type> is <Fixed-Length-String> will '
' accept ANY <Fixed-Length-String> array argument, regardless of Declared Length. '
' However, since Fixed-Length-Strings have no alignment, the starting position of '
' an element and the starting position of its proxy will always be the same.      '
'*********************************************************************************'
Private Sub InitByProxy(ByRef ProxyElements() As LONG_PTR)
    Dim pcDims As LongPtr
    pcDims = VarPtr(m_cDims)
    
    Dim i As Long
    
    For i = 0& To 13&
        ProxyElements(i) = pcDims
    Next i
End Sub

Public Property Get RefInt(ByVal Target As LongPtr) As Integer
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefInt = m_RefInt(0&)
End Property
Public Property Let RefInt(ByVal Target As LongPtr, ByVal RefInt As Integer)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefInt(0&) = RefInt
End Property

Public Property Get RefLng(ByVal Target As LongPtr) As Long
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefLng = m_RefLng(0&)
End Property
Public Property Let RefLng(ByVal Target As LongPtr, ByVal RefLng As Long)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefLng(0&) = RefLng
End Property

Public Property Get RefSng(ByVal Target As LongPtr) As Single
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefSng = m_RefSng(0&)
End Property
Public Property Let RefSng(ByVal Target As LongPtr, ByVal RefSng As Single)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefSng(0&) = RefSng
End Property

Public Property Get RefDbl(ByVal Target As LongPtr) As Double
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefDbl = m_RefDbl(0&)
End Property
Public Property Let RefDbl(ByVal Target As LongPtr, ByVal RefDbl As Double)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefDbl(0&) = RefDbl
End Property

Public Property Get RefCur(ByVal Target As LongPtr) As Currency
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefCur = m_RefCur(0&)
End Property
Public Property Let RefCur(ByVal Target As LongPtr, ByVal RefCur As Currency)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefCur(0&) = RefCur
End Property

Public Property Get RefDate(ByVal Target As LongPtr) As Date
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefDate = m_RefDate(0&)
End Property
Public Property Let RefDate(ByVal Target As LongPtr, ByVal RefDate As Date)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefDate(0&) = RefDate
End Property

Public Property Get RefStr(ByVal Target As LongPtr) As String
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefStr = m_RefStr(0&)
End Property
Public Property Let RefStr(ByVal Target As LongPtr, ByRef RefStr As String)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefStr(0&) = RefStr
End Property

Public Property Get RefObj(ByVal Target As LongPtr) As Object
    If IsInitialized Then Else Initialize
    m_pvData = Target
    Set RefObj = m_RefObj(0&)
End Property
Public Property Set RefObj(ByVal Target As LongPtr, ByVal RefObj As Object)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    Set m_RefObj(0&) = RefObj
End Property

Public Property Get RefBool(ByVal Target As LongPtr) As Boolean
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefBool = m_RefBool(0&)
End Property
Public Property Let RefBool(ByVal Target As LongPtr, ByVal RefBool As Boolean)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefBool(0&) = RefBool
End Property

Public Property Get RefVar(ByVal Target As LongPtr) As Variant
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefVar = m_RefVar(0&)
End Property
Public Property Let RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefVar(0&) = RefVar
End Property
Public Property Set RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    Set m_RefVar(0&) = RefVar
End Property

Public Property Get RefUnk(ByVal Target As LongPtr) As IUnknown
    If IsInitialized Then Else Initialize
    m_pvData = Target
    Set RefUnk = m_RefUnk(0&)
End Property
Public Property Set RefUnk(ByVal Target As LongPtr, ByVal RefUnk As IUnknown)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    Set m_RefUnk(0&) = RefUnk
End Property

'Public Property Get RefDec(ByVal Target As LongPtr) As Variant
'    If IsInitialized Then Else Initialize
'    m_pvData = Target
'    RefDec = m_RefDec(0&)
'End Property
'Public Property Let RefDec(ByVal Target As LongPtr, ByVal RefDec As Variant)
'    If IsInitialized Then Else Initialize
'    m_pvData = Target
'    m_RefDec(0&) = RefDec
'End Property _

Public Property Get RefByte(ByVal Target As LongPtr) As Byte
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefByte = m_RefByte(0&)
End Property
Public Property Let RefByte(ByVal Target As LongPtr, ByVal RefByte As Byte)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefByte(0&) = RefByte
End Property

    Public Property Get RefLngLng(ByVal Target As LongPtr) As LongLong
        If IsInitialized Then Else Initialize
        m_pvData = Target
        RefLngLng = m_RefLngLng(0&)
    End Property
#If Win64 = 0 Then
    Public Property Let RefLngLng(ByVal Target As LongPtr, ByRef RefLngLng As LongLong)
#Else
    Public Property Let RefLngLng(ByVal Target As LongPtr, ByVal RefLngLng As LongLong)
#End If
        If IsInitialized Then Else Initialize
        m_pvData = Target
        m_RefLngLng(0&) = RefLngLng
    End Property

Public Property Get RefLngPtr(ByVal Target As LongPtr) As LongPtr
    If IsInitialized Then Else Initialize
    m_pvData = Target
    RefLngPtr = m_RefLngPtr(0&)
End Property
Public Property Let RefLngPtr(ByVal Target As LongPtr, ByVal RefLngPtr As LongPtr)
    If IsInitialized Then Else Initialize
    m_pvData = Target
    m_RefLngPtr(0&) = RefLngPtr
End Property

