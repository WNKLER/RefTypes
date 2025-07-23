Attribute VB_Name = "RefTypes"
Option Private Module
Option Explicit

Private Enum Context
#If Win64 Then
    [_]   '// 0 on x64; undefined on x86.
#End If
    Win64 '// 0 when [_] is undefined; otherwise, [_] + 1.
    PtrSize = 4 + (4 * Win64)
    VarSize = 8 + (2 * PtrSize)
End Enum

#If VBA7 = 0 Then
    Private Enum LONG_PTR: [_]: End Enum
     Public Enum LongPtr:  [_]: End Enum '// Must be Public for Enum-typed Public Property
#End If

Private Type HalfPtr
    HalfPtr(-Win64 To 0) As Integer
End Type

Private Type Initializer
    Initializer(-1 To 0) As HalfPtr
End Type

Private Type Descriptor
    cDims         As Integer
    fFeatures     As Integer
    cbElements    As Long
    IsInitialized As Boolean
    pvData        As LongPtr
    cElements     As Long
    lLbound       As Long
End Type

'Private Type Vector          '// Declare Static Fields individually. This avoids defining
'    Element()  As Any        '// a separate Vector UDT for each `Element()` Type.
'    Descriptor As Descriptor
'End Type _

Private Sub InitInitializer(ByRef Initializer() As LONG_PTR)
    Const First As Long = -1
    Const Last  As Long = -0
    
    Initializer(Last) = VarPtr(Initializer(First)) + (2 * PtrSize)
End Sub

Private Sub Init(ByRef Descriptor As Descriptor)
    Const FADF_AUTO      As Integer = &H1
    Const FADF_FIXEDSIZE As Integer = &H10
    
    Static This            As Initializer '// Proxy for `Init_Element()`
    Static Init_Element()  As LongPtr     '// Static Init As Vector
    Static Init_Descriptor As Descriptor
    
    With Init_Descriptor
        If .IsInitialized = False Then
            InitInitializer This.Initializer '// Point `Init_Element()` to `Init_Descriptor`
            .lLbound = 0
            .cElements = 1
            .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
            .cDims = 1
            .IsInitialized = True
        End If
        
        .pvData = VarPtr(Descriptor) - PtrSize
        Init_Element(0) = .pvData + PtrSize
    End With
    
    With Descriptor
        .lLbound = 0
        .cElements = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .cDims = 1
        .IsInitialized = True
    End With
End Sub

Public Property Get RefInt(ByVal Target As LongPtr) As Integer
    Static Vector_Element() As Integer, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefInt = Vector_Element(0&)
End Property
Public Property Let RefInt(ByVal Target As LongPtr, ByVal RefInt As Integer)
    Static Vector_Element() As Integer, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefInt
End Property

Public Property Get RefLng(ByVal Target As LongPtr) As Long
    Static Vector_Element() As Long, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefLng = Vector_Element(0&)
End Property
Public Property Let RefLng(ByVal Target As LongPtr, ByVal RefLng As Long)
    Static Vector_Element() As Long, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefLng
End Property

Public Property Get RefSng(ByVal Target As LongPtr) As Single
    Static Vector_Element() As Single, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefSng = Vector_Element(0&)
End Property
Public Property Let RefSng(ByVal Target As LongPtr, ByVal RefSng As Single)
    Static Vector_Element() As Single, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefSng
End Property

Public Property Get RefDbl(ByVal Target As LongPtr) As Double
    Static Vector_Element() As Double, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefDbl = Vector_Element(0&)
End Property
Public Property Let RefDbl(ByVal Target As LongPtr, ByVal RefDbl As Double)
    Static Vector_Element() As Double, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefDbl
End Property

Public Property Get RefCur(ByVal Target As LongPtr) As Currency
    Static Vector_Element() As Currency, Vector_Descriptor As Descriptor    '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefCur = Vector_Element(0&)
End Property
Public Property Let RefCur(ByVal Target As LongPtr, ByVal RefCur As Currency)
    Static Vector_Element() As Currency, Vector_Descriptor As Descriptor    '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefCur
End Property

Public Property Get RefDate(ByVal Target As LongPtr) As Date
    Static Vector_Element() As Date, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefDate = Vector_Element(0&)
End Property
Public Property Let RefDate(ByVal Target As LongPtr, ByVal RefDate As Date)
    Static Vector_Element() As Date, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefDate
End Property

Public Property Get RefStr(ByVal Target As LongPtr) As String
    Static Vector_Element() As String, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    RefStr = Vector_Element(0&)
End Property
Public Property Let RefStr(ByVal Target As LongPtr, ByRef RefStr As String)
    Static Vector_Element() As String, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefStr
End Property

Public Property Get RefObj(ByVal Target As LongPtr) As Object
    Static Vector_Element() As Object, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    Set RefObj = Vector_Element(0&)
End Property
Public Property Set RefObj(ByVal Target As LongPtr, ByVal RefObj As Object)
    Static Vector_Element() As Object, Vector_Descriptor As Descriptor      '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    Set Vector_Element(0&) = RefObj
End Property

Public Property Get RefBool(ByVal Target As LongPtr) As Boolean
    Static Vector_Element() As Boolean, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefBool = Vector_Element(0&)
End Property
Public Property Let RefBool(ByVal Target As LongPtr, ByVal RefBool As Boolean)
    Static Vector_Element() As Boolean, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefBool
End Property

Public Property Get RefVar(ByVal Target As LongPtr) As Variant
    Static Vector_Element() As Variant, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor: Vector_Descriptor.cbElements = VarSize
    Vector_Descriptor.pvData = Target
    If TypeOf Vector_Element(0&) Is IUnknown Then
        Set RefVar = Vector_Element(0&)
    Else
        RefVar = Vector_Element(0&)
    End If
End Property
Public Property Let RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    Static Vector_Element() As Variant, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor: Vector_Descriptor.cbElements = VarSize
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefVar
End Property
Public Property Set RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    Static Vector_Element() As Variant, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor: Vector_Descriptor.cbElements = VarSize
    Vector_Descriptor.pvData = Target
    Set Vector_Element(0&) = RefVar
End Property

Public Property Get RefUnk(ByVal Target As LongPtr) As IUnknown
    Static Vector_Element() As IUnknown, Vector_Descriptor As Descriptor    '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    Set RefUnk = Vector_Element(0&)
End Property
Public Property Set RefUnk(ByVal Target As LongPtr, ByVal RefUnk As IUnknown)
    Static Vector_Element() As IUnknown, Vector_Descriptor As Descriptor    '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = PtrSize
    Vector_Descriptor.pvData = Target
    Set Vector_Element(0&) = RefUnk
End Property

Public Property Get RefByte(ByVal Target As LongPtr) As Byte
    Static Vector_Element() As Byte, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefByte = Vector_Element(0&)
End Property
Public Property Let RefByte(ByVal Target As LongPtr, ByVal RefByte As Byte)
    Static Vector_Element() As Byte, Vector_Descriptor As Descriptor        '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefByte
End Property

#If Win64 = 1 Then
    Public Property Get RefLngLng(ByVal Target As LongPtr) As LongLong
        Static Vector_Element() As LongLong, Vector_Descriptor As Descriptor    '// Static Vector As Vector
        If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
        Vector_Descriptor.pvData = Target
        RefLngLng = Vector_Element(0&)
    End Property
    Public Property Let RefLngLng(ByVal Target As LongPtr, ByVal RefLngLng As LongLong)
        Static Vector_Element() As LongLong, Vector_Descriptor As Descriptor    '// Static Vector As Vector
        If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
        Vector_Descriptor.pvData = Target
        Vector_Element(0&) = RefLngLng
    End Property
#End If

Public Property Get RefLngPtr(ByVal Target As LongPtr) As LongPtr
    Static Vector_Element() As LongPtr, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    RefLngPtr = Vector_Element(0&)
End Property
Public Property Let RefLngPtr(ByVal Target As LongPtr, ByVal RefLngPtr As LongPtr)
    Static Vector_Element() As LongPtr, Vector_Descriptor As Descriptor     '// Static Vector As Vector
    If Vector_Descriptor.IsInitialized Then Else Init Vector_Descriptor     ': Vector_Descriptor.cbElements = LenB(Vector_Element(0&))
    Vector_Descriptor.pvData = Target
    Vector_Element(0&) = RefLngPtr
End Property

