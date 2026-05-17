Attribute VB_Name = "EnumerateDispatch"
'@IgnoreModule MultipleDeclarations, HungarianNotation, UseMeaningfulName, AssignedByValParameter, FunctionReturnValueDiscarded, UnassignedVariableUsage, VariableNotAssigned, IntegerDataType, UDTMemberNotUsed
'@Folder("Module")
'@ModuleDescription("Enumerator Module using the Dispatch Interface.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Enumerator Module using Dispatch Interface
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.05.16
'
' Methods
' Enumerate(iterable, callback, count [, base])  Sets IEnumVARIANT ...
'
' Enumerator works correctly for nested loops with mixed objects as well as for
' nested loops with mixed enumerators.
'
' Code to be included in the iterable object:
'
' Public Function Enumerate() As IEnumVARIANT
'    Set Enumerate = EnumerateDispatch.Enumerate(Me, ...)
'
' Timings (ms) for n = 10.000
' Iterable object with IDispatch interface (For Each)      27.90 (API 26.66  ms)
' VB Collection - VB enumerator (For Each)                  0.21
' VB Array - VB enumerator (For Each)                       0.14
' VB Array - VB loop (For)                                  0.08
'
' The original ideas for a custom enumerator using a typelib and redefining
' the IEnumVARIANT interface routines in a standard module originate from
' Hardcore Visual Basic 5.0 by Bruce McKinney.
'
' The implementation without using a typelib is based on work by Dexter
' Freivald who's original code was for 32 bits and was using late binding.
' https://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib
'
' An alternative method is to define an enumeration procedure in the iterable
' object by copying its items to an embedded VB Collection and subsequently
' exposing the IEnumVARIANT interface of this VB Collection. Another alternative
' is to let the object export the items to an array. The latter is what is used
' by the VB Dictionary. Both of these alternative these methods export the items
' "at once" whereas the Enumerate exports the enumerated items "one by one".

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler Directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Const API = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

' https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cotaskmemalloc
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" ( _
    ByVal cb As Long _
) As LongPtr

' https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cotaskmemfree
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" ( _
    ByVal pv As LongPtr _
)
' https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-dispcallfunc
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvResult As Variant _
) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The vbNullPtr constant is the null pointer (not VB defined).
Private Const vbNullPtr As LongPtr = 0

' The memory size of intrinsic data types.
Private Enum VARSIZE
    vbSizeInteger = 2
    vbSizeLong = 4
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
End Enum

' Selected HRESULT constants.
Private Enum HRESULT
    S_OK = &H0                          ' Operation successful, returns True
    S_FALSE = &H1                       ' Operation successful, returns False
    E_NOTIMPL = &H80004001              ' Not implemented
    E_NOINTERFACE = &H80004002          ' No such interface supported
    E_POINTER = &H80004003              ' Pointer that is not valid
    E_ABORT = &H80004004                ' Operation aborted
    E_FAIL = &H80004005                 ' Unspecified failure
    E_OUTOFMEMORY = &H8007000E          ' Failed to allocate necessary memory
    E_INVALIDARG = &H80070057           ' One of the arguments is not valid
End Enum

' Selectes VARTYPE constants
Private Const VT_I4      As Integer = 3

Private Const CC_STDCALL As Integer = 4

Private Const DISPATCH_METHOD       As Integer = &H1
Private Const DISPATCH_PROPERTYGET  As Integer = &H2

' Selected VBA errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorOutOfMemory = 7
    vbErrorObjectRequired = 424
End Enum

' GUID is the UDT for the Global Unique Identifier.
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type DISPPARAMS
    rgvarg As LongPtr
    rgdispidNamedArgs As LongPtr
    cArgs As Long
    cNamedArgs As Long
End Type

' The IEnumVARIANT status is captured in an UDT.
Private Type TENUM
    pvTable As LongPtr
    caller As Object
    dispid  As Long
    nRef As Long
    First As Long
    Last As Long
    Current As Long
End Type

#If API = False Then
' Variant ByRef construct for memory access by address.
Private Const VT_BYREF As Integer = &H4000
Private Type CONSTRUCT
    vt As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Ignore NonReturningFunction
Public Function Enumerate( _
    ByVal iterable As Object, _
    ByVal callback As String, _
    ByVal count As Long, _
    Optional ByVal base As Long = 1 _
) As IEnumVARIANT

    If iterable Is Nothing Then Err.Raise vbErrorObjectRequired

    ' Initialize the vTable with the redefined IUnknown/IEnumVARIANT functions.
    Static vTable(0 To 6) As LongPtr
    If vTable(0) = vbNullPtr Then
        vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
        vTable(1) = VBA.CLngPtr(AddressOf IUnknown_AddRef)
        vTable(2) = VBA.CLngPtr(AddressOf IUnknown_Release)
        vTable(3) = VBA.CLngPtr(AddressOf IEnumVARIANT_Next)
        vTable(4) = VBA.CLngPtr(AddressOf IEnumVARIANT_Skip)
        vTable(5) = VBA.CLngPtr(AddressOf IEnumVARIANT_Reset)
        vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
#If API = False Then
        ' Initialize the Variant ByRef construct.
        InitializeVarByRef
#End If
    End If

    ' Construct the synthetic IEnumVARIANT object.
    Dim obj As TENUM
    With obj
        .pvTable = VarPtr(vTable(0))
        ' Test whether the callback function exists.
        .dispid = Resolve(iterable, callback)
        Set .caller = iterable
        .First = base
        .Last = base + count - 1
        .nRef = 1
        .Current = .First
    End With

    Dim MemoryBlock As LongPtr: MemoryBlock = CoTaskMemAlloc(LenB(obj))
    If MemoryBlock = vbNullPtr Then
        Err.Raise vbErrorOutOfMemory
    End If
    CopyMemory ByVal MemoryBlock, obj, LenB(obj)
    '@Ignore ValueRequired
    CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr

    ' When obj goes out of scope nref for iterable is decreased.
    ' KeepAlive compensates by increasing the nref for iterable.
    Set KeepAlive(MemoryBlock) = iterable

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Queries a COM object for a pointer to one of its interfaces."
Private Function IUnknown_QueryInterface( _
    ByRef obj As TENUM, _
    ByRef riid As GUID, _
    ByVal ppvObj As LongPtr _
) As Long

    If ppvObj = vbNullPtr Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If

    If IsIID_IUnknown(riid) Or IsIID_IEnumVARIANT(riid) Then
        CopyMemory ByVal ppvObj, VarPtr(obj), vbSizeLongPtr
        IUnknown_AddRef obj
        IUnknown_QueryInterface = S_OK
    Else
        IUnknown_QueryInterface = E_NOINTERFACE
    End If

End Function


'@Description "Increments the reference count for an interface pointer to a COM object."
Private Function IUnknown_AddRef(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef + 1
    IUnknown_AddRef = obj.nRef

End Function


'@Description "Decrements the reference count for an interface pointer to a COM object."
Private Function IUnknown_Release(ByRef obj As TENUM) As Long

    obj.nRef = obj.nRef - 1
    IUnknown_Release = obj.nRef

    ' Free the memory block used for the synthetic IEnumVARIANT object.
    If obj.nRef = 0 Then
        Set KeepAlive(VarPtr(obj)) = Nothing
        CoTaskMemFree VarPtr(obj)
    End If

End Function


'@Description "Retrieves the next item in the enumeration sequence."
Private Function IEnumVARIANT_Next( _
    ByRef obj As TENUM, _
    ByVal celt As Long, _
    ByVal rgVar As LongPtr, _
    ByVal pceltFetched As LongPtr _
) As Long

    If rgVar = vbNullPtr Then
        IEnumVARIANT_Next = E_INVALIDARG
        Exit Function
    End If

    ' Set pceltFetched to 0 if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, 0, vbSizeLong
#Else
        CopyLngByRef pceltFetched, 0, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Only continue with loop for celt > 0.
    If celt <= 0 Then
        If celt < 0 Then
            IEnumVARIANT_Next = E_INVALIDARG
        Else
            IEnumVARIANT_Next = S_OK
        End If
        Exit Function
    End If

    ' Get the next item(s) from the iterable object.
    Dim i As Long, NumberFetched As Long
    For i = obj.Current To obj.Last
        If Invoke(obj.caller, obj.dispid, i, rgVar) < 0 Then
            IEnumVARIANT_Next = E_FAIL
            Exit Function
        End If
        NumberFetched = NumberFetched + 1
        If NumberFetched = celt Then Exit For
        ' Advance the pointer to the next element in the destination array.
        rgVar = rgVar + vbSizeVariant
    Next
    obj.Current = obj.Current + NumberFetched

    ' Set pceltFetched to NumberFetched if the pointer is provided.
    If pceltFetched <> vbNullPtr Then
#If API Then
        CopyMemory ByVal pceltFetched, NumberFetched, vbSizeLong
#Else
        CopyLngByRef pceltFetched, NumberFetched, VarByRef.vt, VarByRef.ref
#End If
    End If

    ' Return S_OK if the number of fetched items matches the requested number.
    If NumberFetched = celt Then
        IEnumVARIANT_Next = S_OK
    Else
        IEnumVARIANT_Next = S_FALSE
    End If

End Function


'@Description "Skips over a number of elements in the enumeration sequence."
Private Function IEnumVARIANT_Skip(ByRef obj As TENUM, ByVal celt As Long) As Long

    Select Case True
    Case celt = 0
        IEnumVARIANT_Skip = S_OK
    Case celt < 0
        IEnumVARIANT_Skip = E_INVALIDARG
    Case celt <= obj.Last - obj.Current + 1
        obj.Current = obj.Current + celt
        IEnumVARIANT_Skip = S_OK
    Case Else
        ' For overshoot set one-past-end.
        obj.Current = obj.Last + 1
        IEnumVARIANT_Skip = S_FALSE
    End Select

End Function


'@Description "Resets the enumeration sequence to the beginning."
Private Function IEnumVARIANT_Reset(ByRef obj As TENUM) As Long

    obj.Current = obj.First
    IEnumVARIANT_Reset = S_OK

End Function


'@Description "Creates a copy of the current state of enumeration."
Private Function IEnumVARIANT_Clone(ByRef obj As TENUM, ByVal ppEnum As LongPtr) As Long

    If ppEnum = vbNullPtr Then
        IEnumVARIANT_Clone = E_INVALIDARG
        Exit Function
    End If

    ' UDT assignment AddRefs caller — dispid is Long, no extra management needed.
    Dim Copy As TENUM: Copy = obj
    Copy.nRef = 1

    Dim MemoryBlock As LongPtr: MemoryBlock = CoTaskMemAlloc(LenB(obj))
    If MemoryBlock = vbNullPtr Then
        IEnumVARIANT_Clone = E_OUTOFMEMORY
        Exit Function
    End If
    CopyMemory ByVal MemoryBlock, Copy, LenB(obj)
    CopyMemory ByVal ppEnum, MemoryBlock, vbSizeLongPtr
    IEnumVARIANT_Clone = S_OK

    ' When Copy goes out of scope nref for iterable is decreased.
    ' KeepAlive compensates by increasing the nref for iterable.
    Set KeepAlive(MemoryBlock) = Copy.caller

End Function


'@Description "Returns True of id is IID_IUnknown GUID."
Private Function IsIID_IUnknown(ByRef id As GUID) As Boolean

'    Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
    IsIID_IUnknown = _
        (id.Data1 = &H0) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


'@Description "Returns True if id is IID_IEnumVARIANT GUID."
Private Function IsIID_IEnumVARIANT(ByRef id As GUID) As Boolean

'    Const IID_IEnumVARIANT As String = "{00020404-0000-0000-C000-000000000046}"
    IsIID_IEnumVARIANT = _
        (id.Data1 = &H20404) And _
        (id.Data2 = &H0) And _
        (id.Data3 = &H0) And _
        (id.Data4(0) = &HC0) And _
        (id.Data4(1) = &H0) And _
        (id.Data4(2) = &H0) And _
        (id.Data4(3) = &H0) And _
        (id.Data4(4) = &H0) And _
        (id.Data4(5) = &H0) And _
        (id.Data4(6) = &H0) And _
        (id.Data4(7) = &H46)

End Function


'@Description "Keep alive the iterable object stored in heap memory."
Private Property Set KeepAlive(ByVal block As LongPtr, ByVal RHS As Object)
' Increase or decrease reference count for the iterable object.

    Static Table As Collection
    If Table Is Nothing Then Set Table = New Collection

    ' The key is the address of the allocated memory block.
    Dim Key As String: Key = CStr(block)

    If RHS Is Nothing Then
        On Error Resume Next
        ' Ignore if not found.
        Table.Remove Key
        On Error GoTo 0
    Else
        On Error Resume Next
        ' Replace if present.
        Table.Remove Key
        Err.Clear
        Table.Add RHS, Key
        On Error GoTo 0
    End If

End Property


'@Description "IDispatch:GetIDsOfNames for specified object and name."
Private Function Resolve( _
    ByVal obj As Object, _
    ByVal member As String _
) As Long

    ' IDispatch::GetIDsOfNames — vtable slot 5
    ' HRESULT GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*)

    Const slot As Long = 5
    Const oVft As LongPtr = slot * vbSizeLongPtr
    Const cvt As Long = 5

    ' Each pv(i) must point to a VARIANT.
    ' Store every argument value in a Variant.
    ' VarPtr of that Variant goes in pv().
    ' VarPtr returns LongPtr -> VT_I4 on 32-bit, VT_I8 on 64-bit automatically,
    ' VarType() gives the right type for pointer arguments on both bitnesses.
    ' Do NOT use intermediate LongPtr variables and then VarPtr those — that hands
    ' DispCallFunc a pointer to a raw LongPtr, not to a VARIANT, so it reads the
    ' low 2 bytes of the pointer value as the vt field (garbage).

    Dim iid As GUID                 ' IID_NULL — zero-initialised by VBA.

    Static Init As Boolean
    If Init = False Then
        Static Names(0) As LongPtr
        Names(0) = StrPtr(member)   ' LPOLESTR array.
        Static cNames As Long: cNames = 1
        Static lcid As Long: lcid = 0
        Static dispid As Long

        Static var(0 To 4) As Variant
        var(0) = VarPtr(iid)        ' REFIID    — pointer to IID_NULL
        var(1) = VarPtr(Names(0))   ' LPOLESTR* — pointer to name array
        var(2) = cNames             ' UINT      — VT_I4, value = 1
        var(3) = lcid               ' LCID      — VT_I4, value = 0
        var(4) = VarPtr(dispid)     ' DISPID*   — pointer to output Long

        Static vt(0 To 4) As Integer
        vt(0) = VarType(var(0))     ' VT_I4 / VT_I8
        vt(1) = VarType(var(1))     ' VT_I4 / VT_I8
        vt(2) = VarType(var(2))     ' VT_I4
        vt(3) = VarType(var(3))     ' VT_I4
        vt(4) = VarType(var(4))     ' VT_I4 / VT_I8

        Static pv(0 To 4) As LongPtr
        pv(0) = VarPtr(var(0))
        pv(1) = VarPtr(var(1))
        pv(2) = VarPtr(var(2))
        pv(3) = VarPtr(var(3))
        pv(4) = VarPtr(var(4))

        Init = True
    End If

    ' Adjust for dynamic arguments.
    Names(0) = StrPtr(member)
    var(0) = VarPtr(iid)

    ' Call DispCallFunc API.
    Dim hr As Long
    Dim dummy As Variant
    hr = DispCallFunc(ObjPtr(obj), oVft, CC_STDCALL, VT_I4, cvt, vt(0), pv(0), dummy)

    If hr < 0 Then
        Err.Raise vbObjectError Or &H5001, "IDispatch", "GetIDsOfNames failed: 0x" & Hex$(hr)
        Exit Function
    End If

    Resolve = dispid

End Function


'@Description "IDispatch:Invoke for specified object, dispid and index."
Private Function Invoke( _
    ByVal obj As Object, _
    ByVal dispid As Long, _
    ByVal index As Long, _
    ByVal rgVar As LongPtr _
) As Long

    ' IDispatch::Invoke — vtable slot 6
    ' HRESULT Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*)

    Const slot As Long = 6
    Const oVft As LongPtr = slot * vbSizeLongPtr
    Const cvt As Long = 8

    Dim iid As GUID ' IID_NULL — zero-initialised by VBA.

    Static Init As Boolean
    If Init = False Then
        Static arg As Variant: arg = CLng(index)
        Static dp As DISPPARAMS
        With dp
            .cArgs = 1
            .cNamedArgs = 0
            .rgdispidNamedArgs = 0
            .rgvarg = VarPtr(arg) ' VBA Variant is a VARIANTARG
        End With
        Static lcid As Long: lcid = 0
        Static wFlags As Long: wFlags = DISPATCH_PROPERTYGET ' Long -> VT_I4, avoids VT_I2

        Static var(0 To 7) As Variant
        var(0) = dispid             ' DISPID      — VT_I4
        var(1) = VarPtr(iid)        ' REFIID      — pointer to IID_NULL
        var(2) = lcid               ' LCID        — VT_I4
        var(3) = wFlags             ' WORD flags  — VT_I4
        var(4) = VarPtr(dp)         ' DISPPARAMS* — pointer to dp
        var(5) = rgVar              ' VARIANT*    — pointer to result
        var(6) = vbNullPtr          ' EXCEPINFO*  — NULL
        var(7) = vbNullPtr          ' UINT*       — NULL

        Static vt(0 To 7) As Integer
        vt(0) = VarType(var(0))     ' VT_I4
        vt(1) = VarType(var(1))     ' VT_I4 / VT_I8
        vt(2) = VarType(var(2))     ' VT_I4
        vt(3) = VarType(var(3))     ' VT_I4
        vt(4) = VarType(var(4))     ' VT_I4 / VT_I8
        vt(5) = VarType(var(5))     ' VT_I4 / VT_I8
        vt(6) = VarType(var(6))     ' VT_I4 / VT_I8
        vt(7) = VarType(var(7))     ' VT_I4 / VT_I8

        Static pv(0 To 7) As LongPtr
        pv(0) = VarPtr(var(0))
        pv(1) = VarPtr(var(1))
        pv(2) = VarPtr(var(2))
        pv(3) = VarPtr(var(3))
        pv(4) = VarPtr(var(4))
        pv(5) = VarPtr(var(5))
        pv(6) = VarPtr(var(6))
        pv(7) = VarPtr(var(7))

        Init = True
    End If

    ' Adjust for arguments provided.
    arg = CLng(index)
    var(0) = dispid
    var(1) = VarPtr(iid)
    var(5) = rgVar

    ' Call DispCallFunc API.
    Dim dummy As Variant
    Dim hr As Long
    hr = DispCallFunc(ObjPtr(obj), oVft, CC_STDCALL, VT_I4, cvt, vt(0), pv(0), dummy)

    If hr < 0 Then
        Invoke = hr
        Exit Function
    End If

    Invoke = S_OK

End Function


#If API = False Then
'@Description "Initializes the Variant ByRef construct."
Private Sub InitializeVarByRef()

    VarByRef.vt = VarPtr(VarByRef.ref)
    CopyMemory VarByRef.vt, VBA.vbInteger Or VT_BYREF, vbSizeInteger

End Sub
#End If


#If API = False Then
'@Description "Copies a Long to a memory address using the Variant ByRef construct."
Private Sub CopyLngByRef( _
    ByVal address As LongPtr, _
    ByVal Value As Long, _
    ByRef vt As Variant, _
    ByRef ref As Variant _
)

    VarByRef.ref = address
    vt = VBA.vbLong Or VT_BYREF
    ref = Value

End Sub
#End If
