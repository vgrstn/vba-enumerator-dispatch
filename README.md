# vba-enumerator-dispatch
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for `IEnumVARIANT` interface implementation — no typelib required, dispatch binding via `DispCallFunc`.

Implements the full `IEnumVARIANT` interface (`Next`, `Skip`, `Reset`, `Clone`) in a standard module using `AddressOf` and a heap-allocated vtable. Items are retrieved one by one via `IDispatch::Invoke` called directly through `DispCallFunc`, so the iterable Class does not need to implement any interface.

> For the early-binding variant (fastest, requires `IEnumerator` interface) see [vba-enumerator](https://github.com/vgrstn/vba-enumerator).
> For the late-binding variant (no interface required, uses `CallByName`) see [vba-enumerator-late-binding](https://github.com/vgrstn/vba-enumerator-late-binding).

---

## 📦 Features

- **`For Each` without a typelib** — pure VBA, no external dependencies
- **No interface required** — items retrieved via `IDispatch::Invoke`; the iterable Class exposes any indexed property or method
- **Nested loops** — works correctly for nested `For Each` with mixed objects and mixed enumerators
- **Fast variant copy** — uses a Variant ByRef construct (~5× faster than `VariantCopy` API); switch to API mode with `#Const API = True`
- **Full COM lifecycle** — `QueryInterface`, `AddRef`, `Release`, `Clone` all correctly implemented
- **DISPID resolved once** — `IDispatch::GetIDsOfNames` called at construction time only; no name lookup in the hot path
- x86 / x64 compatible via `LongPtr` and `#If Win64`
- Pure VBA, zero dependencies, Rubberduck-friendly annotations

---

## 📁 Files

| File | Type | Description |
|---|---|---|
| `EnumerateDispatch.bas` | Module | `Enumerate(iterable, callback, count [, base])` — the main entry point |

Each file has a corresponding `_WithAttributes` version (e.g. `EnumerateDispatch_WithAttributes.bas`) with [Rubberduck](https://rubberduckvba.com/) annotations removed and VB attributes baked in. Import the `_WithAttributes` files if you are not using Rubberduck.

---

## ⚙️ Public Interface

### `EnumerateDispatch` module

| Member | Description |
|---|---|
| `Enumerate(iterable, callback, count [, base])` | Returns a synthetic `IEnumVARIANT` for the iterable object. `callback` is the name of the indexed property or method (e.g. `"Item"`). `count` is the number of items. `base` is the first index (default `1`). Raises an error if `iterable` is `Nothing` or `callback` is not callable on `iterable`. |

---

## 🚀 Quick Start

**1. Expose an indexed property in your Class:**

```vb
Public Property Get Item(ByVal index As Long) As Variant
    If VBA.IsObject(this.Items(index)) Then
        Set Item = this.Items(index)
    Else
        Item = this.Items(index)
    End If
End Property
```

**2. Add the `Enumerate` function:**

```vb
'@Enumerator
Public Function Enumerate() As IEnumVARIANT
    Set Enumerate = EnumerateDispatch.Enumerate(Me, "Item", this.Count)
End Function
```

**3. Use `For Each`:**

```vb
Dim obj As MyClass
Set obj = New MyClass
' ... populate obj ...

Dim v As Variant
For Each v In obj
    Debug.Print v
Next
```

---

## ⏱️ Performance

Timings for `n = 10,000` items (Immediate Window):

| Method | Time (ms) |
|---|---|
| Custom enumerator — `For Each` (VarByRef) | 27.90 |
| Custom enumerator — `For Each` (API) | 26.66 |
| VB Collection — `For Each` | 0.21 |
| VB Array — `For Each` | 0.14 |
| VB Array — `For i` | 0.08 |

The dominant cost is `DispCallFunc` dispatch per element — slower than the early-bound `IEnumerator.Item(i)` call in [vba-enumerator](https://github.com/vgrstn/vba-enumerator) (2.89 ms) and slightly slower than `VBA.CallByName` in [vba-enumerator-late-binding](https://github.com/vgrstn/vba-enumerator-late-binding) (19.58 ms). The DISPID is resolved once at construction; the hot path pays only the `DispCallFunc` call overhead per element. The VarByRef and API variant copy paths are nearly equal for this variant because the `pceltFetched` write is skipped by `For Each` (`pceltFetched = NULL`).

---

## ⚙️ Compiler directive

| Directive | Default | Description |
|---|---|---|
| `#Const API` | `False` | Write `pceltFetched` via Variant ByRef construct (fast) or `CopyMemory` API |

`#Const API = True` uses `CopyMemory` to write `pceltFetched`. `#Const API = False` uses the Variant ByRef construct. VBA's `For Each` always passes `pceltFetched = NULL`, so neither branch executes in normal use — the difference is negligible.

---

## 🧠 How it works

VBA's `For Each` requires the iterable object to expose `IEnumVARIANT` via `_NewEnum`. Rather than using a typelib to define `IEnumVARIANT`, this module:

1. Allocates a block of heap memory (`CoTaskMemAlloc`) large enough to hold the enumerator state (`TENUM` UDT)
2. Builds a vtable of function pointers (`AddressOf`) for the seven COM methods: `QueryInterface`, `AddRef`, `Release`, `Next`, `Skip`, `Reset`, `Clone`
3. Writes the vtable pointer into the first field of `TENUM` — making it a valid COM object
4. Resolves the DISPID for the callback property once via `IDispatch::GetIDsOfNames` and stores it in `TENUM.dispid`
5. Overwrites the return value of `Enumerate` with the heap pointer — returning the synthetic object as `IEnumVARIANT`
6. Keeps the iterable object alive via a `Static Collection` keyed by the heap address, compensating for reference count changes when the local `TENUM` goes out of scope

Based on work by Dexter Freivald (32-bit, late binding) and ideas from *Hardcore Visual Basic 5.0* by Bruce McKinney.

---

## 🧠 Implementation notes

### `TENUM` — synthetic COM object layout

```vb
Private Type TENUM
    pvTable  As LongPtr  ' MUST be first — COM reads vtable pointer at offset 0
    caller   As Object   ' late-bound reference to the iterable object
    dispid   As Long     ' DISPID resolved once at construction — no BSTR management
    nRef     As Long     ' COM reference count
    First    As Long     ' index of first item
    Last     As Long     ' index of last item
    Current  As Long     ' index of current position
End Type
```

`pvTable` is first because COM requires a pointer to the vtable at offset 0 of any COM object. `caller As Object` is a late-bound reference — no interface is required. `dispid As Long` holds the DISPID resolved at construction time; unlike the late-binding version there is no BSTR to allocate or free. There is no `Step` field — this module is ascending-only.

### vtable — built once per session

```vb
Static vTable(0 To 6) As LongPtr
If vTable(0) = vbNullPtr Then
    vTable(0) = VBA.CLngPtr(AddressOf IUnknown_QueryInterface)
    ...
    vTable(6) = VBA.CLngPtr(AddressOf IEnumVARIANT_Clone)
End If
```

The `Static` array persists for the lifetime of the VBA session. Subsequent calls to `Enumerate` reuse the same vtable without rebuilding it. Slots 0–2 are the three `IUnknown` methods; slots 3–6 are the four `IEnumVARIANT` methods.

### Return value trick

```vb
CopyMemory ByVal VarPtr(Enumerate), MemoryBlock, vbSizeLongPtr
```

`Enumerate` returns `IEnumVARIANT`. VBA stores the return value as an object pointer at `VarPtr(Enumerate)`. Overwriting those bytes with the heap block address makes VBA believe that address is a valid COM object — which it is, because the vtable pointer sits at offset 0. This is why `Enumerate` carries `'@Ignore NonReturningFunction`; the return value is set by raw memory write, not by a `Set Enumerate = ...` assignment.

### `KeepAlive` — reference count management

`CopyMemory ByVal MemoryBlock, obj, LenB(obj)` copies the `TENUM` struct to the heap as raw bytes — it copies the `caller` pointer without calling `AddRef`. When `obj` goes out of scope at function exit, VBA decrements the reference count on `obj.caller`, which could destroy the iterable. `KeepAlive` compensates:

```vb
Set KeepAlive(MemoryBlock) = iterable   ' hold one tracked reference
```

A `Static Collection` inside the property holds the reference, keyed by the heap block address. When `IUnknown_Release` sees `nRef = 0`:

```vb
Set KeepAlive(VarPtr(obj)) = Nothing   ' release tracked reference
CoTaskMemFree VarPtr(obj)              ' free heap block
```

The same pattern applies in `IEnumVARIANT_Clone`.

### `Resolve` — DISPID resolved once

```vb
hr = DispCallFunc(ObjPtr(obj), 5 * vbSizeLongPtr, CC_STDCALL, VT_I4, 5, vt(0), pv(0), dummy)
```

`IDispatch::GetIDsOfNames` sits at vtable slot 5 (byte offset `5 * vbSizeLongPtr`). The DISPID is written to a `Static dispid As Long` whose address is stored in `pv(4)` during first-call initialisation. A Static Init pattern (`If Init = False Then`) builds the argument arrays once; subsequent calls only update the two dynamic arguments (`Names(0)` and `var(0)` for the local `iid`). `Resolve` is called once per `Enumerate` invocation — never in the hot path.

### `Invoke` — hot path

```vb
hr = DispCallFunc(ObjPtr(obj), 6 * vbSizeLongPtr, CC_STDCALL, VT_I4, 8, vt(0), pv(0), dummy)
```

`IDispatch::Invoke` sits at vtable slot 6 (byte offset `6 * vbSizeLongPtr`). `DispCallFunc` marshals the eight `Invoke` arguments from the `var`/`vt`/`pv` Static arrays and writes the result `VARIANT` directly to `rgVar` — the destination address in `IEnumVARIANT_Next`. Per iteration the Adjust section updates four dynamic arguments (`arg`, `var(0)`, `var(1)`, `var(5)`) before the call; the Static arrays are not rebuilt.

### `oVft` — vtable byte offsets

```vb
Const slot As Long = 6
Const oVft As LongPtr = slot * vbSizeLongPtr   ' 24 on x86 / 48 on x64
```

`DispCallFunc`'s `oVft` parameter is a **byte offset** from the start of the vtable. The IDispatch vtable layout (zero-based):

| Slot | Method |
|---|---|
| 0 | QueryInterface |
| 1 | AddRef |
| 2 | Release |
| 3 | GetTypeInfoCount |
| 4 | GetTypeInfo |
| 5 | GetIDsOfNames |
| 6 | Invoke |

Named `Const` values rather than inline arithmetic make the intent clear and keep the code correct on both platforms without a conditional.

### `IEnumVARIANT_Next` — hot path

```vb
For i = obj.Current To obj.Last
    If Invoke(obj.caller, obj.dispid, i, rgVar) < 0 Then
        IEnumVARIANT_Next = E_FAIL
        Exit Function
    End If
    NumberFetched = NumberFetched + 1
    If NumberFetched = celt Then Exit For
    rgVar = rgVar + vbSizeVariant
Next
obj.Current = obj.Current + NumberFetched
```

Per iteration: one `Invoke` call (one `DispCallFunc` internally) that writes the result directly to `rgVar`. `rgVar` is advanced by `vbSizeVariant` (16 bytes x86 / 24 bytes x64) for multi-element fetches; VBA's `For Each` always requests one item at a time (`celt = 1`), so the inner `Exit For` fires immediately. `obj.Current` is updated in one step after the loop.

### `IEnumVARIANT_Skip` — `Select Case True`

```vb
Select Case True
Case celt = 0:  IEnumVARIANT_Skip = S_OK
Case celt < 0:  IEnumVARIANT_Skip = E_INVALIDARG
Case celt <= obj.Last - obj.Current + 1
    obj.Current = obj.Current + celt
    IEnumVARIANT_Skip = S_OK
Case Else
    obj.Current = obj.Last + 1
    IEnumVARIANT_Skip = S_FALSE
End Select
```

`obj.Current` is only mutated after the bounds check — state is never corrupted on an invalid or overshooting call. `celt = 0` is an explicit fast-path. On overshoot, `Current` is placed one step past `Last`, matching the post-exhaustion state that `Next` leaves.

### `IEnumVARIANT_Clone` — simpler than late-binding

```vb
' UDT assignment AddRefs caller — dispid is Long, no extra management needed.
Dim Copy As TENUM: Copy = obj
Copy.nRef = 1
```

`Copy = obj` copies all fields. VBA automatically calls `AddRef` on the embedded `caller As Object` during the UDT assignment. `dispid` is a plain `Long` — no BSTR allocation or independent copy is needed (contrast with the late-binding version which requires `SysAllocString` for the clone's callback string). The clone starts at `nRef = 1` and captures the enumeration position at the moment of cloning.

### Variant ByRef construct

```vb
Private Type CONSTRUCT
    vt  As Variant
    ref As Variant
End Type
Private VarByRef As CONSTRUCT
```

`InitializeVarByRef` sets `vt` to `VT_INTEGER | VT_BYREF` pointing at `ref`. `CopyLngByRef` uses this mechanism (`vbLong | VT_BYREF`) to write `pceltFetched` without a `CopyMemory` API call. Unlike the late-binding version, `CopyVarByRef` is not needed — item retrieval goes through `Invoke` which writes the result `VARIANT` directly to `rgVar` via `DispCallFunc`.

### GUID comparison — field-by-field

```vb
' Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
IsIID_IUnknown = (id.Data1 = &H0) And (id.Data2 = &H0) And ... And (id.Data4(7) = &H46)
```

Direct integer comparison on the `GUID` UDT fields — no string parsing. The commented `Const` documents the expected GUID string without runtime cost.

### `VARSIZE` — x86/x64 portability

```vb
#If Win64 Then
    vbSizeLongPtr = 8
    vbSizeVariant = 24
#Else
    vbSizeLongPtr = 4
    vbSizeVariant = 16
#End If
```

A `Variant` is 16 bytes on x86 and 24 bytes on x64. All `CopyMemory` sizes, pointer arithmetic, and `oVft` constants use these values — no hard-coded numbers appear in the code.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
