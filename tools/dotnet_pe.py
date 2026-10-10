"""Dependency-free ECMA-335 (PE/CLI) reader — proves what a packaged .NET assembly actually IS.

PROMPT-030.  The Windows Portable died at start-up with

    RuntimeError: Failed to resolve Python.Runtime.Loader.Initialize from
                  H:\\...\\_internal\\pythonnet\\runtime\\Python.Runtime.dll

raised by ``clr_loader/netfx.py::_get_callable`` whenever the native ``pyclr_get_function()`` returns
NULL.  That one string is the *only* thing the netfx backend reports: its cdef
(``clr_loader/ffi/netfx.py``) declares ``pyclr_initialize`` / ``pyclr_create_appdomain`` /
``pyclr_get_function`` / ``pyclr_close_appdomain`` / ``pyclr_finalize`` and **no error accessor**, so
every distinct failure — a missing file, a blocked file, a too-old .NET Framework, a wrong-architecture
host, a renamed non-.NET payload, or a signature the CLR refuses to turn into a delegate — is collapsed
into "Failed to resolve".  ``tools/build_portable.py`` could previously only prove that the file EXISTS
and hashes correctly, which is why the build passed and the exe still crashed.

This module closes the part of that gap that does not need a Windows CLR: it reads the managed metadata
out of the PE itself, on any platform, with nothing but the standard library, and answers the questions
that decide whether ``pyclr_get_function`` can possibly succeed —

  * is this a managed image at all (CLI header present)?
  * what is its **assembly identity** (name / version / culture / public key token)?  A DLL that merely
    happens to be named ``Python.Runtime.dll`` is not the pythonnet runtime.
  * does the type ``Python.Runtime.Loader`` exist, and does it declare a **static** ``Initialize`` whose
    signature is exactly ``int32 (native int, int32)`` — the ``entry_point`` typedef that
    ``clr_loader/ffi/netfx.py`` hands to ``Marshal.GetFunctionPointerForDelegate``?  Anything else makes
    the delegate creation throw and ``pyclr_get_function`` return NULL.
  * which assemblies does it reference, and what does ``TargetFrameworkAttribute`` say it targets?  That
    is what turns "does this PC have a suitable .NET Framework" from a guess into a comparison against a
    concrete minimum release number.
  * what is the PE architecture and is it IL-only / 32-bit-required?  For the native, mixed-mode host
    (``clr_loader/ffi/dlls/<arch>/ClrLoader.dll``) it also lists the exported symbols, so a package that
    shipped the x86 host into an amd64 process — or a host missing ``pyclr_get_function`` — is caught
    here rather than at double-click time.

Deliberately small: only the tables this decision needs are decoded, unknown ones are skipped by their
declared row size, and every parse failure raises :class:`DotNetPEError` instead of returning a
half-filled object that a build gate might trust.
"""
from __future__ import annotations

import binascii
import hashlib
import struct
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

__all__ = [
    "DotNetPEError",
    "AssemblyIdentity",
    "MethodInfo",
    "ManagedImage",
    "inspect",
    "signature_matches_entry_point",
    "ENTRY_POINT_SIGNATURE",
    "ENTRY_POINT_TYPE",
    "ENTRY_POINT_METHOD",
]

#: The entry point pythonnet asks clr_loader for (``pythonnet/__init__.py::load``).
ENTRY_POINT_TYPE = "Python.Runtime.Loader"
ENTRY_POINT_METHOD = "Initialize"
#: clr_loader's cdef: ``typedef int (*entry_point)(void* buffer, int size);`` and its own contract in
#: ``clr_loader/types.py``: "The function must be ``static``, and it must have the signature
#: ``int Func(IntPtr ptr, int size)``."
ENTRY_POINT_SIGNATURE = "int32 (native int, int32)"

_METADATA_SIGNATURE = 0x424A5342
_IMAGE_FILE_MACHINE = {0x014C: "x86", 0x8664: "amd64", 0x0200: "ia64", 0xAA64: "arm64"}

# COMIMAGE_FLAGS (ECMA-335 II.25.3.3.1)
_FLAG_ILONLY = 0x00000001
_FLAG_32BITREQUIRED = 0x00000002
_FLAG_STRONGNAMESIGNED = 0x00000008
_FLAG_32BITPREFERRED = 0x00020000

# MethodDef.Flags (ECMA-335 II.23.1.10) — the bits that decide delegate compatibility.
METHOD_FLAG_STATIC = 0x0010

_METHOD_VISIBILITY_MASK = 0x0007
_METHOD_VISIBILITY = {
    0: "compiler-controlled",
    1: "private",
    2: "family-and-assembly",
    3: "assembly",
    4: "family",
    5: "family-or-assembly",
    6: "public",
}

# Element types (ECMA-335 II.23.1.16)
_ET = {
    0x01: "void", 0x02: "bool", 0x03: "char", 0x04: "int8", 0x05: "uint8",
    0x06: "int16", 0x07: "uint16", 0x08: "int32", 0x09: "uint32", 0x0A: "int64",
    0x0B: "uint64", 0x0C: "float32", 0x0D: "float64", 0x0E: "string",
    0x16: "typedref", 0x18: "native int", 0x19: "native uint", 0x1C: "object",
}
_ET_SIMPLE = frozenset(_ET)
_ET_TOKED = {0x11: "valuetype", 0x12: "class", 0x15: "genericinst", 0x1D: "szarray",
             0x14: "array", 0x0F: "ptr", 0x10: "byref", 0x13: "var", 0x1E: "mvar"}


class DotNetPEError(Exception):
    """The file is not the managed image the caller expected.  Always carries the reason."""


# --------------------------------------------------------------------------- metadata table schemas
# Simple index columns name their target table; coded indexes name an entry of _CODED; "u1"/"u2"/"u4" are
# fixed width; "str"/"guid"/"blob" are heap indexes whose width comes from the #~ HeapSizes byte.
_CODED: Dict[str, Tuple[int, Sequence[Optional[int]]]] = {
    "TypeDefOrRef": (2, (0x02, 0x01, 0x1B)),
    "HasConstant": (2, (0x04, 0x08, 0x17)),
    "HasCustomAttribute": (5, (0x06, 0x04, 0x01, 0x02, 0x08, 0x09, 0x0A, 0x00, 0x0E, 0x17,
                               0x14, 0x11, 0x1A, 0x1B, 0x20, 0x23, 0x26, 0x27, 0x28, 0x2A,
                               0x2C, 0x2B)),
    "HasFieldMarshal": (1, (0x04, 0x08)),
    "HasDeclSecurity": (2, (0x02, 0x06, 0x20)),
    "MemberRefParent": (3, (0x02, 0x01, 0x1A, 0x06, 0x1B)),
    "HasSemantics": (1, (0x14, 0x17)),
    "MethodDefOrRef": (1, (0x06, 0x0A)),
    "MemberForwarded": (1, (0x04, 0x06)),
    "Implementation": (2, (0x26, 0x23, 0x27)),
    "CustomAttributeType": (3, (None, None, 0x06, 0x0A, None)),
    "ResolutionScope": (2, (0x00, 0x1A, 0x23, 0x01)),
    "TypeOrMethodDef": (1, (0x02, 0x06)),
}

_TABLES: Dict[int, Tuple[str, Sequence[Tuple[str, object]]]] = {
    0x00: ("Module", (("Generation", "u2"), ("Name", "str"), ("Mvid", "guid"),
                      ("EncId", "guid"), ("EncBaseId", "guid"))),
    0x01: ("TypeRef", (("ResolutionScope", "ResolutionScope"), ("Name", "str"), ("Namespace", "str"))),
    0x02: ("TypeDef", (("Flags", "u4"), ("Name", "str"), ("Namespace", "str"),
                       ("Extends", "TypeDefOrRef"), ("FieldList", 0x04), ("MethodList", 0x06))),
    0x04: ("Field", (("Flags", "u2"), ("Name", "str"), ("Signature", "blob"))),
    0x06: ("MethodDef", (("RVA", "u4"), ("ImplFlags", "u2"), ("Flags", "u2"), ("Name", "str"),
                         ("Signature", "blob"), ("ParamList", 0x08))),
    0x08: ("Param", (("Flags", "u2"), ("Sequence", "u2"), ("Name", "str"))),
    0x09: ("InterfaceImpl", (("Class", 0x02), ("Interface", "TypeDefOrRef"))),
    0x0A: ("MemberRef", (("Class", "MemberRefParent"), ("Name", "str"), ("Signature", "blob"))),
    0x0B: ("Constant", (("Type", "u1"), ("Pad", "u1"), ("Parent", "HasConstant"), ("Value", "blob"))),
    0x0C: ("CustomAttribute", (("Parent", "HasCustomAttribute"), ("Type", "CustomAttributeType"),
                               ("Value", "blob"))),
    0x0D: ("FieldMarshal", (("Parent", "HasFieldMarshal"), ("NativeType", "blob"))),
    0x0E: ("DeclSecurity", (("Action", "u2"), ("Parent", "HasDeclSecurity"), ("PermissionSet", "blob"))),
    0x0F: ("ClassLayout", (("PackingSize", "u2"), ("ClassSize", "u4"), ("Parent", 0x02))),
    0x10: ("FieldLayout", (("Offset", "u4"), ("Field", 0x04))),
    0x11: ("StandAloneSig", (("Signature", "blob"),)),
    0x12: ("EventMap", (("Parent", 0x02), ("EventList", 0x14))),
    0x14: ("Event", (("EventFlags", "u2"), ("Name", "str"), ("EventType", "TypeDefOrRef"))),
    0x15: ("PropertyMap", (("Parent", 0x02), ("PropertyList", 0x17))),
    0x17: ("Property", (("Flags", "u2"), ("Name", "str"), ("Type", "blob"))),
    0x18: ("MethodSemantics", (("Semantics", "u2"), ("Method", 0x06), ("Association", "HasSemantics"))),
    0x19: ("MethodImpl", (("Class", 0x02), ("MethodBody", "MethodDefOrRef"),
                          ("MethodDeclaration", "MethodDefOrRef"))),
    0x1A: ("ModuleRef", (("Name", "str"),)),
    0x1B: ("TypeSpec", (("Signature", "blob"),)),
    0x1C: ("ImplMap", (("MappingFlags", "u2"), ("MemberForwarded", "MemberForwarded"),
                       ("ImportName", "str"), ("ImportScope", 0x1A))),
    0x1D: ("FieldRVA", (("RVA", "u4"), ("Field", 0x04))),
    0x20: ("Assembly", (("HashAlgId", "u4"), ("MajorVersion", "u2"), ("MinorVersion", "u2"),
                        ("BuildNumber", "u2"), ("RevisionNumber", "u2"), ("Flags", "u4"),
                        ("PublicKey", "blob"), ("Name", "str"), ("Culture", "str"))),
    0x21: ("AssemblyProcessor", (("Processor", "u4"),)),
    0x22: ("AssemblyOS", (("OSPlatformId", "u4"), ("OSMajorVersion", "u4"), ("OSMinorVersion", "u4"))),
    0x23: ("AssemblyRef", (("MajorVersion", "u2"), ("MinorVersion", "u2"), ("BuildNumber", "u2"),
                           ("RevisionNumber", "u2"), ("Flags", "u4"), ("PublicKeyOrToken", "blob"),
                           ("Name", "str"), ("Culture", "str"), ("HashValue", "blob"))),
    0x24: ("AssemblyRefProcessor", (("Processor", "u4"), ("AssemblyRef", 0x23))),
    0x25: ("AssemblyRefOS", (("OSPlatformId", "u4"), ("OSMajorVersion", "u4"), ("OSMinorVersion", "u4"),
                             ("AssemblyRef", 0x23))),
    0x26: ("File", (("Flags", "u4"), ("Name", "str"), ("HashValue", "blob"))),
    0x27: ("ExportedType", (("Flags", "u4"), ("TypeDefId", "u4"), ("TypeName", "str"),
                            ("TypeNamespace", "str"), ("Implementation", "Implementation"))),
    0x28: ("ManifestResource", (("Offset", "u4"), ("Flags", "u4"), ("Name", "str"),
                                ("Implementation", "Implementation"))),
    0x29: ("NestedClass", (("NestedClass", 0x02), ("EnclosingClass", 0x02))),
    0x2A: ("GenericParam", (("Number", "u2"), ("Flags", "u2"), ("Owner", "TypeOrMethodDef"),
                            ("Name", "str"))),
    0x2B: ("MethodSpec", (("Method", "MethodDefOrRef"), ("Instantiation", "blob"))),
    0x2C: ("GenericParamConstraint", (("Owner", 0x2A), ("Constraint", "TypeDefOrRef"))),
}


@dataclass(frozen=True)
class AssemblyIdentity:
    """An Assembly (0x20) or AssemblyRef (0x23) row, reduced to what identifies it."""
    name: str
    version: Tuple[int, int, int, int]
    culture: str
    public_key_token: str          # hex, "" when the assembly is not strong-named

    @property
    def version_text(self) -> str:
        return ".".join(str(part) for part in self.version)

    def __str__(self) -> str:                          # noqa: D105 – mirrors .NET's display form
        culture = f", Culture={self.culture or 'neutral'}"
        token = f", PublicKeyToken={self.public_key_token or 'null'}"
        return f"{self.name}, Version={self.version_text}{culture}{token}"


@dataclass(frozen=True)
class MethodInfo:
    """A MethodDef row plus the decoded shape clr_loader has to agree with."""
    name: str
    flags: int
    signature: str                 # e.g. "int32 (native int, int32)"

    @property
    def is_static(self) -> bool:
        return bool(self.flags & METHOD_FLAG_STATIC)

    @property
    def visibility(self) -> str:
        return _METHOD_VISIBILITY.get(self.flags & _METHOD_VISIBILITY_MASK, "unknown")


@dataclass(frozen=True)
class ManagedImage:
    """Everything a build gate needs to know about one PE file."""
    path: str
    size: int
    sha256: str
    machine: str                   # "amd64" / "x86" / …
    pe32_plus: bool
    is_dll: bool
    is_managed: bool
    cli_runtime_version: str       # from the metadata root, e.g. "v4.0.30319"
    cli_flags: int
    il_only: bool
    requires_32bit: bool
    strong_name_signed: bool
    exports: Tuple[str, ...]       # native export names (a mixed-mode host such as ClrLoader.dll has them)
    assembly: Optional[AssemblyIdentity]
    assembly_refs: Tuple[AssemblyIdentity, ...]
    target_frameworks: Tuple[str, ...]
    type_names: Tuple[str, ...]    # every TypeDef, fully qualified
    methods: Dict[str, Tuple[MethodInfo, ...]]

    # ------------------------------------------------------------------ convenience predicates
    def has_type(self, fqn: str) -> bool:
        """True when the image declares ``fqn`` (namespace-qualified) as one of its own types."""
        return fqn in self.type_names

    def find_method(self, type_fqn: str, method_name: str) -> Optional[MethodInfo]:
        """The first ``method_name`` declared on ``type_fqn``, or None when absent."""
        for method in self.methods.get(type_fqn, ()):
            if method.name == method_name:
                return method
        return None

    @property
    def arch_text(self) -> str:
        """Bitness the CLR will actually apply: IL-only images follow the host process."""
        if not self.il_only:
            return self.machine
        if self.requires_32bit:
            return "x86"
        return "any-cpu"


def signature_matches_entry_point(image: ManagedImage,
                                  type_fqn: str = ENTRY_POINT_TYPE,
                                  method: str = ENTRY_POINT_METHOD) -> Tuple[bool, str]:
    """Does ``image`` declare the exact entry point clr_loader must bind?

    Returns ``(ok, reason)``; ``reason`` is always printable and names the specific mismatch, because the
    whole point of PROMPT-030 is that "Failed to resolve …" on its own explains nothing.
    """
    if not image.is_managed:
        return False, f"{Path(image.path).name} is not a managed (.NET) image — it has no CLI header"
    if not image.has_type(type_fqn):
        return False, f"type {type_fqn} is not declared by this assembly (assembly={image.assembly})"
    found = image.find_method(type_fqn, method)
    if found is None:
        return False, f"{type_fqn} exists but declares no method named {method}"
    if not found.is_static:
        return False, (f"{type_fqn}.{method} is not static — clr_loader/types.py requires "
                       "``static int Func(IntPtr, int)``")
    if found.signature != ENTRY_POINT_SIGNATURE:
        return False, (f"{type_fqn}.{method} has signature '{found.signature}', clr_loader needs "
                       f"'{ENTRY_POINT_SIGNATURE}'")
    return True, f"{type_fqn}.{method} is static {found.signature} ({found.visibility})"


# --------------------------------------------------------------------------- PE / CLI parsing
def _read(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except OSError as exc:
        raise DotNetPEError(f"cannot read {path}: {exc}") from exc


class _Cursor:
    """Bounds-checked sequential reader — a truncated PE must raise, never return zeros."""

    def __init__(self, blob: bytes, offset: int = 0):
        self._blob = blob
        self.pos = offset

    def take(self, n: int) -> bytes:
        end = self.pos + n
        if end > len(self._blob):
            raise DotNetPEError(f"unexpected end of data at {self.pos}+{n} (have {len(self._blob)})")
        out = self._blob[self.pos:end]
        self.pos = end
        return out

    def u1(self) -> int:
        return self.take(1)[0]

    def u2(self) -> int:
        return struct.unpack("<H", self.take(2))[0]

    def u4(self) -> int:
        return struct.unpack("<I", self.take(4))[0]

    def compressed(self) -> int:
        """ECMA-335 II.23.2 compressed unsigned integer."""
        first = self.u1()
        if not first & 0x80:
            return first
        if not first & 0x40:
            return ((first & 0x3F) << 8) | self.u1()
        return ((first & 0x1F) << 24) | (self.u1() << 16) | (self.u1() << 8) | self.u1()

    def cstring(self) -> str:
        out = bytearray()
        while True:
            byte = self.take(1)
            if byte == b"\x00":
                break
            out += byte
        return out.decode("utf-8", "replace")


def _sections(data: bytes, pe: int) -> Tuple[List[Tuple[int, int, int, int]], int, int, int]:
    machine, nsec = struct.unpack_from("<HH", data, pe + 4)
    optsize = struct.unpack_from("<H", data, pe + 20)[0]
    chars = struct.unpack_from("<H", data, pe + 22)[0]
    magic = struct.unpack_from("<H", data, pe + 24)[0]
    table = pe + 24 + optsize
    out = []
    for i in range(nsec):
        base = table + 40 * i
        vsize, vaddr, rsize, raddr = struct.unpack_from("<IIII", data, base + 8)
        out.append((vaddr, vsize, raddr, rsize))
    return out, machine, magic, chars


def _rva_to_offset(sections: Sequence[Tuple[int, int, int, int]], rva: int) -> int:
    for vaddr, vsize, raddr, rsize in sections:
        if vaddr <= rva < vaddr + max(vsize, rsize):
            return raddr + (rva - vaddr)
    raise DotNetPEError(f"RVA {hex(rva)} is not inside any PE section")


def _data_directory(data: bytes, pe: int, magic: int, index: int) -> Tuple[int, int]:
    base = pe + 24 + (112 if magic == 0x20B else 96)
    count = struct.unpack_from("<I", data, base - 4)[0]
    if index >= count:
        return 0, 0
    return struct.unpack_from("<II", data, base + 8 * index)


def _exports(data: bytes, sections, pe: int, magic: int) -> Tuple[str, ...]:
    rva, size = _data_directory(data, pe, magic, 0)
    if not rva or not size:
        return ()
    off = _rva_to_offset(sections, rva)
    nnames = struct.unpack_from("<I", data, off + 24)[0]
    names_rva = struct.unpack_from("<I", data, off + 32)[0]
    if not nnames or not names_rva:
        return ()
    table = _rva_to_offset(sections, names_rva)
    out = []
    for i in range(min(nnames, 4096)):                     # a sane ceiling on a malformed image
        name_rva = struct.unpack_from("<I", data, table + 4 * i)[0]
        out.append(_Cursor(data, _rva_to_offset(sections, name_rva)).cstring())
    return tuple(sorted(out))


@dataclass
class _Heaps:
    strings: bytes
    blobs: bytes
    guids: bytes
    wide_strings: bool
    wide_blobs: bool
    wide_guids: bool

    def string(self, index: int) -> str:
        if not index:
            return ""
        end = self.strings.find(b"\x00", index)
        if end < 0:
            end = len(self.strings)
        return self.strings[index:end].decode("utf-8", "replace")

    def blob(self, index: int) -> bytes:
        if not index:
            return b""
        cur = _Cursor(self.blobs, index)
        return cur.take(cur.compressed())


def _decode_type(cur: _Cursor, depth: int = 0) -> str:
    """One Type in a signature blob.  Enough of ECMA-335 II.23.2.10 to name the entry point's shape."""
    if depth > 8:
        raise DotNetPEError("signature nesting too deep")
    while True:                                            # custom modifiers are decoration, skip them
        mark = cur.pos
        element = cur.u1()
        if element in (0x1F, 0x20):                        # CMOD_REQD / CMOD_OPT
            cur.compressed()
            continue
        cur.pos = mark
        element = cur.u1()
        break
    if element in _ET_SIMPLE:
        return _ET[element]
    if element == 0x45:                                    # PINNED
        return _decode_type(cur, depth + 1)
    if element in (0x11, 0x12, 0x1D, 0x0F, 0x10, 0x13, 0x1E):
        inner = _decode_type(cur, depth + 1) if element in (0x1D, 0x0F, 0x10, 0x13, 0x1E) else None
        if inner is None:
            cur.compressed()
            return _ET_TOKED[element]
        return f"{_ET_TOKED[element]}({inner})"
    if element == 0x15:                                    # GENERICINST
        kind = cur.u1()
        cur.compressed()
        arity = cur.compressed()
        for _ in range(arity):
            cur.compressed()
        return f"{_ET_TOKED.get(kind, hex(kind))}<{arity}>"
    if element == 0x14:                                    # ARRAY
        _decode_type(cur, depth + 1)
        rank = cur.compressed()
        sizes = cur.compressed()
        for _ in range(sizes):
            cur.compressed()
        bounds = cur.compressed()
        for _ in range(bounds):
            cur.compressed()
        return f"array(rank={rank})"
    raise DotNetPEError(f"unsupported signature element type {hex(element)}")


def _decode_method_signature(blob: bytes) -> str:
    """``int32 (native int, int32)``-style rendering of a MethodDef signature blob."""
    if not blob:
        raise DotNetPEError("method has no signature blob")
    cur = _Cursor(blob)
    convention = cur.u1()
    if convention & 0x10:                                  # GENERIC — a type/method arity prefix follows
        cur.compressed()
    arity = cur.compressed()
    ret = _decode_type(cur)
    params = [_decode_type(cur) for _ in range(arity)]
    return f"{ret} ({', '.join(params)})"


def _public_key_token(public_key: bytes) -> str:
    """The 8-byte token .NET displays: the tail of SHA1 over the public key, byte-reversed."""
    if not public_key:
        return ""
    digest = hashlib.sha1(public_key).digest()            # noqa: S324 – this IS the .NET algorithm
    return binascii.hexlify(digest[-8:][::-1]).decode()


def _target_frameworks(heaps: _Heaps, rows: Dict[int, List[dict]], assembly_rows: int) -> Tuple[str, ...]:
    """Decode ``[TargetFramework("...")]`` from the assembly's custom attributes."""
    out: List[str] = []
    for attr in rows.get(0x0C, []):
        parent = attr["Parent"]
        if parent[0] != 0x20 or parent[1] - 1 >= assembly_rows:
            continue
        ctor = attr["Type"]
        owner = _attribute_owner_name(ctor, rows)
        if owner != "System.Runtime.Versioning.TargetFrameworkAttribute":
            continue
        blob = heaps.blob(attr["Value"][1] if isinstance(attr["Value"], tuple) else attr["Value"])
        if len(blob) >= 2 and blob[0] == 0x01 and blob[1] == 0x00:
            cur = _Cursor(blob, 2)
            try:
                length = cur.compressed()
                out.append(cur.take(length).decode("utf-8", "replace"))
            except DotNetPEError:
                continue
    return tuple(dict.fromkeys(out))


def _attribute_owner_name(ctor, rows: Dict[int, List[dict]]) -> str:
    """Fully qualified owner of a CustomAttributeType coded index (MethodDef or MemberRef)."""
    if ctor[0] == 0x06:                                    # MethodDef in this image
        return ""
    if ctor[0] != 0x0A:
        return ""
    index = ctor[1] - 1
    members = rows.get(0x0A, [])
    if index < 0 or index >= len(members):
        return ""
    parent = members[index]["Class"]
    if parent[0] != 0x01:                                  # TypeRef
        return ""
    refs = rows.get(0x01, [])
    pindex = parent[1] - 1
    if pindex < 0 or pindex >= len(refs):
        return ""
    ref = refs[pindex]
    namespace, name = ref["Namespace"], ref["Name"]
    return f"{namespace}.{name}" if namespace else name


def inspect(path) -> ManagedImage:
    """Read one PE file and prove what it is.  Raises :class:`DotNetPEError` with a printable reason."""
    path = Path(path)
    data = _read(path)
    if len(data) < 64 or data[:2] != b"MZ":
        raise DotNetPEError(f"{path.name} is not a PE image (no MZ header)")
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe:pe + 4] != b"PE\0\0":
        raise DotNetPEError(f"{path.name} has no PE signature")
    sections, machine, magic, chars = _sections(data, pe)
    machine_text = _IMAGE_FILE_MACHINE.get(machine, f"unknown({hex(machine)})")

    cli_rva, cli_size = _data_directory(data, pe, magic, 14)
    if not cli_rva or cli_size < 72:
        raise DotNetPEError(f"{path.name} has no CLI header — it is a native DLL, not a .NET assembly")

    cli = _rva_to_offset(sections, cli_rva)
    flags = struct.unpack_from("<I", data, cli + 16)[0]
    meta_rva, meta_size = struct.unpack_from("<II", data, cli + 8)
    if not meta_rva or not meta_size:
        raise DotNetPEError(f"{path.name} has a CLI header but no metadata directory")

    root = _Cursor(data, _rva_to_offset(sections, meta_rva))
    if root.u4() != _METADATA_SIGNATURE:
        raise DotNetPEError(f"{path.name} metadata signature is wrong")
    root.u2(); root.u2(); root.u4()                        # major, minor, reserved
    version = root.take(root.u4()).split(b"\x00")[0].decode("utf-8", "replace")
    root.u2()                                               # flags
    streams = {}
    for _ in range(root.u2()):
        offset, size = root.u4(), root.u4()
        name = root.cstring()
        while root.pos % 4:                                 # stream names are 4-byte aligned
            root.take(1)
        streams[name] = (meta_rva + offset, size)

    if "#~" not in streams and "#-" not in streams:
        raise DotNetPEError(f"{path.name} has no metadata tables stream")
    tables_rva, _ = streams.get("#~") or streams["#-"]
    cur = _Cursor(data, _rva_to_offset(sections, tables_rva))
    cur.u4(); cur.u1(); cur.u1()                            # reserved, major, minor
    heap_sizes = cur.u1()
    cur.u1()                                                # reserved (always 1)
    valid = cur.u4() | (cur.u4() << 32)
    cur.u4(); cur.u4()                                      # sorted mask
    counts: Dict[int, int] = {}
    for bit in range(64):
        if valid >> bit & 1:
            counts[bit] = cur.u4()

    heaps = _Heaps(
        strings=_stream_bytes(data, sections, streams, "#Strings"),
        blobs=_stream_bytes(data, sections, streams, "#Blob"),
        guids=_stream_bytes(data, sections, streams, "#GUID"),
        wide_strings=bool(heap_sizes & 0x01),
        wide_guids=bool(heap_sizes & 0x02),
        wide_blobs=bool(heap_sizes & 0x04),
    )

    def index_width(kind) -> int:
        if kind == "str":
            return 4 if heaps.wide_strings else 2
        if kind == "guid":
            return 4 if heaps.wide_guids else 2
        if kind == "blob":
            return 4 if heaps.wide_blobs else 2
        if isinstance(kind, int):
            return 4 if counts.get(kind, 0) >= 0x10000 else 2
        if kind in ("u1", "u2", "u4"):
            return {"u1": 1, "u2": 2, "u4": 4}[kind]
        bits, targets = _CODED[kind]
        largest = max((counts.get(t, 0) for t in targets if t), default=0)
        return 4 if largest >= (1 << (16 - bits)) else 2

    widths = {t: [(n, k, index_width(k)) for n, k in schema]
              for t, (_, schema) in _TABLES.items() if t in counts}
    rows: Dict[int, List[dict]] = {}
    for table in sorted(counts):
        schema = widths.get(table)
        raw = counts[table]
        if schema is None:                                  # a table this reader does not model
            raise DotNetPEError(f"{path.name} uses metadata table {hex(table)} which is not modelled")
        parsed = []
        for _ in range(raw):
            row = {}
            for name, kind, w in schema:
                value = int.from_bytes(cur.take(w), "little")
                if kind == "str":
                    value = heaps.string(value)
                elif kind == "blob":
                    row[name] = value                       # keep the heap index for later decoding
                    continue
                elif kind == "guid":
                    pass
                elif isinstance(kind, str) and kind in _CODED:
                    bits, targets = _CODED[kind]
                    tag = value & ((1 << bits) - 1)
                    target = targets[tag] if tag < len(targets) else None
                    value = (target, value >> bits)
                row[name] = value
            parsed.append(row)
        rows[table] = parsed

    assembly = None
    assembly_row = (rows.get(0x20) or [None])[0]
    if assembly_row is not None:
        assembly = AssemblyIdentity(
            name=assembly_row["Name"],
            version=(assembly_row["MajorVersion"], assembly_row["MinorVersion"],
                     assembly_row["BuildNumber"], assembly_row["RevisionNumber"]),
            culture=assembly_row["Culture"],
            public_key_token=_public_key_token(heaps.blob(assembly_row["PublicKey"])),
        )
    refs = tuple(
        AssemblyIdentity(
            name=r["Name"],
            version=(r["MajorVersion"], r["MinorVersion"], r["BuildNumber"], r["RevisionNumber"]),
            culture=r["Culture"],
            public_key_token=_public_key_token(heaps.blob(r["PublicKeyOrToken"])),
        )
        for r in rows.get(0x23, [])
    )

    methods_by_type: Dict[str, List[MethodInfo]] = {}
    type_names: List[str] = []
    typedefs = rows.get(0x02, [])
    methoddefs = rows.get(0x06, [])
    for i, typedef in enumerate(typedefs):
        namespace, name = typedef["Namespace"], typedef["Name"]
        if name.startswith("<"):                            # compiler-generated
            continue
        fqn = f"{namespace}.{name}" if namespace else name
        type_names.append(fqn)
        first = typedef["MethodList"]
        last = typedefs[i + 1]["MethodList"] if i + 1 < len(typedefs) else len(methoddefs) + 1
        bucket: List[MethodInfo] = []
        for index in range(first, last):
            row = methoddefs[index - 1]
            try:
                signature = _decode_method_signature(heaps.blob(row["Signature"]))
            except DotNetPEError:
                signature = "undecodable"
            bucket.append(MethodInfo(name=row["Name"], flags=row["Flags"], signature=signature))
        methods_by_type[fqn] = tuple(bucket)

    return ManagedImage(
        path=str(path),
        size=len(data),
        sha256=hashlib.sha256(data).hexdigest(),
        machine=machine_text,
        pe32_plus=(magic == 0x20B),
        is_dll=bool(chars & 0x2000),
        is_managed=True,
        cli_runtime_version=version,
        cli_flags=flags,
        il_only=bool(flags & _FLAG_ILONLY),
        requires_32bit=bool(flags & (_FLAG_32BITREQUIRED | _FLAG_32BITPREFERRED)),
        strong_name_signed=bool(flags & _FLAG_STRONGNAMESIGNED),
        exports=_exports(data, sections, pe, magic),
        assembly=assembly,
        assembly_refs=refs,
        target_frameworks=_target_frameworks(heaps, rows, len(rows.get(0x20, []))),
        type_names=tuple(type_names),
        methods=methods_by_type,
    )


def _stream_bytes(data: bytes, sections, streams: Dict[str, Tuple[int, int]], name: str) -> bytes:
    if name not in streams:
        return b""
    rva, size = streams[name]
    offset = _rva_to_offset(sections, rva)
    return data[offset:offset + size]


def describe(image: ManagedImage) -> str:
    """A short, log-safe multi-line summary — what the build prints next to the SHA256."""
    lines = [
        f"assembly       {image.assembly}" if image.assembly else "assembly       <none>",
        f"machine        {image.machine} (PE32+={image.pe32_plus}, IL-only={image.il_only}, "
        f"32-bit-required={image.requires_32bit}) → {image.arch_text}",
        f"metadata       {image.cli_runtime_version}, flags={hex(image.cli_flags)}, "
        f"strong-name={image.strong_name_signed}",
        f"targets        {', '.join(image.target_frameworks) or '<no TargetFrameworkAttribute>'}",
        f"references     {', '.join(f'{r.name} {r.version_text}' for r in image.assembly_refs) or '<none>'}",
        f"types          {len(image.type_names)}",
    ]
    if image.exports:
        lines.append(f"exports        {', '.join(image.exports)}")
    return "\n".join(lines)
