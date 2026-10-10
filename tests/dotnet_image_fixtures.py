"""Synthetic ECMA-335/PE images for PROMPT-030 STRUCTURAL tests.

These fixtures model assembly identity, metadata, target framework, architecture and export-name tables so
``tools/dotnet_pe.py`` and the cross-platform validator checks can be tested without .NET or installed DLLs.
They are not substitutes for executing a real runtime: in particular, ``clr_loader_image()`` writes the
``pyclr_*`` export names and placeholder function RVAs, but contains no callable native/mixed-mode
``ClrLoader`` implementation. Never use it in a test that claims Windows DLL loadability or symbol resolution.

PROMPT-030 deliberately makes the production gate stronger than presence/hash checks. Keep these synthetic
images for structural and malformed-image cases; Windows live-probe tests must stage the real installed
runtime payloads. No platform, .NET installation or repository file is required for the structural tests.
"""
from __future__ import annotations

import struct
from typing import Dict, Sequence, Tuple

# MethodDef.Flags (ECMA-335 II.23.1.10)
METHOD_PUBLIC = 0x0006
METHOD_STATIC = 0x0010
METHOD_HIDE_BY_SIG = 0x0080
#: A C# ``static class`` is emitted abstract + sealed.
TYPE_ABSTRACT = 0x00000080
TYPE_SEALED = 0x00000100
TYPE_PUBLIC = 0x00000001

# COMIMAGE_FLAGS (ECMA-335 II.25.3.3.1)
FLAG_ILONLY = 0x00000001
FLAG_32BITREQUIRED = 0x00000002
FLAG_STRONGNAMESIGNED = 0x00000008

_MACHINE = {"x86": 0x014C, "amd64": 0x8664}
_ELEMENT = {
    "void": 0x01, "bool": 0x02, "int32": 0x08, "uint32": 0x09, "int64": 0x0A,
    "string": 0x0E, "native int": 0x18, "native uint": 0x19, "object": 0x1C,
}

#: Reader-visible metadata measured from the real pythonnet 3.0.5 runtime with tools/dotnet_pe.py. The
#: generated image models these fields only; it does not contain the runtime's executable IL method bodies.
REAL_PYTHON_RUNTIME = dict(
    assembly_name="Python.Runtime",
    version=(3, 0, 5, 0),
    namespace="Python.Runtime",
    type_name="Loader",
    method_name="Initialize",
    signature="int32 (native int, int32)",
    target_framework=".NETStandard,Version=v2.0",
    references=("netstandard", "System.Reflection.Emit", "System.Reflection.Emit.ILGeneration"),
    il_only=True,
    requires_32bit=False,
    machine="x86",
    strong_name=True,
)
#: Structural values measured from the real clr_loader 0.2.10 netfx host. These flags and names model its PE
#: metadata only; build_managed_pe() does NOT synthesize the native C++/CLI implementation.
REAL_CLR_LOADER = dict(
    assembly_name="ClrLoader",
    version=(1, 0, 0, 0),
    namespace="ClrLoader",
    type_name="Host",
    method_name="Unused",
    target_framework=".NETFramework,Version=v4.7.2",
    references=("mscorlib", "System", "System.Core"),
    il_only=False,
    requires_32bit=False,
    machine="amd64",
    strong_name=False,
    exports=("pyclr_initialize", "pyclr_create_appdomain", "pyclr_get_function",
             "pyclr_close_appdomain", "pyclr_finalize"),
)


class _Heap:
    """A metadata heap: index 0 is always the empty entry, entries are 4-byte padded at the end."""

    def __init__(self) -> None:
        self._buf = bytearray(b"\x00")
        self._index: Dict[bytes, int] = {}

    def string(self, text: str) -> int:
        raw = text.encode("utf-8")
        if not raw:
            return 0
        if raw in self._index:
            return self._index[raw]
        at = len(self._buf)
        self._buf += raw + b"\x00"
        self._index[raw] = at
        return at

    def blob(self, payload: bytes) -> int:
        if not payload:
            return 0
        if payload in self._index:
            return self._index[payload]
        at = len(self._buf)
        size = len(payload)
        if size < 0x80:
            prefix = bytes([size])
        elif size < 0x4000:
            prefix = bytes([0x80 | (size >> 8), size & 0xFF])
        else:
            prefix = bytes([0xC0 | (size >> 24), (size >> 16) & 0xFF, (size >> 8) & 0xFF, size & 0xFF])
        self._buf += prefix + payload
        self._index[payload] = at
        return at

    def render(self) -> bytes:
        out = bytes(self._buf)
        return out + b"\x00" * ((4 - len(out) % 4) % 4)


def _signature_blob(signature: str, has_this: bool = False) -> bytes:
    """``"int32 (native int, int32)"`` → the MethodDef signature blob clr_loader has to agree with."""
    ret, _, rest = signature.partition("(")
    params = [p.strip() for p in rest.rstrip(")").split(",") if p.strip()]
    out = bytearray([0x20 if has_this else 0x00, len(params)])
    out.append(_ELEMENT[ret.strip()])
    for param in params:
        out.append(_ELEMENT[param])
    return bytes(out)


def _target_framework_blob(target_framework: str) -> bytes:
    """A CustomAttribute value blob: 0x0001 prolog, one fixed SerString argument, no named arguments."""
    text = target_framework.encode("utf-8")
    return b"\x01\x00" + bytes([len(text)]) + text + b"\x00"


def build_managed_pe(assembly_name: str = "Python.Runtime",
                     version: Tuple[int, int, int, int] = (3, 0, 5, 0),
                     namespace: str = "Python.Runtime",
                     type_name: str = "Loader",
                     method_name: str = "Initialize",
                     method_flags: int = METHOD_PUBLIC | METHOD_STATIC | METHOD_HIDE_BY_SIG,
                     signature: str = "int32 (native int, int32)",
                     target_framework: str = ".NETStandard,Version=v2.0",
                     references: Sequence[str] = ("netstandard",),
                     il_only: bool = True,
                     requires_32bit: bool = False,
                     machine: str = "x86",
                     strong_name: bool = False,
                     exports: Sequence[str] = (),
                     include_method: bool = True,
                     include_type: bool = True,
                     metadata_version: str = "v4.0.30319") -> bytes:
    """Emit a parseable PE/CLI metadata image whose reader-visible properties are controlled by arguments.

    The image is suitable for structural parser/validator tests, not for loading into the CLR. In particular,
    it does not emit executable IL method bodies. ``include_type=False`` / ``include_method=False`` produce an
    image that parses but does not declare the entry point the PROMPT-030 gate has to reject.
    """
    strings = _Heap()
    blobs = _Heap()

    # ---------------------------------------------------------------- #Strings / #Blob content
    s_module = strings.string("<Module>")
    s_type = strings.string(type_name)
    s_ns = strings.string(namespace)
    s_method = strings.string(method_name)
    s_attr = strings.string("TargetFrameworkAttribute")
    s_attr_ns = strings.string("System.Runtime.Versioning")
    s_ctor = strings.string(".ctor")
    s_asm = strings.string(assembly_name)

    b_method = blobs.blob(_signature_blob(signature))
    b_ctor = blobs.blob(_signature_blob("void (string)", has_this=True))
    b_attr = blobs.blob(_target_framework_blob(target_framework))
    # 8 zero bytes stand in for a real public key; the token is SHA1(...)[-8:] reversed, as .NET does.
    b_pubkey = blobs.blob(b"\x00" * 8) if strong_name else 0

    # ---------------------------------------------------------------- rows (all indexes 2 bytes: tiny image)
    ref_names = ["mscorlib"] + [name for name in references if name != "mscorlib"]
    assembly_ref_rows = [struct.pack("<HHHHIHHHH", 2, 0, 0, 0, 0, 0, strings.string(name), 0, 0)
                         for name in ref_names]
    module_row = struct.pack("<HHHHH", 0, s_module, 1, 0, 0)
    # TypeRef #1 = TargetFrameworkAttribute, scoped to AssemblyRef #1 (mscorlib).
    typeref_row = struct.pack("<HHH", (1 << 2) | 2, s_attr, s_attr_ns)
    typedef_rows = [struct.pack("<IHHHHH", 0, s_module, 0, 0, 1, 1)]
    if include_type:
        typedef_rows.append(struct.pack(
            "<IHHHHH",
            TYPE_ABSTRACT | TYPE_SEALED | TYPE_PUBLIC, s_type, s_ns, 0, 1, 1 if include_method else 2,
        ))
    methoddef_rows = []
    if include_type and include_method:
        methoddef_rows.append(struct.pack("<IHHHHH", 0, 0, method_flags, s_method, b_method, 1))
    memberref_row = struct.pack("<HHH", (1 << 3) | 1, s_ctor, b_ctor)     # parent = TypeRef #1
    customattr_row = struct.pack("<HHH", (1 << 5) | 14, (1 << 3) | 3, b_attr)   # Assembly #1, MemberRef #1
    assembly_row = struct.pack("<IHHHHIHHH", 0x8004, *version,
                               (FLAG_STRONGNAMESIGNED if strong_name else 0), b_pubkey, s_asm, 0)

    counts = {0x00: 1, 0x01: 1, 0x02: len(typedef_rows), 0x06: len(methoddef_rows),
              0x0A: 1, 0x0C: 1, 0x20: 1, 0x23: len(assembly_ref_rows)}
    tables = {
        0x00: [module_row],
        0x01: [typeref_row],
        0x02: typedef_rows,
        0x06: methoddef_rows,
        0x0A: [memberref_row],
        0x0C: [customattr_row],
        0x20: [assembly_row],
        0x23: assembly_ref_rows,
    }
    valid = 0
    for table in sorted(counts):
        valid |= 1 << table
    rows_header = b"".join(struct.pack("<I", counts[t]) for t in sorted(counts))
    tables_blob = b"".join(b"".join(tables[t]) for t in sorted(counts))
    heap_sizes = 0
    tilde = (struct.pack("<IBBBB", 0, 2, 0, heap_sizes, 1)
             + struct.pack("<Q", valid) + struct.pack("<Q", 0)
             + rows_header + tables_blob)

    guid_heap = bytes(range(16))
    stream_payloads = [(b"#~\x00", tilde), (b"#Strings\x00", strings.render()),
                       (b"#GUID\x00", guid_heap), (b"#Blob\x00", blobs.render())]

    version_bytes = metadata_version.encode("utf-8") + b"\x00"
    version_bytes += b"\x00" * ((4 - len(version_bytes) % 4) % 4)
    root_size = 16 + len(version_bytes) + 2 + 2
    for name, payload in stream_payloads:
        padded = name + b"\x00" * ((4 - len(name) % 4) % 4)
        root_size += 8 + len(padded)
    offset = root_size
    stream_headers = b""
    bodies = b""
    for name, payload in stream_payloads:
        padded = name + b"\x00" * ((4 - len(name) % 4) % 4)
        stream_headers += struct.pack("<II", offset, len(payload)) + padded
        bodies += payload + b"\x00" * ((4 - len(payload) % 4) % 4)
        offset += len(payload) + ((4 - len(payload) % 4) % 4)
    metadata = (struct.pack("<IHHI", 0x424A5342, 1, 1, 0) + struct.pack("<I", len(version_bytes))
                + version_bytes + struct.pack("<HH", 0, len(stream_payloads))
                + stream_headers + bodies)

    # ---------------------------------------------------------------- PE image
    section_alignment, file_alignment, headers_size = 0x2000, 0x200, 0x200
    base_rva = section_alignment

    cli_flags = (FLAG_ILONLY if il_only else 0) | (FLAG_32BITREQUIRED if requires_32bit else 0) \
        | (FLAG_STRONGNAMESIGNED if strong_name else 0)
    metadata_rva = base_rva + 72
    body = bytearray()
    body += struct.pack("<IHHIIII", 72, 2, 5, metadata_rva, len(metadata), cli_flags, 0)
    body += b"\x00" * (72 - len(body))                      # the remaining CLI header directories
    body += metadata

    export_rva = 0
    if exports:
        while len(body) % 4:
            body.append(0)
        export_rva = base_rva + len(body)
        names = list(exports)
        # directory(40) + function RVAs(4n) + name RVAs(4n) + ordinals(2n) + the name strings
        strings_rva = export_rva + 40 + 4 * len(names) + 4 * len(names) + 2 * len(names)
        dll_name = b"ClrLoader.dll\x00"
        name_blob = bytearray(dll_name)
        name_rvas = []
        cursor = strings_rva + len(dll_name)
        for name in names:
            name_rvas.append(cursor)
            encoded = name.encode("utf-8") + b"\x00"
            name_blob += encoded
            cursor += len(encoded)
        directory = bytearray()
        directory += struct.pack("<II", 0, 0)               # Characteristics, TimeDateStamp
        directory += struct.pack("<HH", 0, 0)               # Major/MinorVersion
        directory += struct.pack("<I", strings_rva)         # Name
        directory += struct.pack("<I", 1)                   # OrdinalBase
        directory += struct.pack("<II", len(names), len(names))
        directory += struct.pack("<I", export_rva + 40)                     # AddressOfFunctions
        directory += struct.pack("<I", export_rva + 40 + 4 * len(names))    # AddressOfNames
        directory += struct.pack("<I", export_rva + 40 + 8 * len(names))    # AddressOfNameOrdinals
        assert len(directory) == 40, len(directory)
        body += directory
        body += b"".join(struct.pack("<I", base_rva) for _ in names)        # one shared stub RVA
        body += b"".join(struct.pack("<I", rva) for rva in name_rvas)
        body += b"".join(struct.pack("<H", i) for i in range(len(names)))
        body += bytes(name_blob)

    virtual_size = len(body)
    raw = bytes(body) + b"\x00" * ((file_alignment - len(body) % file_alignment) % file_alignment)
    size_of_image = section_alignment + -(-virtual_size // section_alignment) * section_alignment

    plus = machine == "amd64"
    optional = bytearray()

    def u1(value: int) -> None:
        optional.extend(struct.pack("<B", value))

    def u2(value: int) -> None:
        optional.extend(struct.pack("<H", value))

    def u4(value: int) -> None:
        optional.extend(struct.pack("<I", value))

    def u8(value: int) -> None:
        optional.extend(struct.pack("<Q", value))

    u2(0x20B if plus else 0x10B)                            # Magic
    u1(8); u1(0)                                            # linker version
    u4(virtual_size); u4(0); u4(0)                          # SizeOfCode / initialized / uninitialized
    u4(0)                                                   # AddressOfEntryPoint
    u4(base_rva)                                            # BaseOfCode
    if not plus:
        u4(0)                                               # BaseOfData (PE32 only)
    u8(0x180000000) if plus else u4(0x400000)               # ImageBase
    u4(section_alignment); u4(file_alignment)
    u2(6); u2(0)                                            # operating system version
    u2(0); u2(0)                                            # image version
    u2(6); u2(0)                                            # subsystem version
    u4(0)                                                   # Win32VersionValue
    u4(size_of_image); u4(headers_size); u4(0)              # SizeOfImage / SizeOfHeaders / CheckSum
    u2(3); u2(0x100)                                        # Subsystem / DllCharacteristics
    for value in (0x100000, 0x1000, 0x100000, 0x1000):      # stack + heap reserve/commit
        u8(value) if plus else u4(value)
    u4(0)                                                   # LoaderFlags
    u4(16)                                                  # NumberOfRvaAndSizes
    assert len(optional) == (112 if plus else 96), len(optional)

    directories = bytearray(16 * 8)
    if exports:
        struct.pack_into("<II", directories, 0 * 8, export_rva, 40)
    struct.pack_into("<II", directories, 14 * 8, base_rva, 72)
    optional += directories

    coff = struct.pack("<HHIIIHH", _MACHINE[machine], 1, 0, 0, 0, len(optional), 0x2002)
    section = struct.pack("<8sIIIIIIHHI", b".text\x00\x00\x00", virtual_size, base_rva, len(raw),
                          headers_size, 0, 0, 0, 0, 0x60000020)
    dos = bytearray(headers_size - 4 - len(coff) - len(optional) - len(section))
    dos[0:2] = b"MZ"
    image = bytearray(dos)
    struct.pack_into("<I", image, 0x3C, len(dos))
    image += b"PE\x00\x00" + coff + bytes(optional) + section + raw
    return bytes(image)


def python_runtime_image(**overrides) -> bytes:
    """Build a structural PE/CLI fixture for Python.Runtime metadata, not a loadable runtime DLL."""
    params = dict(REAL_PYTHON_RUNTIME)
    params.update(overrides)
    return build_managed_pe(**params)


def clr_loader_image(machine: str = "amd64", **overrides) -> bytes:
    """Build a structural export-table fixture, not a callable ``ClrLoader.dll``."""
    params = dict(REAL_CLR_LOADER)
    params.update(overrides)
    params["machine"] = machine
    if machine == "x86":
        params["requires_32bit"] = overrides.get("requires_32bit", True)
    return build_managed_pe(**params)
