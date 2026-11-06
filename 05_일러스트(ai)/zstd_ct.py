"""Git for Windows에 들어 있는 libzstd.dll을 ctypes로 써서 zstd 압축을 풀고 다시 압축함 (파이썬 패키지 설치 없이)."""
import ctypes, os

_CANDIDATES = [os.environ.get('LIBZSTD_DIR', ''),
               r'C:\Program Files\Git\mingw64\bin',
               r'C:\Program Files (x86)\Git\mingw64\bin',
               os.path.expandvars(r'%LOCALAPPDATA%\Programs\Git\mingw64\bin')]
_DIR = next((d for d in _CANDIDATES if d and os.path.exists(os.path.join(d, 'libzstd.dll'))), None)
if _DIR is None:
    raise ImportError('libzstd.dll을 못 찾음 — Git for Windows를 설치하거나 LIBZSTD_DIR 환경변수에 폴더를 넣으세요')
os.add_dll_directory(_DIR)
_z = ctypes.CDLL(os.path.join(_DIR, 'libzstd.dll'))


class _In(ctypes.Structure):
    _fields_ = [('src', ctypes.c_void_p), ('size', ctypes.c_size_t), ('pos', ctypes.c_size_t)]


class _Out(ctypes.Structure):
    _fields_ = [('dst', ctypes.c_void_p), ('size', ctypes.c_size_t), ('pos', ctypes.c_size_t)]


_z.ZSTD_createDCtx.restype = ctypes.c_void_p
_z.ZSTD_freeDCtx.argtypes = [ctypes.c_void_p]
_z.ZSTD_decompressStream.restype = ctypes.c_size_t
_z.ZSTD_decompressStream.argtypes = [ctypes.c_void_p, ctypes.POINTER(_Out), ctypes.POINTER(_In)]
_z.ZSTD_isError.restype = ctypes.c_uint
_z.ZSTD_isError.argtypes = [ctypes.c_size_t]
_z.ZSTD_getErrorName.restype = ctypes.c_char_p
_z.ZSTD_getErrorName.argtypes = [ctypes.c_size_t]
_z.ZSTD_compressBound.restype = ctypes.c_size_t
_z.ZSTD_compressBound.argtypes = [ctypes.c_size_t]
_z.ZSTD_compress.restype = ctypes.c_size_t
_z.ZSTD_compress.argtypes = [ctypes.c_void_p, ctypes.c_size_t, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_int]


def _check(code):
    if _z.ZSTD_isError(code):
        raise RuntimeError(_z.ZSTD_getErrorName(code).decode())
    return code


def decompress(data: bytes):
    """프레임이 끝난 뒤 남는 바이트 수도 함께 돌려줌 (뒤에 다른 자료가 붙어 있는지 확인용)."""
    dctx = _z.ZSTD_createDCtx()
    src = ctypes.create_string_buffer(data, len(data))
    inb = _In(ctypes.cast(src, ctypes.c_void_p), len(data), 0)
    chunk = 1 << 24
    buf = ctypes.create_string_buffer(chunk)
    parts = []
    try:
        while True:
            out = _Out(ctypes.cast(buf, ctypes.c_void_p), chunk, 0)
            ret = _check(_z.ZSTD_decompressStream(dctx, ctypes.byref(out), ctypes.byref(inb)))
            parts.append(buf.raw[:out.pos])
            if ret == 0 and out.pos < chunk:
                break          # 프레임 하나 끝
            if inb.pos >= inb.size and out.pos == 0:
                raise RuntimeError('입력이 끝났는데 프레임이 안 끝남')
    finally:
        _z.ZSTD_freeDCtx(dctx)
    return b''.join(parts), len(data) - inb.pos


def compress(data: bytes, level: int = 3) -> bytes:
    cap = _z.ZSTD_compressBound(len(data))
    dst = ctypes.create_string_buffer(cap)
    src = ctypes.create_string_buffer(data, len(data))
    n = _check(_z.ZSTD_compress(dst, cap, src, len(data), level))
    return dst.raw[:n]
