"""One-time checksum-verified import of the delivered Gridline source."""
import base64
import hashlib
import json
import lzma
from pathlib import Path, PurePosixPath

root = Path(__file__).resolve().parents[1]
encoded = ''.join((root / '.bootstrap' / f'part-{i:02d}.b64').read_text(encoding='ascii') for i in range(14))
archive = base64.b64decode(encoded, validate=True)
expected = '66a7541fd0d33fa98dcf7f87a7e34a061a574e1544a7b73f163dc5b63a4b89aa'
if hashlib.sha256(archive).hexdigest() != expected:
    raise RuntimeError('Source archive SHA-256 mismatch; refusing to import')
decoder = lzma.LZMADecompressor(memlimit=256 * 1024 * 1024)
decoded = decoder.decompress(archive, max_length=8 * 1024 * 1024)
if not decoder.eof or decoder.unused_data:
    raise RuntimeError('Invalid or oversized source archive')
files = json.loads(decoded)
if not isinstance(files, dict) or len(files) != 22:
    raise RuntimeError('Unexpected source manifest')
for name, content in files.items():
    path = PurePosixPath(name)
    if path.is_absolute() or '..' in path.parts or not path.parts or path.parts[0] in {'.git', '.github', '.bootstrap'}:
        raise RuntimeError(f'Unsafe import path: {name}')
    if not isinstance(content, str):
        raise RuntimeError(f'Non-text source entry: {name}')
    target = root.joinpath(*path.parts)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding='utf-8')
print(f'Imported {len(files)} source files; archive SHA-256 verified: {expected}')
