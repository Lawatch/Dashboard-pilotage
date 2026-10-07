"""MS-OVBA 2.4.1 compression (CompressedContainer)."""
import struct

def _copy_token_help(decompressed_current, decompressed_chunk_start):
    difference = decompressed_current - decompressed_chunk_start
    bit_count = max((difference - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    offset_mask = (~length_mask) & 0xFFFF
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length

def _matching(data, chunk_start, cur, end):
    best_len = 0; best_off = 0
    _, _, _, max_len = _copy_token_help(cur, chunk_start)
    limit = min(max_len, end - cur)
    if limit < 3:
        return 0, 0
    # search backwards
    first = data[cur]
    c = cur - 1
    while c >= chunk_start:
        if data[c] == first:
            l = 1
            while l < limit and data[c + l] == data[cur + l]:
                l += 1
            if l > best_len:
                best_len = l; best_off = cur - c
                if l == limit:
                    break
        c -= 1
    if best_len >= 3:
        return best_off, best_len
    return 0, 0

def _compress_chunk(data, start, end):
    out = bytearray()
    cur = start
    while cur < end:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if cur >= end:
                break
            off, ln = _matching(data, start, cur, end)
            if ln:
                length_mask, offset_mask, bit_count, _ = _copy_token_help(cur, start)
                token = ((off - 1) << (16 - bit_count)) | (ln - 3)
                out += struct.pack('<H', token)
                flags |= 1 << bit
                cur += ln
            else:
                out.append(data[cur])
                cur += 1
        out[flag_pos] = flags
    return out

def compress(data: bytes) -> bytes:
    data = bytes(data)
    res = bytearray(b'\x01')
    pos = 0
    n = len(data)
    while pos < n:
        end = min(pos + 4096, n)
        body = _compress_chunk(data, pos, end)
        if len(body) > 4096:
            if end - pos != 4096:
                raise ValueError('incompressible partial chunk')
            # uncompressed chunk (exactly 4096 bytes)
            raw = data[pos:end]
            header = 0x3000 | (4096 - 1)  # flag 0, signature 0b011
            res += struct.pack('<H', header) + raw
        else:
            header = 0xB000 | (len(body) + 2 - 3)  # flag 1, signature 0b011
            res += struct.pack('<H', header) + body
        pos = end
    return bytes(res)
