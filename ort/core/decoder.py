"""
core/decoder.py - Decoding utility functions
role: Module for decoding local files extracted from OOXML containers

author: Jiyoon Kim
date: 2025-05-06

description:
    decode_utf8() - Decodes a hex string into UTF-8 string
    decompress_deflate_hex() - Attempts to decompress hex string using DEFLATE (handles corrupt segments)
    decode_local_file_data() - Applies decompression and decoding to .xml/.rels local files
"""
import binascii
import zlib

def decode_utf8(hex_str):
    """Decode a hex string as UTF-8, ignoring decoding errors"""
    try:
        raw_bytes = binascii.unhexlify(hex_str)
        return raw_bytes.decode('utf-8', 'ignore')
    except Exception:
        return ""

def decompress_deflate_hex(hex_string):
    """Decompress a hex string assumed to be compressed with raw DEFLATE (no headers)"""
    compressed_data = bytes.fromhex(hex_string)
    decompressor = zlib.decompressobj(-zlib.MAX_WBITS)

    try:
        return decompressor.decompress(compressed_data).decode("utf-8")
    except UnicodeDecodeError:
        for i in range(1, len(compressed_data)):
            try:
                data = zlib.decompress(compressed_data[:-i], -zlib.MAX_WBITS)
                return data.decode("utf-8")
            except:
                continue
    return ""

def decode_local_file_data(json_file_list):
    """Decompress and decode all local file data entries, including media."""
    for item in json_file_list:
        name = item.get('local_file_name', '')
        hex_data = item.get('local_file_data', '')

        if not hex_data:
            continue

        try:
            # Try to decompress raw DEFLATE data
            decompressed = decompress_deflate_hex(hex_data)
            item['local_file_data'] = decompressed
        except zlib.error:
            # If decompression fails for XML, try trimming
            if name.endswith('.xml') or name.endswith('.rels'):
                for i in range(2, len(hex_data), 2):
                    try:
                        decompressed = decompress_deflate_hex(hex_data[:-i])
                        item['local_file_data'] = decompressed
                        break
                    except zlib.error:
                        continue
            else:
                # 🛑 media 파일 등은 원래 binary로 두기
                pass
        except UnicodeDecodeError:
            if name.endswith('.xml') or name.endswith('.rels'):
                for i in range(2, len(hex_data), 2):
                    try:
                        item['local_file_data'] = decode_utf8(hex_data[:-i])
                        break
                    except UnicodeDecodeError:
                        continue

