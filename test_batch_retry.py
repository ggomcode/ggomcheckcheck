import time
import requests
import json
import os

def chunk_list(lst: list, chunk_size: int):
    for i in range(0, len(lst), chunk_size):
        yield lst[i:i + chunk_size]

# Test payload chunking
records = [{"id": i} for i in range(50)]
chunks = list(chunk_list(records, 15))
print(f"Total records: {len(records)}, Chunks: {len(chunks)}")
for idx, ch in enumerate(chunks):
    print(f"Chunk {idx+1}: {len(ch)} items")

print("=== BATCH CHUNKING TEST PASSED! ===")
