import sys
import os

sys.path.insert(0, r"c:\Users\lichao\Office-Word-MCP-Server")
print("System path:\n", sys.path)
print("Current working directory:\n", os.getcwd())

try:
    import word_document_server
    print("word_document_server package found at:\n", word_document_server.__file__)
    from word_document_server.com_backend import WordBackend
    print("COMBackend imported successfully")
except ImportError as e:
    print("Import error:\n", e)
    sys.exit(1)

import json

try:
    with WordBackend(file_path=r"c:\Users\lichao\Office-Word-MCP-Server\tests\test_docs\test_document.docx") as backend:
        status = backend.get_protection_status()
        print(json.dumps(status))
except Exception as e:
    print("Runtime error:\n", e)
    sys.exit(1)