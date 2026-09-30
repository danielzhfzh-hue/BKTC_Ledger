#!/usr/bin/env python3
"""Write the Release sidecar consumed by the in-app update verifier."""
import hashlib
import sys
from pathlib import Path


for argument in sys.argv[1:]:
    path = Path(argument)
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    path.with_name(path.name + ".sha256").write_text(
        f"{digest}  {path.name}\n", encoding="ascii"
    )
