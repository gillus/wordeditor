"""Check two .docx files are equivalent: same parts, and canonically identical XML. Usage: verify_same.py a.docx b.docx"""
import sys
import zipfile
from xml.etree.ElementTree import canonicalize


def main(a, b):
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    assert sorted(za.namelist()) == sorted(zb.namelist()), "package parts differ"
    for name in za.namelist():
        da, db = za.read(name), zb.read(name)
        if da == db:
            continue
        assert name.endswith(".xml") or name.endswith(".rels"), f"binary part changed: {name}"
        ca, cb = canonicalize(da.decode("utf-8")), canonicalize(db.decode("utf-8"))
        if ca != cb:
            i = next(i for i, (x, y) in enumerate(zip(ca, cb)) if x != y)
            raise AssertionError(f"{name} differs at {i}:\n  {ca[i-150:i+150]}\n  {cb[i-150:i+150]}")
    print("verify_same OK")


if __name__ == "__main__":
    main(*sys.argv[1:])
