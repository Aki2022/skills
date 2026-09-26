#!/usr/bin/env python3
"""Patch a cloned app bundle's display name, identifier, and URL schemes."""

import plistlib
import sys


def main():
    if len(sys.argv) != 4:
        print(
            "Usage: patch_plist.py <plist> <display-name> <bundle-identifier>",
            file=sys.stderr,
        )
        return 64

    path, display, ident = sys.argv[1:]
    with open(path, "rb") as fh:
        info = plistlib.load(fh)

    info["CFBundleDisplayName"] = display
    info["CFBundleIdentifier"] = ident

    url_types = []
    for entry in info.get("CFBundleURLTypes", []):
        schemes = []
        for scheme in entry.get("CFBundleURLSchemes", []):
            if scheme in ("http", "https"):
                continue
            suffix = ".seat2" if scheme.startswith("msauth.") else "-seat2"
            schemes.append(scheme + suffix)
        if not schemes:
            continue
        entry["CFBundleURLSchemes"] = schemes
        url_types.append(entry)

    if url_types:
        info["CFBundleURLTypes"] = url_types
    elif "CFBundleURLTypes" in info:
        del info["CFBundleURLTypes"]

    with open(path, "wb") as fh:
        plistlib.dump(info, fh)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
