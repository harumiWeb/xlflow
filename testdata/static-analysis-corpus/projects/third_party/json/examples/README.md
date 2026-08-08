# Examples

This directory contains importable VBA modules that demonstrate the public API from [package/JSON.cls](../package/JSON.cls).

## Modules

* [BasicRead.bas](BasicRead.bas): Parses JSON, reads object fields, traverses nested nodes, and reads arrays by index.
* [TokenIteration.bas](TokenIteration.bas): Iterates arrays of objects with token helpers for high-volume reads.
* [StringifyValues.bas](StringifyValues.bas): Serializes Dictionaries, Collections, arrays, and parsed JSON nodes.

## Using the Examples

1. Import [package/JSON.cls](../package/JSON.cls) into your VBA project.
2. Import one or more `.bas` files from this directory.
3. Run the public procedures from the VBA editor.

The examples use `Debug.Print`, so output appears in the Immediate Window. Open it with `Ctrl + G` in the VBA editor.
