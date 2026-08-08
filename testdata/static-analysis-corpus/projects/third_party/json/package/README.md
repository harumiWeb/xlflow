# Package

This directory contains the production VBA class for the project.

## Contents

* [JSON.cls](JSON.cls): The complete standalone JSON parser and writer. It includes zero-copy parsing, typed accessors, token iteration, raw field access, and Stringify support.

## Importing

1. Open your Excel, PowerPoint, Word, or Access file.
2. Press `Alt + F11` to open the VBA editor.
3. Choose **File > Import File...**.
4. Select `JSON.cls`.
5. Save your document as a macro-enabled file such as `.xlsm`, `.pptm`, `.docm`, or `.accdb`.

## Runtime Notes

* `JSON.cls` is a predeclared class, so factory-style calls such as `Set doc = JSON.Parse(text)` work after import.
* No external references are required for parsing, traversal, raw access, or writing.
* The class supports 32-bit and 64-bit Office through conditional declarations in `package/JSON.cls`.
* Child nodes returned by `Node`, `NodeAt`, `TokenNode`, and `NodeFromToken` depend on the root parsed document, so keep the root `JSON` variable alive while using them.
