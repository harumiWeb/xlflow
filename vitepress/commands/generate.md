# xlflow generate

Generate project artifacts such as test modules.

## xlflow generate test

Create a new VBA test module under the configured module source directory.

### Usage

```bash
xlflow generate test <module-name> [--json]
```

### Arguments

| Argument      | Description                                     |
| ------------- | ----------------------------------------------- |
| `module-name` | Name of the test module (without `.bas` suffix) |

### Description

`generate test` scaffolds a new standard module with:

- `Attribute VB_Name` and `Option Explicit`
- Lifecycle hook stubs: `BeforeAll`, `AfterAll`, `BeforeEach`, `AfterEach`
- A sample test sub to copy and adapt

The file is written to the configured `[src].modules` directory. The command refuses to overwrite an existing file.

### Example

```bash
xlflow generate test OrderServiceTests
```

This creates `src/modules/OrderServiceTests.bas` with the following content:

```vb
Attribute VB_Name = "OrderServiceTests"
Option Explicit

Public Sub BeforeAll()
End Sub

Public Sub AfterAll()
End Sub

Public Sub BeforeEach()
End Sub

Public Sub AfterEach()
End Sub

Public Sub Test_Sample()
    XlflowAssert.AssertTrue True, "replace with real assertions"
End Sub
```

## Related

- [test](./test)
- [module install](./module-install)
