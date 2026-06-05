Attribute VB_Name = "SampleTests"
'@Folder("Tests")
Option Explicit

' xlflow tests are public parameterless Sub procedures whose names match
' Test* or *_Test.  xlflow discovers them automatically at run time.
'
' Use XlflowAssert helpers to raise clear, JSON-friendly failures.
'
' Tags: add '@Tag("name")' comment lines directly above a test sub
' and run only matching tests with  xlflow test --tag <name>.
'
' Hooks: BeforeAll / AfterAll / BeforeEach / AfterEach are optional
' reserved names.  They must be public parameterless Subs and they
' affect only tests in the same module.
'
' Keep tests independent.  Use BeforeEach / AfterEach for isolation
' and BeforeAll for expensive one-time setup.

Public Sub BeforeAll()
    ' Runs once before the first test in this module.
End Sub

Public Sub AfterAll()
    ' Runs once after the last test in this module.
End Sub

Public Sub BeforeEach()
    ' Runs before every test in this module.
End Sub

Public Sub AfterEach()
    ' Runs after every test in this module.
End Sub

'@Tag("smoke")
Public Sub Test_Sample_Pass()
    ' A passing test demonstrates the AssertEquals helper.
    XlflowAssert.AssertEquals 1 + 1, 2, "basic arithmetic should work"
End Sub
