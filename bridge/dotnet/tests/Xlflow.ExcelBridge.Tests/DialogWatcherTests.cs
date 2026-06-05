using Xlflow.ExcelBridge.Windows;

namespace Xlflow.ExcelBridge.Tests;

public sealed class DialogWatcherTests
{
    [Fact]
    public void ClassifyRuntimeDialogFromLocalizedText()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic",
            text: ["実行時エラー '5':", "bridge runtime boom"],
            buttons: [Button(11, "終了"), Button(12, "デバッグ")]);

        Assert.Equal(DialogKind.Runtime, DialogFingerprint.Classify(candidate, null));
    }

    [Fact]
    public void ClassifyCompileDialogFromCompileMessage()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic",
            text: ["Compile error:", "Sub or Function not defined"],
            buttons: [Button(11, "OK")]);

        Assert.Equal(DialogKind.Compile, DialogFingerprint.Classify(candidate, null));
    }

    [Fact]
    public void ClassifyLocalizedCompileDialogFromSyntaxErrorText()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic for Applications",
            text: ["コンパイル エラー:", "構文エラー"],
            buttons: [Button(11, "OK"), Button(12, "ヘルプ")]);

        Assert.Equal(DialogKind.Compile, DialogFingerprint.Classify(candidate, null));
    }

    [Fact]
    public void ClassifyInputBoxFromEditChild()
    {
        var candidate = Candidate(
            title: "Input",
            children: [new WindowElement(20, "Edit", "")],
            buttons: [Button(11, "OK"), Button(12, "Cancel")]);

        Assert.Equal(DialogKind.InputBox, DialogFingerprint.Classify(candidate, null));
    }

    [Fact]
    public void CompileActionDoesNotUseFirstButtonFallback()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic",
            text: ["Compile error"],
            buttons: [Button(11, "Help")]);

        var action = DialogActionSelector.Select(DialogKind.Compile, candidate, DialogActionPolicy.SuppressVbaError);

        Assert.Equal(DialogAction.None, action);
    }

    [Fact]
    public void CompileActionClosesLocalizedOkButton()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic for Applications",
            text: ["コンパイル エラー:", "構文エラー"],
            buttons: [Button(11, "OK"), Button(12, "ヘルプ")]);

        var action = DialogActionSelector.Select(DialogKind.Compile, candidate, DialogActionPolicy.SuppressVbaError);

        Assert.Equal("compile_close", action.Name);
        Assert.Equal(11, action.TargetHwnd);
    }

    [Fact]
    public void RuntimeActionPrefersEndToAvoidLeavingVbeInBreakMode()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic",
            text: ["Run-time error '5'"],
            buttons: [Button(11, "End"), Button(12, "Debug")]);

        var action = DialogActionSelector.Select(DialogKind.Runtime, candidate, DialogActionPolicy.SuppressVbaError);

        Assert.Equal("runtime_end", action.Name);
        Assert.Equal(11, action.TargetHwnd);
    }

    [Fact]
    public void RuntimeActionMatchesLocalizedAcceleratorSuffix()
    {
        var candidate = Candidate(
            title: "Microsoft Visual Basic",
            text: ["実行時エラー '5'"],
            buttons: [Button(11, "終了(&E)"), Button(12, "デバッグ(&D)")]);

        var action = DialogActionSelector.Select(DialogKind.Runtime, candidate, DialogActionPolicy.SuppressVbaError);

        Assert.Equal("runtime_end", action.Name);
        Assert.Equal(11, action.TargetHwnd);
    }

    [Fact]
    public void SuppressPolicyCancelsOnlyExplicitNativeCancel()
    {
        var candidate = Candidate(
            title: "Input",
            children: [new WindowElement(20, "Edit", "")],
            buttons: [Button(11, "OK"), Button(12, "キャンセル(&C)")]);

        var action = DialogActionSelector.Select(DialogKind.InputBox, candidate, DialogActionPolicy.SuppressVbaError);

        Assert.Equal("native_cancel", action.Name);
        Assert.Equal(12, action.TargetHwnd);
    }

    [Fact]
    public void NativeDialogActionUsesOnlyExplicitCancel()
    {
        var candidate = Candidate(
            title: "Message",
            buttons: [Button(11, "Yes"), Button(12, "No")]);

        var action = DialogActionSelector.Select(DialogKind.MsgBox, candidate, DialogActionPolicy.CancelSupportedNativeUi);

        Assert.Equal(DialogAction.None, action);
    }

    private static WindowCandidate Candidate(
        string title,
        string className = "#32770",
        string[]? text = null,
        WindowElement[]? buttons = null,
        WindowElement[]? children = null)
    {
        return new WindowCandidate(
            Hwnd: 1,
            Pid: 100,
            ThreadId: 200,
            OwnerHwnd: 2,
            RootOwnerHwnd: 3,
            Title: title,
            ClassName: className,
            Visible: true,
            ProcessImage: "EXCEL.EXE",
            OwnerChain: [2, 3],
            Text: text ?? [],
            Buttons: buttons ?? [],
            Children: children ?? buttons ?? []);
    }

    private static WindowElement Button(long hwnd, string text)
    {
        return new WindowElement(hwnd, "Button", text);
    }
}
