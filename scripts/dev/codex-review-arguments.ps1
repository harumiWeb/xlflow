Set-StrictMode -Version Latest

function New-BackslashSequence {
    param(
        [int]$Count
    )

    if ($Count -le 0) {
        return ""
    }

    return (1..$Count | ForEach-Object { '\' }) -join ""
}

function ConvertTo-ProcessArgument {
    param(
        [AllowEmptyString()]
        [string]$Value,
        [switch]$ForceQuotes
    )

    if ($null -eq $Value -or $Value.Length -eq 0) {
        return '""'
    }

    if (-not $ForceQuotes -and $Value -notmatch '[\s"]') {
        return $Value
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append([char]34)
    $backslashCount = 0

    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq [char]92) {
            $backslashCount++
            continue
        }

        if ($character -eq [char]34) {
            [void]$builder.Append(
                (New-BackslashSequence -Count ($backslashCount * 2 + 1))
            )
            [void]$builder.Append([char]34)
            $backslashCount = 0
            continue
        }

        [void]$builder.Append(
            (New-BackslashSequence -Count $backslashCount)
        )
        [void]$builder.Append($character)
        $backslashCount = 0
    }

    [void]$builder.Append(
        (New-BackslashSequence -Count ($backslashCount * 2))
    )
    [void]$builder.Append([char]34)

    return $builder.ToString()
}

function ConvertTo-CmdProcessArgument {
    param(
        [AllowEmptyString()]
        [string]$Value,
        [Parameter(Mandatory)]
        [string]$PercentVariableName
    )

    if ($null -eq $Value -or $Value.Length -eq 0) {
        return '""'
    }

    # cmd.exe expands percent variables before it honors quotes. Replace each
    # literal percent with a process environment token whose value is '%'; the
    # token is expanded once and the resulting percent is then passed through
    # the batch shim without a second variable expansion.
    $percentToken = "%{0}%" -f $PercentVariableName
    $encodedValue = $Value.Replace('%', $percentToken)
    $requiresQuotes = $encodedValue -match '[\s"&%!|<>()^]'

    return ConvertTo-ProcessArgument `
        -Value $encodedValue `
        -ForceQuotes:$requiresQuotes
}
