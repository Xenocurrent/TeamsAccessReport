@{
    Severity     = @('Error', 'Warning', 'Information')
    ExcludeRules = @(
        # Positionella parametrar tillåts för interna enradiga hjälpanrop (Get-TarProperty $x 'Name').
        'PSAvoidUsingPositionalParameters'
    )
    Rules        = @{
        PSUseCompatibleSyntax      = @{ Enable = $true; TargetVersions = @('7.2') }
        PSUseCompatibleCmdlets     = @{ Compatibility = @('core-6.1.0-windows', 'core-6.1.0-linux') }
        PSPlaceOpenBrace           = @{ Enable = $true; OnSameLine = $true; NewLineAfter = $true; IgnoreOneLineBlock = $true }
        PSPlaceCloseBrace          = @{ Enable = $true; NewLineAfter = $false; IgnoreOneLineBlock = $true; NoEmptyLineBefore = $false }
        PSUseConsistentIndentation = @{ Enable = $true; IndentationSize = 4; Kind = 'space'; PipelineIndentation = 'IncreaseIndentationForFirstPipeline' }
        PSUseConsistentWhitespace  = @{ Enable = $true; CheckOpenBrace = $true; CheckOpenParen = $true; CheckOperator = $false; CheckSeparator = $true; CheckInnerBrace = $true; CheckPipe = $true }
        PSAvoidUsingCmdletAliases  = @{ Enable = $true }
    }
}
