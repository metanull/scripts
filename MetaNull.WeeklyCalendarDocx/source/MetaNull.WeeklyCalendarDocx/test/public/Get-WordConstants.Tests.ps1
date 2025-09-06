Describe "Testing public module function Get-WordConstants" -Tag "UnitTest" {
    Context "General tests for Word Constants retrieval" {
        BeforeAll {
            $ModuleRoot = $PSCommandPath | Split-Path -Parent | Split-Path -Parent | Split-Path -Parent
            $ScriptName = $PSCommandPath | Split-Path -Leaf
            $Visibility = $PSCommandPath | Split-Path -Parent | Split-Path -Leaf
            $SourceDirectory = Resolve-Path (Join-Path $ModuleRoot "source\$Visibility")
            $TestDirectory = Resolve-Path (Join-Path $ModuleRoot "test\$Visibility")

            $FunctionPath = Join-Path $SourceDirectory ($ScriptName -replace '\.Tests\.ps1$', '.ps1')
    
            # Create a Stub for the module function to test
            Function Get-WordConstants {
                . $FunctionPath @args | write-Output
            }
        }

        It "Should return a hashtable of all constants with no parameters" {
            $Result = Get-WordConstants
            $Result | Should -BeOfType 'System.Collections.Hashtable'
            $Result.Count | Should -BeGreaterThan 0
            $Result.ContainsKey('WD_LINE_STYLE_NONE') | Should -Be $true
        }

        It "Should return a hashtable of all constants with -GetAll" {
            $Result = Get-WordConstants -All
            $Result | Should -BeOfType 'System.Collections.Hashtable'
            $Result.Count | Should -BeGreaterThan 0
            $Result.ContainsKey('WD_LINE_STYLE_NONE') | Should -Be $true
        }

        It "Should return the correct value for a known constant with -ConstantName" {
            $Result = Get-WordConstants -ConstantName 'WD_LINE_STYLE_NONE'
            $Result | Should -Be 0
        }

        It "Should return a list of constant names with -ListAvailable" {
            $Result = Get-WordConstants -ListAvailable
            $Result.GetType() | Should -Be 'System.Object[]'
            $Result.Count | Should -BeGreaterThan 0
            $Result | Should -Contain 'WD_LINE_STYLE_NONE'
        }
    }
}
