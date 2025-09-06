Describe "Testing public module function Get-ISOWeekNumber" -Tag "UnitTest" {
    Context "General tests for ISO Week Number calculation" {
        BeforeAll {
            $ModuleRoot = $PSCommandPath | Split-Path -Parent | Split-Path -Parent | Split-Path -Parent
            $ScriptName = $PSCommandPath | Split-Path -Leaf
            $Visibility = $PSCommandPath | Split-Path -Parent | Split-Path -Leaf
            $SourceDirectory = Resolve-Path (Join-Path $ModuleRoot "source\$Visibility")
            $TestDirectory = Resolve-Path (Join-Path $ModuleRoot "test\$Visibility")

            $FunctionPath = Join-Path $SourceDirectory ($ScriptName -replace '\.Tests\.ps1$', '.ps1')
    
            # Create a Stub for the module function to test
            Function Get-ISOWeekNumber {
                . $FunctionPath @args | write-Output
            }
        }

        It "Should return the expected ISO Week number for 5 august 1978" {
            $Result = Get-ISOWeekNumber -Date (Get-Date -Year 1978 -Month 8 -Day 5 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be 31
        }

        It "Should return the expected ISO Week number for 5 January 1979" {
            $Result = Get-ISOWeekNumber -Date (Get-Date -Year 1979 -Month 1 -Day 5 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be 1
        }

        It "Should return the expected ISO Week number for 27 August 2008" {
            $Result = Get-ISOWeekNumber -Date (Get-Date -Year 2008 -Month 8 -Day 27 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be 35
        }

        It "Should return the expected ISO Week number for 27 January 2014" {
            $Result = Get-ISOWeekNumber -Date (Get-Date -Year 2014 -Month 1 -Day 27 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be 5
        }
    }
}
