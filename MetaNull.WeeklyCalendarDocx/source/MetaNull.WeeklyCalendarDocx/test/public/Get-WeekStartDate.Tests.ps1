Describe "Testing public module function Get-WeekStartDate" -Tag "UnitTest" {
    Context "General tests for Week Start Date calculation" {
        BeforeAll {
            $ModuleRoot = $PSCommandPath | Split-Path -Parent | Split-Path -Parent | Split-Path -Parent
            $ScriptName = $PSCommandPath | Split-Path -Leaf
            $Visibility = $PSCommandPath | Split-Path -Parent | Split-Path -Leaf
            $SourceDirectory = Resolve-Path (Join-Path $ModuleRoot "source\$Visibility")
            $TestDirectory = Resolve-Path (Join-Path $ModuleRoot "test\$Visibility")

            $FunctionPath = Join-Path $SourceDirectory ($ScriptName -replace '\.Tests\.ps1$', '.ps1')
    
            # Create a Stub for the module function to test
            Function Get-WeekStartDate {
                . $FunctionPath @args | write-Output
            }
        }

        It "Should return the expected Week Start Date for 5 august 1978" {
            $Result = Get-WeekStartDate -Date (Get-Date -Year 1978 -Month 8 -Day 5 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be (Get-Date -Year 1978 -Month 7 -Day 31 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
        }

        It "Should return the expected Week Start Date for 5 January 1979" {
            $Result = Get-WeekStartDate -Date (Get-Date -Year 1979 -Month 1 -Day 5 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be (Get-Date -Year 1979 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
        }

        It "Should return the expected Week Start Date for 27 August 2008" {
            $Result = Get-WeekStartDate -Date (Get-Date -Year 2008 -Month 8 -Day 27 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be (Get-Date -Year 2008 -Month 8 -Day 25 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
        }

        It "Should return the expected Week Start Date for 27 January 2014" {
            $Result = Get-WeekStartDate -Date (Get-Date -Year 2014 -Month 1 -Day 27 -Hour 5 -Minute 0 -Second 0 -Millisecond 0)
            $Result | Should -Be (Get-Date -Year 2014 -Month 1 -Day 27 -Hour 0 -Minute 0 -Second 0 -Millisecond 0)
        }
    }
}
