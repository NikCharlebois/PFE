[CmdletBinding()]
param(
    [Parameter()]
    [string]
    $SharePointStubsModule = (Join-Path -Path $PSScriptRoot `
                                         -ChildPath "..\Stubs\SharePoint\15.0.4805.1000\Microsoft.SharePoint.PowerShell.psm1" `
                                         -Resolve)
)

Import-Module -Name (Join-Path -Path $PSScriptRoot `
                                -ChildPath "..\UnitTestHelper.psm1" `
                                -Resolve)

$Global:TestHelper = New-UnitTestHelper -SharePointStubModule $SharePointStubsModule `
                                              -Cmdlet "Get-RegistryCredential"

Describe -Name $Global:TestHelper.DescribeHeader -Fixture {

    InModuleScope -ModuleName $Global:TestHelper.ModuleName -ScriptBlock {
        Invoke-Command -ScriptBlock $Global:TestHelper.InitializeScript -NoNewScope

        $mockPassword = ConvertTo-SecureString -String "password" -AsPlainText -Force
        $mockCredential = New-Object -TypeName System.Management.Automation.PSCredential `
                                     -ArgumentList @("DOMAIN\username", $mockPassword)

        Mock -CommandName Get-ItemProperty -MockWith {
            return @{
                UserName = "contoso\john.smith"
                Password = "01000000d08c9ddf0115d1118c7a00c04fc297eb01000000f74bf2f14b6cdf4586266f16f1dfac130000000002000000000010660000000100002000000074cd73f55cadf2725bcb64431d9edef3035d12505cdbf5d93501a128e1c06385000000000e8000000002000020000000c91bd5204db65776114afa3441a6d91b8bb10e109357a36a08ceb4c2ae7d795310000000c689444aecd318866642e4d767f3b37f4000000024d0de35fd828a0b9e0a8ce4d4717a6bc58103287ea0b19b668cb04de85ba3a7e412fa038452bea13b994db890a7511befbf4b7b738f91822dab8aa3869a4bed"
            }
        }

        Mock -CommandName ConvertTo-SecureString -MockWith {
            return $mockPassword
        }

        Context -Name "When the credential object is found" -Fixture {
            Mock -CommandName CheckForExistingRegistryCredential -MockWith {
                return $true
            }
            $testParams = @{
                ApplicationName = "MyTestApplication"
                OrgName = "Contoso"
                AccountDescription = "JohnSmith"
            }

            It "Should return contoso\john.smith" {
                (Get-RegistryCredential @testParams).UserName | Should Be "contoso\john.smith"
            }
        }

        Context -Name "When the credential don't exist" -Fixture {
            Mock -CommandName CheckForExistingRegistryCredential -MockWith {
                return $false
            }
            $testParams = @{
                ApplicationName = "MyTestApplication"
                OrgName = "Contoso"
                AccountDescription = "JohnSmith"
            }

            It "Should return contoso\john.smith" {
                { Get-RegistryCredential @testParams } | Should Throw "Could not locate credential object at 'HKCU:\Software\MyTestApplication\Contoso\Credentials\JohnSmith'"
            }
        }
    }
}

Invoke-Command -ScriptBlock $Global:TestHelper.CleanupScript -NoNewScope
