param(
    [string]$SvgFilePath='E:\Temp\Visio\icon-collection-master\azure-patterns\*.svg',
    [string]$StencilName='azure-patterns.vssx'
)

# Dot source the New-VisioStencil cmdlet
. $PSScriptRoot\Visio\New-VisioStencil.ps1

# Declare a name extractpr command
$nameExtractor = {param($name) ($name.Replace('-',' ')) }

# Generate stencil
New-VisioStencil (Get-ChildItem $SvgFilePath) -StencilPath $StencilName -MasterNameExtractor $nameExtractor