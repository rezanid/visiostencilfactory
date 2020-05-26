# Dot source the New-VisioStencil cmdlet
. .\Visio\New-VisioStencil.ps1

# Generate stencil
New-VisioStencil (Get-ChildItem "*.svg") -StencilPath "Stencil1.vssx"