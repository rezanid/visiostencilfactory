. .\Visio\New-VisioStencil.ps1
$svgFilePath = "E:\Temp\Visio\icon-collection-master\azure-icons\*.svg" 
$nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
$groupsOfSvgFiles = Get-ChildItem $svgFilePath |
    Group-Object -Property @{
        Expression = {$_.BaseName.Substring(0,$_.BaseName.IndexOf('-'))}
    }
$groupsOfSvgFiles |
    ForEach-Object $_ -Begin {
        $i = 0
        Write-Host "Stencils to be created: $($groupsOfSvgFiles.Count)"
    } -Process {
        Write-Host "Creating stencil $($_.Name).vssx with $($_.Group.Count) masters..."
        Write-Progress -Id 1 -Activity "Creating stencils..." -Status "File $($i + 1) of $($groupsOfSvgFiles.Count)" -PercentComplete ($i / $groupsOfSvgFiles.Count * 100)
        New-VisioStencil $_.Group -StencilPath "$($_.Name).vssx" -MasterNameExtractor $nameExtractor
        Write-Host "Stencil $($_.Name).vssx completed."
        $i++
    } -End {
        Write-Progress -Id 1 -Activity "Creating stencils..." -Completed
    }