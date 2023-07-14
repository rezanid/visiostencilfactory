param([string]$svgFilePath='PATH-To-ICONS\Architecture-Service-Icons_04282023\*\64\*.svg')

# This requires the Architecture-Service-Icons to be unziped into a directory (on Linux/UNIX based systems) 
# and had the name modified as follows:
#   sudo apt-get install rename
#   unzip unzip AWS\ Architecture\ Icons\ Asset-Package.zip
#   find ./Architecture-Service-Icons_04282023/ -type f -name "*:*" -exec rename -n 's/:/-/g' {} +
#
# This is required because some of the file names have a colon (':') in them (re:POST), 
# and Windows will not extract those to the file system. Once that is done, the below will work on Windows.
# Not tested on any other system. I used WSL (Ubuntu 22.04 LTS) to unzip and rename.

# Dot source the New-VisioStencil cmdlet
$scriptPath = $PWD.Path + "\visio\New-VisioStencil.ps1"
. $scriptPath

# Declare a name extractor command
$nameExtractor = {
    param($name) 
    $name = $name -replace 'Arch_Amazon-', ''  # Remove 'Arch_Amazon-'
    $name = $name -replace 'Arch_AWS-', ''  # Remove 'Arch_AWS-'
    $name = $name -replace 'Arch_', ''  # Remove 'Arch_AWS-'
    $name = $name -replace '_64', ''       # Remove '_64.svg'
    $name = $name -replace '-', ' '             # Replace '-' with ' '
    $name = $name -replace '', ':' # Replace '' with a ':'
    return $name
}

# Group files into categories
$groupsOfSvgFiles = Get-ChildItem $svgFilePath -Recurse |
    Group-Object -Property @{
        Expression = {$_.Directory.Parent.Name}
    }

# Generate Visio stencils, group by group
$groupsOfSvgFiles |
    ForEach-Object -Begin {
        $i = 0
        Write-Host "Stencils to be created: $($groupsOfSvgFiles.Count)"
    } -Process {
        Write-Host "Creating stencil AWS-$($_.Name).vssx with $($_.Group.Count) masters..."
        Write-Progress -Id 1 -Activity "Creating stencils..." -Status "File $($i + 1) of $($groupsOfSvgFiles.Count)" -PercentComplete ($i / $groupsOfSvgFiles.Count * 100)
        New-VisioStencil $_.Group -StencilPath "AWS-$($_.Name).vssx" -MasterNameExtractor $nameExtractor
        Write-Host "Stencil AWS-$($_.Name).vssx completed."
        $i++
    } -End {
        Write-Host "Creating stencils..."
    }
