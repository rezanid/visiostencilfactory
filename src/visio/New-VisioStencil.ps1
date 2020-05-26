using namespace Microsoft.Office.Interop.Visio

function New-VisioStencil {
    <#
    .SYNOPSIS
        Creates a new VisioStencil (vssx file) using Visio Automation by using a collection of SVG files.
    .DESCRIPTION
        The Create-VisioStencil cmdlet creates from a collection of SVG files.
    .EXAMPLE
    C:\PS>dir "*.svg" | select -first 3 | New-VisioStencil -StencilPath "TestStencil1.vssx"

    .EXAMPLE
    C:\PS>New-VisioStencil (Get-ChildItem "*.svg" | Select-Object -First 5) -StencilPath "E:\Temp\Visio\TestStencil2.vssx"

    .EXAMPLE
    C:\PS>$nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
    C:\PS>New-VisioStencil (Get-ChildItem "*.svg" | Select-Object -First 5) -StencilPath "E:\Temp\Visio\TestStencil2.vssx" -MasterNameExtractor $nameExtractor

    .EXAMPLE
    C:\PS>$nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
    C:\PS>$groupsOfSvgFiles = Get-ChildItem "*.svg" |
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
                New-VisioStencil $_.Group -StencilPath "$($_.Name).vssx" -MasterNameExtractor $nameExtractor -Verbose
                Write-Host "Stencil $($_.Name).vssx completed."
                $i++
            } -End {
                Write-Progress -Id 1 -Activity "Creating stencils..." -Completed
            }
        #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [AllowNull()]
        [String[]]
        $SvgPath,
        [string]
        $StencilPath,
        [AllowNull()]
        [scriptblock]
        $MasterNameExtractor
    )
    begin { 
        $stopWatch = [System.Diagnostics.StopWatch]::StartNew()
        Write-Debug "Starting Visio application and creating a new stencil..."
        $visioApp = New-Object -ComObject Visio.Application
        $visioStencil = $VisioApp.Documents.Add("vssx")
        $visioMasters = $VisioStencil.Masters
        Write-Verbose "Empty stencil created successfully"
        $i = 0
    }
    process { 
        $showProgress = $SvgPath.Count -gt 1
        foreach ($svgfile in $SvgPath) {
            if ($showProgress) { 
                Write-Progress -ParentId 1 -Id 100 -PercentComplete ($i / $SvgPath.Count * 100) -Activity "Creating masters..." -Status "Master $($i + 1) of $($SvgPath.Count)" 
            }

            # Validate file exists
            if (-Not (Test-Path $svgfile)) { throw "File does not exist" }
            if (-Not (Test-Path $svgfile -PathType Leaf)) { throw "The Path argument must be a file. Folder paths are not allowed."}
            
            # Extract the filename without extension
            [System.IO.FileInfo]$fileInfo = $svgfile
            $masterName = $fileInfo.BaseName
            
            if ($MasterNameExtractor -ne $null) { 
                Write-Verbose "Extracting master name from $($masterName)..."
                $masterName = [string]$MasterNameExtractor.Invoke($masterName)
                Write-Verbose "Master name extracted: $($masterName)"
            }
            $newMaster = $visioMasters.Add()
            $newMaster.Name = $masterName
            $shape = $newMaster.Import($svgfile)
            $oldWidth = $shape.CellsU("Width").ResultIU

            # Set the width to 0.5 inch and change the heigth respectively.
            $shape.CellsU("Width").ResultIU = 0.5
            $shape.CellsU("Height").ResultIU = $shape.CellsU("Height").ResultIU * (0.5 / $oldWidth)

            $shape.CellsSrc([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowGroup, [VisCellIndices]::visGroupSelectMode).FormulaU = "0"

            try {
                Set-ShapeData($shape)
            }
            catch {
                Write-Warning "Failed to set all properties for the master shape: $masterName, file: $svgfile"
            }

            $i++
        }
    }
    end {
        if ($showProgress) { Write-Progress -Id 100 -Activity "Creating masters..." -Completed }
        $visioStencil.SaveAs($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($StencilPath)) | Out-Null
        $visioStencil.Close() | Out-Null
        $visioApp.Quit() | Out-Null
        $stopWatch.Stop()
        Write-Host "Stencil created and stored successfully (elapsed: $($stopWatch.Elapsed))."
    }
}

#region Private functions...
function Set-ShapeData($shape) {
    # Add 12 connection points evenly distributed around the shape's box.
    $shape.AddSection([VisSectionIndices]::visSectionConnectionPts) | Out-Null
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    # 3 connection point on the left
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 0, [VisCellIndices]::visCnnctX).FormulaU = "0"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 0, [VisCellIndices]::visCnnctY).FormulaU = "0.25*Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 1, [VisCellIndices]::visCnnctX).FormulaU = "0"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 1, [VisCellIndices]::visCnnctY).FormulaU = "0.5*Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 2, [VisCellIndices]::visCnnctX).FormulaU = "0"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 2, [VisCellIndices]::visCnnctY).FormulaU = "0.75*Height"
    # 3 connection points on the bottom
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 3, [VisCellIndices]::visCnnctX).FormulaU = "0.25*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 3, [VisCellIndices]::visCnnctY).FormulaU = "0"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 4, [VisCellIndices]::visCnnctX).FormulaU = "0.5*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 4, [VisCellIndices]::visCnnctY).FormulaU = "0"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 5, [VisCellIndices]::visCnnctX).FormulaU = "0.75*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 5, [VisCellIndices]::visCnnctY).FormulaU = "0"
    # 3 connection points on the top
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 6, [VisCellIndices]::visCnnctX).FormulaU = "0.25*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 6, [VisCellIndices]::visCnnctY).FormulaU = "Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 7, [VisCellIndices]::visCnnctX).FormulaU = "0.5*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 7, [VisCellIndices]::visCnnctY).FormulaU = "Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 8, [VisCellIndices]::visCnnctX).FormulaU = "0.75*Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 8, [VisCellIndices]::visCnnctY).FormulaU = "Height"
    # 3 connection points on the right
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 9, [VisCellIndices]::visCnnctX).FormulaU = "Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 9, [VisCellIndices]::visCnnctY).FormulaU = "0.25*Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 10, [VisCellIndices]::visCnnctX).FormulaU = "Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 10, [VisCellIndices]::visCnnctY).FormulaU = "0.5*Height"
    $shape.AddRow([VisSectionIndices]::visSectionConnectionPts, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 11, [VisCellIndices]::visCnnctX).FormulaU = "Width"
    $shape.CellsSRC([VisSectionIndices]::visSectionConnectionPts, 11, [VisCellIndices]::visCnnctY).FormulaU = "0.75*Height"
    
    # Make dynamic connector connecto to the bounding box instead of the shape's geometry.
    # Ref.: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-xp/aa200985(v=office.10)
    # Ref.: https://docs.microsoft.com/en-us/office/client-developer/visio/shapefixedcode-cell-shape-layout-section
    $shape.CellsSRC([VisSectionIndices]::visSectionObject,[VisRowIndices]::visRowShapeLayout, [VisCellIndices]::visSLOConFixedCode).FormulaU = [VisCellVals]::visSLOFixedNoFoldToShape

    # Add a control point centered under the shape, to allow the user to control shape's text position.
    $shape.AddSection([VisSectionIndices]::visSectionControls) | Out-Null
    $shape.AddRow([VisSectionIndices]::visSectionControls, [VisRowIndices]::visRowLast, [VisRowTags]::visTagDefault) | Out-Null
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlX).FormulaU = "Width*0.5"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlY).FormulaU = "-0.5*(ABS(SIN(Angle))*TxtWidth+ABS(COS(Angle))*TxtHeight)"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlXDyn).FormulaU = "Width*0.5"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlYDyn).FormulaU = "Height*0.5"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlXCon).FormulaU = "(Controls.Row_1>Width*0.5)*2+2+IF(OR(HideText,STRSAME(SHAPETEXT(TheText),`"`")),5,0)"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlYCon).FormulaU = "(Controls.Row_1.Y>Height*0.5)*2+2"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlGlue).FormulaU = "TRUE"
    $shape.CellsSRC([VisSectionIndices]::visSectionControls, 0, [VisCellIndices]::visCtlTip).FormulaU = "`"Reposition text`""

    # Set shapes's text font and color.
    $shape.CellsSRC([VisSectionIndices]::visSectionCharacter, 0, [VisCellIndices]::visCharacterSize).FormulaU = "9 pt"
    $shape.CellsSRC([VisSectionIndices]::visSectionCharacter, 0, [VisCellIndices]::visCharacterColor).FormulaU = "IF(LUM(THEMEVAL())>205,0,THEMEVAL(`"TextColor`",0))"

    # Set shape's text transformation properties.
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormWidth).FormulaU = "IF(TextDirection=0,TEXTWIDTH(TheText),TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormHeight).FormulaU = "IF(TextDirection=1,TEXTWIDTH(TheText),TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormHeight).FormulaU = "IF(TextDirection=1,TEXTWIDTH(TheText),TEXTHEIGHT(TheText,TEXTWIDTH(TheText)))"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormAngle).FormulaU = "IF(BITXOR(FlipX,FlipY),1,-1)*Angle"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormPinX).FormulaU = "Controls.Row_1"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormPinY).FormulaU = "Controls.Row_1.Y"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormLocPinX).FormulaU = "TxtWidth*0.5"
    $shape.CellsSRC([VisSectionIndices]::visSectionObject, [VisRowIndices]::visRowTextXForm, [VisCellIndices]::visXFormLocPinY).FormulaU = "TxtHeight*0.5"
}
#endregion

#region Execution examples...
# Example 1:
# dir "*.svg" | select -first 3 | New-VisioStencil -StencilName "TestStencil1"

# Example 2:
# New-VisioStencil (Get-ChildItem "*.svg" | Select-Object -First 5) -StencilPath "E:\Temp\Visio\TestStencil2.vssx"

# Example 3:
# Set-Location 'E:\Temp\Visio\Official Azure Icon Set'
# $nameExtractor = {param($name) ($name | Select-String "^\w+?-\d+?-(.+)").Matches[0].Groups[1].Value.Replace('-',' ') }
# $groupsOfSvgFiles = Get-ChildItem "*.svg" |
#     Group-Object -Property @{
#         Expression = {$_.BaseName.Substring(0,$_.BaseName.IndexOf('-'))}
#     }
# $groupsOfSvgFiles |
#     ForEach-Object $_ -Begin {
#         $i = 0
#         Write-Host "Stencils to be created: $($groupsOfSvgFiles.Count)"
#     } -Process {
#         Write-Host "Creating stencil $($_.Name).vssx with $($_.Group.Count) masters..."
#         Write-Progress -Id 1 -Activity "Creating stencils..." -Status "File $($i + 1) of $($groupsOfSvgFiles.Count)" -PercentComplete ($i / $groupsOfSvgFiles.Count * 100)
#         New-VisioStencil $_.Group -StencilPath "$($_.Name).vssx" -MasterNameExtractor $nameExtractor -Verbose
#         Write-Host "Stencil $($_.Name).vssx completed."
#         $i++
#     } -End {
#         Write-Progress -Id 1 -Activity "Creating stencils..." -Completed
#     }\
#endregion