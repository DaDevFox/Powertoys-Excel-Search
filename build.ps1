param (
    [string]$Version = "",
    [string]$ProjName = ""
)
if($ProjName -eq ""){
    $projectFile = (get-childitem .\*.csproj -Recurse)[0]
    Write-Host "Detected .csproj: $projectFile"
    $_, $ProjName = ($projectFile.Name -match '.*Community\.PowerToys\.Run\.Plugin\.(.*)\.csproj') , $Matches[1]
    Write-Host "Detected project name: $ProjName"
}

$success = $true

if (Get-Process -Name "powertoys" -ErrorAction SilentlyContinue) {
    Write-Host "Stopping PowerToys..."
    $restart_needed = $true
    Stop-Process -Name "powertoys"
}
else {
    $restart_needed = $false
}

if ($?) {
    Write-Host "`nStarting build...`n"

    if ($Version -ne "") {
        Write-Host "Updating release version to $Version..."

        $plugin_json = Get-Content -Path ".\plugin.json" -Raw | ConvertFrom-Json
        $plugin_json.Version = $Version
        $plugin_json | ConvertTo-Json | Set-Content -Path ".\plugin.json"

        Write-Host "Updated version number in plugin.json."

        [xml]$project_xml = Get-Content -Path ".\Community.PowerToys.Run.Plugin.$ProjName.csproj"
        $project_xml.Project.PropertyGroup.Version = $Version
        $project_xml.Save(".\Community.PowerToys.Run.Plugin.$ProjName.csproj")

        Write-Host "Updated verion number in project-file.`n"
    }

    dotnet msbuild .\Community.PowerToys.Run.Plugin.$ProjName.sln /property:GenerateFullPaths=true /consoleloggerparameters:NoSummary /p:Configuration=Release /p:Platform="Any CPU"

    if ($?) {
        $source = ".\bin\x64\Release\net8.0-windows\"
        $plugin_folder = $env:LOCALAPPDATA + "\Microsoft\PowerToys\PowerToys Run\Plugins\$ProjName\"

        Write-Host "`nCopying data into plugins folder..."

        if (Test-Path -Path $plugin_folder) {
            Write-Host "Removing old plugin version..."
            Remove-Item $plugin_folder -Recurse
        }
        Copy-Item $source $plugin_folder -Recurse

        Write-Host "Copied new release into plugin folder."
        Write-Host "`nCreating release-archive..."

        if (Test-Path -Path "Release.zip") {
            Write-Host "Removing old release-archive..."
            Remove-Item "Release.zip"
        }

        Compress-Archive -Path $source -DestinationPath "Release.zip"

        Write-Host "Release-archive created."

        #Remove-Item $source -Recurse

        Write-Host "`nBuild complete."
    }
    else {
        Write-Host "`nBuild was unsuccessful."
        $success = $false
    }

    if ($restart_needed) {
        Write-Host "`nRestarting PowerToys..."
        Start-Process -FilePath "C:\Program Files\PowerToys\PowerToys.exe"
    }
}
else {
    Write-Host "`nCouldn't stop PowerToys. Abort build."
    $success = $false
}

if ($success -eq $false) {
    exit 1
}