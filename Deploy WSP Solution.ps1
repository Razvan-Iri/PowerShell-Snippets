[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$webAppUrl = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Site Collection URL", "Site Collection URL")
$CurrentDate = Get-Date
$CurrentDate = $CurrentDate.ToString('MM-dd-yyyy||hh-mm-ss||')
$ErrorPath = "$PSScriptRoot\log.txt"
$WebPartPath = "$PSScriptRoot\SDC_PowerBI.wsp"
$SolutionNAme = "SDC_PowerBI.wsp"


Try{
    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue 
    Write-Host "Miscrosoft SharePoint Snapin has loaded"
    "$CurrentDate Miscrosoft SharePoint Snapin has loaded" | Out-File -Append -FilePath $ErrorPath
}Catch{
    Write-Host $_.Exception.Message
    $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
}
Start-Sleep 5

Try{
    Try{
    $check = Get-SPSolution -Identity:$SolutionNAme -ErrorAction SilentlyContinue 
    Write-Host "Solution is present on Site Collection"
    "$CurrentDate Solution is present on Site Collection"| Out-File -Append -FilePath $ErrorPath
    }Catch{
    Write-Host $_.Exception.Message
    $CurrentDate + $_.Exception.Message| Out-File -Append -FilePath $ErrorPath
    }
    
    if($check.Name -eq $SolutionNAme -and $check.Deployed -eq $True){
        Write-Host "Solution is already installed on Site Collection"
        "$CurrentDate Solution is already installed on Site Collection" | Out-File -Append -FilePath $ErrorPath
       Try{
        Uninstall-SPSolution $SolutionNAme -AllWebApplications -confirm:$true -ErrorAction SilentlyContinue 
        Write-Host "Solution has been uninstalled from the Site Collection"
        "$CurrentDate Solution has been uninstalled from the Site Collection" | Out-File -Append -FilePath $ErrorPath
       }Catch {
        Write-Host "Unable to uninstall solution"
        Write-Host $_.Exception.Message
        $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
        }
    }
    elseif($check.Name -eq $SolutionNAme-and $check.Deployed -eq $False ) {
        Write-Host "Solution is already added on Site Collection"
        "$CurrentDate Solution is already added on Site Collection" | Out-File -Append -FilePath $ErrorPath
      Try{
        Remove-SPSolution -Identity:$SolutionNAme -ErrorAction SilentlyContinue 
        Write-Host "Solution has been removed from the Site Collection"
       "$CurrentDate Solution has been removed from the Site Collection" | Out-File -Append -FilePath $ErrorPath
      }Catch{
        Write-Host "Unable to remove solution"
        $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
      }
    }
}Catch{
    Write-Host "Solution is not installed on Site Collection"
    $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
}

Try {
    Add-SPSolution -LiteralPath $webPartPath -ErrorAction SilentlyContinue 
    Write-Host "Solution has been addded"
    "$CurrentDate Solution has been addded" | Out-File -Append -FilePath $ErrorPath
}Catch{
    Write-Host "Solution has not been added"
    $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
}
Start-Sleep 5
Try {
    Install-SPSolution -Identity $SolutionNAme -WebApplication $webAppUrl -GACDeployment -ErrorAction SilentlyContinue 
    Write-Host "Solution has been installed"    
    "$CurrentDate Solution has been installed" | Out-File -Append -FilePath $ErrorPath
}
Catch {
    Write-Host "Solution has not been installed"
    $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
}
Start-Sleep 20
Try{
    Enable-SPFeature -Identity 60b2b5ac-5159-40f2-a4de-735084a0445f -force -Url http://romada-adcsh1:1919 -ErrorAction Stop
    Write-Host "Solution Feature has been installed" 
    "$CurrentDate Solution Feature has been installed" | Out-File -Append -FilePath $ErrorPath
}
Catch{
    Write-Host "Solution Feature has not been installed"
    $CurrentDate + $_.Exception.Message | Out-File -Append -FilePath $ErrorPath
}



