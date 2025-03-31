Clear-Host

function Show-Menu {
    Write-Output "----------------------------------------------"
    Write-Output "                 PC SETUP [ PRIYANSHU ]               "
    Write-Output "----------------------------------------------"
    Write-Output "[1] Install Normal Software - from GitHub"
    Write-Output "[2] Install MS Office - from GitHub"
    Write-Output "[3] Activation Tool - Using Custom URL"
    Write-Output "[0] Exit"
    Write-Output "----------------------------------------------"
    
    $choice = Read-Host "Choose a menu option using your keyboard [1,2,3,0]"

    switch ($choice) {
        '1' {
            Write-Output "Installing Normal Software..."
            powershell -ep Bypass -c "iwr 'https://raw.githubusercontent.com/Priyanshu8494/my-choco-install-script/main/setup-new-pc.ps1' | iex"
        }
        '2' {
            Write-Output "Installing MS Office..."
            powershell -ep Bypass -c "iwr 'https://raw.githubusercontent.com/Priyanshu8494/ms-office-install-script/main/setup-office.ps1' | iex"
        }
        '3' {
            Write-Output "Running Activation Tool..."
            irm https://get.activated.win | iex
        }
        '0' {
            Write-Output "Exiting. Goodbye!"; exit
        }
        default {
            Write-Output "Invalid selection. Please try again."
        }
    }

    Pause
    Show-Menu
}

Show-Menu
