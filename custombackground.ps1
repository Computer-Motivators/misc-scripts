param(
    [switch]$KeepCurrentWallpaper
)

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

# [ RUN STRING ]
# 
# Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/Computer-Motivators/misc-scripts/main/custombackground.ps1'))
# (Optional) Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; $s="$env:TEMP\bg.ps1"; iwr 'https://raw.githubusercontent.com/Computer-Motivators/misc-scripts/main/custombackground.ps1' -OutFile $s; & $s -KeepCurrentWallpaper
# 

# Gather system information
$hostname = $env:COMPUTERNAME
$ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notlike '*Loopback*'} | Select-Object -ExpandProperty IPAddress -First 1)
$macAddress = (Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty MacAddress -First 1)
$serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber

# Add RustDesk ID system information
if (Test-Path "C:\Program Files\Computer_Motivators_Support_Client\Computer_Motivators_Support_Client.exe") {
    $rustdeskId = & "C:\Program Files\Computer_Motivators_Support_Client\Computer_Motivators_Support_Client.exe" --get-id | Out-String
} else {
    $rustdeskId = "Not Installed"
}
$rustdeskId = $rustdeskId.Trim() # If there is an extra newline in the RustDesk ID, trim it

# [ OPTIONS ]
$skipHostnames = @("Server-Ignore") # Blacklisted computer hostnames
$baseWallpaperImage = "https://next.physcorp.com/s/w86T47ssbneXqYg/download/background-light-alt.png"
$baseLogoImage = "https://next.physcorp.com/s/7rMNLtCMeJtN9nC/download/Long%20Logo%20Background%20Pill.png"
# Define the text to overlay
$supportEmail = "support@computermotivators.com"
$supportLink = "www.computermotivators.com"
$supportPhone = "(210) PC-WIZRD"
$text = @"
=== Computer Info ===
Name: $hostname
Serial: $serialNumber
Support ID: $rustdeskId

=== Network Info ===
Last IP: $ipAddress
MAC: $macAddress

=== Support 9am-5pm M-F ===
$supportEmail
$supportLink
$supportPhone
"@

# Check if the computer name is in the skip list and exit if true
if ($skipHostnames -contains $hostname) {
    Write-Output "Skipping script execution on $hostname."
    exit
}

# Save the output image to a hidden location in the user's AppData folder
$outputWallpaperImage = "$env:APPDATA\Microsoft\Windows\Themes\custom-background.png"
$transcodedWallpaper = "$env:APPDATA\Microsoft\Windows\Themes\TranscodedWallpaper"

# Check if the paths to any images are URLs and download them to a temp folder
$tempFolder = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "CustomBackground")
if (-not (Test-Path $tempFolder)) {
    New-Item -ItemType Directory -Path $tempFolder | Out-Null
}

if ($baseWallpaperImage -match "^https?://") {
    $baseWallpaperImageLocal = [System.IO.Path]::Combine($tempFolder, "baseWallpaperImage.png")
    Invoke-WebRequest -Uri $baseWallpaperImage -OutFile $baseWallpaperImageLocal
    $baseWallpaperImage = $baseWallpaperImageLocal
}

if ($baseLogoImage -match "^https?://") {
    $baseLogoImageLocal = [System.IO.Path]::Combine($tempFolder, "baseLogoImage.png")
    Invoke-WebRequest -Uri $baseLogoImage -OutFile $baseLogoImageLocal
    $baseLogoImage = $baseLogoImageLocal
}

# Accept a script argument to keep the current wallpaper and just apply the overlay text
if ($KeepCurrentWallpaper) {
    $baseWallpaperImage = $transcodedWallpaper
}

# Replace [System.Windows.Forms.SystemInformation] with an alternative method to get screen dimensions
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ScreenResolution {
    [DllImport("user32.dll")]
    public static extern int GetSystemMetrics(int nIndex);
}
"@
$screenWidth = [ScreenResolution]::GetSystemMetrics(0)  # SM_CXSCREEN
$screenHeight = [ScreenResolution]::GetSystemMetrics(1)  # SM_CYSCREEN

# Add error handling for image loading
if (-not (Test-Path $baseWallpaperImage)) {
    Write-Error "Base wallpaper image not found: $baseWallpaperImage"
    exit
}

if (-not (Test-Path $baseLogoImage)) {
    Write-Error "Base logo image not found: $baseLogoImage"
    exit
}

# Load the base image and create a new image with text using System.Drawing
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Ensure resizedImage is initialized properly
$resizedImage = New-Object System.Drawing.Bitmap($screenWidth, $screenHeight)
$graphics = [System.Drawing.Graphics]::FromImage($resizedImage)

# Add logic to determine text color based on the average color of the base wallpaper image
function Get-AverageColor {
    param([System.Drawing.Bitmap]$image)

    $totalR = 0
    $totalG = 0
    $totalB = 0
    $pixelCount = 0

    for ($x = 0; $x -lt $image.Width; $x++) {
        for ($y = 0; $y -lt $image.Height; $y++) {
            $pixel = $image.GetPixel($x, $y)
            $totalR += $pixel.R
            $totalG += $pixel.G
            $totalB += $pixel.B
            $pixelCount++
        }
    }

    $averageR = $totalR / $pixelCount
    $averageG = $totalG / $pixelCount
    $averageB = $totalB / $pixelCount

    return [System.Drawing.Color]::FromArgb($averageR, $averageG, $averageB)
}

# Load the base image and calculate the average color
$image = [System.Drawing.Image]::FromFile($baseWallpaperImage)
$bitmap = New-Object System.Drawing.Bitmap $image
$averageColor = Get-AverageColor -image $bitmap
$image.Dispose()
$bitmap.Dispose()

# Determine whether the text color should be white or black based on the average color
$brightnessThreshold = 128  # Adjust threshold as needed
$averageBrightness = ($averageColor.R * 0.299) + ($averageColor.G * 0.587) + ($averageColor.B * 0.114)
if ($averageBrightness -ge $brightnessThreshold) {
    $brushColor = [System.Drawing.Color]::FromArgb(160, 0, 0, 0)  # Black with transparency
} else {
    $brushColor = [System.Drawing.Color]::FromArgb(225, 255, 255, 255)  # White with transparency
}

# Draw the base image and logo only if they are successfully loaded
$image = [System.Drawing.Image]::FromFile($baseWallpaperImage)
$baseLogo = [System.Drawing.Image]::FromFile($baseLogoImage)

if ($image -and $baseLogo) {
    $graphics.DrawImage($image, 0, 0, $screenWidth, $screenHeight)

    # Maintain aspect ratio for the logo and calculate scaled dimensions
    $logoMaxWidth = 256  # Maximum width for the logo
    $logoMaxHeight = 128  # Maximum height for the logo
    $logoAspectRatio = $baseLogo.Width / $baseLogo.Height
    if ($logoMaxWidth / $logoAspectRatio -le $logoMaxHeight) {
        $logoWidth = $logoMaxWidth
        $logoHeight = [int]($logoMaxWidth / $logoAspectRatio)
    } else {
        $logoHeight = $logoMaxHeight
        $logoWidth = [int]($logoMaxHeight * $logoAspectRatio)
    }

    # Set the logo position in the top-right corner
    $logoX = $screenWidth - $logoWidth - 20  # Position with a margin from the right
    $logoY = 20  # Position with a margin from the top

    # Draw the logo with 50% opacity
    $logoAttributes = New-Object System.Drawing.Imaging.ImageAttributes
    $opacityMatrix = New-Object System.Drawing.Imaging.ColorMatrix
    $opacityMatrix.Matrix33 = 0.5  # 50% opacity
    $logoAttributes.SetColorMatrix($opacityMatrix, [System.Drawing.Imaging.ColorMatrixFlag]::Default, [System.Drawing.Imaging.ColorAdjustType]::Bitmap)
    $graphics.DrawImage($baseLogo, [System.Drawing.Rectangle]::FromLTRB($logoX, $logoY, $logoX + $logoWidth, $logoY + $logoHeight), 0, 0, $baseLogo.Width, $baseLogo.Height, [System.Drawing.GraphicsUnit]::Pixel, $logoAttributes)

    # Set text properties
    $font = New-Object System.Drawing.Font("Consolas", 12, [System.Drawing.FontStyle]::Bold)  # Industrial-style font

    # Update the brush initialization
    $brush = New-Object System.Drawing.SolidBrush($brushColor)

    # Create a StringFormat object for right-aligned text
    $stringFormat = New-Object System.Drawing.StringFormat
    $stringFormat.Alignment = [System.Drawing.StringAlignment]::Far  # Aligns text to the right
    $stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Near  # Aligns text at the top

    # Define the position for the text below the logo
    $textPositionX = $screenWidth - 40  # Align text with a margin from the right
    $textPositionY = $logoY + $logoHeight + 10  # Position text just below the logo with a margin
    $textPosition = New-Object System.Drawing.PointF($textPositionX, $textPositionY)

    # Draw the text on the image using StringFormat for right-alignment
    $graphics.DrawString($text, $font, $brush, $textPosition, $stringFormat)
}

# Dispose of graphics and save the image
$graphics.Dispose()
$resizedImage.Save($outputWallpaperImage, [System.Drawing.Imaging.ImageFormat]::Png)
$resizedImage.Dispose()
$image.Dispose()
$baseLogo.Dispose()

# Clean up temp folder
if (Test-Path $tempFolder) {
    Remove-Item -Path $tempFolder -Recurse -Force
}

# Update the TranscodedWallpaper to ensure the new background is applied immediately
Copy-Item -Path $outputWallpaperImage -Destination $transcodedWallpaper -Force

# Set the new image as the wallpaper using registry update
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop\" -Name "Wallpaper" -Value $outputWallpaperImage

# Refresh the wallpaper by updating the registry and broadcasting the change
$shell = New-Object -ComObject WScript.Shell
$shell.RegWrite("HKCU\Control Panel\Desktop\Wallpaper", $outputWallpaperImage)
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
[Wallpaper]::SystemParametersInfo(0x0014, 0, $outputWallpaperImage, 0x0001)
