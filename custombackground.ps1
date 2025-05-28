# [ RUN STRING ]
# 
# Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/Computer-Motivators/misc-scripts/main/custombackground.ps1'))
# 

# Gather system information
$hostname = $env:COMPUTERNAME
$ipAddress = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object {$_.InterfaceAlias -notlike '*Loopback*'} | Select-Object -ExpandProperty IPAddress -First 1)
$macAddress = (Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Select-Object -ExpandProperty MacAddress -First 1)
$serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
$screenWidth = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Width
$screenHeight = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize.Height

# Add RustDesk ID system information
if (Test-Path "C:\Program Files\Computer_Motivators_Support_Client\Computer_Motivators_Support_Client.exe") {
    $rustdeskId = & "C:\Program Files\Computer_Motivators_Support_Client\Computer_Motivators_Support_Client.exe" --get-id | Out-String
} else {
    $rustdeskId = "N/A"
}

# [ OPTIONS ]
$skipHostnames = @("Server-Ignore") # Blacklisted computer hostnames
$baseWallpaperImage = "https://next.physcorp.com/s/w86T47ssbneXqYg/download/background-light-alt.png"
$baseLogoImage = "https://next.physcorp.com/s/HE3HxtCTkgP5YWt/download/Long%20Logo.png"
# Define the text to overlay
$supportEmail = "support@computermotivators.com"
$supportLink = "https://help.computermotivators.com"
$text = @"
=== Computer Info ===
Name: $hostname
Serial: $serialNumber

=== Network Info ===
Last IP: $ipAddress
MAC: $macAddress

=== Support ===
Email: $supportEmail
Web: $supportLink
ID: $rustdeskId
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
param(
    [switch]$KeepCurrentWallpaper
)

if ($KeepCurrentWallpaper) {
    $baseWallpaperImage = $transcodedWallpaper
}

# Load the base image and create a new image with text using System.Drawing
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
$image = [System.Drawing.Image]::FromFile($baseWallpaperImage)
$baseLogo = [System.Drawing.Image]::FromFile($baseLogoImage)

# Resize image to screen resolution
$resizedImage = New-Object System.Drawing.Bitmap($screenWidth, $screenHeight)
$graphics = [System.Drawing.Graphics]::FromImage($resizedImage)
$graphics.DrawImage($image, 0, 0, $screenWidth, $screenHeight)

# Maintain aspect ratio for the logo and calculate scaled dimensions
$logoMaxWidth = 160  # Maximum width for the logo
$logoMaxHeight = 96  # Maximum height for the logo
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
$brushColor = [System.Drawing.Color]::FromArgb(160, 0, 0, 0)
$brush = New-Object System.Drawing.SolidBrush($brushColor)

# Create a StringFormat object for right-aligned text
$stringFormat = New-Object System.Drawing.StringFormat
$stringFormat.Alignment = [System.Drawing.StringAlignment]::Far  # Aligns text to the right
$stringFormat.LineAlignment = [System.Drawing.StringAlignment]::Near  # Aligns text at the top

# Define the position for the text below the logo
$textPositionX = $screenWidth - 20  # Align text with a margin from the right
$textPositionY = $logoY + $logoHeight + 10  # Position text just below the logo with a margin
$textPosition = New-Object System.Drawing.PointF($textPositionX, $textPositionY)

# Draw the text on the image using StringFormat for right-alignment
$graphics.DrawString($text, $font, $brush, $textPosition, $stringFormat)
$graphics.Dispose()

# Save the new image
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
