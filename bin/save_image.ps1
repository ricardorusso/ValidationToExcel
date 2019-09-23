#'D:\FileEx\MonitorCSW .xlsx'
#'D:\FileEx\test3.png' 
$testfile = $args[0]

$excel = New-Object -COM "Excel.Application"		# Create new COM object
$excel.displayalerts = $false
$excel.visible = $false
$excel.usercontrol = $false				# Disable user interaction with Excel
$Workbook=$excel.Workbooks.open($testfile)		# Open XLS file in Excel
$Worksheet=$Workbook.Worksheets.Item(1)
$Worksheet.Activate() | Out-Null

add-type -an system.windows.forms


$rng = $Worksheet.range("A5","C18")
$pic = $rng.Copy()

Add-Type -AssemblyName System.Windows.Forms
$clipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
if ($clipboard.ContainsImage()) {
    $filename= $args[1]      
    [System.Drawing.Bitmap]$clipboard.getimage().Save($filename, [System.Drawing.Imaging.ImageFormat]::Png)
    Write-Output "clipboard content saved as $filename"
} else {
    Write-Output "clipboard does not contains image data"
}
$excel.Quit();