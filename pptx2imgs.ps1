param(
    [Parameter(Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('FullName')]
    [String] $ppFile
)


$ErrorActionPreference = "Stop"
Add-Type -AssemblyName Microsoft.Office.InterOp.PowerPoint    

function main() {    
    $fileName = [IO.Path]::GetFileNameWithoutExtension((Split-Path -Leaf $ppFile))
    $destDir = Join-Path (Get-Item $ppFile).DirectoryName $fileName
    if(Test-Path $destDir) {
        New-Item $destDir -ItemType Directory -Force
    }
    pptx2JPG $destDir
    resizeImg $destDir
    rename $destDir
    JPG2jpg $destDir
}

function pptx2JPG($destDir) {
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = "msoTrue"
    $app.DisplayAlerts = "ppAlertsAll"
    try {
        $pptPres = $app.Presentations.Open($ppFile)    
        try {
            $pptPres.SaveAs($destDir, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::PpSaveAsJPG)
        } finally {
            $pptPres.Close()
        }
    } finally {
        $app.Quit()
        [gc]::collect()
        [gc]::WaitForPendingFinalizers()
        [gc]::collect() 
    }    
}

function resizeImg($destDir) {
    [void][Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $width = 640
    $height = 480
    foreach($f in Get-ChildItem $destDir) {
        # ???T?C?Y??????p?X
        $tmp = Join-Path $destDir ("re" + ([System.IO.Path]::GetFileNameWithoutExtension($f.FullName)) + ".JPG")
        $image = New-Object System.Drawing.Bitmap($f.FullName)    
        # ???T?C?Y???????T?C?Y??canvas?????
        $canvas = New-Object System.Drawing.Bitmap($width, $height)
        # ?????canvas?????`????
        $graphics = [System.Drawing.Graphics]::FromImage($canvas)
        $graphics.DrawImage($image, (New-Object System.Drawing.Rectangle(0, 0, $canvas.Width, $canvas.Height)))
        $canvas.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        # ???
        $graphics.Dispose()
        $canvas.Dispose()
        $image.Dispose()
        # ???T?C?Y?O?????????
        Remove-Item $f.FullName
    }    
}

function rename($destDir) {
    foreach($f in Get-ChildItem $destDir) {
        $tmp = Join-Path $destDir $f
        $new = $tmp.Replace("re?X???C?h", "img")
        Rename-Item -Path $tmp -NewName $new            
    }
}

function JPG2jpg($destDir) {
    $oldExt = "JPG"
    $newExt = "jpg"
    
    $matStr = '.' + $oldExt
    $oldStr = '\.'+ $oldExt + '$'
    $newStr = '.' + $newExt
    
    Get-ChildItem -Path $destDir | Where-Object {$_.Extension -eq $matStr} | Rename-Item -NewName { $_.Name -replace $oldStr, $newStr }
}

# ???s
main
