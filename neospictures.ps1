# NeosVRのスクリーンショットにsaveexif.dllが保存したExifからlocationName(ワールド名)を出力
Add-Type -AssemblyName System.Drawing
$targetFolder = (Get-Location).Path
$imageList = Get-ChildItem -Path $targetFolder -Filter "*.jpg"
try {
    foreach($image in $imageList)
    {
        $fileName = $targetFolder + "\" + $image.Name
        $bitmapObj = New-Object Drawing.Bitmap($fileName)
        $byteAry = ($bitmapObj.PropertyItems | Where-Object{$_.Id -eq 37510}).Value
        if($null -ne $byteAry)
        {
            $jsonString = [System.Text.Encoding]::Unicode.GetString($byteAry).TrimEnd("`0")
            $jsonObj = ConvertFrom-Json $jsonString
            Write-Output($jsonObj.locationName)
        }
        $bitmapObj.Dispose()
    }
}
catch [Exception]{
    $error[0] | Out-string | write-host
}
