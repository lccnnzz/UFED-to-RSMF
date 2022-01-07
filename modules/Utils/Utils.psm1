function Format-Timestamp{
    param(
        [Parameter(Mandatory)] $Timestamp
    )
    $regexUTC       = [Regex]::new("\(UTC[+|-][0-12]\)")
    $match          = $timestamp.Replace($regexUTC.Match($timestamp).Value, "")
    $ts = [datetime]::ParseExact(($match -replace "[^0-9/\:\s]"),"dd/MM/yyyy HH:mm:ss",$Null)
    $ISOTimestamp   = Get-Date $ts -Format "yyyy-MM-ddTHH:mm:ss"
    return $ISOTimestamp
}

function Get-StringHash{
    param(
        [Parameter(Mandatory)] $stringToHash,
        [Parameter()] $Algorithm       
    )
    $stringAsStream = [System.IO.MemoryStream]::new()
    $writer = [System.IO.StreamWriter]::new($stringAsStream)
    $writer.write($stringToHash)
    $writer.Flush()
    $stringAsStream.Position = 0
    $hash = Get-FileHash -InputStream $stringAsStream -Algorithm $Algorithm | Select-Object Hash
    return $hash
}

function ConcatHash{
    param(
        [Parameter()] [System.Collections.ArrayList] $Values,
        [Parameter()]
        [ValidateSet('SHA1', 'SHA256', 'SHA384', "SHA512", "MD5")]
        [String] $Algorithm = 'MD5'
    )
    
    $rawString = ""
    foreach ($value in $Values){
        switch (($value.GetType()).Name){
            'String'{$Svalue = $value.replace(" ", ""); break}
            # Clear whitespaces

            'Int32' {$Svalue = [string]$value; break}
            # Casts integer value to string

            'DateTime' {$Svalue = (Get-Date $value -format "yyyy-MM-ddTHH:mm:ss"); break}
            # Returns formatted timestamp as string
        }
        $rawString += $Svalue
    }
    $shash = (Get-StringHash -stringToHash $rawString -Algorithm $Algorithm)
    return $shash.Hash
}