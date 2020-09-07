#
#Source: https://github.com/MichaelRyom/Nordnet2SimplyWall.st
#Created by Michael Ryom
#

$data = Import-Csv $( dir | sort LastWriteTime -Descending | Out-GridView -PassThru ) -Delimiter "`t"

$output = @()
foreach ( $line in $data ) {
    $Details = "" | Select-Object Ticker,USDate,Shares,Price,Cost,Type

    $Details.Ticker = $line.Værdipapirer
    $Details.USDate = get-date $line.Handelsdag -Format "MM/dd/yyyy"
    $Details.Shares = $line.Antal -replace "\." -replace ",","."
    $Details.Price = [decimal]($line.Kurs -replace "\." -replace ",",".")
    $Details.Cost = if ( [decimal]($line.Beløb -replace "\." -replace ",",".") -match "-" ){
            -[decimal]($line.Beløb -replace "\." -replace ",",".")
        }else {
            [decimal]($line.Beløb -replace "\." -replace ",",".")
        }
    $Details.Type = if ( $line.Transaktionstype -like "KØBT" ){
            "Buy"
        }elseif ( $line.Transaktionstype -like "SOLGT" ) {
            "Sell"
        }
    $output += $Details

}

$output | Where-Object {$_.type -match "buy" -or $_.type -match "sell"} | ConvertTo-Csv -Delimiter "," -NoTypeInformation -OutVariable staged | Out-Null
$staged -replace '"' | Out-File SimplyWallSt.csv -Encoding utf8 -NoClobber
