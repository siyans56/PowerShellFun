

Function translate {

[uint16]$first =Read-Host -Prompt "Enter first portion of IP"
[uint16]$second = Read-Host -Prompt "Enter second portion of IP"
[uint16]$third = Read-Host -Prompt "Enter third portion of IP"
[uint16]$fourth = Read-Host -Prompt "Enter fourth portion of IP"
[uint16]$mask = Read-Host -Prompt "Enter mask"



$mask = 23
$total
$count = [math]::pow(2,32-$mask)
$runtimes = ($count/256)
$i = $fourth

For ($runs = 0; $runs -le $runtimes-1; $runs++) {
    For ($i; $i -le 255; $i++) {
        "$first" + "." + "$second" + "." + "$third" + "." + "$i"
        $total++
        }
    $i = 0
    $third++
    }

While ($i -le $fourth-1) {
    "$first" + "." + "$second" + "." + "$third" + "." + "$i"
    $i++
    $total++
    }

#"$total IPs listed"
#Read-Host -Prompt "Press Enter to close window"

}

translate | Out-File -append -filepath "C:\Users\P2824589\Documents\output.csv"
"Exported to C:\Users\P2824589\Documents\output.csv"
$again = Read-Host -Prompt "Type 1 to run again"

while ($again) { 
translate | Out-File -append -filepath "C:\Users\P2824589\Documents\output.csv"
"Exported to C:\Users\P2824589\Documents\output.csv"
}
Start-Sleep -Seconds 1.5