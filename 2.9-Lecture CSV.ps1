$CurrentDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$csv ="$CurrentDir\file.csv"

if (Test-Path $csv -PathType leaf)
	{
	$importTable = @()
	$data = import-csv -path $csv

	foreach($item in $data){
		$hash = @{
        	societe = $item.societe
	        Address = $item.name
	        prenom = $item.prenom
	        telephone = $item.telephone
    		}
	$objTemp = new-object psobject -property $hash
	$importTable += $objTemp
	}
	$importTable | ogv -passthru
}
else
{
echo "csv file is missing"
echo $csv
sleep 4
}
