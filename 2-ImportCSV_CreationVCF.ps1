$CurrentDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$csv ="$CurrentDir\file.csv"


if (!(Test-Path "$CurrentDir\result" -PathType leaf))
	{
	mkdir "$CurrentDir\result"
	echo "$CurrentDir\result is created"
	}

if (Test-Path $csv -PathType leaf)
	{
	$data = import-csv -path $csv

	$importTable = @()
	foreach($item in $data){
		$hash = @{
        	societe = $item.societe
        	service = $item.service
        	intitule = $item.intitule
		nom = $item.nom
		prenom = $item.prenom
		email = $item.email
		telephone = $item.telephone
		}



###############################
		$filename="result\contact_"+$item.nom+"_"+$item.prenom+"_.vcf"
		[string]$filecontent=""
		$filecontent+="BEGIN:VCARD"+"`r`n"
		$filecontent+="VERSION:2.1"+"`r`n"
		$filecontent+="N;LANGUAGE=fr;CHARSET=Windows-1252:$(($item).nom);$(($item).prenom) "+"`r`n"
		$filecontent+="FN;CHARSET=Windows-1252:$(($item).nom) $(($item).prenom)"+"`r`n"
		$filecontent+="ORG:$(($item).societe)"+"`r`n"
		$filecontent+="TITLE:$(($item).intitule)"+"`r`n"
		$filecontent+="TEL;CELL;VOICE:$(($item).telephone)"+"`r`n"
		$filecontent+="EMAIL:$(($item).email)"+"`r`n"
		$filecontent+="X-MS-OL-DEFAULT-POSTAL-ADDRESS:0"+"`r`n"
		$filecontent+='X-MS-OL-DESIGN;CHARSET=utf-8:<card xmlns="http://schemas.microsoft.com/office/outlook/12/electronicbusinesscards" ver="1.0" layout="left" bgcolor="ffffff"><img xmlns="" align="fit" area="16" use="cardpicture"/><fld xmlns="" prop="name" align="left" dir="ltr" style="b" color="000000" size="10"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="telcell" align="left" dir="ltr" color="d48d2a" size="8"><label align="right" color="626262">Mobile</label></fld><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/></card>'+"`r`n"
		$pourrev=Get-Date -Format "yyyyMMddThhmmssZ"
		$filecontent+="REV:$pourrev"+"`r`n"
		$filecontent+="END:VCARD"+"`r`n"
		#
		$filecontent | out-file -Encoding "ASCII" -filepath $filename
#################################
		
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

