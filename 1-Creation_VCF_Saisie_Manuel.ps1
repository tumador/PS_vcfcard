$societe = Read-Host -Prompt 'Saississez la societe :'
$poste = Read-Host -Prompt 'Saississez l intitulé du poste :'
$nom = Read-Host -Prompt 'Saississez le nom :'
$prenom = Read-Host -Prompt 'Saississez le prenom :'
$email = Read-Host -Prompt 'Saississez l adresse email :'
$telmobile = Read-Host -Prompt 'Saississez le numéro mobile +33-1-23-45-67-89 :'
$telfixe = Read-Host -Prompt 'Saississez le numéro fixe +33-1-23-45-67-89 :'
$adressepostale = Read-Host -Prompt 'Saississez l adresse postale :'
$url = Read-Host -Prompt 'Saississez l URL :'

$filename="contact_"+$nom+"_"+$prenom+"_.vcf"

$filecontent+="BEGIN:VCARD"+"`r`n"
$filecontent+="VERSION:2.1"+"`r`n"
$filecontent+="N;LANGUAGE=fr;CHARSET=Windows-1252:$nom;$prenom "+"`r`n"
$filecontent+="FN;CHARSET=Windows-1252:$nom $prenom"+"`r`n"
$filecontent+="ORG:$societe"+"`r`n"
$filecontent+="TITLE:$poste"+"`r`n"
$filecontent+="TEL;CELL;VOICE:$telmobile"+"`r`n"
$filecontent+="TEL;WORK;VOICE:$telfixe"+"`r`n"
$filecontent+="ORG:$societe"+"`r`n"
$filecontent+="TITLE:$poste"+"`r`n"
$filecontent+="EMAIL;PREF;INTERNET:$email"+"`r`n"
$filecontent+="URL;WORK:$url"+"`r`n"

$filecontent+="X-MS-OL-DEFAULT-POSTAL-ADDRESS:0"+"`r`n"
$filecontent+='X-MS-OL-DESIGN;CHARSET=utf-8:<card xmlns="http://schemas.microsoft.com/office/outlook/12/electronicbusinesscards" ver="1.0" layout="left" bgcolor="ffffff"><img xmlns="" align="fit" area="16" use="cardpicture"/><fld xmlns="" prop="name" align="left" dir="ltr" style="b" color="000000" size="10"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="telcell" align="left" dir="ltr" color="d48d2a" size="8"><label align="right" color="626262">Mobile</label></fld><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/><fld xmlns="" prop="blank" size="8"/></card>'+"`r`n"
$pourrev=Get-Date -Format "yyyyMMddThhmmssZ"+"`r`n"
$filecontent+="REV:$pourrev"+"`r`n"
$filecontent+="END:VCARD"+"`r`n"

$filecontent | out-file -Encoding "ASCII" -filepath $filename