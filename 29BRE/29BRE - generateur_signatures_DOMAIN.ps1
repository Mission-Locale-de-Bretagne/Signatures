[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Connect-ExchangeOnline


# Cible le ou les utilisateurs concernÃ©s

$mailboxes = Get-ExoMailBox -Filter {UserPrincipalName -like "*.rapicault@mission-locale-brest.org" -and RecipientTypeDetails -eq 'UserMailbox' -and CustomAttribute15 -eq "29BRE"} | Select-Object firstname,lastname,title,phone,mobilephone,userprincipalname,streetaddress,postalcode,city,Office,Company
#$CompanyName = Get-AzureADUser -Filter "UserPrincipalName eq 'vmarie@mlfougeres.onmicrosoft.com'" | Select-Object -ExpandProperty CompanyName

# Chemin vers le template HTML
$templateSignatureHTML = Get-Content -Path "C:\Users\AliceQUERIC\ARMLB\Pays de Brest - t5_informatique\Administration courante\Utilisateurs\Signatures emails\29BRE-template-signature.html" -raw

# Boucle pour chaque utilisateur
foreach ($mailbox in $mailboxes) { 
    $user = Get-User -Identity $mailbox.UserPrincipalName | Select-Object FirstName, LastName, Title, Phone, MobilePhone, UserPrincipalName, StreetAddress, PostalCode, City, Office, Company, Department
	$signatureHTML = $templateSignatureHTML 
	# VÃ©rification qu'il s'agit bien d'un utilisateur
 
	if ($user.firstname -and $user.FirstName -ne "Scan" -and -not $user.FirstName.Contains("Compte administratif") -and -not $user.LastName.Contains("Reply") -and $user.FirstName -ne "Test") {
        $address = ""
        $street = ""
        $postalcode = ""
        $city = ""
        $phone = "" 

		# Selon le nom de l'antenne, on spécifie l'adresse de celle-ci.
        if ($user.Office.Contains("Siège administratif") -and $user.Department -ne "Aller vers")
        {
            $address = "Mission Locale du Pays de Brest - Siège Administratif"
            $street = "7 Rue Keravel - BP 71028"
            $postalcode = "29210 "
            $city = "Brest Cedex 1"
            $phone = "02 98 43 51 00" 
        }
        elseif ($user.Office.Contains("Aller vers"))
        {
            $address = "Mission Locale du Pays de Brest - Antenne Aller Vers"
            $street = "7 Rue Keravel"
            $postalcode = "29200"
            $city = "Brest"
            $phone = "02 98 83 84 58" 
        }
        elseif ($user.Office.Contains("Jaurès"))
        {
            $address = "Mission Locale du Pays de Brest - Antenne Jaurès"
            $street = "253 rue Jean Jaurès"
            $postalcode = "29200"
            $city = "Brest"
            $phone = "02 98 41 06 90" 
        }
        elseif ($user.Office.Contains("Rive Droite"))
        {
            $address = "Mission Locale du Pays de Brest - Antenne Rive Droite"
            $street = "45 rue Dupuy de Lôme"
            $postalcode = "29200"
            $city = "Brest"
            $phone = "02 98 49 70 85" 
        }
        elseif ($user.Office.Contains("Bellevue"))
        {
            $address = "Mission Locale du Pays de Brest - Antenne Bellevue"
            $street = "Place Napoléon III - CC B2"
            $postalcode = "29200"
            $city = "Brest"
            $phone = "02 98 47 25 53" 
        }
        # Cas des salariés présents sur deux antennes selon le jour de la semaine.
        # Dans de tel cas, la syntaxe dans Entra ID est la suivante : "Nom_Antenne_1 (Jours_Antenne_1) / Nom_Antenne_2 (Jours_Antenne_2)"
        elseif ($user.Office.Contains("/"))
        {

            $antenne_1 = $user.Office.split("/")[0].Split("(")[0]#.Split(" ")[1]
            $jours_antenne_1 = $user.Office.split("/")[0].Split("(")[1].Split(")")[0]

            $antenne_2 = $user.Office.split("/")[1].Trim(" ").Split("(")[0]
            $jours_antenne_2 = $user.Office.split("/")[1].Split("(")[1].Trim(")")

            $address = "Mission Locale du Pays de Brest"

            $postalcode = "$antenne_1 ($jours_antenne_1) :"

            if ($antenne_1.Contains("Landerneau"))
            {
                $postalcode = "$postalcode 02 98 21 52 29"
            } 
            elseif ($antenne_1.Contains("Lannilis"))
            {
                $postalcode = "$postalcode 02 98 04 14 54"
            } 
            elseif ($antenne_1.Contains("Plabennec"))
            {
                $postalcode = "$postalcode 02 30 06 00 33"
            } 
            elseif ($antenne_1.Contains("Lesneven"))
            {
                $postalcode = "$postalcode 02 98 21 19 32"
            } 
            elseif ($antenne_1.Contains("Lanrivoar"))
            {
                $postalcode = "$postalcode 02 98 32 43 05"
            } 
            elseif ($antenne_1.Contains("Crozon"))
            {
                $postalcode = "$postalcode 02 98 26 23 21"
            } 
            elseif ($antenne_1.Contains("Pont de Buis"))
            {
                $postalcode = "$postalcode 02 98 73 07 95"
            } 
            elseif ($antenne_1.Contains("Chateaulin"))
            {
                $postalcode = "$postalcode 02 98 16 14 27"
            } 

            $phone = "$antenne_2 ($jours_antenne_2) :"
            if ($antenne_2.Contains("Landerneau"))
            {
                $phone = "$phone 02 98 21 52 29"
            } 
            elseif ($antenne_2.Contains("Lannilis"))
            {
                $phone = "$phone 02 98 04 14 54"
            } 
            elseif ($antenne_2.Contains("Plabennec"))
            {
                $phone = "$phone 02 30 06 00 33"
            } 
            elseif ($antenne_2.Contains("Lesneven"))
            {
                $phone = "$phone 02 98 21 19 32"
            } 
            elseif ($antenne_2.Contains("Lanrivoar"))
            {
                $phone = "$phone 02 98 32 43 05"
            } 
            elseif ($antenne_2.Contains("Crozon"))
            {
                $phone = "$phone 02 98 26 23 21"
            } 
            elseif ($antenne_2.Contains("Pont de Buis"))
            {
                $phone = "$phone 02 98 73 07 95"
            } 
            elseif ($antenne_2.Contains("Chateaulin"))
            {
                $phone = "$phone 02 98 16 14 27"
            }  
        }
        # Cas spécifique des antennes CEJ et SEE
        elseif ($user.Department -ne $null -and ($user.Department.Contains("CEJ") -or $user.Department.Contains("SEE")))
        {
            if ($user.Department.Contains("CEJ"))
            {
                #Write-Host $user.lastname $user.FirstName "est dans le service CEJ"
                $address = "Mission Locale du Pays de Brest - Service Contrat Engagement Jeune"
                $street = "9 rue de Vendée"
                $postalcode = "29200"
                $city = "Brest"
                $phone = "02 98 83 84 50" 
            }
            elseif ($user.Department.Contains("SEE"))
            {
                #Write-Host $user.lastname $user.FirstName "est dans le service SEE"
                $address = "Mission Locale du Pays de Brest - Service Emploi Entreprises"
                $street = "7 rue de Vendée"
                $postalcode = "29200"
                $city = "Brest"
                $phone = "02 98 43 51 30" 
            }
        }
        # Le reste concerne les antennes rurales
        else {

            $address = "Mission Locale du Pays de Brest"
            $postalcode = $user.Office.Split("(")[0].Trim(" ")

            Write-Host $user.FirstName $user.LastName "est sur" $postalcode

            if ($user.Office.Contains("Landerneau"))
            {
                $phone = "02 98 21 52 29"
            } 
            elseif ($user.Office.Contains("Lannilis"))
            {
                $phone = "02 98 04 14 54"
            } 
            elseif ($user.Office.Contains("Plabennec"))
            {
                $phone = "02 30 06 00 33"
            } 
            elseif ($user.Office.Contains("Lesneven"))
            {
                $phone = "02 98 21 19 32"
            } 
            elseif ($user.Office.Contains("Lanrivoar"))
            {
                $phone = "02 98 32 43 05"
            } 
            elseif ($user.Office.Contains("Crozon"))
            {
                $phone = "02 98 26 23 21"
            } 
            elseif ($user.Office.Contains("Pont de Buis"))
            {
                $phone = "02 98 73 07 95"
            } 
            elseif ($user.Office.Contains("Chateaulin"))
            {
                $phone = "02 98 16 14 27"
            } 
            elseif ($user.Office.Contains("IFAC"))
            {
                $phone = "02 98 16 14 27"
            } 
        }

		# Remplacement des tags dans le template par les valeurs correspondantes
		$signatureHTML = $signatureHTML.Replace("{First name}", $user.firstname) 
		$signatureHTML = $signatureHTML.Replace("{Last name}", $user.lastname) 
		$signatureHTML = $signatureHTML.Replace("{Title}", $user.title)
		$signatureHTML = $signatureHTML.Replace("{Address}", $address)
		$signatureHTML = $signatureHTML.Replace("{Street}", $street) 
		$signatureHTML = $signatureHTML.Replace("{PostalCode}", $postalcode) 
		$signatureHTML = $signatureHTML.Replace("{City}", $city)
        $signatureHTML = $signatureHTML.Replace("{Phone}", $phone)
        # Exclusion des téléphones mobiles pour les salariées qui en font la demande expresse
        if ($user.LastName.Contains("GLEVAREC") -or $user.LastName.Contains("MAHE")) {
            $signatureHTML = $signatureHTML.Replace("{MobilePhone}", "")
        }
        else {
		    $signatureHTML = $signatureHTML.Replace("{MobilePhone}", $user.mobilephone)
        }
        
        # Modification du logo pour les personnes dont le poste est financé par le FSE.
        if ($user.Department -eq "SEE" -or $user.LastName -eq "RAPICAULT" -or $user.LastName -eq "VERDES" -or $user.LastName -eq "SOUSSEING") {
            $signatureHTML = $signatureHTML.Replace("BrestV2.png", "Brest_FSE_2.png")
            $signatureHTML = $signatureHTML.Replace('width="150"', 'width="300"')
            $signatureHTML = $signatureHTML.Replace('height="150"', 'height="200"')
        }

        # ecriture de fichiers HTML pour vérification
        $filename = "C:\Users\AliceQUERIC\ARMLB\Pays de Brest - t5_informatique\Administration courante\Utilisateurs\Signatures emails\test_signatures\" + $user.firstname + "_" + $user.LastName + ".html"
        #Out-File -FilePath $filename -InputObject $signatureHTML

        # Selon l'itération, on regénère l'ensemble des signatures ou juste les perosnnes impactées par la génération souhaitée
        if ($user.LastName -eq "RAPICAULT") {
            #Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true
            Out-File -FilePath $filename -InputObject $signatureHTML
        }
        # Ou bien toutes les signatures mais ca risque de raler
        #Set-MailboxMessageConfiguration -Identity $user.userPrincipalName -signatureHTML $signatureHTML -AutoAddSignature $true -AutoAddSignatureOnReply $true
        #Out-File -FilePath $filename -InputObject $signatureHTML


	}
}

Disconnect-ExchangeOnline