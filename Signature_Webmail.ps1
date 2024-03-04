# Si le format change il faudra changer la valeur des variables $FormatImg1 et $FormatImg2
################################## Partie a modifier ####################################

$Image1 = "Bandeau_Webmail.jpg"

#########################################################################################

# Création fichier de logs
$DateLog = Get-Date -Format "dd_MM_yyyy_HH_mm_ss"
$NomFichier = "Signature_Webmail-$DateLog.txt"
$Log = ".\Logs\$NomFichier"
New-Item -path $Log -ItemType File -Force

# Infos de connexion au serveur Exchange
$Admin_AD = "<Nom_Utilisateur>"
$MDPAdmin_AD = ConvertTo-SecureString "<Mot_de_Passe>" -AsPlainText -Force

# Creation d'une session Remote Powershell vers le Exchange Management Shell
$credentials = New-Object System.Management.Automation.PSCredential($Admin_AD, $MDPAdmin_AD)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<Adresse_Serveur_Exchange>/PowerShell/ -Authentication Kerberos -Credential $credentials

# Parcourir la liste des utilisateurs
$CSVFile = "\Chemin\vers\le\fichier\contenant\la\liste\des\utilisateur.csv"

##### Le sript peut egalement etre modifie en utilisant la commande Get-ADUser afin d'avoir la liste des utilisateurs #####

$CSVData = Import-CSV -Path $CSVFile -Delimiter ";" -Encoding Default

# Chemin vers le modele de signature
$modelePath = ".\Ressources\Signature_Webmail.html"

$CheminImage1 = ".\Ressources\Images\$Image1"
$ConversionImage1 = [Convert]::ToBase64String((Get-Content -Path $CheminImage1 -Encoding Byte))

# Permet d'ajouter les informations des comtpes AD necessaires pour la signature
$Tableau = @()

if (Test-Path $CSVFile) {
    # Parcourir chaque ligne du fichier CSV
    foreach ($Employe in $CSVData) {
        $Login = $Employe.Identifiant

        # Verifie que l'employe possede un identifiant
        if ($null -ne $Login) {
                $Utilisateur = Get-ADUser -Identity $Login -Properties officephone, streetaddress, department, displayname, postalcode, city, mail, title

                $Tableau += $Utilisateur
        }
    }
    # Creer un session Powershell vers le serveur Exchange
    Import-PSSession $Session -DisableNameChecking

    # Lit chaque ligne du tableau et recupere toutes les infos
    foreach ($Ligne in $Tableau) {
        $Nom = $Ligne.DisplayName
        $Telephone = $Ligne.OfficePhone
        $Adresse = $Ligne.StreetAddress
        $Fonction = $Ligne.Title
        $Mail = $Ligne.Mail
        $Etablissement = $Ligne.Department
        $CodePostale = $Ligne.PostalCode
        $Ville = $Ligne.City

        # Charger le modele de signature
        $modele = Get-Content $modelePath -Encoding UTF8

        # Ajoute les informations de l'utilisateur dans le fichier HTML de la signature
        $signature = $modele -replace '\{Nom\}', $Nom -replace '\{Tel\}', $Telephone -replace '\{Mail\}', $Mail -replace '\{Fonction\}', $Fonction -replace '\{Etablissement\}', $Etablissement -replace '\{Adresse\}', $Adresse -replace '\{CodePostale\}', $CodePostale -replace '\{Ville\}', $Ville -replace '\{Bandeau\}', $ConversionImage1

        Write-Output "Application de la signature pour $Nom" | Out-File -FilePath "$Log" -Append
        # Definit le fichier HTML comme signature de l'utilisateur
        Set-MailboxMessageConfiguration -Identity "$Mail" -SignatureHTML "$signature" -AutoAddSignature $true
    }

    # Fermeture de la session powershell a distance
    Remove-PSSession $Session
}
