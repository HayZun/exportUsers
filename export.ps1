param([switch]$Elevated)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

#get-date
$currentdate = (Get-Date -Format "dd/MM/yyyy").Replace("/","_") + "_users"

#getmonth
$month = (Get-Culture).DateTimeFormat.GetMonthName((Get-Date).Month)

#get retrieve data
$ouList = @("OU=Users,OU=2.xxxxxx,DC=xxxxxx,DC=com")

# Recherchez les utilisateurs dans les emplacements sp�cifi�s
$users = $ouList | ForEach-Object {
    Get-ADUser -SearchBase $_ -Filter * -Properties passwordlastset, passwordneverexpires, enabled |
    Where-Object {$_.DistinguishedName -notlike '*ou=lock,*'}
}

# Triez les utilisateurs par nom
$sortedUsers = $users | Sort-Object Name

# S�lectionnez les propri�t�s n�cessaires et exportez-les vers un fichier CSV
$sortedUsers | Select-Object @{
    Name = "Nom"
    Expression = {$_.Name}
},
@{
    Name = "date_dernier_changement_de_mot_de_passe"
    Expression = {$_.passwordlastset}
},
@{
    Name = "Statut du compte"
    Expression = {if ($_.enabled) {"Actif"} else {"Inactif"}}
},
@{
    Name = "Section"
    Expression = {
        if ($_ -match "OU=Utilisateur") {"cient"}
        else {"Autre"}
    }
}| Export-csv -Path "" -NotypeInformation -Encoding utf8

# Sender and Recipient Info
$MailFrom = ""
$MailTo = ""
#$MailTo = ""
# Sender Credentials
$Username = ""
$Password = "
# Server Info
$SmtpServer = ""
$SmtpPort = ""
# Message stuff
$MessageSubject = "$month - Etats mots de passe utilisateur"
$Message = New-Object System.Net.Mail.MailMessage $MailFrom,$MailTo
$Message.IsBodyHTML = $true
$Message.Subject = $MessageSubject

#cc
$copy = New-Object MailAddress("email1","email2");
$message.CC.Add($copy);

$Message.Body = @"
<p> text <p>
"@

#attachment
$attachment = new-object System.Net.Mail.Attachment "C:\Users\"
$message.Attachments.Add("C:\Users\")

# Construct the SMTP client object, credentials, and send
$Smtp = New-Object Net.Mail.SmtpClient($SmtpServer,$SmtpPort)
$Smtp.EnableSsl = $true
$Smtp.Credentials = New-Object System.Net.NetworkCredential($Username,$Password)
$Smtp.Send($Message)
