This tool exports users, with a mail adress, of a specific OU to a CSV file and uploads this CSV to stuffbase to keep the users up to date.

example configuration in Configuration.ini:

[General]
Domain=wackershauser.de
OUs=OU=User-Accounts,OU=Benutzer,OU=Werk1,OU=UFT,DC=wackershauser,DC=de;OU=User-Accounts,OU=Benutzer,OU=Werk2,OU=UFT,DC=wackershauser,DC=de
DoNotExportMails=j.doe@uft.de;montage.mobil@uft.de;praktikant@uft.de;schichtfuehrer@uft.de;sortierung@uft.de;versand@uft.de
[Mapping aus AD]
Identifikation=mail
Vorname=givenName
Nachname=sn
Jobtitel=title
Abteilung=department
Unternehmensstandort=l
Email=mail
Telefonnummer=ipPhone
Mobiltelefonnummer=mobile
Nutzername=sAMAccountName
[Upload]
Token=NjMwY2ExMWIxNTdkODAwNDY1YjM1ZWYyOnNzRTBaWUlacTZMZUhQdXNFME9uem1JWlFrZEktflF3M3Z6Rld6dm1nUi5ha3BTT3g7Nm9NVUNyfVNYTGhLSlc=
Mapping=externalID,profile-field:firstName,profile-field:lastName,profile-field:jobtitel,profile-field:department,profile-field:standort,eMail,profile-field:telefonnummer,profile-field:mobiltelefonnummer,userName
ImportTag=uft
Dry=false

multiple OUs, or excluded useres have to be seperated by ;
