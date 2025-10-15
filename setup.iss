; Script d'installation pour Gestion Moyenne Semestre 9 IC3 IT/ESTP
; Créé avec Inno Setup 6

[Setup]
; Informations de l'application
AppName=Gestion Moyenne Semestre 9 IC3 IT
AppVersion=1.0.0
AppPublisher=ESTP - IC3 IT
AppPublisherURL=https://www.estp.fr
AppSupportURL=https://www.estp.fr
AppUpdatesURL=https://www.estp.fr

; Dossier d'installation par défaut
DefaultDirName={autopf}\GestionMoyenneIC3IT
DefaultGroupName=ESTP IC3 IT
AllowNoIcons=yes

; Dossier de sortie de l'installateur
OutputDir=Installer
OutputBaseFilename=Setup_GestionMoyenne_IC3IT_v1.0
SetupIconFile=photo AI-modified.ico

; Compression
Compression=lzma2/max
SolidCompression=yes

; Interface moderne
WizardStyle=modern

; Privilèges
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Langue
ShowLanguageDialog=no

; Divers
DisableProgramGroupPage=yes
DisableReadyPage=no

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "Créer un raccourci sur le bureau"; GroupDescription: "Raccourcis:"

[Files]
; Copier tous les fichiers publiés
Source: "bin\Release\net8.0-windows\win-x64\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Raccourci dans le menu Démarrer
Name: "{group}\Gestion Moyenne IC3 IT"; Filename: "{app}\GESTION_MOYENNE_SEMESTRE_9_IC3__IT.exe"
Name: "{group}\Désinstaller"; Filename: "{uninstallexe}"

; Raccourci sur le bureau (si coché)
Name: "{autodesktop}\Gestion Moyenne IC3 IT"; Filename: "{app}\GESTION_MOYENNE_SEMESTRE_9_IC3__IT.exe"; Tasks: desktopicon

[Run]
; Proposer de lancer l'application après installation
Filename: "{app}\GESTION_MOYENNE_SEMESTRE_9_IC3__IT.exe"; Description: "Lancer Gestion Moyenne IC3 IT"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Supprimer les fichiers créés par l'application
Type: filesandordirs; Name: "{app}\Exports"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  MsgBox('Bienvenue dans l''installation de' + #13#10 + 
         'Gestion Moyenne Semestre 9 IC3 IT/ESTP' + #13#10#13#10 +
         'Cliquez sur Suivant pour continuer.', 
         mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    MsgBox('Installation terminée avec succès !' + #13#10#13#10 + 
           'Vous pouvez maintenant utiliser l''application' + #13#10 +
           'pour gérer vos moyennes du Semestre 9.', 
           mbInformation, MB_OK);
  end;
end;
