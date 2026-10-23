; Instalador do Campos Cotação — Inno Setup 6.
; Ao compilar, publica automaticamente o código atual em Release com todas as dependências.
; Requer o SDK do .NET 10 instalado e o comando dotnet disponível no PATH.
; Essa pasta contém o aplicativo, suas bibliotecas e o .NET Desktop Runtime.
; O computador de destino não precisa baixar nem instalar o .NET separadamente.
; Compile este arquivo no Inno Setup para gerar installer/Output/InstaladorCamposCotacao.exe.

#define MyAppName "Campos Cotação"
#define MyAppPublisher "Campos"
#define MyAppExeName "COTACAO_INSUMO.exe"
#define MyAppIcon "Resources\Calculator.ico"
#define ReleaseDir "bin\Release\Instalador\win-x64"

; Aguarda a publicação e impede o empacotamento de arquivos antigos em caso de erro.
#define PublishArgs "publish " + AddQuotes(SourcePath + "COTACAO_INSUMO.csproj") + " -c Release -r win-x64 --self-contained true -o " + AddQuotes(SourcePath + ReleaseDir) + " -p:PublishSingleFile=false -p:PublishTrimmed=false"
#if Exec("dotnet.exe", PublishArgs, SourcePath, 1) != 0
  #error Falha ao publicar o aplicativo atual. Corrija o erro do dotnet publish e compile novamente; a geração do instalador foi interrompida.
#endif

#if !FileExists(SourcePath + ReleaseDir + "\COTACAO_INSUMO.exe")
  #error A publicação não gerou COTACAO_INSUMO.exe na pasta Release esperada.
#endif
#if !FileExists(SourcePath + ReleaseDir + "\coreclr.dll") || !FileExists(SourcePath + ReleaseDir + "\hostpolicy.dll") || !FileExists(SourcePath + ReleaseDir + "\System.Windows.Forms.dll")
  #error A publicação não gerou o runtime completo necessário para o instalador.
#endif

#define MyAppVersion GetVersionNumbersString(SourcePath + ReleaseDir + "\" + MyAppExeName)

[Setup]
AppId={{940C0D54-C14B-479D-BF81-49350228BCA3}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Campos Cotação
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
OutputDir=installer\Output
OutputBaseFilename=InstaladorCamposCotacao
SetupIconFile={#MyAppIcon}
UninstallDisplayIcon={app}\{#MyAppIcon}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
SetupLogging=yes

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar um atalho na Área de Trabalho"; GroupDescription: "Atalhos adicionais:"; Flags: unchecked

[Files]
; Copia integralmente a publicação Release, preservando suas subpastas.
Source: "{#ReleaseDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Inclui o ícone mesmo quando ele não foi copiado pela publicação.
Source: "{#MyAppIcon}"; DestDir: "{app}\Resources"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppIcon}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\{#MyAppIcon}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar {#MyAppName}"; Flags: nowait postinstall skipifsilent
