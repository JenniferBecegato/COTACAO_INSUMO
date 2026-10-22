; Instalador do Campos Cotação — Inno Setup 6.
; Antes de compilar, atualize a publicação Release com todas as dependências:
; dotnet publish COTACAO_INSUMO.csproj -c Release -r win-x64 --self-contained true -o bin/Release/Instalador/win-x64 -p:PublishSingleFile=false -p:PublishTrimmed=false
; Essa pasta contém o aplicativo, suas bibliotecas e o .NET Desktop Runtime.
; O computador de destino não precisa baixar nem instalar o .NET separadamente.
; Compile este arquivo no Inno Setup para gerar installer/Output/InstaladorCamposCotacao.exe.

#define MyAppName "Campos Cotação"
#define MyAppPublisher "Campos"
#define MyAppExeName "COTACAO_INSUMO.exe"
#define ReleaseDir "bin\Release\Instalador\win-x64"

#if !FileExists(SourcePath + ReleaseDir + "\COTACAO_INSUMO.exe")
  #error Execute o comando dotnet publish indicado no início deste arquivo para preparar a pasta Release.
#endif
#if !FileExists(SourcePath + ReleaseDir + "\coreclr.dll") || !FileExists(SourcePath + ReleaseDir + "\hostpolicy.dll") || !FileExists(SourcePath + ReleaseDir + "\System.Windows.Forms.dll")
  #error A publicação precisa incluir o runtime. Execute dotnet publish com --self-contained true conforme indicado no início deste arquivo.
#endif

#define MyAppVersion GetFileVersion(SourcePath + ReleaseDir + "\" + MyAppExeName)

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
SetupIconFile=Resources\Calculator.ico
UninstallDisplayIcon={app}\Resources\Calculator.ico
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
Source: "Resources\Calculator.ico"; DestDir: "{app}\Resources"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\Resources\Calculator.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\Resources\Calculator.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Executar {#MyAppName}"; Flags: nowait postinstall skipifsilent
