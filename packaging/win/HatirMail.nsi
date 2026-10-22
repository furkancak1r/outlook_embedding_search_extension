; HatirMail setup program (NSIS 3, Unicode). Build on macOS: sh packaging/build_win.sh (it calls makensis).
; It wraps the tested scripts and does not repeat their work:
;   - unpacks the package files (build/pkg/HatirMail) into $PLUGINSDIR (a temp folder that NSIS deletes at the end),
;   - runs install.ps1 from there (the same steps as the zip install: files, offline wheels, model check, HKCU
;     registration, COM activation check; -Machine is added when UAC is off and the user is an administrator),
;   - runs firewall.ps1: a Windows Firewall rule blocks outbound connections of the worker (python.exe of the install
;     folder). It asks for administrator approval (UAC) once; without it, HatırMail works without the rule.
;   - adds "HatırMail" to Windows Settings > Apps (HKCU; uninstall.exe in the install folder runs uninstall.ps1,
;     which also removes the firewall rule).
; Per user (RequestExecutionLevel user); only the firewall step asks for administrator approval. Not signed: Windows
; SmartScreen asks once.
; Isolated test: HatirMail-Kurulum.exe /S /TESTDIR=<folder> installs the files and the firewall rule for that folder
; only (no registry key, no Apps entry). The uninstaller touches the registry only for the install folder that the
; Apps entry names; for any other folder (a test install) it removes the files and the rule of that folder only.
; ponytail: no file list here; install.ps1 and uninstall.ps1 own the file layout.

Unicode true
ManifestDPIAware true
SetCompressor /SOLID lzma
SetCompressorDictSize 64
RequestExecutionLevel user

!ifndef PKG
  !error "Give -DPKG=<staged package folder> (build/pkg/HatirMail)"
!endif
!ifndef OUTFILE
  !error "Give -DOUTFILE=<exe path>"
!endif
!ifndef VERSION
  !error "Give -DVERSION=<x.y.z>"
!endif

!include MUI2.nsh
!include LogicLib.nsh
!include FileFunc.nsh
!include x64.nsh

Name "HatırMail"
OutFile "${OUTFILE}"
InstallDir "$LOCALAPPDATA\HatirMail"
ShowInstDetails show
ShowUninstDetails show
BrandingText "HatırMail ${VERSION}"
VIProductVersion "${VERSION}.0"
VIAddVersionKey /LANG=1055 "ProductName" "HatırMail"
VIAddVersionKey /LANG=1055 "FileDescription" "HatırMail kurulumu"
VIAddVersionKey /LANG=1055 "FileVersion" "${VERSION}"
VIAddVersionKey /LANG=1055 "ProductVersion" "${VERSION}"
VIAddVersionKey /LANG=1055 "LegalCopyright" "HatırMail"

!define UNKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\HatirMail"
!define PSEXE '"$WINDIR\Sysnative\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass'

Var TestDir   ; /TESTDIR=: isolated test install (files only)
Var Machine   ; "1": also register the COM class per machine (UAC off + administrator)
Var Reg       ; uninstaller: "1" = this folder is the registered install

; ---- pages (own Turkish texts: plain language; some NSIS default texts have spelling errors) ---------------------
!define MUI_ABORTWARNING
!define MUI_ABORTWARNING_TEXT "Kurulumdan çıkmak istiyor musunuz? HatırMail kurulmaz."
!define MUI_UNABORTWARNING
!define MUI_UNABORTWARNING_TEXT "Kaldırmadan çıkmak istiyor musunuz?"   ; set before the first page (MUI2 sets defaults there)
!define MUI_WELCOMEPAGE_TITLE "HatırMail kurulumu"
!define MUI_WELCOMEPAGE_TEXT "HatırMail, Outlook'ta hatırladığınız birkaç ayrıntıyla e-postalarınızı bulur. Yapay zekâ yalnız bu bilgisayarda çalışır; postalarınız hiçbir yere gönderilmez.$\r$\n$\r$\nKurulum yalnız sizin kullanıcı hesabınıza yapılır. Sonunda Windows bir kez yönetici onayı isteyebilir: bu onayla bir güvenlik duvarı kuralı HatırMail'in arama işleminin internete çıkmasını engeller. Onay vermezseniz HatırMail yine çalışır.$\r$\n$\r$\nKurulum sırasında yaklaşık 2,5 GB, sonra 1,3 GB boş disk alanı gerekir.$\r$\n$\r$\nDevam etmeden önce Outlook'u kapatın. Sonra İleri'ye basın."
!insertmacro MUI_PAGE_WELCOME
!define MUI_PAGE_HEADER_TEXT "Kuruluyor"
!define MUI_PAGE_HEADER_SUBTEXT "HatırMail kuruluyor. Bu birkaç dakika sürebilir."
!define MUI_INSTFILESPAGE_FINISHHEADER_TEXT "Kurulum bitti"
!define MUI_INSTFILESPAGE_FINISHHEADER_SUBTEXT "HatırMail kuruldu."
!define MUI_INSTFILESPAGE_ABORTHEADER_TEXT "Kurulum bitmedi"
!define MUI_INSTFILESPAGE_ABORTHEADER_SUBTEXT "HatırMail kurulamadı. Ayrıntılar aşağıda."
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TITLE "HatırMail kuruldu"
!define MUI_FINISHPAGE_TEXT "Outlook'u açın. Giriş sekmesinde HatırMail ile Ara düğmesine basın.$\r$\n$\r$\nİlk kullanımda İndeks sekmesinde İndekslemeyi başlat'a basın. Büyük bir posta kutusunda ilk indeksleme saatler sürebilir; bu sırada Outlook'u kullanabilirsiniz.$\r$\n$\r$\nKaldırmak için: Windows Ayarlar > Uygulamalar > HatırMail."
!insertmacro MUI_PAGE_FINISH

!define MUI_UNCONFIRMPAGE_TEXT_TOP "HatırMail bu bilgisayardan kaldırılır. Yerel arama indeksi, ayarlar ve kayıtlı aramalar da silinir. Outlook'taki postalarınız değişmez.$\r$\n$\r$\nDevam etmeden önce Outlook'u kapatın."
!insertmacro MUI_UNPAGE_CONFIRM
!define MUI_PAGE_HEADER_TEXT "Kaldırılıyor"
!define MUI_PAGE_HEADER_SUBTEXT "HatırMail kaldırılıyor."
!define MUI_INSTFILESPAGE_FINISHHEADER_TEXT "Kaldırma bitti"
!define MUI_INSTFILESPAGE_FINISHHEADER_SUBTEXT "HatırMail kaldırıldı."
; NSIS 3.13 InstallFiles.nsh unsets ...ABORTWARNING_TEXT after the page, not ...ABORTHEADER_TEXT: unset it here.
!undef MUI_INSTFILESPAGE_ABORTHEADER_TEXT
!undef MUI_INSTFILESPAGE_ABORTHEADER_SUBTEXT
!define MUI_INSTFILESPAGE_ABORTHEADER_TEXT "Kaldırma bitmedi"
!define MUI_INSTFILESPAGE_ABORTHEADER_SUBTEXT "HatırMail kaldırılamadı. Ayrıntılar aşağıda."
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Turkish"

; ---- helpers -----------------------------------------------------------------------------------------------------
; Wait until Outlook is closed: the add-in DLL is loaded in it, and the scripts refuse to change the registration
; while it runs (the same Get-Process check as in install.ps1). Exit code 0 = Outlook runs.
!macro WaitOutlook UN
  Function ${UN}WaitOutlook
    ${Do}
      nsExec::Exec '${PSEXE} -Command "if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }"'
      Pop $0
      ${If} $0 != "0"
        ${Break}
      ${EndIf}
      ${If} ${Silent}
        SetErrorLevel 3
        Abort "Outlook açık. Outlook'u kapatıp yeniden deneyin."
      ${EndIf}
      MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Outlook açık. Outlook'u kapatın, sonra Yeniden Dene'ye basın." IDRETRY retry
      SetErrorLevel 3
      Abort "Outlook açık olduğu için işlem yapılmadı."
      retry:
    ${Loop}
  FunctionEnd
!macroend
!insertmacro WaitOutlook ""
!insertmacro WaitOutlook "un."

; 64-bit PowerShell (Sysnative: this setup program is 32-bit); output goes to the details list. $0 = exit code.
!macro RunPs SCRIPT ARGS
  nsExec::ExecToLog /OEM '${PSEXE} -File "${SCRIPT}" ${ARGS}'
  Pop $0
!macroend

Function .onInit
  ${IfNot} ${RunningX64}
    MessageBox MB_ICONSTOP "HatırMail 64 bit Windows ister."
    Abort
  ${EndIf}
  ${GetParameters} $R0
  ClearErrors
  ${GetOptions} $R0 "/TESTDIR=" $TestDir
  ${If} ${Errors}
    StrCpy $TestDir ""
  ${EndIf}
  ${If} $TestDir != ""
    StrCpy $INSTDIR $TestDir
  ${EndIf}
  ; UAC off + administrator: Outlook runs with a full administrator token, and COM reads only HKLM classes for it.
  ; Then -Machine (HKLM\Software\Classes, HatırMail keys only). This process has the same full token: no prompt.
  StrCpy $Machine ""
  SetRegView 64
  ReadRegDWORD $0 HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" "EnableLUA"
  SetRegView lastused
  UserInfo::GetAccountType
  Pop $1
  ${If} $0 == "0"
  ${AndIf} $1 == "Admin"
    StrCpy $Machine "1"
  ${EndIf}
FunctionEnd

Section "HatırMail"
  ${If} $TestDir == ""
    Call WaitOutlook
  ${EndIf}
  InitPluginsDir
  SetOutPath "$PLUGINSDIR\pkg"
  DetailPrint "Dosyalar açılıyor…"
  SetDetailsPrint listonly
  File /r "${PKG}/*"
  SetDetailsPrint both
  ${If} $TestDir != ""
    DetailPrint "Test kurulumu (yalnız dosyalar, kayıt defteri değişmez): $INSTDIR"
    !insertmacro RunPs "$PLUGINSDIR\pkg\install.ps1" '-InstallDir "$INSTDIR" -SkipRegistration'
  ${ElseIf} $Machine == "1"
    DetailPrint "UAC kapalı ve hesap yönetici: eklenti makine düzeyinde de kaydedilir (yalnız HatırMail anahtarları)."
    !insertmacro RunPs "$PLUGINSDIR\pkg\install.ps1" '-InstallDir "$INSTDIR" -Machine'
  ${Else}
    !insertmacro RunPs "$PLUGINSDIR\pkg\install.ps1" '-InstallDir "$INSTDIR"'
  ${EndIf}
  SetOutPath "$TEMP"   ; leave $PLUGINSDIR: NSIS deletes it at the end
  ${If} $0 != "0"
    DetailPrint "install.ps1 çıkış kodu: $0"
    SetErrorLevel 2
    Abort "Kurulum bitmedi. Yukarıdaki ayrıntıların ekran görüntüsünü gönderin."
  ${EndIf}
  ; Firewall rule: not a condition of the install (HatırMail works without it). 1 = no rule, 2 = no effect here.
  DetailPrint "Güvenlik duvarı kuralı ekleniyor. Windows yönetici onayı isteyebilir."
  !insertmacro RunPs "$INSTDIR\firewall.ps1" '-InstallDir "$INSTDIR" -ParentWindow $HWNDPARENT'
  ${If} $0 == "0"
    DetailPrint "Güvenlik duvarı kuralı etkin: HatırMail'in arama işlemi internete çıkamaz."
  ${ElseIf} $0 == "2"
    DetailPrint "Güvenlik duvarı kuralı eklendi, ama engelin çalıştığı doğrulanamadı."
    ${IfNot} ${Silent}
      MessageBox MB_ICONINFORMATION "Güvenlik duvarı kuralı eklendi, ama engelin çalıştığı doğrulanamadı. Windows Güvenlik Duvarı kapalı olabilir ya da bir şirket ayarı yerel kuralları uygulamıyor olabilir.$\r$\n$\r$\nHatırMail çalışır. Bilgisayarınızın yöneticisine danışın."
    ${EndIf}
  ${Else}
    DetailPrint "Güvenlik duvarı kuralı eklenmedi (yönetici onayı yok ya da hata, çıkış kodu $0)."
    ${IfNot} ${Silent}
      MessageBox MB_ICONINFORMATION "Güvenlik duvarı kuralı eklenmedi: yönetici onayı verilmedi ya da bir hata oldu (ayrıntılar kurulum listesinde).$\r$\n$\r$\nHatırMail çalışır. Kuralı sonra eklemek için kurulum programını yeniden çalıştırın; indeksiniz korunur."
    ${EndIf}
  ${EndIf}
  WriteUninstaller "$INSTDIR\uninstall.exe"
  ${If} $TestDir == ""
    ${GetSize} "$INSTDIR" "/S=0K" $1 $2 $3
    WriteRegStr HKCU "${UNKEY}" "DisplayName" "HatırMail"
    WriteRegStr HKCU "${UNKEY}" "DisplayVersion" "${VERSION}"
    WriteRegStr HKCU "${UNKEY}" "Publisher" "HatırMail"
    WriteRegStr HKCU "${UNKEY}" "InstallLocation" "$INSTDIR"
    WriteRegStr HKCU "${UNKEY}" "UninstallString" '"$INSTDIR\uninstall.exe"'
    WriteRegStr HKCU "${UNKEY}" "QuietUninstallString" '"$INSTDIR\uninstall.exe" /S'
    WriteRegStr HKCU "${UNKEY}" "HatirMailMachine" "$Machine"
    WriteRegDWORD HKCU "${UNKEY}" "EstimatedSize" $1
    WriteRegDWORD HKCU "${UNKEY}" "NoModify" 1
    WriteRegDWORD HKCU "${UNKEY}" "NoRepair" 1
  ${EndIf}
SectionEnd

; ---- uninstaller: uninstall.ps1 removes the files and the add-in registration -----------------------------------
Section "Uninstall"
  ReadRegStr $1 HKCU "${UNKEY}" "InstallLocation"
  StrCpy $Reg ""
  ${If} $1 == $INSTDIR   ; NSIS compares strings without case
    StrCpy $Reg "1"
  ${EndIf}
  ${If} $Reg == "1"
    Call un.WaitOutlook
    ReadRegStr $Machine HKCU "${UNKEY}" "HatirMailMachine"
  ${EndIf}
  ${If} ${FileExists} "$INSTDIR\uninstall.ps1"
    ; Run a copy: uninstall.ps1 deletes itself and the folder that it runs from.
    InitPluginsDir
    CopyFiles /SILENT "$INSTDIR\uninstall.ps1" "$PLUGINSDIR\uninstall.ps1"
    ${If} $Reg != "1"
      DetailPrint "Kayıtlı kurulum değil: yalnız dosyalar ve bu klasörün güvenlik duvarı kuralı silinir, kayıt defteri değişmez."
      !insertmacro RunPs "$PLUGINSDIR\uninstall.ps1" '-InstallDir "$INSTDIR" -SkipRegistration -ParentWindow $HWNDPARENT'
    ${ElseIf} $Machine == "1"
      !insertmacro RunPs "$PLUGINSDIR\uninstall.ps1" '-InstallDir "$INSTDIR" -Machine -ParentWindow $HWNDPARENT'
    ${Else}
      !insertmacro RunPs "$PLUGINSDIR\uninstall.ps1" '-InstallDir "$INSTDIR" -ParentWindow $HWNDPARENT'
    ${EndIf}
    ${If} $0 != "0"
      DetailPrint "uninstall.ps1 çıkış kodu: $0"
      SetErrorLevel 2
      Abort "Kaldırma bitmedi. Yukarıdaki ayrıntıların ekran görüntüsünü gönderin."
    ${EndIf}
  ${EndIf}
  Delete "$INSTDIR\uninstall.exe"
  RMDir "$INSTDIR"   ; only when empty
  ${If} $Reg == "1"
    DeleteRegKey HKCU "${UNKEY}"
  ${EndIf}
SectionEnd
