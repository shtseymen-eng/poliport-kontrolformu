; Poliport Kontrol Formu Arsivi - NSIS Kurulum Scripti
; NSIS ile derlemek icin: makensis installer.nsi

!define APP_NAME "Poliport Kontrol Formu Arsivi"
!define APP_EXE  "PoliportArsiv.exe"
!define APP_VER  "1.0"
!define PUBLISHER "S. SEYMEN"
!define INSTALL_DIR "$LOCALAPPDATA\Poliport Kontrol Formu"

Name "${APP_NAME}"
OutFile "PoliportArsiv_Setup_v${APP_VER}.exe"
InstallDir "${INSTALL_DIR}"
InstallDirRegKey HKCU "Software\${APP_NAME}" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma

; Modern UI
!include "MUI2.nsh"

!define MUI_ABORTWARNING
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_WELCOMEPAGE_TITLE "Poliport Kontrol Formu Arşivi Kurulumu"
!define MUI_WELCOMEPAGE_TEXT "Bu sihirbaz Poliport Kontrol Formu Arşivi uygulamasını bilgisayarınıza kuracak.$\r$\n$\r$\nDevam etmek için İleri'ye tıklayın."
!define MUI_FINISHPAGE_RUN "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Uygulamayı şimdi başlat"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Turkish"

Section "Ana Uygulama" SecMain
    SetOutPath "$INSTDIR"
    File "${APP_EXE}"
    File "icon.ico"

    ; Masaustu kisayolu
    CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\icon.ico"

    ; Baslat menusu
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\icon.ico"
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\Kaldir.lnk" "$INSTDIR\Uninstall.exe"

    ; Kayit defteri (kullanici seviyesi - admin gerekmez)
    WriteRegStr HKCU "Software\${APP_NAME}" "InstallDir" "$INSTDIR"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayName" "${APP_NAME}"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayIcon" "$INSTDIR\icon.ico"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "Publisher" "${PUBLISHER}"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}" "DisplayVersion" "${APP_VER}"

    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Uninstall"
    Delete "$INSTDIR\${APP_EXE}"
    Delete "$INSTDIR\icon.ico"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"
    Delete "$DESKTOP\${APP_NAME}.lnk"
    RMDir /r "$SMPROGRAMS\${APP_NAME}"
    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"
    DeleteRegKey HKCU "Software\${APP_NAME}"
SectionEnd
