!define APP_NAME "SmartRingQC"
!define APP_VERSION "1.0"
!define APP_EXE "SmartRingQC.exe"

Name "${APP_NAME} ${APP_VERSION}"
OutFile "SmartRingQC_Installer.exe"
InstallDir "$PROGRAMFILES\${APP_NAME}"
InstallDirRegKey HKCU "Software\${APP_NAME}" ""

Section "Install"
  SetOutPath "$INSTDIR"
  File "dist\SmartRingQC.exe"
  File /r "SmartRingQC_App\captures"
  File /r "SmartRingQC_App\database"
  File /r "SmartRingQC_App\reports"
  File "SmartRingQC_App\HOW_TO_USE.txt"

  ; Desktop shortcut
  CreateShortcut "$DESKTOP\${APP_NAME}.lnk" \
    "$INSTDIR\${APP_EXE}" "" \
    "$INSTDIR\${APP_EXE}" 0

  ; Start menu shortcut
  CreateDirectory "$SMPROGRAMS\${APP_NAME}"
  CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" \
    "$INSTDIR\${APP_EXE}"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\${APP_EXE}"
  Delete "$DESKTOP\${APP_NAME}.lnk"
  RMDir /r "$SMPROGRAMS\${APP_NAME}"
  RMDir /r "$INSTDIR"
SectionEnd
