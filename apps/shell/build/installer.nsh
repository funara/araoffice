; installer.nsh - NSIS script for ARAOffice Windows Registration & File Associations
; Registers ARAOffice with Windows "Default Apps", "Open With", and specific file icons

!macro customWelcomePage
  !insertmacro MUI_PAGE_WELCOME
!macroend


!macro customInstall
  DetailPrint "Registering ARAOffice default apps capabilities and file associations..."

  ; 1. Register ARAOffice under Windows RegisteredApplications
  WriteRegStr SHELL_CONTEXT "Software\RegisteredApplications" "ARAOffice" "Software\ARAOffice\Capabilities"

  ; 2. Define ARAOffice Capabilities (for Windows Settings -> Apps -> Default Apps)
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities" "ApplicationName" "ARAOffice"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities" "ApplicationDescription" "ARAOffice - Modern AI-powered Office Suite for Documents, Spreadsheets, Presentations, and PDFs"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities" "ApplicationIcon" "$appExe,0"

  ; File associations under Capabilities
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".docx" "ARAOffice.docx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".doc" "ARAOffice.docx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".xlsx" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".xls" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".csv" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".pptx" "ARAOffice.pptx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".ppt" "ARAOffice.pptx"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".pdf" "ARAOffice.pdf"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".md" "ARAOffice.md"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".markdown" "ARAOffice.md"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".drawio" "ARAOffice.drawio"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".dio" "ARAOffice.drawio"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".mmd" "ARAOffice.mmd"
  WriteRegStr SHELL_CONTEXT "Software\ARAOffice\Capabilities\FileAssociations" ".mermaid" "ARAOffice.mmd"

  ; 3. Register Applications\ARAOffice.exe (used by Windows "Open With" list & Default Apps)
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe" "FriendlyAppName" "ARAOffice"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe" "ApplicationCompany" "ARAOffice"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\DefaultIcon" "" "$appExe,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\shell\open" "FriendlyAppName" "ARAOffice"
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\shell\open\command" "" '"$appExe" "%1"'

  ; SupportedTypes list for ARAOffice.exe
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".docx" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".doc" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".xlsx" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".xls" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".csv" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".pptx" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".ppt" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".pdf" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".md" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".markdown" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".drawio" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".dio" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".mmd" ""
  WriteRegStr SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe\SupportedTypes" ".mermaid" ""

  ; 4. Ensure ProgIDs have correct DefaultIcon and AppUserModelId
  ; DOCX
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx" "" "ARA Docs"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx" "FriendlyTypeName" "ARA Docs"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx\DefaultIcon" "" "$INSTDIR\resources\docx.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.docx\shell\open" "FriendlyAppName" "ARAOffice"

  ; XLSX
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx" "" "ARA Sheets"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx" "FriendlyTypeName" "ARA Sheets"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx\DefaultIcon" "" "$INSTDIR\resources\xlsx.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx\shell\open" "FriendlyAppName" "ARAOffice"

  ; PPTX
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx" "" "ARA Slides"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx" "FriendlyTypeName" "ARA Slides"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx\DefaultIcon" "" "$INSTDIR\resources\pptx.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pptx\shell\open" "FriendlyAppName" "ARAOffice"

  ; PDF
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf" "" "ARA PDF"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf" "FriendlyTypeName" "ARA PDF"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf\DefaultIcon" "" "$INSTDIR\resources\pdf.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.pdf\shell\open" "FriendlyAppName" "ARAOffice"

  ; MD
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md" "" "ARA Notes"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md" "FriendlyTypeName" "ARA Notes"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md\DefaultIcon" "" "$INSTDIR\resources\md.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.md\shell\open" "FriendlyAppName" "ARAOffice"
  ; Diagram
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio" "" "ARA Diagrams"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio" "FriendlyTypeName" "ARA Diagrams"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio\DefaultIcon" "" "$INSTDIR\resources\diagram.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.drawio\shell\open" "FriendlyAppName" "ARAOffice"

  ; Mermaid
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd" "" "ARA Diagram Code"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd" "FriendlyTypeName" "ARA Diagram Code"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd" "AppUserModelID" "com.araoffice.app"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd\DefaultIcon" "" "$INSTDIR\resources\diagram-code.ico,0"
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd\shell\open\command" "" '"$appExe" "%1"'
  WriteRegStr SHELL_CONTEXT "Software\Classes\ARAOffice.mmd\shell\open" "FriendlyAppName" "ARAOffice"

  ; 5. Register OpenWithList and OpenWithProgids for all extensions
  WriteRegNone SHELL_CONTEXT "Software\Classes\.docx\OpenWithProgids" "ARAOffice.docx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.docx\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.doc\OpenWithProgids" "ARAOffice.docx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.doc\OpenWithList\ARAOffice.exe" "" ""

  WriteRegNone SHELL_CONTEXT "Software\Classes\.xlsx\OpenWithProgids" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.xlsx\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.xls\OpenWithProgids" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.xls\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.csv\OpenWithProgids" "ARAOffice.xlsx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.csv\OpenWithList\ARAOffice.exe" "" ""

  WriteRegNone SHELL_CONTEXT "Software\Classes\.pptx\OpenWithProgids" "ARAOffice.pptx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.pptx\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.ppt\OpenWithProgids" "ARAOffice.pptx"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.ppt\OpenWithList\ARAOffice.exe" "" ""

  WriteRegNone SHELL_CONTEXT "Software\Classes\.pdf\OpenWithProgids" "ARAOffice.pdf"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.pdf\OpenWithList\ARAOffice.exe" "" ""

  WriteRegNone SHELL_CONTEXT "Software\Classes\.md\OpenWithProgids" "ARAOffice.md"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.md\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.markdown\OpenWithProgids" "ARAOffice.md"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.markdown\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.drawio\OpenWithProgids" "ARAOffice.drawio"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.drawio\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.dio\OpenWithProgids" "ARAOffice.drawio"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.dio\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.mmd\OpenWithProgids" "ARAOffice.mmd"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.mmd\OpenWithList\ARAOffice.exe" "" ""
  WriteRegNone SHELL_CONTEXT "Software\Classes\.mermaid\OpenWithProgids" "ARAOffice.mmd"
  WriteRegStr SHELL_CONTEXT "Software\Classes\.mermaid\OpenWithList\ARAOffice.exe" "" ""

  ; 6. Refresh shell notification
  System::Call "shell32::SHChangeNotify(i 0x08000000, i 0x1000, i 0, i 0)"
!macroend

!macro customUnInstall
  ; Remove RegisteredApplications and Capabilities
  DeleteRegValue SHELL_CONTEXT "Software\RegisteredApplications" "ARAOffice"
  DeleteRegKey SHELL_CONTEXT "Software\ARAOffice"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\Applications\ARAOffice.exe"

  ; Remove ProgIDs
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.docx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.xlsx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.pptx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.pdf"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.md"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.drawio"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\ARAOffice.mmd"

  ; Remove OpenWith entries
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.docx\OpenWithProgids" "ARAOffice.docx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.docx\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.doc\OpenWithProgids" "ARAOffice.docx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.doc\OpenWithList\ARAOffice.exe"

  DeleteRegValue SHELL_CONTEXT "Software\Classes\.xlsx\OpenWithProgids" "ARAOffice.xlsx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.xlsx\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.xls\OpenWithProgids" "ARAOffice.xlsx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.xls\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.csv\OpenWithProgids" "ARAOffice.xlsx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.csv\OpenWithList\ARAOffice.exe"

  DeleteRegValue SHELL_CONTEXT "Software\Classes\.pptx\OpenWithProgids" "ARAOffice.pptx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.pptx\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.ppt\OpenWithProgids" "ARAOffice.pptx"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.ppt\OpenWithList\ARAOffice.exe"

  DeleteRegValue SHELL_CONTEXT "Software\Classes\.pdf\OpenWithProgids" "ARAOffice.pdf"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.pdf\OpenWithList\ARAOffice.exe"

  DeleteRegValue SHELL_CONTEXT "Software\Classes\.md\OpenWithProgids" "ARAOffice.md"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.md\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.markdown\OpenWithProgids" "ARAOffice.md"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.markdown\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.drawio\OpenWithProgids" "ARAOffice.drawio"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.drawio\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.dio\OpenWithProgids" "ARAOffice.drawio"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.dio\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.mmd\OpenWithProgids" "ARAOffice.mmd"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.mmd\OpenWithList\ARAOffice.exe"
  DeleteRegValue SHELL_CONTEXT "Software\Classes\.mermaid\OpenWithProgids" "ARAOffice.mmd"
  DeleteRegKey SHELL_CONTEXT "Software\Classes\.mermaid\OpenWithList\ARAOffice.exe"

  ; Refresh shell notification
  System::Call "shell32::SHChangeNotify(i 0x08000000, i 0x1000, i 0, i 0)"
!macroend
