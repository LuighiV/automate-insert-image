# Setup project for editdocuments project in WIX

This project creates the installer for the Insert image in Document application which resides on 
editdocuments project. 

The installer is designed to be localizable for two languages: english and spanish.

For the creation it was mainly based on:

- Youtube tutorial: [How To Create Windows Installer MSI - .Net Core Wix](https://www.youtube.com/watch?v=6Yf-eDsRrnM)
- Repository reference: [Windows Installer Wix DotNet Core](https://github.com/angelsix/youtube/tree/develop/C%23%20General/Windows%20Installer%20Wix%20DotNet%20Core)

For the localization it was required to embed the mst language database as otherwise, an error occurs,
advertising about a wrong transfors when executing in machines which doesn't have the target 
language installed.

The reference how to obtain  the report is in: https://stackoverflow.com/q/18795351
and requires the following command:

```
msiexec /i IIDInstaller.msi /l*vx install.log
```

The result when installing in those machines which don't have the language database is:
```
MSI (c) (04:10) [23:46:56:094]: Looking for storage transform: 3082
MSI (c) (04:10) [23:46:56:094]: Note: 1: 2203 2: 3082 3: -2147287038 
DEBUG: Error 2203:  Database: 3082. Cannot open database file. System error -2147287038
1: 2203 2: 3082 3: -2147287038 
Error al aplicar las transformaciones. Compruebe que las rutas de las transformaciones especificadas son válidas.
3082
MSI (c) (04:10) [23:46:56:095]: Note: 1: 1708 
MSI (c) (04:10) [23:46:56:095]: Product: Insert image in Document -- Installation failed.

MSI (c) (04:10) [23:46:56:096]: Windows Installer installed the product. Product Name: Insert image in Document. Product Version: 1.0.0.3. Product Language: 1033. Manufacturer: Luighi Viton-Zorrilla. Installation success or error status: 1624.

MSI (c) (04:10) [23:46:56:098]: MainEngineThread is returning 1624
```

To solve it, I followed the references:
- https://stackoverflow.com/a/22405700
- http://www.installsite.org/pages/en/msi/articles/embeddedlang/
- https://www.codeproject.com/Articles/103749/Creating-a-Localized-Windows-Installer-Bootstrap-3
- http://www.geektieguy.com/2010/03/13/create-a-multi-lingual-multi-language-msi-using-wix-and-custom-build-scripts/

The procedure to package the msi with the language database requires of the Windows SDK that can
be downloaded from https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/.

In the project, I included the script WiSubStg.vb under tools folder, grabbed from the SDK repository.