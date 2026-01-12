# SheetJS VSCode Extension

[![Visual Studio Marketplace Version](https://img.shields.io/visual-studio-marketplace/v/asadbek.sheetjs-demo)](https://marketplace.visualstudio.com/items?itemName=asadbek.sheetjs-demo)
[![Visual Studio Marketplace Downloads](https://img.shields.io/visual-studio-marketplace/d/asadbek.sheetjs-demo)](https://marketplace.visualstudio.com/items?itemName=asadbek.sheetjs-demo)
[![Visual Studio Marketplace Rating](https://img.shields.io/visual-studio-marketplace/r/asadbek.sheetjs-demo)](https://marketplace.visualstudio.com/items?itemName=asadbek.sheetjs-demo)


[![View on Marketplace](https://img.shields.io/badge/View%20on-Marketplace-blue?style=for-the-badge&logo=visualstudiocode)](https://marketplace.visualstudio.com/items?itemName=asadbek.sheetjs-demo)

---

<img src="https://git.sheetjs.com/asadbek064/sheetjs-vscode-extension/raw/branch/main/asset/sheetjs-vscode-extension-demo.gif" alt="SheetJS VSCode Extension Preview" width="600"/>

## SpreadSheet Viewer

Powered by [SheetJS](http://sheetjs.com/) a powerful VSCode extension that lets you view spreadsheets right in your editor. Works with **XLSX**, **XLS**, **CSV**, **ODS** and 30+ other formats.


## Key Features

- Caches workbooks and sheets to avoid re-parsing files
- Loads sheets on-demand when switching between them
- Automatically reloads when files are edited externally
- Handles mega large files with pagination
- Toggle viewer on/off for specific file extensions via command palette or context menu

## Supported File Formats

| [Supported File Formats](https://docs.sheetjs.com/docs/miscellany/formats/) |
| ---------------------- |
| *.xlsx                 |
| *.xlsm                 |
| *.xlsb                 |
| *.xls                  |
| *.xlw                  |
| *.xlr                  |
| *.numbers              |
| *.csv                  |
| *.dif                  |
| *.slk                  |
| *.sylk                 |
| *.prn                  |
| *.et                   |
| *.ods                  |
| *.fods                 |
| *.uos                  |
| *.dbf                  |
| *.wk1                  |
| *.wk3                  |
| *.wks                  |
| *.wk2                  |
| *.wk4                  |
| *.123                  |
| *.wq1                  |
| *.wq2                  |
| *.wb1                  |
| *.wb2                  |
| *.wb3                  |
| *.qpw                  |
| *.xlr                  |
| *.eth                  |

## Usage

### Disabling/Enabling the Viewer

You can easily disable the SheetJS viewer for specific file extensions:

<img src="https://git.sheetjs.com/asadbek064/sheetjs-vscode-extension/raw/branch/main/asset/toggle_context_menu.png" alt="SheetJS VSCode Extension Preview" width="600"/>
<img src="https://git.sheetjs.com/asadbek064/sheetjs-vscode-extension/raw/branch/main/asset/toggle_ext_via_palette.png" alt="SheetJS VSCode Extension Preview" width="600"/>

**Command Palette** (Ctrl/Cmd+Shift+P):
- `SheetJS: Disable Viewer for Current File Extension` - Switches to default text editor for that extension
- `SheetJS: Enable Viewer for Current File Extension` - Re-enables the viewer

**Context Menu**: Right-click any spreadsheet file in the Explorer to access the same commands.

**Built-in VSCode**: Right-click any file and select "Open With..." to choose between SheetJS Viewer and other editors.

## Getting Started
Want to integrate SheetJS in your own VSCode extension? Check out our [detailed tutorial](https://docs.sheetjs.com/docs/) to learn how to implement these capabilities in your projects.

## Development

To run the extension in development mode, install dependencies with `pnpm install` and press F5 in VSCode. This opens a new Extension Development Host window where you can test the extension by opening any spreadsheet file.


Build for production with `pnpm run package`.

## Publishing
```bash
npx vsce login foo
npx vsce publish
```

## Learn More
For more information on using this extension and integrating SheetJS capabilities in your own projects, visit our [documentation](https://docs.sheetjs.com/docs/).

---
_Created by Asadbek Karimov  | [contact@asadk.dev](mailto:contact@asadk.dev) | [asadk.dev](https://asadk.dev)_
