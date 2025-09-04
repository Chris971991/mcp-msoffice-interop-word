# VBA Automation Extension for MCP Word Server

This extension adds comprehensive VBA (Visual Basic for Applications) automation capabilities to the MCP Word server, enabling programmatic creation, manipulation, and execution of VBA code in Microsoft Word documents.

## Features Overview

### VBA Module Management
- **Create/Delete VBA Modules**: Support for all module types (standard, class, form, document)
- **Code Manipulation**: Add, modify, delete VBA procedures and functions
- **Import/Export**: Import modules from files and export to files
- **Project Management**: Set project properties, manage references
- **Protection**: Password protect/unprotect VBA projects

### ActiveX Control Integration
- **Control Creation**: Add various ActiveX controls (buttons, text boxes, combo boxes, etc.)
- **Property Management**: Set and get control properties programmatically
- **Event Handlers**: Bind VBA code to control events
- **UserForms**: Create custom dialog forms with multiple controls
- **Layout Management**: Group controls, set tab order

### Macro Execution & Testing
- **Run Macros**: Execute VBA macros with parameters
- **Testing Framework**: Test macros with validation
- **Debugging Support**: Debug information and breakpoint management
- **Error Handling**: Comprehensive error tracking and reporting
- **Immediate Execution**: Execute VBA code snippets instantly

### Document Integration
- **Event Handlers**: Document-level and application-level events
- **Auto Macros**: AutoExec, AutoOpen, AutoClose, etc.
- **Custom Properties**: Add document properties accessible from VBA
- **Startup Code**: Create self-executing documents

### Template & Deployment
- **Macro-Enabled Documents**: Save as .docm, .dotm formats
- **Templates**: Create document templates with embedded VBA
- **Add-ins**: Create and install Word add-ins
- **Digital Signing**: Sign VBA projects (where supported)
- **Project Export/Import**: Full VBA project deployment

## Available Tools

### VBA Module Tools
- `word_vba_createModule` - Create new VBA modules
- `word_vba_deleteModule` - Delete VBA modules
- `word_vba_getModuleCode` - Retrieve VBA code from modules
- `word_vba_setModuleCode` - Set complete VBA code in modules
- `word_vba_addProcedure` - Add procedures to modules
- `word_vba_deleteProcedure` - Delete procedures from modules
- `word_vba_listModules` - List all VBA modules
- `word_vba_importModule` - Import modules from files
- `word_vba_exportModule` - Export modules to files
- `word_vba_addReference` - Add library references
- `word_vba_removeReference` - Remove library references
- `word_vba_listReferences` - List all references
- `word_vba_setProjectProperties` - Set VBA project properties
- `word_vba_protectProject` - Protect/unprotect VBA project

### ActiveX Control Tools
- `word_activex_addControl` - Add ActiveX controls to document
- `word_activex_deleteControl` - Delete ActiveX controls
- `word_activex_setProperties` - Set control properties
- `word_activex_getProperties` - Get control properties
- `word_activex_addEventHandler` - Add VBA event handlers
- `word_activex_listControls` - List all ActiveX controls
- `word_activex_createUserForm` - Create UserForms
- `word_activex_addControlToForm` - Add controls to UserForms
- `word_activex_setTabOrder` - Set control tab order
- `word_activex_groupControls` - Group multiple controls

### VBA Execution Tools
- `word_vba_runMacro` - Execute VBA macros
- `word_vba_testMacro` - Test macros with validation
- `word_vba_debugCode` - Debug VBA code
- `word_vba_compileProject` - Compile VBA project
- `word_vba_addDocumentEvent` - Add document event handlers
- `word_vba_addApplicationEvent` - Add application event handlers
- `word_vba_createAutoMacro` - Create auto-executing macros
- `word_vba_getErrorInfo` - Get VBA error information
- `word_vba_executeImmediate` - Execute code in immediate window
- `word_vba_listMacros` - List available macros
- `word_vba_recordMacro` - Start recording macros
- `word_vba_stopRecording` - Stop macro recording

### Template & Deployment Tools
- `word_vba_saveAsMacroEnabled` - Save as macro-enabled format
- `word_vba_createTemplate` - Create document templates
- `word_vba_setMacroSecurity` - Set macro security levels
- `word_vba_signProject` - Sign VBA projects digitally
- `word_vba_createSelfExecuting` - Create self-executing documents
- `word_vba_exportProject` - Export entire VBA project
- `word_vba_importProject` - Import VBA project
- `word_vba_createAddIn` - Create Word add-ins
- `word_vba_installAddIn` - Install/uninstall add-ins
- `word_vba_createRibbon` - Create custom ribbon interfaces
- `word_vba_addCustomProperty` - Add custom document properties

## Example Usage Scenarios

### 1. Create a Document with Interactive Button

```json
{
  "tool": "word_vba_createModule",
  "parameters": {
    "moduleName": "ButtonActions",
    "moduleType": "standard",
    "code": "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub"
  }
}
```

```json
{
  "tool": "word_activex_addControl",
  "parameters": {
    "controlType": "commandButton",
    "name": "btnHello",
    "caption": "Click Me!",
    "width": 100,
    "height": 30
  }
}
```

```json
{
  "tool": "word_activex_addEventHandler",
  "parameters": {
    "controlName": "btnHello",
    "eventName": "Click",
    "vbaCode": "Call HelloWorld"
  }
}
```

### 2. Create a Self-Executing Document

```json
{
  "tool": "word_vba_createSelfExecuting",
  "parameters": {
    "filePath": "C:/temp/AutoDoc.docm",
    "startupCode": "MsgBox \"Welcome to this automated document!\"\nActiveDocument.Range.Text = \"This document was created with VBA automation!\"",
    "hideCode": true,
    "password": "mypassword"
  }
}
```

### 3. Create a UserForm with Multiple Controls

```json
{
  "tool": "word_activex_createUserForm",
  "parameters": {
    "formName": "DataEntry",
    "caption": "Enter Your Information",
    "width": 400,
    "height": 300
  }
}
```

```json
{
  "tool": "word_activex_addControlToForm",
  "parameters": {
    "formName": "DataEntry",
    "controlType": "label",
    "controlName": "lblName",
    "caption": "Name:",
    "left": 10,
    "top": 10
  }
}
```

```json
{
  "tool": "word_activex_addControlToForm",
  "parameters": {
    "formName": "DataEntry",
    "controlType": "textBox",
    "controlName": "txtName",
    "left": 60,
    "top": 10,
    "width": 200
  }
}
```

### 4. Create and Test a Custom Macro

```json
{
  "tool": "word_vba_createModule",
  "parameters": {
    "moduleName": "Calculations",
    "code": "Function AddNumbers(a As Double, b As Double) As Double\n    AddNumbers = a + b\nEnd Function"
  }
}
```

```json
{
  "tool": "word_vba_testMacro",
  "parameters": {
    "macroName": "Calculations.AddNumbers",
    "testData": [5, 3],
    "expectedResult": 8
  }
}
```

### 5. Create a Complete Business Document Automation

```json
{
  "tool": "word_vba_createModule",
  "parameters": {
    "moduleName": "ReportGenerator",
    "code": "Sub GenerateReport()\n    Dim doc As Document\n    Set doc = ActiveDocument\n    \n    ' Create header\n    doc.Range(0, 0).Text = \"Monthly Report\" & vbCrLf & vbCrLf\n    \n    ' Add current date\n    doc.Range.InsertAfter \"Generated on: \" & Format(Now, \"mmmm dd, yyyy\") & vbCrLf & vbCrLf\n    \n    ' Create table for data\n    Dim tbl As Table\n    Set tbl = doc.Tables.Add(doc.Range, 3, 2)\n    tbl.Cell(1, 1).Range.Text = \"Item\"\n    tbl.Cell(1, 2).Range.Text = \"Value\"\n    tbl.Cell(2, 1).Range.Text = \"Sales\"\n    tbl.Cell(2, 2).Range.Text = \"$10,000\"\n    tbl.Cell(3, 1).Range.Text = \"Expenses\"\n    tbl.Cell(3, 2).Range.Text = \"$8,000\"\n    \n    MsgBox \"Report generated successfully!\"\nEnd Sub"
  }
}
```

```json
{
  "tool": "word_activex_addControl",
  "parameters": {
    "controlType": "commandButton",
    "name": "btnGenerate",
    "caption": "Generate Report",
    "width": 120,
    "height": 30
  }
}
```

```json
{
  "tool": "word_activex_addEventHandler",
  "parameters": {
    "controlName": "btnGenerate",
    "eventName": "Click",
    "vbaCode": "Call GenerateReport"
  }
}
```

## Security Considerations

- **Macro Security**: VBA automation requires appropriate macro security settings
- **Trust Access**: "Trust access to the VBA project object model" must be enabled
- **Code Review**: Always review generated VBA code before execution
- **Permissions**: Ensure proper file system permissions for import/export operations
- **Digital Signing**: Consider signing VBA projects for deployment

## Limitations

- Some VBA features may require manual user interaction (security prompts)
- Ribbon customization requires additional Open XML manipulation
- Immediate window access is limited through COM automation
- Macro recording is simplified compared to native Word recording
- Digital signing may require external certificate management tools

## Prerequisites

- Microsoft Word with VBA support
- Appropriate macro security settings enabled
- "Trust access to VBA project object model" enabled in Trust Center
- Administrative privileges may be required for some operations

This comprehensive VBA automation extension enables the creation of sophisticated, interactive Word documents with full programmatic control over VBA code, ActiveX controls, and document behavior.