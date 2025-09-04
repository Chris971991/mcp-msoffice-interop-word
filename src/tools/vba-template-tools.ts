import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";
import { debug } from "../utils/debug.js";

// File format enums for saving
const MacroFileFormat = z.enum([
  "docm",     // Word Macro-Enabled Document
  "dotm",     // Word Macro-Enabled Template
  "docx",     // Word Document (will remove macros)
  "dotx",     // Word Template (will remove macros)
  "doc",      // Word 97-2003 Document
  "dot"       // Word 97-2003 Template
]);

// --- Tool: Save as Macro-Enabled Document ---
const saveAsMacroEnabledSchema = z.object({
  filePath: z.string().describe("Path to save the document"),
  format: MacroFileFormat.optional().default("docm").describe("File format for saving"),
  createBackup: z.boolean().optional().default(false).describe("Whether to create a backup of the original")
});

async function saveAsMacroEnabledTool(args: z.infer<typeof saveAsMacroEnabledSchema>): Promise<CallToolResult> {
  try {
    await wordService.saveAsMacroEnabled(args.filePath, args.format, args.createBackup);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully saved document as macro-enabled format at '${args.filePath}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in saveAsMacroEnabledTool:", error);
    return {
      content: [{ type: "text", text: `Failed to save as macro-enabled: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create Document Template ---
const createDocumentTemplateSchema = z.object({
  templatePath: z.string().describe("Path to save the template"),
  includeMacros: z.boolean().optional().default(true).describe("Whether to include macros in the template"),
  protectTemplate: z.boolean().optional().default(false).describe("Whether to protect the template"),
  password: z.string().optional().describe("Password for template protection")
});

async function createDocumentTemplateTool(args: z.infer<typeof createDocumentTemplateSchema>): Promise<CallToolResult> {
  try {
    await wordService.createDocumentTemplate(
      args.templatePath,
      args.includeMacros,
      args.protectTemplate,
      args.password
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created document template at '${args.templatePath}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createDocumentTemplateTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create template: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Macro Security Settings ---
const setMacroSecuritySchema = z.object({
  level: z.enum(["veryHigh", "high", "medium", "low"]).describe("Macro security level"),
  trustAccessVBOM: z.boolean().optional().describe("Whether to trust access to VBA object model")
});

async function setMacroSecurityTool(args: z.infer<typeof setMacroSecuritySchema>): Promise<CallToolResult> {
  try {
    await wordService.setMacroSecurity(args.level, args.trustAccessVBOM);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully set macro security to '${args.level}'${args.trustAccessVBOM ? ' with VBA object model access' : ''}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in setMacroSecurityTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set macro security: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Sign VBA Project ---
const signVbaProjectSchema = z.object({
  certificatePath: z.string().optional().describe("Path to digital certificate file"),
  certificateName: z.string().optional().describe("Name of installed certificate to use"),
  timestamp: z.boolean().optional().default(true).describe("Whether to add timestamp to signature")
});

async function signVbaProjectTool(args: z.infer<typeof signVbaProjectSchema>): Promise<CallToolResult> {
  try {
    await wordService.signVbaProject(args.certificatePath, args.certificateName, args.timestamp);
    return {
      content: [{ 
        type: "text", 
        text: "Successfully signed VBA project with digital certificate" 
      }],
    };
  } catch (error: any) {
    debug.error("Error in signVbaProjectTool:", error);
    return {
      content: [{ type: "text", text: `Failed to sign VBA project: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create Self-Executing Document ---
const createSelfExecutingDocSchema = z.object({
  filePath: z.string().describe("Path to save the self-executing document"),
  startupCode: z.string().describe("VBA code to execute on document open"),
  hideCode: z.boolean().optional().default(false).describe("Whether to hide/protect the VBA code"),
  password: z.string().optional().describe("Password for VBA project protection")
});

async function createSelfExecutingDocTool(args: z.infer<typeof createSelfExecutingDocSchema>): Promise<CallToolResult> {
  try {
    await wordService.createSelfExecutingDocument(
      args.filePath,
      args.startupCode,
      args.hideCode,
      args.password
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created self-executing document at '${args.filePath}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createSelfExecutingDocTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create self-executing document: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Export VBA Project ---
const exportVbaProjectSchema = z.object({
  exportPath: z.string().describe("Directory path to export all VBA modules to"),
  includeReferences: z.boolean().optional().default(true).describe("Whether to export reference information"),
  includeProjectInfo: z.boolean().optional().default(true).describe("Whether to export project properties")
});

async function exportVbaProjectTool(args: z.infer<typeof exportVbaProjectSchema>): Promise<CallToolResult> {
  try {
    const exportedFiles = await wordService.exportVbaProject(
      args.exportPath,
      args.includeReferences,
      args.includeProjectInfo
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully exported VBA project to '${args.exportPath}'\nExported files:\n${exportedFiles.join('\n')}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in exportVbaProjectTool:", error);
    return {
      content: [{ type: "text", text: `Failed to export VBA project: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Import VBA Project ---
const importVbaProjectSchema = z.object({
  importPath: z.string().describe("Directory path containing VBA modules to import"),
  clearExisting: z.boolean().optional().default(false).describe("Whether to clear existing VBA code before importing"),
  importReferences: z.boolean().optional().default(true).describe("Whether to import reference information")
});

async function importVbaProjectTool(args: z.infer<typeof importVbaProjectSchema>): Promise<CallToolResult> {
  try {
    const importedModules = await wordService.importVbaProject(
      args.importPath,
      args.clearExisting,
      args.importReferences
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully imported VBA project from '${args.importPath}'\nImported modules:\n${importedModules.join('\n')}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in importVbaProjectTool:", error);
    return {
      content: [{ type: "text", text: `Failed to import VBA project: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create Add-In ---
const createAddInSchema = z.object({
  addInPath: z.string().describe("Path to save the Word add-in (.dotm)"),
  addInName: z.string().describe("Name of the add-in"),
  description: z.string().optional().describe("Description of the add-in"),
  autoLoad: z.boolean().optional().default(false).describe("Whether to automatically load the add-in")
});

async function createAddInTool(args: z.infer<typeof createAddInSchema>): Promise<CallToolResult> {
  try {
    await wordService.createAddIn(
      args.addInPath,
      args.addInName,
      args.description,
      args.autoLoad
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created Word add-in '${args.addInName}' at '${args.addInPath}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createAddInTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create add-in: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Install Add-In ---
const installAddInSchema = z.object({
  addInPath: z.string().describe("Path to the add-in file to install"),
  install: z.boolean().optional().default(true).describe("True to install, false to uninstall")
});

async function installAddInTool(args: z.infer<typeof installAddInSchema>): Promise<CallToolResult> {
  try {
    await wordService.installAddIn(args.addInPath, args.install);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully ${args.install ? 'installed' : 'uninstalled'} add-in from '${args.addInPath}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in installAddInTool:", error);
    return {
      content: [{ type: "text", text: `Failed to ${args.install ? 'install' : 'uninstall'} add-in: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create Ribbon Customization ---
const createRibbonCustomizationSchema = z.object({
  ribbonXML: z.string().describe("Custom Ribbon XML definition"),
  callbacks: z.record(z.string()).optional().describe("Map of callback names to VBA procedure names")
});

async function createRibbonCustomizationTool(args: z.infer<typeof createRibbonCustomizationSchema>): Promise<CallToolResult> {
  try {
    await wordService.createRibbonCustomization(args.ribbonXML, args.callbacks);
    return {
      content: [{ 
        type: "text", 
        text: "Successfully created custom ribbon interface" 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createRibbonCustomizationTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create ribbon customization: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add Custom Document Property ---
const addCustomDocPropertySchema = z.object({
  name: z.string().describe("Name of the custom property"),
  value: z.any().describe("Value of the property"),
  linkToContent: z.boolean().optional().default(false).describe("Whether to link property to document content")
});

async function addCustomDocPropertyTool(args: z.infer<typeof addCustomDocPropertySchema>): Promise<CallToolResult> {
  try {
    await wordService.addCustomDocumentProperty(args.name, args.value, args.linkToContent);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added custom document property '${args.name}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addCustomDocPropertyTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add custom property: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerVbaTemplateTools(server: McpServer) {
  server.tool(
    "word_vba_saveAsMacroEnabled",
    "Saves the document in a macro-enabled format",
    saveAsMacroEnabledSchema.shape,
    saveAsMacroEnabledTool
  );
  
  server.tool(
    "word_vba_createTemplate",
    "Creates a document template with optional macros and protection",
    createDocumentTemplateSchema.shape,
    createDocumentTemplateTool
  );
  
  server.tool(
    "word_vba_setMacroSecurity",
    "Sets macro security settings",
    setMacroSecuritySchema.shape,
    setMacroSecurityTool
  );
  
  server.tool(
    "word_vba_signProject",
    "Signs the VBA project with a digital certificate",
    signVbaProjectSchema.shape,
    signVbaProjectTool
  );
  
  server.tool(
    "word_vba_createSelfExecuting",
    "Creates a self-executing document with startup VBA code",
    createSelfExecutingDocSchema.shape,
    createSelfExecutingDocTool
  );
  
  server.tool(
    "word_vba_exportProject",
    "Exports the entire VBA project to a directory",
    exportVbaProjectSchema.shape,
    exportVbaProjectTool
  );
  
  server.tool(
    "word_vba_importProject",
    "Imports a VBA project from a directory",
    importVbaProjectSchema.shape,
    importVbaProjectTool
  );
  
  server.tool(
    "word_vba_createAddIn",
    "Creates a Word add-in with VBA functionality",
    createAddInSchema.shape,
    createAddInTool
  );
  
  server.tool(
    "word_vba_installAddIn",
    "Installs or uninstalls a Word add-in",
    installAddInSchema.shape,
    installAddInTool
  );
  
  server.tool(
    "word_vba_createRibbon",
    "Creates custom ribbon interface with VBA callbacks",
    createRibbonCustomizationSchema.shape,
    createRibbonCustomizationTool
  );
  
  server.tool(
    "word_vba_addCustomProperty",
    "Adds a custom document property accessible from VBA",
    addCustomDocPropertySchema.shape,
    addCustomDocPropertyTool
  );
}