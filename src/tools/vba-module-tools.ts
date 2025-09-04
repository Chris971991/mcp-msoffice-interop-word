import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";
import { debug } from "../utils/debug.js";

// VBA Module Types
const VbaModuleType = z.enum([
  "standard",     // vbext_ct_StdModule = 1
  "class",        // vbext_ct_ClassModule = 2
  "form",         // vbext_ct_MSForm = 3
  "document"      // vbext_ct_Document = 100
]);

// --- Tool: Create VBA Module ---
const createVbaModuleSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to create"),
  moduleType: VbaModuleType.optional().default("standard").describe("Type of module to create (standard, class, form, document)"),
  code: z.string().optional().describe("Initial VBA code to add to the module")
});

async function createVbaModuleTool(args: z.infer<typeof createVbaModuleSchema>): Promise<CallToolResult> {
  try {
    await wordService.createVbaModule(args.moduleName, args.moduleType, args.code);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created VBA module '${args.moduleName}' of type '${args.moduleType}'${args.code ? ' with initial code' : ''}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createVbaModuleTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create VBA module: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Delete VBA Module ---
const deleteVbaModuleSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to delete")
});

async function deleteVbaModuleTool(args: z.infer<typeof deleteVbaModuleSchema>): Promise<CallToolResult> {
  try {
    await wordService.deleteVbaModule(args.moduleName);
    return {
      content: [{ type: "text", text: `Successfully deleted VBA module '${args.moduleName}'` }],
    };
  } catch (error: any) {
    debug.error("Error in deleteVbaModuleTool:", error);
    return {
      content: [{ type: "text", text: `Failed to delete VBA module: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Get VBA Module Code ---
const getVbaModuleCodeSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to retrieve code from")
});

async function getVbaModuleCodeTool(args: z.infer<typeof getVbaModuleCodeSchema>): Promise<CallToolResult> {
  try {
    const code = await wordService.getVbaModuleCode(args.moduleName);
    return {
      content: [{ 
        type: "text", 
        text: `VBA Module '${args.moduleName}' code:\n\n${code}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in getVbaModuleCodeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get VBA module code: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set VBA Module Code ---
const setVbaModuleCodeSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to update"),
  code: z.string().describe("Complete VBA code to replace the module contents with")
});

async function setVbaModuleCodeTool(args: z.infer<typeof setVbaModuleCodeSchema>): Promise<CallToolResult> {
  try {
    await wordService.setVbaModuleCode(args.moduleName, args.code);
    return {
      content: [{ type: "text", text: `Successfully updated VBA module '${args.moduleName}' code` }],
    };
  } catch (error: any) {
    debug.error("Error in setVbaModuleCodeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set VBA module code: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add VBA Procedure ---
const addVbaProcedureSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to add the procedure to"),
  procedureCode: z.string().describe("Complete VBA procedure code including Sub/Function declaration and End Sub/Function"),
  position: z.number().optional().describe("Optional line number to insert the procedure at")
});

async function addVbaProcedureTool(args: z.infer<typeof addVbaProcedureSchema>): Promise<CallToolResult> {
  try {
    await wordService.addVbaProcedure(args.moduleName, args.procedureCode, args.position);
    return {
      content: [{ type: "text", text: `Successfully added procedure to VBA module '${args.moduleName}'` }],
    };
  } catch (error: any) {
    debug.error("Error in addVbaProcedureTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add VBA procedure: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Delete VBA Procedure ---
const deleteVbaProcedureSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module containing the procedure"),
  procedureName: z.string().describe("Name of the procedure to delete")
});

async function deleteVbaProcedureTool(args: z.infer<typeof deleteVbaProcedureSchema>): Promise<CallToolResult> {
  try {
    await wordService.deleteVbaProcedure(args.moduleName, args.procedureName);
    return {
      content: [{ type: "text", text: `Successfully deleted procedure '${args.procedureName}' from module '${args.moduleName}'` }],
    };
  } catch (error: any) {
    debug.error("Error in deleteVbaProcedureTool:", error);
    return {
      content: [{ type: "text", text: `Failed to delete VBA procedure: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: List VBA Modules ---
const listVbaModulesSchema = z.object({});

async function listVbaModulesTool(): Promise<CallToolResult> {
  try {
    const modules = await wordService.listVbaModules();
    return {
      content: [{ 
        type: "text", 
        text: modules.length > 0 
          ? `VBA Modules:\n${modules.map(m => `- ${m.name} (${m.type})`).join('\n')}` 
          : "No VBA modules found in the document"
      }],
    };
  } catch (error: any) {
    debug.error("Error in listVbaModulesTool:", error);
    return {
      content: [{ type: "text", text: `Failed to list VBA modules: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Import VBA Module from File ---
const importVbaModuleSchema = z.object({
  filePath: z.string().describe("Path to the VBA module file (.bas, .cls, .frm) to import")
});

async function importVbaModuleTool(args: z.infer<typeof importVbaModuleSchema>): Promise<CallToolResult> {
  try {
    const moduleName = await wordService.importVbaModule(args.filePath);
    return {
      content: [{ type: "text", text: `Successfully imported VBA module '${moduleName}' from '${args.filePath}'` }],
    };
  } catch (error: any) {
    debug.error("Error in importVbaModuleTool:", error);
    return {
      content: [{ type: "text", text: `Failed to import VBA module: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Export VBA Module to File ---
const exportVbaModuleSchema = z.object({
  moduleName: z.string().describe("Name of the VBA module to export"),
  filePath: z.string().describe("Path where the VBA module file should be saved")
});

async function exportVbaModuleTool(args: z.infer<typeof exportVbaModuleSchema>): Promise<CallToolResult> {
  try {
    await wordService.exportVbaModule(args.moduleName, args.filePath);
    return {
      content: [{ type: "text", text: `Successfully exported VBA module '${args.moduleName}' to '${args.filePath}'` }],
    };
  } catch (error: any) {
    debug.error("Error in exportVbaModuleTool:", error);
    return {
      content: [{ type: "text", text: `Failed to export VBA module: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add VBA Reference ---
const addVbaReferenceSchema = z.object({
  guid: z.string().optional().describe("GUID of the reference library"),
  major: z.number().optional().describe("Major version number"),
  minor: z.number().optional().describe("Minor version number"),
  description: z.string().optional().describe("Description or name of the library (alternative to GUID)")
});

async function addVbaReferenceTool(args: z.infer<typeof addVbaReferenceSchema>): Promise<CallToolResult> {
  try {
    await wordService.addVbaReference(args.guid, args.major, args.minor, args.description);
    return {
      content: [{ type: "text", text: `Successfully added VBA reference ${args.description || args.guid}` }],
    };
  } catch (error: any) {
    debug.error("Error in addVbaReferenceTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add VBA reference: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Remove VBA Reference ---
const removeVbaReferenceSchema = z.object({
  description: z.string().describe("Description or name of the library reference to remove")
});

async function removeVbaReferenceTool(args: z.infer<typeof removeVbaReferenceSchema>): Promise<CallToolResult> {
  try {
    await wordService.removeVbaReference(args.description);
    return {
      content: [{ type: "text", text: `Successfully removed VBA reference '${args.description}'` }],
    };
  } catch (error: any) {
    debug.error("Error in removeVbaReferenceTool:", error);
    return {
      content: [{ type: "text", text: `Failed to remove VBA reference: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: List VBA References ---
const listVbaReferencesSchema = z.object({});

async function listVbaReferencesTool(): Promise<CallToolResult> {
  try {
    const references = await wordService.listVbaReferences();
    return {
      content: [{ 
        type: "text", 
        text: references.length > 0 
          ? `VBA References:\n${references.map(r => `- ${r.description} (${r.guid})`).join('\n')}` 
          : "No VBA references found"
      }],
    };
  } catch (error: any) {
    debug.error("Error in listVbaReferencesTool:", error);
    return {
      content: [{ type: "text", text: `Failed to list VBA references: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set VBA Project Properties ---
const setVbaProjectPropertiesSchema = z.object({
  name: z.string().optional().describe("VBA project name"),
  description: z.string().optional().describe("VBA project description"),
  helpFile: z.string().optional().describe("Path to help file"),
  helpContextId: z.number().optional().describe("Help context ID")
});

async function setVbaProjectPropertiesTool(args: z.infer<typeof setVbaProjectPropertiesSchema>): Promise<CallToolResult> {
  try {
    await wordService.setVbaProjectProperties(args);
    return {
      content: [{ type: "text", text: "Successfully updated VBA project properties" }],
    };
  } catch (error: any) {
    debug.error("Error in setVbaProjectPropertiesTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set VBA project properties: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Protect/Unprotect VBA Project ---
const protectVbaProjectSchema = z.object({
  password: z.string().describe("Password to protect/unprotect the VBA project"),
  protect: z.boolean().describe("True to protect, false to unprotect")
});

async function protectVbaProjectTool(args: z.infer<typeof protectVbaProjectSchema>): Promise<CallToolResult> {
  try {
    await wordService.protectVbaProject(args.password, args.protect);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully ${args.protect ? 'protected' : 'unprotected'} VBA project` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in protectVbaProjectTool:", error);
    return {
      content: [{ type: "text", text: `Failed to ${args.protect ? 'protect' : 'unprotect'} VBA project: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerVbaModuleTools(server: McpServer) {
  server.tool(
    "word_vba_createModule",
    "Creates a new VBA module in the active Word document",
    createVbaModuleSchema.shape,
    createVbaModuleTool
  );
  
  server.tool(
    "word_vba_deleteModule",
    "Deletes a VBA module from the active Word document",
    deleteVbaModuleSchema.shape,
    deleteVbaModuleTool
  );
  
  server.tool(
    "word_vba_getModuleCode",
    "Gets the VBA code from a specific module",
    getVbaModuleCodeSchema.shape,
    getVbaModuleCodeTool
  );
  
  server.tool(
    "word_vba_setModuleCode",
    "Sets/replaces the complete VBA code in a module",
    setVbaModuleCodeSchema.shape,
    setVbaModuleCodeTool
  );
  
  server.tool(
    "word_vba_addProcedure",
    "Adds a VBA procedure (Sub or Function) to a module",
    addVbaProcedureSchema.shape,
    addVbaProcedureTool
  );
  
  server.tool(
    "word_vba_deleteProcedure",
    "Deletes a VBA procedure from a module",
    deleteVbaProcedureSchema.shape,
    deleteVbaProcedureTool
  );
  
  server.tool(
    "word_vba_listModules",
    "Lists all VBA modules in the active document",
    listVbaModulesSchema.shape,
    listVbaModulesTool
  );
  
  server.tool(
    "word_vba_importModule",
    "Imports a VBA module from a file",
    importVbaModuleSchema.shape,
    importVbaModuleTool
  );
  
  server.tool(
    "word_vba_exportModule",
    "Exports a VBA module to a file",
    exportVbaModuleSchema.shape,
    exportVbaModuleTool
  );
  
  server.tool(
    "word_vba_addReference",
    "Adds a reference to an external library in the VBA project",
    addVbaReferenceSchema.shape,
    addVbaReferenceTool
  );
  
  server.tool(
    "word_vba_removeReference",
    "Removes a reference from the VBA project",
    removeVbaReferenceSchema.shape,
    removeVbaReferenceTool
  );
  
  server.tool(
    "word_vba_listReferences",
    "Lists all references in the VBA project",
    listVbaReferencesSchema.shape,
    listVbaReferencesTool
  );
  
  server.tool(
    "word_vba_setProjectProperties",
    "Sets VBA project properties (name, description, help file, etc.)",
    setVbaProjectPropertiesSchema.shape,
    setVbaProjectPropertiesTool
  );
  
  server.tool(
    "word_vba_protectProject",
    "Protects or unprotects the VBA project with a password",
    protectVbaProjectSchema.shape,
    protectVbaProjectTool
  );
}