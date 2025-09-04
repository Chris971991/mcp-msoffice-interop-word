import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";
import { debug } from "../utils/debug.js";

// --- Tool: Run VBA Macro ---
const runVbaMacroSchema = z.object({
  macroName: z.string().describe("Name of the macro to run (can include module name like 'Module1.MacroName')"),
  parameters: z.array(z.any()).optional().describe("Optional array of parameters to pass to the macro")
});

async function runVbaMacroTool(args: z.infer<typeof runVbaMacroSchema>): Promise<CallToolResult> {
  try {
    const result = await wordService.runVbaMacro(args.macroName, args.parameters);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully executed macro '${args.macroName}'${result ? `\nResult: ${JSON.stringify(result)}` : ''}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in runVbaMacroTool:", error);
    return {
      content: [{ type: "text", text: `Failed to run VBA macro: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Test VBA Macro ---
const testVbaMacroSchema = z.object({
  macroName: z.string().describe("Name of the macro to test"),
  testData: z.any().optional().describe("Optional test data/parameters"),
  expectedResult: z.any().optional().describe("Optional expected result for validation")
});

async function testVbaMacroTool(args: z.infer<typeof testVbaMacroSchema>): Promise<CallToolResult> {
  try {
    const testResult = await wordService.testVbaMacro(args.macroName, args.testData, args.expectedResult);
    return {
      content: [{ 
        type: "text", 
        text: `Macro test '${args.macroName}':\n` +
              `Status: ${testResult.success ? 'PASSED' : 'FAILED'}\n` +
              `${testResult.result ? `Result: ${JSON.stringify(testResult.result)}\n` : ''}` +
              `${testResult.error ? `Error: ${testResult.error}\n` : ''}` +
              `${testResult.executionTime ? `Execution time: ${testResult.executionTime}ms` : ''}`
      }],
    };
  } catch (error: any) {
    debug.error("Error in testVbaMacroTool:", error);
    return {
      content: [{ type: "text", text: `Failed to test VBA macro: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Debug VBA Code ---
const debugVbaCodeSchema = z.object({
  moduleName: z.string().describe("Module containing the code to debug"),
  procedureName: z.string().describe("Procedure name to debug"),
  breakpoints: z.array(z.number()).optional().describe("Optional array of line numbers for breakpoints")
});

async function debugVbaCodeTool(args: z.infer<typeof debugVbaCodeSchema>): Promise<CallToolResult> {
  try {
    const debugInfo = await wordService.debugVbaCode(args.moduleName, args.procedureName, args.breakpoints);
    return {
      content: [{ 
        type: "text", 
        text: `Debug information for '${args.moduleName}.${args.procedureName}':\n` +
              `${JSON.stringify(debugInfo, null, 2)}`
      }],
    };
  } catch (error: any) {
    debug.error("Error in debugVbaCodeTool:", error);
    return {
      content: [{ type: "text", text: `Failed to debug VBA code: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Compile VBA Project ---
const compileVbaProjectSchema = z.object({});

async function compileVbaProjectTool(): Promise<CallToolResult> {
  try {
    const compileResult = await wordService.compileVbaProject();
    return {
      content: [{ 
        type: "text", 
        text: compileResult.success 
          ? "VBA project compiled successfully" 
          : `VBA compilation failed:\n${compileResult.errors?.join('\n') || 'Unknown error'}`
      }],
    };
  } catch (error: any) {
    debug.error("Error in compileVbaProjectTool:", error);
    return {
      content: [{ type: "text", text: `Failed to compile VBA project: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add Document Event Handler ---
const addDocumentEventHandlerSchema = z.object({
  eventName: z.string().describe("Name of the document event (e.g., 'Document_Open', 'Document_Close', 'Document_New')"),
  vbaCode: z.string().describe("VBA code for the event handler (without the Sub declaration)")
});

async function addDocumentEventHandlerTool(args: z.infer<typeof addDocumentEventHandlerSchema>): Promise<CallToolResult> {
  try {
    await wordService.addDocumentEventHandler(args.eventName, args.vbaCode);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added document event handler '${args.eventName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addDocumentEventHandlerTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add document event handler: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add Application Event Handler ---
const addApplicationEventHandlerSchema = z.object({
  eventName: z.string().describe("Name of the application event (e.g., 'DocumentOpen', 'DocumentBeforeSave', 'WindowActivate')"),
  vbaCode: z.string().describe("VBA code for the event handler")
});

async function addApplicationEventHandlerTool(args: z.infer<typeof addApplicationEventHandlerSchema>): Promise<CallToolResult> {
  try {
    await wordService.addApplicationEventHandler(args.eventName, args.vbaCode);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added application event handler '${args.eventName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addApplicationEventHandlerTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add application event handler: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create Auto-Executing Macro ---
const createAutoMacroSchema = z.object({
  autoType: z.enum(["AutoExec", "AutoNew", "AutoOpen", "AutoClose", "AutoExit"]).describe("Type of auto-executing macro"),
  vbaCode: z.string().describe("VBA code to execute automatically")
});

async function createAutoMacroTool(args: z.infer<typeof createAutoMacroSchema>): Promise<CallToolResult> {
  try {
    await wordService.createAutoMacro(args.autoType, args.vbaCode);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created ${args.autoType} macro` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createAutoMacroTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create auto macro: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Get VBA Error Information ---
const getVbaErrorInfoSchema = z.object({});

async function getVbaErrorInfoTool(): Promise<CallToolResult> {
  try {
    const errorInfo = await wordService.getVbaErrorInfo();
    return {
      content: [{ 
        type: "text", 
        text: errorInfo ? `VBA Error Information:\n${JSON.stringify(errorInfo, null, 2)}` : "No VBA errors detected"
      }],
    };
  } catch (error: any) {
    debug.error("Error in getVbaErrorInfoTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get VBA error information: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Clear VBA Immediate Window ---
const clearVbaImmediateWindowSchema = z.object({});

async function clearVbaImmediateWindowTool(): Promise<CallToolResult> {
  try {
    await wordService.clearVbaImmediateWindow();
    return {
      content: [{ type: "text", text: "Successfully cleared VBA Immediate window" }],
    };
  } catch (error: any) {
    debug.error("Error in clearVbaImmediateWindowTool:", error);
    return {
      content: [{ type: "text", text: `Failed to clear VBA Immediate window: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Execute VBA in Immediate Window ---
const executeVbaImmediateSchema = z.object({
  vbaCode: z.string().describe("VBA code to execute in the Immediate window")
});

async function executeVbaImmediateTool(args: z.infer<typeof executeVbaImmediateSchema>): Promise<CallToolResult> {
  try {
    const result = await wordService.executeVbaImmediate(args.vbaCode);
    return {
      content: [{ 
        type: "text", 
        text: `Immediate window execution:\n${result || '(No output)'}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in executeVbaImmediateTool:", error);
    return {
      content: [{ type: "text", text: `Failed to execute in Immediate window: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: List Available Macros ---
const listAvailableMacrosSchema = z.object({});

async function listAvailableMacrosTool(): Promise<CallToolResult> {
  try {
    const macros = await wordService.listAvailableMacros();
    return {
      content: [{ 
        type: "text", 
        text: macros.length > 0 
          ? `Available Macros:\n${macros.map(m => `- ${m.name} (${m.module})`).join('\n')}` 
          : "No macros found in the document"
      }],
    };
  } catch (error: any) {
    debug.error("Error in listAvailableMacrosTool:", error);
    return {
      content: [{ type: "text", text: `Failed to list available macros: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Record Macro ---
const recordMacroSchema = z.object({
  macroName: z.string().describe("Name for the recorded macro"),
  description: z.string().optional().describe("Description of what the macro does")
});

async function recordMacroTool(args: z.infer<typeof recordMacroSchema>): Promise<CallToolResult> {
  try {
    await wordService.startRecordingMacro(args.macroName, args.description);
    return {
      content: [{ 
        type: "text", 
        text: `Started recording macro '${args.macroName}'. Use word_vba_stopRecording to stop.` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in recordMacroTool:", error);
    return {
      content: [{ type: "text", text: `Failed to start recording macro: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Stop Recording Macro ---
const stopRecordingMacroSchema = z.object({});

async function stopRecordingMacroTool(): Promise<CallToolResult> {
  try {
    const macroName = await wordService.stopRecordingMacro();
    return {
      content: [{ 
        type: "text", 
        text: `Successfully stopped recording and saved macro${macroName ? ` '${macroName}'` : ''}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in stopRecordingMacroTool:", error);
    return {
      content: [{ type: "text", text: `Failed to stop recording macro: ${error.message}` }],
      isError: true,
    };
  }
}

// --- ShowUserFormModal Function Creation ---
const createShowUserFormModalFunctionSchema = z.object({});

async function createShowUserFormModalFunctionTool(): Promise<CallToolResult> {
  try {
    await wordService.createShowUserFormModalFunction();
    return {
      content: [{ 
        type: "text", 
        text: "Successfully created ShowUserFormModal VBA function. This function can now be called from other VBA code to display UserForms modally and collect data." 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createShowUserFormModalFunctionTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create ShowUserFormModal function: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Generic Form Event Handlers ---
const addFormEventHandlersSchema = z.object({
  formName: z.string().describe("Name of the UserForm to add event handlers to")
});

async function addFormEventHandlersTool(args: z.infer<typeof addFormEventHandlersSchema>): Promise<CallToolResult> {
  try {
    await wordService.addStandardFormEventHandlers(args.formName);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added standard event handlers to UserForm '${args.formName}'. The form now has OK/Cancel button handlers and proper closure handling.` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addFormEventHandlersTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add form event handlers: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerVbaExecutionTools(server: McpServer) {
  server.tool(
    "word_vba_runMacro",
    "Runs a VBA macro in the active document",
    runVbaMacroSchema.shape,
    runVbaMacroTool
  );
  
  server.tool(
    "word_vba_testMacro",
    "Tests a VBA macro with optional test data and validation",
    testVbaMacroSchema.shape,
    testVbaMacroTool
  );
  
  server.tool(
    "word_vba_debugCode",
    "Debugs VBA code with breakpoints",
    debugVbaCodeSchema.shape,
    debugVbaCodeTool
  );
  
  server.tool(
    "word_vba_compileProject",
    "Compiles the VBA project and reports any errors",
    compileVbaProjectSchema.shape,
    compileVbaProjectTool
  );
  
  server.tool(
    "word_vba_addDocumentEvent",
    "Adds a document-level event handler",
    addDocumentEventHandlerSchema.shape,
    addDocumentEventHandlerTool
  );
  
  server.tool(
    "word_vba_addApplicationEvent",
    "Adds an application-level event handler",
    addApplicationEventHandlerSchema.shape,
    addApplicationEventHandlerTool
  );
  
  server.tool(
    "word_vba_createAutoMacro",
    "Creates an auto-executing macro (AutoExec, AutoOpen, etc.)",
    createAutoMacroSchema.shape,
    createAutoMacroTool
  );
  
  server.tool(
    "word_vba_getErrorInfo",
    "Gets current VBA error information",
    getVbaErrorInfoSchema.shape,
    getVbaErrorInfoTool
  );
  
  server.tool(
    "word_vba_clearImmediate",
    "Clears the VBA Immediate window",
    clearVbaImmediateWindowSchema.shape,
    clearVbaImmediateWindowTool
  );
  
  server.tool(
    "word_vba_executeImmediate",
    "Executes VBA code in the Immediate window",
    executeVbaImmediateSchema.shape,
    executeVbaImmediateTool
  );
  
  server.tool(
    "word_vba_listMacros",
    "Lists all available macros in the document",
    listAvailableMacrosSchema.shape,
    listAvailableMacrosTool
  );
  
  server.tool(
    "word_vba_recordMacro",
    "Starts recording a new macro",
    recordMacroSchema.shape,
    recordMacroTool
  );
  
  server.tool(
    "word_vba_stopRecording",
    "Stops recording the current macro",
    stopRecordingMacroSchema.shape,
    stopRecordingMacroTool
  );
  
  server.tool(
    "word_vba_createShowUserFormModalFunction",
    "Creates the ShowUserFormModal VBA function for form display",
    createShowUserFormModalFunctionSchema.shape,
    createShowUserFormModalFunctionTool
  );
  
  server.tool(
    "word_vba_addFormEventHandlers",
    "Adds standard OK/Cancel event handlers to a UserForm",
    addFormEventHandlersSchema.shape,
    addFormEventHandlersTool
  );
}