import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { wordService } from "../word/word-service.js";
import { debug } from "../utils/debug.js";

// ActiveX Control Types
const ActiveXControlType = z.enum([
  "commandButton",
  "textBox",
  "label",
  "checkBox",
  "optionButton",
  "comboBox",
  "listBox",
  "toggleButton",
  "spinButton",
  "scrollBar",
  "image",
  "frame"
]);

// --- Tool: Add ActiveX Control ---
const addActiveXControlSchema = z.object({
  controlType: ActiveXControlType.describe("Type of ActiveX control to add"),
  name: z.string().describe("Name/ID for the control"),
  caption: z.string().optional().describe("Caption text for the control (for buttons, labels, etc.)"),
  left: z.number().optional().default(0).describe("Left position in points"),
  top: z.number().optional().default(0).describe("Top position in points"),
  width: z.number().optional().default(100).describe("Width in points"),
  height: z.number().optional().default(25).describe("Height in points"),
  anchorToRange: z.boolean().optional().default(false).describe("Whether to anchor the control to the current selection")
});

async function addActiveXControlTool(args: z.infer<typeof addActiveXControlSchema>): Promise<CallToolResult> {
  try {
    const control = await wordService.addActiveXControl(
      args.controlType,
      args.name,
      args.caption,
      args.left,
      args.top,
      args.width,
      args.height,
      args.anchorToRange
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added ActiveX ${args.controlType} control '${args.name}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addActiveXControlTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add ActiveX control: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Delete ActiveX Control ---
const deleteActiveXControlSchema = z.object({
  controlName: z.string().describe("Name/ID of the control to delete")
});

async function deleteActiveXControlTool(args: z.infer<typeof deleteActiveXControlSchema>): Promise<CallToolResult> {
  try {
    await wordService.deleteActiveXControl(args.controlName);
    return {
      content: [{ type: "text", text: `Successfully deleted ActiveX control '${args.controlName}'` }],
    };
  } catch (error: any) {
    debug.error("Error in deleteActiveXControlTool:", error);
    return {
      content: [{ type: "text", text: `Failed to delete ActiveX control: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set ActiveX Control Properties ---
const setActiveXControlPropertiesSchema = z.object({
  controlName: z.string().describe("Name/ID of the control"),
  properties: z.record(z.any()).describe("Object with properties to set (e.g., {Caption: 'Click Me', Enabled: true, BackColor: 0xFF0000})")
});

async function setActiveXControlPropertiesTool(args: z.infer<typeof setActiveXControlPropertiesSchema>): Promise<CallToolResult> {
  try {
    await wordService.setActiveXControlProperties(args.controlName, args.properties);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully updated properties for ActiveX control '${args.controlName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in setActiveXControlPropertiesTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set ActiveX control properties: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Get ActiveX Control Properties ---
const getActiveXControlPropertiesSchema = z.object({
  controlName: z.string().describe("Name/ID of the control")
});

async function getActiveXControlPropertiesTool(args: z.infer<typeof getActiveXControlPropertiesSchema>): Promise<CallToolResult> {
  try {
    const properties = await wordService.getActiveXControlProperties(args.controlName);
    return {
      content: [{ 
        type: "text", 
        text: `Properties for ActiveX control '${args.controlName}':\n${JSON.stringify(properties, null, 2)}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in getActiveXControlPropertiesTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get ActiveX control properties: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add Event Handler to ActiveX Control ---
const addActiveXEventHandlerSchema = z.object({
  controlName: z.string().describe("Name/ID of the control"),
  eventName: z.string().describe("Name of the event (e.g., 'Click', 'Change', 'DblClick')"),
  vbaCode: z.string().describe("VBA code for the event handler (without the Sub declaration)")
});

async function addActiveXEventHandlerTool(args: z.infer<typeof addActiveXEventHandlerSchema>): Promise<CallToolResult> {
  try {
    await wordService.addActiveXEventHandler(args.controlName, args.eventName, args.vbaCode);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added ${args.eventName} event handler for control '${args.controlName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addActiveXEventHandlerTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add event handler: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: List ActiveX Controls ---
const listActiveXControlsSchema = z.object({});

async function listActiveXControlsTool(): Promise<CallToolResult> {
  try {
    const controls = await wordService.listActiveXControls();
    return {
      content: [{ 
        type: "text", 
        text: controls.length > 0 
          ? `ActiveX Controls:\n${controls.map(c => `- ${c.name} (${c.type})`).join('\n')}` 
          : "No ActiveX controls found in the document"
      }],
    };
  } catch (error: any) {
    debug.error("Error in listActiveXControlsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to list ActiveX controls: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Create UserForm ---
const createUserFormSchema = z.object({
  formName: z.string().describe("Name of the UserForm to create"),
  caption: z.string().optional().describe("Caption for the UserForm"),
  width: z.number().optional().default(400).describe("Width of the form in points"),
  height: z.number().optional().default(300).describe("Height of the form in points")
});

async function createUserFormTool(args: z.infer<typeof createUserFormSchema>): Promise<CallToolResult> {
  try {
    await wordService.createUserForm(args.formName, args.caption, args.width, args.height);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created UserForm '${args.formName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createUserFormTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create UserForm: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Add Control to UserForm ---
const addControlToUserFormSchema = z.object({
  formName: z.string().describe("Name of the UserForm"),
  controlType: ActiveXControlType.describe("Type of control to add"),
  controlName: z.string().describe("Name for the control"),
  caption: z.string().optional().describe("Caption text for the control"),
  left: z.number().optional().default(10).describe("Left position in the form"),
  top: z.number().optional().default(10).describe("Top position in the form"),
  width: z.number().optional().default(100).describe("Width of the control"),
  height: z.number().optional().default(25).describe("Height of the control")
});

async function addControlToUserFormTool(args: z.infer<typeof addControlToUserFormSchema>): Promise<CallToolResult> {
  try {
    await wordService.addControlToUserForm(
      args.formName,
      args.controlType,
      args.controlName,
      args.caption,
      args.left,
      args.top,
      args.width,
      args.height
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added ${args.controlType} '${args.controlName}' to UserForm '${args.formName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addControlToUserFormTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add control to UserForm: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Set Control Tab Order ---
const setControlTabOrderSchema = z.object({
  controlName: z.string().describe("Name of the control"),
  tabIndex: z.number().describe("Tab index (0-based order)")
});

async function setControlTabOrderTool(args: z.infer<typeof setControlTabOrderSchema>): Promise<CallToolResult> {
  try {
    await wordService.setControlTabOrder(args.controlName, args.tabIndex);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully set tab index ${args.tabIndex} for control '${args.controlName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in setControlTabOrderTool:", error);
    return {
      content: [{ type: "text", text: `Failed to set control tab order: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Tool: Group Controls ---
const groupControlsSchema = z.object({
  controlNames: z.array(z.string()).describe("Array of control names to group together"),
  groupName: z.string().optional().describe("Optional name for the group")
});

async function groupControlsTool(args: z.infer<typeof groupControlsSchema>): Promise<CallToolResult> {
  try {
    await wordService.groupControls(args.controlNames, args.groupName);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully grouped ${args.controlNames.length} controls${args.groupName ? ` as '${args.groupName}'` : ''}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in groupControlsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to group controls: ${error.message}` }],
      isError: true,
    };
  }
}

// --- UserForm Schemas and Tools ---
const showUserFormSchema = z.object({
  formName: z.string().describe("Name of the UserForm to show")
});

async function showUserFormTool(args: z.infer<typeof showUserFormSchema>): Promise<CallToolResult> {
  try {
    const formData = await wordService.showUserForm(args.formName);
    return {
      content: [{ 
        type: "text", 
        text: `UserForm '${args.formName}' displayed. Collected data: ${JSON.stringify(formData, null, 2)}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in showUserFormTool:", error);
    return {
      content: [{ type: "text", text: `Failed to show UserForm '${args.formName}': ${error.message}` }],
      isError: true,
    };
  }
}

const addUserFormSubmitLogicSchema = z.object({
  formName: z.string().describe("Name of the UserForm"),
  targetMacro: z.string().describe("Name of the macro to call when form is submitted")
});

async function addUserFormSubmitLogicTool(args: z.infer<typeof addUserFormSubmitLogicSchema>): Promise<CallToolResult> {
  try {
    await wordService.addUserFormSubmitLogic(args.formName, args.targetMacro);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully added submit logic to UserForm '${args.formName}' targeting macro '${args.targetMacro}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in addUserFormSubmitLogicTool:", error);
    return {
      content: [{ type: "text", text: `Failed to add submit logic: ${error.message}` }],
      isError: true,
    };
  }
}

const listUserFormsSchema = z.object({});

async function listUserFormsTool(): Promise<CallToolResult> {
  try {
    const forms = await wordService.listUserForms();
    return {
      content: [{ 
        type: "text", 
        text: `UserForms in document: ${forms.length > 0 ? forms.join(', ') : 'None found'}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in listUserFormsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to list UserForms: ${error.message}` }],
      isError: true,
    };
  }
}

const deleteUserFormSchema = z.object({
  formName: z.string().describe("Name of the UserForm to delete")
});

async function deleteUserFormTool(args: z.infer<typeof deleteUserFormSchema>): Promise<CallToolResult> {
  try {
    await wordService.deleteUserForm(args.formName);
    return {
      content: [{ 
        type: "text", 
        text: `Successfully deleted UserForm '${args.formName}'` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in deleteUserFormTool:", error);
    return {
      content: [{ type: "text", text: `Failed to delete UserForm '${args.formName}': ${error.message}` }],
      isError: true,
    };
  }
}

const getUserFormControlsSchema = z.object({
  formName: z.string().describe("Name of the UserForm")
});

async function getUserFormControlsTool(args: z.infer<typeof getUserFormControlsSchema>): Promise<CallToolResult> {
  try {
    const controls = await wordService.getUserFormControls(args.formName);
    return {
      content: [{ 
        type: "text", 
        text: `Controls on UserForm '${args.formName}': ${JSON.stringify(controls, null, 2)}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in getUserFormControlsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to get UserForm controls: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Generic Form Creation ---
const createFormWithControlsSchema = z.object({
  formName: z.string().describe("Name of the UserForm to create"),
  caption: z.string().describe("Caption/title of the form"),
  width: z.number().min(300).describe("Width of the form in pixels (minimum 300)"),
  height: z.number().min(200).describe("Height of the form in pixels (minimum 200)"),
  controls: z.array(z.object({
    type: z.enum(["textBox", "label", "commandButton", "checkBox", "optionButton", "comboBox", "listBox"]).describe("Type of control"),
    name: z.string().describe("Name of the control"),
    caption: z.string().optional().describe("Caption/text for the control"),
    left: z.number().describe("Left position in pixels"),
    top: z.number().describe("Top position in pixels"),
    width: z.number().describe("Width in pixels"),
    height: z.number().describe("Height in pixels")
  })).describe("Array of controls to add to the form"),
  addStandardButtons: z.boolean().optional().default(true).describe("Whether to add OK/Cancel buttons automatically")
});

async function createFormWithControlsTool(args: z.infer<typeof createFormWithControlsSchema>): Promise<CallToolResult> {
  try {
    await wordService.createFormWithControls(
      args.formName,
      args.caption,
      args.width,
      args.height,
      args.controls,
      args.addStandardButtons
    );
    return {
      content: [{ 
        type: "text", 
        text: `Successfully created UserForm '${args.formName}' with ${args.controls.length} controls. ${args.addStandardButtons ? 'Standard OK/Cancel buttons and event handlers added.' : 'No standard buttons added.'}` 
      }],
    };
  } catch (error: any) {
    debug.error("Error in createFormWithControlsTool:", error);
    return {
      content: [{ type: "text", text: `Failed to create form: ${error.message}` }],
      isError: true,
    };
  }
}

// --- Register Tools ---
export function registerActiveXControlTools(server: McpServer) {
  server.tool(
    "word_activex_addControl",
    "Adds an ActiveX control to the Word document",
    addActiveXControlSchema.shape,
    addActiveXControlTool
  );
  
  server.tool(
    "word_activex_deleteControl",
    "Deletes an ActiveX control from the document",
    deleteActiveXControlSchema.shape,
    deleteActiveXControlTool
  );
  
  server.tool(
    "word_activex_setProperties",
    "Sets properties of an ActiveX control",
    setActiveXControlPropertiesSchema.shape,
    setActiveXControlPropertiesTool
  );
  
  server.tool(
    "word_activex_getProperties",
    "Gets properties of an ActiveX control",
    getActiveXControlPropertiesSchema.shape,
    getActiveXControlPropertiesTool
  );
  
  server.tool(
    "word_activex_addEventHandler",
    "Adds a VBA event handler for an ActiveX control",
    addActiveXEventHandlerSchema.shape,
    addActiveXEventHandlerTool
  );
  
  server.tool(
    "word_activex_listControls",
    "Lists all ActiveX controls in the document",
    listActiveXControlsSchema.shape,
    listActiveXControlsTool
  );
  
  server.tool(
    "word_activex_createUserForm",
    "Creates a new UserForm for the document",
    createUserFormSchema.shape,
    createUserFormTool
  );
  
  server.tool(
    "word_activex_addControlToForm",
    "Adds a control to a UserForm",
    addControlToUserFormSchema.shape,
    addControlToUserFormTool
  );
  
  server.tool(
    "word_activex_setTabOrder",
    "Sets the tab order for a control",
    setControlTabOrderSchema.shape,
    setControlTabOrderTool
  );
  
  server.tool(
    "word_activex_groupControls",
    "Groups multiple controls together",
    groupControlsSchema.shape,
    groupControlsTool
  );
  
  // --- UserForm Tools ---
  server.tool(
    "word_userform_show",
    "Shows a UserForm modally and returns collected data",
    showUserFormSchema.shape,
    showUserFormTool
  );
  
  server.tool(
    "word_userform_addSubmitLogic",
    "Adds submit logic to a UserForm",
    addUserFormSubmitLogicSchema.shape,
    addUserFormSubmitLogicTool
  );
  
  server.tool(
    "word_userform_list",
    "Lists all UserForms in the current document",
    listUserFormsSchema.shape,
    listUserFormsTool
  );
  
  server.tool(
    "word_userform_delete",
    "Deletes a UserForm",
    deleteUserFormSchema.shape,
    deleteUserFormTool
  );
  
  server.tool(
    "word_userform_getControls",
    "Gets all controls on a UserForm",
    getUserFormControlsSchema.shape,
    getUserFormControlsTool
  );
  
  server.tool(
    "word_userform_createWithControls",
    "Creates a UserForm with multiple controls in one operation (completely generic)",
    createFormWithControlsSchema.shape,
    createFormWithControlsTool
  );
}