// Test script to validate ActiveX control naming and UserForm fixes
const { wordService } = require('./dist/word/word-service.js');

async function testActiveXControlNaming() {
  console.log("Testing ActiveX control naming fix...");
  
  try {
    // Create a new document
    await wordService.createDocument();
    console.log("✓ Document created");
    
    // Add a control with specific name
    const controlName = "btnCustomizeProposal";
    await wordService.addActiveXControl("commandButton", controlName, "Customize Proposal", 100, 100, 150, 30);
    console.log(`✓ Added control with name: ${controlName}`);
    
    // List controls to verify the name was set correctly
    const controls = await wordService.listActiveXControls();
    console.log("Controls found:", controls);
    
    const foundControl = controls.find(c => c.name === controlName);
    if (foundControl) {
      console.log(`✓ SUCCESS: Control found with correct name: ${foundControl.name}`);
    } else {
      console.log(`✗ FAILED: Control with name '${controlName}' not found. Available controls:`, controls.map(c => c.name));
    }
    
    // Test event handler binding
    try {
      await wordService.addActiveXEventHandler(controlName, "Click", "MsgBox \"Button clicked!\"");
      console.log("✓ Event handler added successfully");
    } catch (eventError) {
      console.log("✗ Event handler failed:", eventError.message);
    }
    
  } catch (error) {
    console.log("✗ ActiveX control test failed:", error.message);
  }
}

async function testUserFormCreation() {
  console.log("\nTesting UserForm creation fix...");
  
  try {
    const formName = "TestForm";
    
    // Test creating UserForm (this should no longer throw COM dispatch errors)
    await wordService.createUserForm(formName, "Test Form Caption", 400, 300);
    console.log(`✓ UserForm '${formName}' created successfully`);
    
    // Add a control to the form
    await wordService.addControlToUserForm(formName, "textBox", "txtName", "", 10, 10, 200, 25);
    console.log("✓ Control added to UserForm");
    
    // Add a submit button
    await wordService.addControlToUserForm(formName, "commandButton", "btnSubmit", "Submit", 10, 50, 100, 25);
    console.log("✓ Submit button added to UserForm");
    
    // List UserForms
    const forms = await wordService.listUserForms();
    console.log("UserForms found:", forms);
    
    if (forms.includes(formName)) {
      console.log("✓ SUCCESS: UserForm creation and management working");
    } else {
      console.log("✗ FAILED: UserForm not found in list");
    }
    
    // Test getting form controls
    const formControls = await wordService.getUserFormControls(formName);
    console.log("Form controls:", formControls);
    
  } catch (error) {
    console.log("✗ UserForm test failed:", error.message);
  }
}

async function runTests() {
  console.log("=== MCP Word Server Fix Validation ===\n");
  
  await testActiveXControlNaming();
  await testUserFormCreation();
  
  console.log("\n=== Tests Complete ===");
  
  // Clean up - close document without saving
  try {
    const doc = await wordService.getActiveDocument();
    await wordService.closeDocument(doc, 0); // Don't save
    console.log("✓ Test document closed");
  } catch (cleanupError) {
    console.log("Note: Could not clean up test document:", cleanupError.message);
  }
}

// Only run if this file is executed directly
if (require.main === module) {
  runTests().catch(console.error);
}

module.exports = { testActiveXControlNaming, testUserFormCreation };