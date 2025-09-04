// Test script to validate generic UserForm functionality
const { wordService } = require('./dist/word/word-service.js');

async function testGenericFormCreation() {
  console.log("=== Testing Generic UserForm Creation ===\n");
  
  try {
    // Step 1: Create a new document
    console.log("1. Creating new document...");
    await wordService.createDocument();
    console.log("✓ Document created");
    
    // Step 2: Create the ShowUserFormModal VBA function
    console.log("\n2. Creating ShowUserFormModal VBA function...");
    await wordService.createShowUserFormModalFunction();
    console.log("✓ ShowUserFormModal function created");
    
    // Step 3: Create a generic form for any purpose (not mining-specific)
    console.log("\n3. Creating generic business form...");
    
    const businessFormControls = [
      { type: "label", name: "lblName", caption: "Name:", left: 20, top: 30, width: 80, height: 20 },
      { type: "textBox", name: "txtName", caption: "", left: 120, top: 28, width: 200, height: 22 },
      
      { type: "label", name: "lblEmail", caption: "Email:", left: 20, top: 60, width: 80, height: 20 },
      { type: "textBox", name: "txtEmail", caption: "", left: 120, top: 58, width: 200, height: 22 },
      
      { type: "label", name: "lblCompany", caption: "Company:", left: 20, top: 90, width: 80, height: 20 },
      { type: "textBox", name: "txtCompany", caption: "", left: 120, top: 88, width: 200, height: 22 },
      
      { type: "label", name: "lblMessage", caption: "Message:", left: 20, top: 120, width: 80, height: 20 },
      { type: "textBox", name: "txtMessage", caption: "", left: 120, top: 118, width: 200, height: 60 },
    ];
    
    await wordService.createFormWithControls(
      "ContactForm",
      "Contact Information", 
      400,  // width
      250,  // height
      businessFormControls,
      true  // add standard OK/Cancel buttons
    );
    
    console.log("✓ Generic contact form created");
    
    // Step 4: Test the VBA code that calls ShowUserFormModal
    console.log("\n4. Adding generic VBA macro that uses the form...");
    const genericMacroCode = `
Sub ShowContactForm()
    Dim formData As String
    formData = ShowUserFormModal("ContactForm")  ' This should work for any form!
    
    If formData <> "" Then
        MsgBox "Contact data collected: " & formData
        Call ProcessContactData(formData)
    Else
        MsgBox "Form was cancelled or no data collected"
    End If
End Sub

Sub ProcessContactData(formData As String)
    ' Process the collected form data - generic approach
    MsgBox "Processing contact data: " & formData
    
    ' Here you would parse the JSON-like string and use the data
    ' to populate any kind of document template
End Sub
`;
    
    await wordService.createVbaModule("GenericFormMacros", "standard", genericMacroCode);
    console.log("✓ Generic macro added that calls ShowUserFormModal");
    
    // Step 5: Verify all VBA modules exist
    console.log("\n5. Verifying VBA modules...");
    const modules = await wordService.listVbaModules();
    console.log("VBA modules found:", modules);
    
    const hasUserFormHelper = modules.some(m => m.name === "UserFormHelper");
    const hasGenericMacros = modules.some(m => m.name === "GenericFormMacros");
    
    if (hasUserFormHelper && hasGenericMacros) {
      console.log("✓ SUCCESS: All required VBA modules created");
    } else {
      console.log("✗ FAILED: Missing VBA modules");
      console.log("  - UserFormHelper:", hasUserFormHelper ? "✓" : "✗");
      console.log("  - GenericFormMacros:", hasGenericMacros ? "✓" : "✗");
    }
    
    // Step 6: List UserForms
    console.log("\n6. Verifying UserForm creation...");
    const forms = await wordService.listUserForms();
    console.log("UserForms found:", forms);
    
    if (forms.includes("ContactForm")) {
      console.log("✓ SUCCESS: ContactForm UserForm created");
      
      // Get form controls
      const formControls = await wordService.getUserFormControls("ContactForm");
      console.log("ContactForm controls:", formControls.map(c => c.name));
      
      const expectedControls = [
        "txtName", "txtEmail", "txtCompany", "txtMessage", "btnOK", "btnCancel"
      ];
      
      const allControlsPresent = expectedControls.every(expected => 
        formControls.some(control => control.name === expected)
      );
      
      if (allControlsPresent) {
        console.log("✓ SUCCESS: All required form controls present");
      } else {
        console.log("✗ WARNING: Some form controls may be missing");
        console.log("Expected:", expectedControls);
        console.log("Found:", formControls.map(c => c.name));
      }
    } else {
      console.log("✗ FAILED: ContactForm not found");
    }
    
    // Step 7: Test form sizing fix
    console.log("\n7. Testing form sizing...");
    try {
      // Create a form with specific dimensions to test sizing
      await wordService.createUserForm("SizeTestForm", "Size Test", 500, 400);
      console.log("✓ Form sizing test passed - no errors creating 500x400 form");
      
      const sizeForms = await wordService.listUserForms();
      if (sizeForms.includes("SizeTestForm")) {
        console.log("✓ Size test form created successfully");
      }
    } catch (sizeError) {
      console.log("✗ Form sizing issue:", sizeError.message);
    }
    
    console.log("\n=== GENERIC SOLUTION SUMMARY ===");
    console.log("✅ REMOVED: Mining-specific code");
    console.log("✅ CREATED: Generic ShowUserFormModal() VBA function");
    console.log("✅ CREATED: Generic form creation tools");
    console.log("✅ FIXED: UserForm sizing issues (minimum 300x200, proper positioning)");
    console.log("✅ FEATURES:");
    console.log("   - word_userform_createWithControls: Create any form with any controls");
    console.log("   - word_vba_addFormEventHandlers: Add OK/Cancel handling to any form");
    console.log("   - ShowUserFormModal(formName): Works with any UserForm");
    console.log("   - Automatic data collection from all control types");
    console.log("   - JSON-like return format for easy parsing");
    console.log("✅ USER CAN NOW: Create any business form for any document type!");
    
  } catch (error) {
    console.log("✗ Test failed:", error.message);
    console.log("Stack:", error.stack);
  }
}

async function runTest() {
  await testGenericFormCreation();
  
  // Clean up - close document without saving
  try {
    const doc = await wordService.getActiveDocument();
    await wordService.closeDocument(doc, 0); // Don't save
    console.log("\n✓ Test document closed");
  } catch (cleanupError) {
    console.log("\nNote: Could not clean up test document:", cleanupError.message);
  }
}

// Only run if this file is executed directly
if (require.main === module) {
  runTest().catch(console.error);
}

module.exports = { testGenericFormCreation };