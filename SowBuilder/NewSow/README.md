# SOW Builder Implementation Guide

## Overview
The SOW Builder is a VBA-based tool for Microsoft Word that allows users to generate Statement of Work (SOW) documents using a simple form interface. The tool creates SOW documents based on the HWC-HB-no-MSA-SOW-US template but embeds all text directly in the code rather than relying on external templates.

## Files
- **SOWBuilderSinglePage.bas** - The main module containing document generation functions
- **frmSOWBuilderSingle.frm** - The UserForm module with all UI logic
- **frmSOWBuilderSingle.frx.instructions** - Instructions for setting up the form's UI controls

## Implementation Steps

### 1. Set Up the VBA Project
1. Open Microsoft Word
2. Press Alt+F11 to open the VBA Editor
3. In the VBA Project Explorer (Ctrl+R if not visible), right-click on your project and select:
   - Insert > Module (for the SOWBuilderSinglePage module)
   - Insert > UserForm (for the frmSOWBuilderSingle form)

### 2. Add the Module Code
1. Copy the content of `SOWBuilderSinglePage.bas` into your new module
2. Save the module with the name "SOWBuilderSinglePage"

### 3. Create the UserForm
1. Name your new UserForm "frmSOWBuilderSingle"
2. Set up all controls according to the instructions in `frmSOWBuilderSingle.frx.instructions`
3. Copy the code from `frmSOWBuilderSingle.frm` into the code window of your form

### 4. Create a Macro to Launch the Form
1. Add a new module to your project
2. Add the following code to call the SOW Builder:
```vba
Sub LaunchSOWBuilder()
    SOWBuilderSinglePage.ShowSOWBuilder
End Sub
```

### 5. Test the Implementation
1. Close the VBA Editor
2. Run the macro by pressing Alt+F8, selecting "LaunchSOWBuilder", and clicking "Run"
3. The SOW Builder form should appear, ready for input

## Using the SOW Builder

### Form Sections
1. **General & Contract Information**
   - Fill in client information, WTW details, and contract dates

2. **Compensation Options**
   - Select one of the four compensation options (A-D)
   - Enter fee amount if applicable

3. **Fee Details**
   - Choose billing methodology (milestone or installments)

4. **Commission Details**
   - Add policies and their commission rates if using options B, C, or D

5. **Optional Clauses**
   - Select any additional clauses to include (Auto-Renewal, GDPR)

6. **Additional Notes**
   - Add any special notes to be included in the SOW

### Generating the Document
1. Fill out all required fields
2. Click "Generate Doc" to create the SOW document
3. The document will be created as a new Word document
4. Review and save the document as needed

## Customization

If you need to modify the SOW text or formatting:

1. Edit the relevant section in the `SOWBuilderSinglePage.bas` module
2. Each document section is contained in its own sub-procedure (AddDocumentHeader, AddTermsAndConditionsSection, etc.)
3. Modify the text or formatting as needed
4. Save the module

## Troubleshooting

### References
This implementation requires:
- Microsoft Word Object Library
- Microsoft VBA Runtime

If experiencing errors, check that these references are enabled:
1. In the VBA Editor, go to Tools > References
2. Ensure the above libraries are checked
3. Click OK and try again

### Form Not Displaying Correctly
If the form layout is incorrect:
1. Refer to the detailed control positions in `frmSOWBuilderSingle.frx.instructions`
2. Adjust control positions and sizes as needed

### Document Generation Issues
If the document format is incorrect:
1. Check the `FormatSOWDocument` sub-procedure in `SOWBuilderSinglePage.bas`
2. Adjust formatting code as needed

## Support
For additional support, please contact your internal VBA development team. 
