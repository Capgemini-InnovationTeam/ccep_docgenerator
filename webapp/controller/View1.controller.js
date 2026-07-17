/* global $:readonly */
/* eslint-disable @sap-ux/fiori-tools/sap-no-dom-insertion, @sap-ux/fiori-tools/sap-timeout-usage */
sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
	"sap/m/MessageToast" ,
    "ccepdocgenerator/controller/TSDGenerator",
    "ccepdocgenerator/controller/TSDFollowUp",
    "ccepdocgenerator/controller/FSDGenerator",
    "ccepdocgenerator/controller/FSDFollowUp"
], function (Controller, JSONModel, MessageBox, MessageToast, TSDGenerator, TSDFollowUp, FSDGenerator, FSDFollowUp) {
    "use strict";

    return Controller.extend("ccepdocgenerator.controller.View1",
    Object.assign({

        onInit: function () {

            this.oBusyDialog = new sap.m.BusyDialog({
                title: "Please wait",
                text: "Getting data from Gen AI..."
            });
			this._aSelectedFiles = null;
			this._name = null;
			this._userprompt = null;

			this.openAIModelData = {
				isEditable: false,
				temperature: 0.1,
				max_tokens: 15000,
				top_p: 1

			};

			

			var oModifiedAPIModel = new sap.ui.model.json.JSONModel();
			this.getView().setModel(oModifiedAPIModel, "ModifiedAPIModel");

			this._aSelectedFilestsd = null;
			this._nametsd = null;
			this._userpromptTD = null;

			// Initialize empty prompt models (OData prompt repository service removed)
			var oPromptsModelFSD = new sap.ui.model.json.JSONModel({ results: [], prompts: [] });
			var oPromptsModelTSD = new sap.ui.model.json.JSONModel({ results: [], prompts: [] });
			this.getView().setModel(oPromptsModelFSD, "promptsModel");
			this.getView().setModel(oPromptsModelTSD, "promptsModel2");

			
			this.response = null;
			this.responseapi1Global = null;
			this.RESPONSE = null;
			this.CleancoreResponse = null;
			this.response_cleancore = null;
			this.DevelopmentType = null;
			this.promptdata = null;
			this.CleancoreResponse2 = null;
		
			this.EnhancementResponse = null;
			this.cleancorerecommendation = false
			this.secondResponseData = null;
			this.firstResponseData = null;
			this.standard_Response = null;
			this.custdeliverytype = null;
			this.formresponse = null;
			this.developmetTypeResponses = null;
			this.followUpPromptingFlag = false;
			this.followUpPromptRespons = null;
			this.followUpCreateTsdFlag = false;
			this.Tittle=null;

//TSD Follow up
			this.onGenerateTSDPress2 = TSDFollowUp.onGenerateTSDPress2.bind(this);
this._callSecondAPITSD2  = TSDFollowUp._callSecondAPITSD2.bind(this);
this.onComboBoxSelectionChangeTSD = TSDFollowUp.onComboBoxSelectionChangeTSD.bind(this);

this.onSavePromptSelection = TSDFollowUp.onSavePromptSelection.bind(this);
this.onClosePromptSelection = TSDFollowUp.onClosePromptSelection.bind(this);

    this.openAIModelData = {
    isEditable: false,
    temperature: 0.0,
    max_tokens: 12000,
    top_p: 1
};

this.updateAPIModelFlag = false;
var oApiModel = new sap.ui.model.json.JSONModel(this.openAIModelData);
this.getView().setModel(oApiModel, "apiModel");
		},



		//-----------FSD-----------//
		onFSDPress: function () {
			this.getView().byId("id_TSD").removeStyleClass("customColor");
			this.getView().byId("id_FSD").addStyleClass("customColor");
			this.getView().byId("id_TSDWizard").setVisible(false);
			this.getView().byId("id_FSDWizard").setVisible(true);
		},
		
onEditSysMessage: function (oEvent) {
    var oButton = oEvent.getSource();
    var oParentHBox = oButton.getParent();      // HBox wrapping [TextArea, Button]
    var oTextArea = oParentHBox.getItems()[0];  // the sys message TextArea

    var bIsEditing = oTextArea.getEditable();

    if (!bIsEditing) {
        // Enter edit mode
        oTextArea.setEditable(true);
        oButton.setIcon("sap-icon://save");
        oButton.setTooltip("Click here to save the system message");
        oTextArea.addStyleClass("sysMessageEditing");
        oTextArea.focus();
    } else {
        // Save and lock again
        oTextArea.setEditable(false);
        oButton.setIcon("sap-icon://edit");
        oButton.setTooltip("Click here to edit the text");
        oTextArea.removeStyleClass("sysMessageEditing");
        sap.m.MessageToast.show("System message saved");
    }
},

			handleTypeMissmatch: function (oEvent) {
			var sFileType = oEvent.getParameter("fileType");
			var sFileName = oEvent.getParameter("fileName");

			MessageBox.error("The file '" + sFileName + "' has an invalid type (" + sFileType + "). " +
				"Please upload only .docx or .pdf files.");
		},
		handleFileSizeExceed: function (oEvent) {
			MessageBox.error("The selected file exceeds the maximum allowed size of 10 MB.");
			oEvent.getSource().setValue("");
		},
		_isAllowedFile: function (sFileName, aAllowedExtensions) {
			if (!sFileName) {
				return false;
			}
			var sLowerName = sFileName.toLowerCase();
			var aBlockedExtensions = ["exe", "bat", "cmd", "sh", "js", "vbs", "ps1", "msi", "jar", "com", "scr", "dll"];
			var aParts = sLowerName.split(".");
			if (aParts.length < 2) {
				return false;
			}
			var aInnerParts = aParts.slice(1, -1);
			for (var i = 0; i < aInnerParts.length; i++) {
				if (aBlockedExtensions.indexOf(aInnerParts[i]) !== -1) {
					return false;
				}
			}
			var sExtension = aParts[aParts.length - 1];
			return aAllowedExtensions.indexOf(sExtension) !== -1;
		},
		onEditPrompt: function (oEvent) {
    if (oEvent.getSource().getParent().getAggregation("items")[0].getId().includes("FSD")) {
        this.getView().byId("id_FSDprompt").setEditable(true);
        this.getView().byId("id_TSDprompt").setEditable(false);
        
    }
    if (oEvent.getSource().getParent().getAggregation("items")[0].getId().includes("TSD")) {
        this.getView().byId("id_TSDprompt").setEditable(true);
        this.getView().byId("id_FSDprompt").setEditable(false);
        
    }
},
		onEditPromptParameterPress: function () {
    if (!this._oPromptPopover1) {
        this._oPromptPopover1 = sap.ui.xmlfragment("promptParameter", "ccepdocgenerator.view.fragment.promptParameter", this);
        this.getView().addDependent(this._oPromptPopover1);
        this._oPromptPopover1.setContentWidth("390px");
    }
    this._oPromptPopover1.open();
},
		onPPCancel: function () {
			this._oPromptPopover1.close();
		},
		
		//-----------functions-----------//
		onFileChange: function (oEvent) {

			this._aSelectedFiles = oEvent.getParameter("files");
			this._name = "FSD";
			if (this._aSelectedFiles && this._aSelectedFiles.length > 0) {
				var oFile = this._aSelectedFiles[0];
				if (!this._isAllowedFile(oFile.name, ["docx", "pdf"])) {
					MessageBox.error("The file '" + oFile.name + "' is not an allowed type. Please upload only .docx or .pdf files.");
					oEvent.getSource().setValue("");
					this._aSelectedFiles = null;
					return;
				}
				sap.m.MessageToast.show("File selected. Now click 'Generate FSD' to proceed.");
			} else {
				sap.m.MessageToast.show("Please select a file first.");
			}
		},

	

		

		
		onEditAPIParameters: function () {

			var oModel = this.getView().getModel("apiModel");
			
			oModel.setProperty("/isEditable", true);
		

		},
		
		onResetParameters: function () {

			var oModel = this.getView().getModel("apiModel");

			var oDefaultData = {
				isEditable: true,
				max_tokens: 4096,
				temperature: 0,
				top_p: 1.0

			};

			oModel.setData(oDefaultData);

			// Refresh the model to apply changes
			oModel.refresh();

			
		},
		onSaveAPIParameters: function () {

			var oModel = this.getView().getModel("apiModel");
			var oheightSlider = sap.ui.core.Fragment.byId("promptParameter", "heightSlider");
			var oheightSlider1 = sap.ui.core.Fragment.byId("promptParameter", "heightSlider1");
			var oheightSlider2 = sap.ui.core.Fragment.byId("promptParameter", "heightSlider2");
			
			this.sV = oheightSlider.getValue();
			this.sV1 = oheightSlider1.getValue();
			this.sV2 = oheightSlider2.getValue();

			oModel.setProperty("/max_tokens", this.sV);
			oModel.setProperty("/temperature", this.sV1);
			oModel.setProperty("/top_p", this.sV2);
			

			sap.m.MessageToast.show("Data saved successfully");
			oModel.setProperty("/isEditable", false); // Set back to false after saving
			oModel.refresh();

			
		},
onCloseParametersDialog: function () {
    if (this._oPromptPopover1) {

        var oModel = this.getView().getModel("apiModel");
        var oData = oModel.getData();

        var oModifiedAPIModel = this.getView().getModel("ModifiedAPIModel");
        oModifiedAPIModel.setData(oData);

        var sModelKey = this.byId("id_FSDWizard").getVisible()
            ? this.byId("_IDGenComboBox").getSelectedKey()
            : this.byId("_IDGenComboBox2").getSelectedKey();

        MessageToast.show(
            "Model: " + sModelKey +
            " | Max Tokens: " + oData.max_tokens +
            " | Temperature: " + oData.temperature,
            { duration: 2000 }
        );

        this._oPromptPopover1.close();
    }
},
		//----------TSD---------//
		onTSDPress: function () {

			this.getView().byId("id_FSD").removeStyleClass("customColor");
			this.getView().byId("id_TSD").addStyleClass("customColor");
			this.getView().byId("id_FSDWizard").setVisible(false);
			this.getView().byId("id_TSDWizard").setVisible(true);
		},
		
		onFileChangeTSD: function (oEvent) {
			this._aSelectedFilestsd = oEvent.getParameter("files");
			this._nametsd = "TSD";
		
			if (this._aSelectedFilestsd && this._aSelectedFilestsd.length > 0) {
				const aFiles = this._aSelectedFilestsd; // Get the list of selected files
				const aFileNames = Array.from(aFiles).map(file => file.name.toLowerCase()); // Convert file names to lowercase

				var oInvalidFile = Array.from(aFiles).find(file => !this._isAllowedFile(file.name, ["docx", "pdf"]));
				if (oInvalidFile) {
					MessageBox.error("The file '" + oInvalidFile.name + "' is not an allowed type. Please upload only .docx or .pdf files.");
					oEvent.getSource().setValue("");
					this._aSelectedFilestsd = null;
					return;
				}
		
				// Check if at least one FSD file is present
				const hasFSD = aFileNames.some(name => name.includes("fsd"));
		
				if (aFiles.length !== 1) {
					sap.m.MessageToast.show("Please upload exactly one FSD document.");
				} else if (!hasFSD) {
					sap.m.MessageToast.show("Only an FSD document is allowed. Please check your file selection.");
				} else {
					sap.m.MessageToast.show("FSD document selected. Now click 'Step 2' and then 'Generate TSD' to proceed.");
				}
			} else {
				sap.m.MessageToast.show("Please select a file first.");
			}
		},
	
		
onGenerateTSDPress2: function () {
			 
			this._userpromptTSD = this.byId("id_TSDprompt").getValue();
			// this._developmenttypeTSD = this.byId("id_TSDdelivery_type").getValue();
			if (!this._userpromptTSD) {
				sap.m.MessageToast.show("Please select a valid option first.");
				return;
			}

			var userpromptTSD = this._userpromptTSD
			var development_typeTSD = this.custdeliverytype;
			if (this.followUpPromptingFlag) {
				this._callSecondAPITSD2(userpromptTSD, development_typeTSD, this.followUpPromptResponse);
			} else {
				 
				var standardResponse = this.standard_Response;
				var enhancementResponse = this.developmetTypeResponses;

				let standard_Response2 = {
					standardResponse,
					enhancementResponse
				};
				
				this.standard_Response = JSON.stringify(standard_Response2);
				
				this._callSecondAPITSD2(userpromptTSD, development_typeTSD, this.standard_Response);

			}

		},
		
		_callSecondAPITSD2: function (userpromptTSD, development_typeTSD, standardsresponse_1) {
    // 1. Setup UI and Paths
    this.oBusyDialog.open();
    var that = this;

    // Build the relative path for the BTP Destination
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));

    // This prefix MUST match the 'source' in your xs-app.json
    var sApiUrlTSD2 = appModulePath + "/tsd_api/fsdtotsdagenticcustom2";

    // 2. Helper: build fresh FormData each attempt (FormData can only be sent once)
    function buildFormData() {
        var oFD = new FormData();
        oFD.append("customized_prompt", userpromptTSD);
        oFD.append("input_json", standardsresponse_1);
        oFD.append("development_type", (development_typeTSD || "").toLowerCase());
        return oFD;
    }

    // 3. Helper: fetch fresh CSRF token then POST; retries on 403 up to maxRetries times
   
    function fetchTokenAndPost(attempt, maxRetries) {
        $.ajax({
            url: sApiUrlTSD2,
            method: "GET",
            headers: { "X-CSRF-Token": "Fetch" },
            complete: function (tokenXHR) {
                var sToken = tokenXHR.getResponseHeader("X-CSRF-Token");

                // Enhanced retry logic for missing token using SAP UI5 modalities
                if (!sToken && attempt < maxRetries) {
                    setTimeout(function () {
                        fetchTokenAndPost(attempt + 1, maxRetries);
                    }, 750); // Streamlined retry control
                    sap.m.MessageToast.show("Retrying fetch due to missing token...");
                    return;
                }

                $.ajax({
                    url: sApiUrlTSD2,
                    method: "POST",
                    headers: { "X-CSRF-Token": sToken || "" },
                    timeout: 0,
                    processData: false,
                    contentType: false,
                    mimeType: "multipart/form-data",
                    data: buildFormData(),
                    success: function (response) {
                        that.oBusyDialog.close();

                        // Handle response (Parse string to JSON if needed)
                        var oParsedData;
                        try {
                            oParsedData = (typeof response === "string") ? JSON.parse(response) : response;
                        } catch {
                            sap.m.MessageToast.show("Error parsing agent response.");
                            return;
                        }

                        // Store results in global variables
                        that.followUpPromptResponse = oParsedData;

                        // Logic to split the response into Enhancement and Standard sections
                        if (oParsedData.enhancementResponse) {
                            that.developmetTypeResponses = oParsedData.enhancementResponse;
                            that.standard_Response = oParsedData.standardResponse || "";
                        } else {
                            // Fallback if the structure is flat
                            that.developmetTypeResponses = oParsedData;
                        }

                        that.followUpPromptingFlag = true;

                        // Ensure TSD UI components are updated
                        if (!that.tsdGenerated) {
                            that.tsdGenerated = true;
                            var oPromptRow = that.getView().byId("id_TSDPromptRow");
                            if (oPromptRow) {
                                oPromptRow.setVisible(true);
                            }
                        }


                        // Re-trigger the document generation with the new data
                        that.createTSD();
                    },
                    error: function (jqXHR) {
                    // Retry structured by managed network handling in SAP UI5 framework
                    if (jqXHR.status === 403 && attempt < maxRetries) {
                        setTimeout(function () {
                            fetchTokenAndPost(attempt + 1, maxRetries);
                        }, 750);
                        sap.m.MessageToast.show("Handling 403 error - Retry in progress.");
                    } else {
                        that.oBusyDialog.close();
                        sap.m.MessageToast.show("API call failed after multiple attempts.");
                    }
                    }
                });
            }
        });
    }

    fetchTokenAndPost(1, 3); // up to 3 attempts
},



		onPromptItemDoubleClick: function (oEvent) {
			 
			this.onSavePromptSelection(oEvent);
		},
		onPromptItemDoubleClick2: function (oEvent) {
			 
			this.onSavePromptSelection2(oEvent);
		},
		
		onSavePromptSelection: function (oEvent) {

			var oContext = oEvent.getSource().getBindingContext("promptsModel2");
			var sUserPrompt = oContext.getProperty("name");
			var sPromptText = `${sUserPrompt}`;
			this.byId("id_TSDprompt").setValue(sPromptText);
			this._oPromptRepository.close();
		},
		onSavePromptSelection2: function (oEvent) {
			 
			var oContext = oEvent.getSource().getBindingContext("promptsModel");
			var sUserPrompt = oContext.getProperty("name");
			var sPromptText = `${sUserPrompt}`;
			this.byId("id_FSDprompt").setValue(sPromptText);


			// oButton1.setVisible(!oButton1.getVisible());
			// oButton2.setVisible(!oButton2.getVisible());
			this._oPromptRepository2.close();
		},
		onClosePromptSelection: function () {
			this._oPromptRepository.close();
		},
		onClosePromptSelection2: function () {
			this._oPromptRepository2.close();
		},

	
	
_onFSDGeneratedSuccess: function () {
    var oFsdNudgeBox = this.byId("fsdNudgeBox");
    if (oFsdNudgeBox) { oFsdNudgeBox.setVisible(true); }

    var oStep1 = this.byId("id_FSDStep1");
    if (oStep1) { oStep1.setValidated(true); }

    
},
onAdvancedSettingsPress: function () {
   
    var oFSDWizard = this.byId("id_FSDWizard");
    var oTSDWizard = this.byId("id_TSDWizard");
 
    if (oFSDWizard.getVisible()) {
        var oFSDStep2 = this.byId("id_FSDStep2");
        oFSDWizard.nextStep();          
        oFSDWizard.goToStep(oFSDStep2);
    } else if (oTSDWizard.getVisible()) {
        var oTSDStep2 = this.byId("id_TSDStep2");
        oTSDWizard.nextStep();
        oTSDWizard.goToStep(oTSDStep2);
    }
},


 
	    }, TSDGenerator, TSDFollowUp, FSDGenerator,   
        FSDFollowUp)
);
});