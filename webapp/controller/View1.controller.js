/* global $:readonly */
sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "ccepdocgenerator/controller/TSDGenerator",
    "ccepdocgenerator/controller/TSDFollowUp",
    "ccepdocgenerator/controller/FSDGenerator",
    "ccepdocgenerator/controller/FSDFollowUp"
], function (Controller, JSONModel, MessageBox, TSDGenerator, TSDFollowUp, FSDGenerator, FSDFollowUp) {
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

			this.updateAPIModelFlag = false;
			var oApiModel = new sap.ui.model.json.JSONModel(this.openAIModelData);
			this.getView().setModel(oApiModel, "apiModel");

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

			var oFeedbackModel = new JSONModel({
				feedback: {
					issueType: "",
					priority: "",
					title: "",
					description: "",
					attachedFile: ""
				}
			});
			// Set the model to the view
			this.getView().setModel(oFeedbackModel, "feedbackModel");
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
this.onImportPromtPress = TSDFollowUp.onImportPromtPress.bind(this);
this.onSavePromptSelection = TSDFollowUp.onSavePromptSelection.bind(this);
this.onClosePromptSelection = TSDFollowUp.onClosePromptSelection.bind(this);


		},




		//-----------FSD-----------//
		onFSDPress: function () {
			this.getView().byId("id_TSD").removeStyleClass("customColor");
			this.getView().byId("id_FSD").addStyleClass("customColor");
			this.getView().byId("id_TSDWizard").setVisible(false);
			this.getView().byId("id_FSDWizard").setVisible(true);
		},
		
		onEditSysMessage: function (oEvent) {


			if (oEvent.getSource().getParent().getAggregation("items")[0].getId().includes("FSD")) {
				this.getView().byId("id_FSDsysMess").setEditable(true);
				this.getView().byId("id_TSDsysMess").setEditable(false);
				this.getView().byId("btnFSD1").setVisible(false);
				this.getView().byId("btnFSD2").setVisible(true);
			}
			if (oEvent.getSource().getParent().getAggregation("items")[0].getId().includes("TSD")) {
				this.getView().byId("id_TSDsysMess").setEditable(true);
				this.getView().byId("id_FSDsysMess").setEditable(false);
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
				this.getView().byId("btnFSD1").setVisible(false);
				this.getView().byId("btnFSD2").setVisible(true);
				this.getView().byId("btnTSD1").setVisible(false);
				this.getView().byId("btnTSD2").setVisible(true);




			}
			if (oEvent.getSource().getParent().getAggregation("items")[0].getId().includes("TSD")) {
				this.getView().byId("id_TSDprompt").setEditable(true);
				this.getView().byId("id_FSDprompt").setEditable(false);
				this.getView().byId("btnTSD1").setVisible(false);
				this.getView().byId("btnTSD2").setVisible(true);
			}
		},
		onEditPromptParameterPress: function () {
			if (!this._oPromptPopover1) {
				this._oPromptPopover1 = sap.ui.xmlfragment("promptParameter", "ccepdocgenerator.view.fragment.promptParameter", this);
				this.getView().addDependent(this._oPromptPopover1);
				this._oPromptPopover1.setContentWidth("390px");
			}
			//	this._oPopover1.setModel(contactModel, "contact");
			this._oPromptPopover1.open();
		},
		onPPCancel: function () {
			this._oPromptPopover1.close();
		},
		onFeedbackPress: function () {

			if (!this._oFeedbackPopover1) {
				this._oFeedbackPopover1 = sap.ui.xmlfragment("id_feedback", "ccepdocgenerator.view.fragment.feedback", this);
				this.getView().addDependent(this._oFeedbackPopover1);

			}
			//	this._oPopover1.setModel(contactModel, "contact");
			this._oFeedbackPopover1.open();
		},
		onFeedbackCancel: function () {
			this._oFeedbackPopover1.close();
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
				this._oPromptPopover1.close();

				var oModel = this.getView().getModel("apiModel");
				var oData = oModel.getData();

				var oModifiedAPIModel = this.getView().getModel("ModifiedAPIModel");
				oModifiedAPIModel.setData(oData);

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
	
		
		




		
		
		
		
		


		onComboBoxSelectionChangeTSD: function (oEvent) {
			 
			var selectedTextTsd = oEvent.getSource().getSelectedItem().getText();

			this.byId("id_TSDprompt").setValue(selectedTextTsd);
			this._userpromptTD = this.byId("id_TSDprompt").getValue();
			var oButtontsd1 = this.byId("btnTSD1");
			var oButtontsd2 = this.byId("btnTSD2");

			oButtontsd1.setVisible(!oButtontsd1.getVisible());
			oButtontsd2.setVisible(!oButtontsd2.getVisible());

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

                // Guard: if token is missing, retry once before giving up
                if (!sToken && attempt < maxRetries) {
                    setTimeout(function () { fetchTokenAndPost(attempt + 1, maxRetries); }, 1000);
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
                        // 403 = stale/missing CSRF token — re-fetch and retry
                        if (jqXHR.status === 403 && attempt < maxRetries) {
                            setTimeout(function () { fetchTokenAndPost(attempt + 1, maxRetries); }, 1500);
                        } else {
                            that.oBusyDialog.close();
                            sap.m.MessageToast.show("Follow-up API call failed (attempt " + attempt + "). Please check logs.");
                        }
                    }
                });
            }
        });
    }

    fetchTokenAndPost(1, 3); // up to 3 attempts
},

onImportPromtPress: function () {
			if (!this._oPromptRepository) {
				this._oPromptRepository = sap.ui.xmlfragment("promptParameter2", "ccepdocgenerator.view.fragment.GobalPromptRepositoryFSD", this);
				this.getView().addDependent(this._oPromptRepository);
				this._oPromptRepository.setContentWidth("590px");
			}
			//	this._oPopover1.setModel(contactModel, "contact");
			var oButton1 = this.byId("btnTSD1");

			var oButton2 = this.byId("btnTSD2");

			oButton1.setVisible(false);

			oButton2.setVisible(true);
			this._oPromptRepository.open();

		},
		onImportPromtPress2: function () {
			 
			if (!this._oPromptRepository2) {
				this._oPromptRepository2 = sap.ui.xmlfragment("promptParameter3", "ccepdocgenerator.view.fragment.GobalPromptRepositoryFSD",
					this);
				this.getView().addDependent(this._oPromptRepository2);
				this._oPromptRepository2.setContentWidth("590px");

			}

			// oButton1.setVisible(!oButton1.getVisible());
			// oButton2.setVisible(!oButton2.getVisible());

			var oButton1 = this.byId("btnFSD1");

			var oButton2 = this.byId("btnFSD2");

			oButton1.setVisible(false);

			oButton2.setVisible(true);

			this._oPromptRepository2.open();
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

		onClearPress: function () {


			var issueTypeSelect = sap.ui.core.Fragment.byId("id_feedback", "issueTypeSelect");
			var prioritySelect = sap.ui.core.Fragment.byId("id_feedback", "prioritySelect");
			var titleInput = sap.ui.core.Fragment.byId("id_feedback", "titleInput");
			var descriptionInput = sap.ui.core.Fragment.byId("id_feedback", "descriptionInput");
			var fileUploader = sap.ui.core.Fragment.byId("id_feedback", "fileUploader");

			issueTypeSelect.setSelectedKey(null);
			prioritySelect.setSelectedKey(null);
			titleInput.setValue("");
			descriptionInput.setValue("");
			fileUploader.clear();

			var oSuccessIcon = sap.ui.core.Fragment.byId("id_feedback", "fileUploadSuccessIcon");
			var oErrorStrip = sap.ui.core.Fragment.byId("id_feedback", "fileUploadError");
			var oFileNameText = sap.ui.core.Fragment.byId("id_feedback", "fileNameText");

			oSuccessIcon.setVisible(false);
			oErrorStrip.setVisible(false);
			oFileNameText.setVisible(false);

			sap.m.MessageToast.show("All fields have been cleared.");
		},

		onCancelPress: function () {


			this._oFeedbackPopover1.close();
		},

		OnSubmitPress: function () {

			var issueTypeSelect = sap.ui.core.Fragment.byId("id_feedback", "issueTypeSelect");
			var prioritySelect = sap.ui.core.Fragment.byId("id_feedback", "prioritySelect");
			var titleInput = sap.ui.core.Fragment.byId("id_feedback", "titleInput");
			var descriptionInput = sap.ui.core.Fragment.byId("id_feedback", "descriptionInput");
			var fileUploader = sap.ui.core.Fragment.byId("id_feedback", "fileUploader");


			var issueType = issueTypeSelect.getSelectedItem();
			var title = titleInput.getValue();
			var description = descriptionInput.getValue();


			if (!issueType || !title || !description) {
				sap.m.MessageToast.show("Please fill out all required fields before submitting.");
				return;
			}


			sap.m.MessageToast.show("Thank you for your feedback! Our team will review it shortly.");

			issueTypeSelect.setSelectedKey(null);
			prioritySelect.setSelectedKey(null);
			titleInput.setValue("");
			descriptionInput.setValue("");
			fileUploader.clear();
			var oSuccessIcon = sap.ui.core.Fragment.byId("id_feedback", "fileUploadSuccessIcon");
			var oErrorStrip = sap.ui.core.Fragment.byId("id_feedback", "fileUploadError");
			var oFileNameText = sap.ui.core.Fragment.byId("id_feedback", "fileNameText");
			oSuccessIcon.setVisible(false);
			oErrorStrip.setVisible(false);
			oFileNameText.setVisible(false);

			this._oFeedbackPopover1.close();
		},
		onFileChangeFeedback: function (oEvent) {
			const oFileUploader = oEvent.getSource();
			const oFile = oFileUploader.oFileUpload.files[0];
			const maxFileSize = 10 * 1024 * 1024;

			const oSuccessIcon = sap.ui.core.Fragment.byId("id_feedback", "fileUploadSuccessIcon");
			const oErrorStrip = sap.ui.core.Fragment.byId("id_feedback", "fileUploadError");
			const oFileNameText = sap.ui.core.Fragment.byId("id_feedback", "fileNameText");

			if (oFile && !this._isAllowedFile(oFile.name, ["docx", "pdf", "png", "jpg", "jpeg"])) {

				oSuccessIcon.setVisible(false);
				oErrorStrip.setVisible(true);
				oErrorStrip.setText("Invalid file type. Allowed types: .docx, .pdf, .png, .jpg, .jpeg");
				oFileNameText.setVisible(false);
				oFileUploader.setValue("");
			} else if (oFile && oFile.size > maxFileSize) {

				oSuccessIcon.setVisible(false);
				oErrorStrip.setVisible(true);
				oErrorStrip.setText("File size exceeds 10 MB. Please upload a smaller file.");
				oFileNameText.setVisible(false);
				oFileUploader.setValue("");
			} else if (oFile) {

				oSuccessIcon.setVisible(true);
				oErrorStrip.setVisible(false);
				oFileNameText.setText(oFile.name);
				oFileNameText.setVisible(true);
			} else {

				oSuccessIcon.setVisible(false);
				oErrorStrip.setVisible(false);
				oFileNameText.setVisible(false);
			}
		},
		_onFSDGeneratedSuccess: function () {
    // Show nudge link
    this.byId("fsdNudgeBox").setVisible(true);
 
    // Unlock Step 2 in wizard
    this.byId("id_FSDStep1").setValidated(true);
 
    // Swap buttons — show Regenerate, hide Generate
    this.byId("btnFSD1").setVisible(false);
    this.byId("btnFSD2").setVisible(true);
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