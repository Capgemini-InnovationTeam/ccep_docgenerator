/* global $:readonly */
sap.ui.define([], function () {
    "use strict";

    return {

        
        onComboBoxSelectionChangeTSD: function (oEvent) {
            var selectedTextTsd = oEvent.getSource().getSelectedItem().getText();
            this.byId("id_TSDprompt").setValue(selectedTextTsd);
            this._userpromptTD = this.byId("id_TSDprompt").getValue();

            var oButton1 = this.byId("btnTSD1");
            var oButton2 = this.byId("btnTSD2");
            oButton1.setVisible(!oButton1.getVisible());
            oButton2.setVisible(!oButton2.getVisible());
        },

        
        onGenerateTSDPress2: function () {
            this._userpromptTSD = this.byId("id_TSDprompt").getValue();

            if (!this._userpromptTSD) {
                sap.m.MessageToast.show("Please select a valid option first.");
                return;
            }

            var userpromptTSD      = this._userpromptTSD;
            var development_typeTSD = this.custdeliverytype;

            // Reset flag so CleanCore Recommendations section renders fresh on regeneration
            this.cleancorerecommendation = false;

            if (this.followUpPromptingFlag) {
                
                this._callSecondAPITSD2(userpromptTSD, development_typeTSD, this.followUpPromptResponse);
            } else {
              
                var standardResponse    = this.standard_Response;
                var enhancementResponse = this.developmetTypeResponses;
                var combined            = JSON.stringify({ standardResponse: standardResponse, enhancementResponse: enhancementResponse });
                this.standard_Response  = combined;
                this._callSecondAPITSD2(userpromptTSD, development_typeTSD, combined);
            }
        },

     
_callSecondAPITSD2: function (userpromptTSD, development_typeTSD, standardsresponse_1) {
    this.oBusyDialog.open();

    // 1. Path construction
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));
    var sApiUrlTSD2 = appModulePath + "/tsd_api/fsdtotsdagenticcustom2";
    var sTokenUrl = appModulePath + "/";

    
    var inputJsonStr = standardsresponse_1;
    if (typeof inputJsonStr !== "string") {
        inputJsonStr = JSON.stringify(inputJsonStr);
    }

    var oFormData = new FormData();
    oFormData.append("customized_prompt", userpromptTSD);
    oFormData.append("input_json", inputJsonStr);
    oFormData.append("development_type", (development_typeTSD || "").toLowerCase());
    var that = this;
    $.ajax({
        url: sTokenUrl,
        method: "GET",
        headers: { "X-CSRF-Token": "Fetch" },
        complete: function (jqXHR) {
            var csrfToken = jqXHR.getResponseHeader("X-CSRF-Token");

            $.ajax({
                url: sApiUrlTSD2,
                method: "POST",
                headers: { "X-CSRF-Token": csrfToken || "" },
                processData: false, 
                contentType: false,   
                data: oFormData,
                success: function (response) {
                    that.oBusyDialog.close();

                    
                    var rawString = (typeof response === "string")
                        ? response
                        : JSON.stringify(response);
                    that.followUpPromptResponse = rawString;

                    
                    var oData = (typeof response === "string") ? JSON.parse(response) : response;
                    try {
                        that.developmetTypeResponses = oData.enhancementResponse || oData;
                        that.standard_Response       = oData.standardResponse    || "";
                    } catch {
                        that.developmetTypeResponses = rawString;
                    }

                    that.followUpPromptingFlag = true;

                    if (!that.tsdGenerated) {
                        that.tsdGenerated = true;
                        if (that.getView().byId("id_TSDPromptRow")) {
                            that.getView().byId("id_TSDPromptRow").setVisible(true);
                        }
                    }

                    
                    that.createTSD();
                },
                error: function () {
                    that.oBusyDialog.close();
                    sap.m.MessageToast.show("An unexpected error occurred. Please refresh the page.");
                }
            });
        }
    });
},


        // Open the global prompt repository dialog (TSD side)
        onImportPromtPress: function () {
            if (!this._oPromptRepository) {
                this._oPromptRepository = sap.ui.xmlfragment(
                    "promptParameter2",
                    "ccep.view.fragment.GobalPromptRepository",
                    this
                );
                this.getView().addDependent(this._oPromptRepository);
                this._oPromptRepository.setContentWidth("590px");
            }

            // Swap the edit/save buttons
            this.byId("btnTSD1").setVisible(false);
            this.byId("btnTSD2").setVisible(true);

            this._oPromptRepository.open();
        },

        // Handle double-click on a prompt list item (TSD)
        onPromptItemDoubleClick: function (oEvent) {
            this.onSavePromptSelection(oEvent);
        },

        // Save the selected prompt into the TSD prompt textarea
        onSavePromptSelection: function (oEvent) {
            var oContext     = oEvent.getSource().getBindingContext("promptsModel2");
            var sUserPrompt  = oContext.getProperty("name");
            this.byId("id_TSDprompt").setValue(sUserPrompt);
            this._oPromptRepository.close();
        },

        // Close the prompt repository dialog (TSD)
        onClosePromptSelection: function () {
            if (this._oPromptRepository) {
                this._oPromptRepository.close();
            }
        }
    };
});