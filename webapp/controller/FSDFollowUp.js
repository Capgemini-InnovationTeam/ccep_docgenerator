/* global $:readonly */
/* eslint-disable @sap-ux/fiori-tools/sap-no-dom-insertion, @sap-ux/fiori-tools/sap-timeout-usage */
sap.ui.define([], function () {
	"use strict";

	return {

		// COMBO BOX SELECTION
		onComboBoxSelectionChange: function (oEvent) {
    var selectedText = oEvent.getSource().getSelectedItem().getText();
    this.byId("id_FSDprompt").setValue(selectedText);
    this._userprompt = this.byId("id_FSDprompt").getValue();
},

		// STEP 2 GENERATE (CUSTOMISED PROMPT)
		onGenerateFSDPress2: function () {
    
    this._userprompt = this.byId("id_FSDprompt").getValue();

    if (!this._userprompt) {
        sap.m.MessageToast.show("Please select a valid option first.");
        return;
    }

    this._callSecondAPI();
},
_callSecondAPI: function () {
    var userprompt = this._userprompt;
    var responseapi1 = this.response;

    if (!responseapi1) {
        sap.m.MessageToast.show("No data available to process.");
        return;
    }

    // Read the model the user picked in Step 2's "Select Model" combobox
    var oModelCombo = this.byId("_IDGenComboBox");
    var sModelName = (oModelCombo && oModelCombo.getSelectedKey()) || "GPT_4o";

    // Read temperature / max_tokens from the apiModel (set via Prompt Parameters popover)
    var oApiModel = this.getView().getModel("apiModel");
    var oApiData = oApiModel ? oApiModel.getData() : {};

    var fTemperature = (oApiData.temperature !== undefined && oApiData.temperature !== null)
        ? Number(oApiData.temperature) : 0.0;
    if (fTemperature < 0) { fTemperature = 0; }
    if (fTemperature > 1) { fTemperature = 1; }

    var iMaxTokens = parseInt(oApiData.max_tokens, 10);
    if (!iMaxTokens || iMaxTokens <= 0) { iMaxTokens = null; }       
    if (iMaxTokens && iMaxTokens > 12000) { iMaxTokens = 12000; }    

    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));
    var sApiUrl2 = appModulePath + "/ccep_api/drdtofsdfollowup";
    var sTokenUrl = appModulePath + "/";
    var that = this;
    this.oBusyDialog.open();

    // Helper: build fresh FormData each attempt
    function buildFormData() {
        var oFD = new FormData();
        oFD.append("customized_prompt", userprompt);
        oFD.append("input_json", responseapi1);
        oFD.append("model_name", sModelName);
        oFD.append("temperature", String(fTemperature));
        if (iMaxTokens) {
            oFD.append("max_tokens", String(iMaxTokens));
        }
        return oFD;
    }

    function fetchTokenAndPost(attempt, maxRetries) {
        $.ajax({
            url: sTokenUrl,
            method: "GET",
            headers: { "X-CSRF-Token": "Fetch" },
            complete: function (tokenXHR) {
                var csrfToken = tokenXHR.getResponseHeader("X-CSRF-Token");

                if (!csrfToken && attempt < maxRetries) {
                    setTimeout(function () { fetchTokenAndPost(attempt + 1, maxRetries); }, 1000);
                    return;
                }

                $.ajax({
                    url: sApiUrl2,
                    method: "POST",
                    processData: false,
                    contentType: false,
                    timeout: 0,
                    headers: { "X-CSRF-Token": csrfToken || "" },
                    data: buildFormData(),
                    success: function (response) {
    that.oBusyDialog.close();

    var stringData = (typeof response === "object") ? JSON.stringify(response) : response;
    that.response = stringData;
    that.responseapi1Global = stringData;



    that.createFSD(stringData);
},
                    error: function (jqXHR, textStatus) {
                        if (jqXHR.status === 403 && attempt < maxRetries) {
                            setTimeout(function () { fetchTokenAndPost(attempt + 1, maxRetries); }, 1500);
                        } else if (jqXHR.status === 400) {
                            that.oBusyDialog.close();
                            sap.m.MessageBox.error(
                                "FSD Follow-up rejected the request (400): " +
                                "check model name, temperature (0.0–1.0), or max_tokens (≤12000)."
                            );
                        } else {
                            that.oBusyDialog.close();
                            sap.m.MessageBox.error(
                                "FSD Follow-up failed (attempt " + attempt + "): " + textStatus + " (" + jqXHR.status + ")"
                            );
                        }
                    }
                });
            }
        });
    }

    fetchTokenAndPost(1, 3);
}

	};
});