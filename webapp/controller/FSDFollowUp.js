/* global $:readonly */
sap.ui.define([], function () {
	"use strict";

	return {

		// COMBO BOX SELECTION
		onComboBoxSelectionChange: function (oEvent) {
			var selectedText = oEvent.getSource().getSelectedItem().getText();
			this.byId("id_FSDprompt").setValue(selectedText);
			this._userprompt = this.byId("id_FSDprompt").getValue();

			var oButton1 = this.byId("btnFSD1");
			var oButton2 = this.byId("btnFSD2");
			oButton1.setVisible(!oButton1.getVisible());
			oButton2.setVisible(!oButton2.getVisible());
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
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));
    var sApiUrl2 = appModulePath + "/ccep_api/drd-to-fsd-gpt-4o";
    var sTokenUrl = appModulePath + "/";
    var that = this;
    this.oBusyDialog.open();

    // Helper: build fresh FormData each attempt
    function buildFormData() {
        var oFD = new FormData();
        oFD.append("customized_prompt", userprompt);
        oFD.append("input_json", responseapi1);
        return oFD;
    }

    function fetchTokenAndPost(attempt, maxRetries) {
        $.ajax({
            url: sTokenUrl,
            method: "GET",
            headers: { "X-CSRF-Token": "Fetch" },
            complete: function (tokenXHR) {
                var csrfToken = tokenXHR.getResponseHeader("X-CSRF-Token");

                // Guard: if token is missing, retry once before giving up
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

    fetchTokenAndPost(1, 3); // up to 3 attempts
}

	};
});