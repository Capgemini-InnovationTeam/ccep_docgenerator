sap.ui.define([
    "sap/ui/core/UIComponent",
    "ccepdocgenerator/model/models"
], (UIComponent, models) => {
    "use strict";

    return UIComponent.extend("ccepdocgenerator.Component", {
        metadata: {
            manifest: "json",
            interfaces: [
                "sap.ui.core.IAsyncContentCreation"
            ]
        },

        init() {
            // call the base component's init function
            UIComponent.prototype.init.apply(this, arguments);

            // set the device model
            this.setModel(models.createDeviceModel(), "device");

            // enable routing
            this.getRouter().initialize();
             // Dynamically load external scripts (now served locally, not from CDN)
            this._loadExternalScripts([
                sap.ui.require.toUrl("ccepdocgenerator/thirdparty/FileSaver.min.js"),
                sap.ui.require.toUrl("ccepdocgenerator/thirdparty/docx.min.js")
            ]).then(() => {
                
            }).catch((error) => {
                
            });


            
        },
        _loadExternalScripts(urls) {
            const promises = urls.map((url) => this._loadScript(url));
            return Promise.all(promises);
        },

        _loadScript(src) {
            return new Promise((resolve, reject) => {
                // eslint-disable-next-line @sap-ux/fiori-tools/sap-no-element-creation -- dynamic <script> tag is required to load external libraries at runtime; no UI5 control equivalent exists
                const script = document.createElement("script");
                script.src = src;
                script.async = true;

                script.onload = () => resolve(src);
                script.onerror = () => reject(new Error(`Failed to load script: ${src}`));

                // eslint-disable-next-line @sap-ux/fiori-tools/sap-no-dom-insertion -- required to actually execute the dynamically created script tag
                document.head.appendChild(script);
            });
        }
    });
});