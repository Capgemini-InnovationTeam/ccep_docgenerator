/* global $:readonly */
sap.ui.define([], function () {
    "use strict";

    // GLOBAL HELPERS

    function safeObj(val) {
        return (val && typeof val === "object" && !Array.isArray(val)) ? val : {};
    }

    function safeArr(val) {
        return Array.isArray(val) ? val : (val ? [val] : []);
    }

   
    function safeParse(raw, fallback) {
        if (raw === null || raw === undefined) return fallback !== undefined ? fallback : {};
        if (typeof raw !== "string") return raw; // already parsed
        try { return JSON.parse(raw); } catch { return fallback !== undefined ? fallback : {}; }
    }

   function makeKVRow(docx, key, value) {
    var {
        TableRow,
        TableCell,
        Paragraph,
        TextRun,
        WidthType,
        VerticalAlign
    } = docx;

    var displayValue = (
        value === null ||
        value === undefined ||
        value === ""
    )
        ? ""
        : (Array.isArray(value)
            ? value.join(", ")
            : String(value));

    var valueRun;

    switch (key) {

        case "Category":
            valueRun = new TextRun({
                text: displayValue,
                bold: true,
                size: 20,
                font: "Helvetica",
                color: "FFFFFF",
                shading: {
                    fill: "FF0000"
                }
            });
            break;

        case "Type":
            valueRun = new TextRun({
                text: displayValue,
                bold: true,
                size: 20,
                font: "Helvetica",
                color: "000000",
                shading: {
                    fill: "FFA500"
                }
            });
            break;

        case "Global/Local":
            valueRun = new TextRun({
                text: displayValue,
                bold: true,
                size: 20,
                font: "Helvetica",
                color: "FFFFFF",
                shading: {
                    fill: "008000"
                }
            });
            break;

        default:
            valueRun = new TextRun({
                text: displayValue,
                size: 20,
                font: "Helvetica"
            });
    }

    return new TableRow({
        children: [
            new TableCell({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: key,
                                bold: true,
                                size: 20,
                                font: "Helvetica"
                            })
                        ]
                    })
                ],
                shading: {
                    fill: "#e5e7eb"
                },
                width: {
                    size: 3500,
                    type: WidthType.DXA
                },
                verticalAlign: VerticalAlign.CENTER
            }),

            new TableCell({
                children: [
                    new Paragraph({
                        children: [valueRun]
                    })
                ],
                width: {
                    size: 6500,
                    type: WidthType.DXA
                },
                verticalAlign: VerticalAlign.CENTER
            })
        ]
    });
}

    function safeSection(docxRef, documentChildren, sectionName, fn) {
        try {
            fn();
        } catch (e) {
            try {
                var { Paragraph, TextRun } = docxRef;
                documentChildren.push(new Paragraph({
                    children: [new TextRun({
                        text: "⚠ Section [" + sectionName + "] failed to load - please retry. Error: " + (e && e.message ? e.message : String(e)),
                        bold: true,
                        color: "CC0000",
                        size: 20,
                        font: "Helvetica"
                    })],
                    spacing: { before: 200, after: 200 }
                }));
            } catch { /* last-resort silent */ }
        }
    }

    function makeColumnTable(docx, columnTitles, dataRows, keyMap) {
        var { Table, TableRow, TableCell, Paragraph, TextRun } = docx;
        var headerRow = new TableRow({
            children: columnTitles.map(function (title) {
                return new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 20, font: "Helvetica" })] })],
                    shading: { fill: "#e5e7eb" }
                });
            })
        });
        var rows = safeArr(dataRows).map(function (field) {
            return new TableRow({
                children: columnTitles.map(function (title) {
                    var key = (keyMap && keyMap[title]) ? keyMap[title] : title;
                    var raw = field ? field[key] : undefined;
                    var displayVal = (raw === null || raw === undefined || raw === "")
                        ? ""
                        : (Array.isArray(raw) ? raw.join(", ") : String(raw));
                    return new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: displayVal, size: 20, font: "Helvetica" })] })]
                    });
                })
            });
        });
        return new Table({ rows: [headerRow].concat(rows) });
    }

    // MAIN MIXIN
    var TSDGenerator = {

        onFileChangeTSD: function (oEvent) {
            this._aSelectedFilestsd = oEvent.getParameter("files");
            this._nametsd = "TSD";
            if (this._aSelectedFilestsd && this._aSelectedFilestsd.length > 0) {
                var aFiles = this._aSelectedFilestsd;
                var aFileNames = Array.from(aFiles).map(function (f) { return f.name.toLowerCase(); });
                var hasFSD = aFileNames.some(function (n) { return n.includes("fsd"); });
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

        sanitizeObj: function (obj) {
            if (Array.isArray(obj)) {
                obj.forEach(function (item, i) {
                    if (item === null || item === undefined) {
                        obj[i] = "";
                    } else if (typeof item === "object") {
                        TSDGenerator.sanitizeObj(item);
                    }
                });
            } else if (typeof obj === "object" && obj !== null) {
                Object.keys(obj).forEach(function (k) {
                    if (obj[k] === null || obj[k] === undefined) {
                        obj[k] = "";
                    } else if (typeof obj[k] === "object") {
                        TSDGenerator.sanitizeObj(obj[k]);
                    }
                });
            }
            return obj;
        },

onGenerateTSDPress: function () {

    // 🔥 FOLLOW-UP FLOW
    if (this.tsdGenerated === true) {
        return this.onGenerateTSDPress2();
    }

    var selectedFilesTSD = this._aSelectedFilestsd;
    if (!selectedFilesTSD || selectedFilesTSD.length === 0) {
        sap.m.MessageToast.show("Please select a file first.");
        return;
    }

    // Reset flag so CleanCore section always renders fresh
    this.cleancorerecommendation = false;
    this.oBusyDialog.open();

    // 1. Build app module path logic for Destinations
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));

    // 2. Updated relative URLs matching the NEW xs-app.json prefixes
    var sApiUrlTSD = appModulePath + "/tsd_base_api/developmentType";
    var sApiUrlInstructions = appModulePath + "/cleancore_api/Instructions";
    var sApiUrlTSD_standard = appModulePath + "/standard_api/standardsections";

    var oFormDataTSD = new FormData();
    var oFormDataInstructions = new FormData();

    oFormDataTSD.append("files", selectedFilesTSD[0]);
    oFormDataInstructions.append("files", selectedFilesTSD[0]);
    oFormDataInstructions.append("file", selectedFilesTSD[0]);

    var that = this;

    // 3. Fetch CSRF Token from app base path (API endpoints only support POST)
    var sTokenUrl = appModulePath + "/";
    $.ajax({
        url: sTokenUrl,
        method: "GET",
        headers: { "X-CSRF-Token": "Fetch" },
        complete: function (jqXHR) {
            var csrfToken = jqXHR.getResponseHeader("X-CSRF-Token");
            var commonHeaders = { "X-CSRF-Token": csrfToken || "" };

            // 4. Trigger Parallel API Calls via Destination
            var firstApiCall = $.ajax({ url: sApiUrlTSD, method: "POST", headers: commonHeaders, processData: false, contentType: false, data: oFormDataTSD, timeout: 0 });
            var secondApiCall = $.ajax({ url: sApiUrlInstructions, method: "POST", headers: commonHeaders, processData: false, contentType: false, data: oFormDataInstructions, timeout: 0 });
            var standardApiCall = $.ajax({ url: sApiUrlTSD_standard, method: "POST", headers: commonHeaders, processData: false, contentType: false, data: oFormDataInstructions, timeout: 0 });

            var firstResult = null, secondResult = null, standardResult = null;
            var firstDeferred = $.Deferred(), secondDeferred = $.Deferred(), standardDeferred = $.Deferred();

            firstApiCall
                .done(function (r) { firstResult = r; firstDeferred.resolve(); })
                .fail(function () { firstDeferred.reject("developmentType"); });

            secondApiCall
                .done(function (r) { secondResult = r; secondDeferred.resolve(); })
                .fail(function () { secondDeferred.reject("instructions"); });

            standardApiCall
                .done(function (r) { standardResult = r; standardDeferred.resolve(); })
                .fail(function () { standardResult = null; standardDeferred.resolve(); });

            $.when(firstDeferred, secondDeferred, standardDeferred).then(
                function () {
                    that.oBusyDialog.close();

                    if (!firstResult) {
                        sap.m.MessageBox.error("Failed to determine development type.", { title: "TSD Generation Failed" });
                        return;
                    }

                    // Parse developmentType response
                    var firstResponseData2 = (typeof firstResult === "string") ? JSON.parse(firstResult) : firstResult;
                    that.sanitizeObj(firstResponseData2);
                    that.RESPONSE = firstResponseData2;

                    that.firstResponseData =
                        (firstResponseData2.response3 && firstResponseData2.response3.BespokeDevelopment) ? firstResponseData2.response3.BespokeDevelopment :
                        (firstResponseData2.response3 && firstResponseData2.response3.Enhancement) ? firstResponseData2.response3.Enhancement : [];

                    // Parse CleanCore / Instructions response
                    var secondParsed = (typeof secondResult === "string") ? JSON.parse(secondResult) : secondResult;
                    that.sanitizeObj(secondParsed);
                    that.secondResponseData = secondParsed;

                    // Parse standardsections response & Fallback
                    var fallbackStandard = [{ "TechnicalUnitTestDetails": [{ "Testing approach": "Unit testing approach to be defined", "Completion criteria": "All test cases pass", "Test data and other needs": "Test data to be provided", "Testing dependencies": "No dependencies", "Technical unit test design": "To be defined", "Test methods (for reference)": "Manual testing", "Technical unit test executors": { "Email": "", "Program Under Test": "", "Programmer Name": "GenAI", "Test Date": "", "Tester Name (if different)": "" }, "Technical unit test case scenarios": [{ "ID": "TC001", "Test Objective": "To be defined", "Test Steps/Procedures": "To be defined", "Test Data /Condition": "To be defined", "Expected Results": "To be defined", "Pass": "", "Fail": "", "Priority": "Medium", "Comments": "" }] }] }, { "response2": [{ "brief_description": "Technical specification document generated by GenAI", "BespokeDevelopmentDetails": [{ "Application/Platform": "", "Object Name": "", "Object Type": "", "Object Description": "", "Step By Step Instructions": "", "Global/Local": "", "Business Unit/Country": "", "Localization Reason": "", "Target/Intermediate State": "", "Related JIRA Issues": "", "Additional Information": "" }], "Dependencies": "" }, { "Title": "Technical Specification Document", "ReferenceItems": [{ "Incoming References": "", "Reference Applications & IT Components LeanIX": "", "Reference Business Capabilities LeanIX": "", "Reference Data Objects Precisely": "", "Reference Functional Specification": "", "Reference Integrated Business Scenarios Signavio": "", "Reference Process Signavio": "" }], "RelatedItems": [{ "Dependent Documents": "", "Other Related Documents": "", "Related JIRA Issues": "", "Related Key Design Decisions": "", "Related Requirements": "" }], "classification": [{ "Business Units/Countries": "", "Category": "", "Complexity": "", "Global/Local": "", "Group": "", "Localization Reason": "", "Release": "", "Sub Group": "", "Target/Intermediate State": "", "Type": "" }], "PrinciplesPolices": "", "Assumptions": "" }, { "Data_description": [{ "NewDomains": { "Domain Name": "", "Description": "", "DataType/Length": "", "OutputLength": "", "LowerCase": "", "CheckTable": "" }, "NewDataElements": { "DataElementName": "", "Description": "", "Domain": "", "FieldLabels": "", "ParameterID": "" }, "ExistingTablesAndStructures": { "DataElement": "", "Description": "", "Domain": "", "FieldName": "" }, "Newtablesandstructures": { "TableStructureName": "", "Type": "", "Description": "", "InitialSizeInRows": "", "MaximumSizeInRows": "", "DataClass": "", "Buffering": "", "BufferingType": "", "LogDataChanges": "" }, "Fields": { "DataElement": "", "Description": "", "Domain": "", "FieldName": "", "KeyFields": "" }, "Newindexes": { "Buffering": "", "BufferingType": "", "DataClass": "", "IndexName": "", "IndexType": "", "InitialSizeInRows": "", "LogDataChanges": "", "MaximumSizeInRows": "", "ShortDescription": "", "TableName": "" }, "Indexfields": { "Field Description": "", "Field Name": "" }, "ForeignKey": { "Cardinality Right Side": "", "CardinalityLeftSide": "", "CheckTable": "", "ErrorMessage": "", "ForeignKeyField": "", "ForeignKeyFieldType": "", "ForeignKeyTable": "", "ScreenCheckRequired": "", "ShortText": "" }, "Update method for tables": "" }] }], "response3": { "TR_details": [], "impacted_objects_details": [], "Object_oriented_details": [], "FM_details": [], "Background_job_details": [], "Code_review": [{ "ATC / Onapsis code review report": "", "ESLint / Onapsis code review report": "", "Peer review code": "", "Checklist": "" }], "authorization_objects_details": [] } }];

                    var standardParsed;
                    if (standardResult) {
                        standardParsed = (typeof standardResult === "string") ? JSON.parse(standardResult) : standardResult;
                        that.sanitizeObj(standardParsed);
                    } else {
                        standardParsed = fallbackStandard;
                    }
                    that.standard_Response = standardParsed;

                    // Extract title safely
                    try {
                        that.Tittle = standardParsed[1]["response2"][1]["Title"] || "Technical Specification Document";
                    } catch {
                        that.Tittle = "Technical Specification Document";
                    }

                    that.createTSD();
                    that._onTSDGeneratedSuccess();
                },
                function (failedApi) {
                    that.oBusyDialog.close();
                    sap.m.MessageBox.error("A critical API call failed (" + failedApi + "). Please refresh and try again.");
                }
            );
        }
    });
},

       
        _makeApiCall: function (sRelativeEndpoint, label) {
    var selectedFilesTSD = this._aSelectedFilestsd;
    if (!selectedFilesTSD || selectedFilesTSD.length === 0) {
        sap.m.MessageToast.show("Please select a file first.");
        return Promise.reject("No files selected.");
    }
    this.oBusyDialog.open();
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));
    var sApiUrl = appModulePath + "/tsd_base_api/" + sRelativeEndpoint;

    var oFormData = new FormData();
    oFormData.append("files", selectedFilesTSD[0]);

    var that = this;

    var sTokenUrlBase = appModulePath + "/";
    return new Promise(function (resolve, reject) {
        $.ajax({
            url: sTokenUrlBase,
            method: "GET",
            headers: { "X-CSRF-Token": "Fetch" },
            complete: function (jqXHR) {
                var csrfToken = jqXHR.getResponseHeader("X-CSRF-Token");
                $.ajax({
                    url: sApiUrl,
                    method: "POST",
                    headers: { "X-CSRF-Token": csrfToken || "" },
                    processData: false, // Required for FormData
                    contentType: false,   // Required for FormData
                    data: oFormData,
                    success: function (response) {
                        that.oBusyDialog.close();
                        var responseData = (typeof response === "string") ? JSON.parse(response) : response;
                        that.developmetTypeResponses = responseData;

                        if (!responseData) {
                            reject(new Error("Invalid response from " + label));
                        } else {
                            that.sanitizeObj(responseData);
                            resolve(responseData);
                        }
                    },
                    error: function (err) {
                        sap.m.MessageToast.show("API call [" + label + "] failed.");
                        that.oBusyDialog.close();
                        reject(err);
                    }
                });
            }
        });
    });
},
callfunction_enhancement: function () { return this._makeApiCall("enhancement", "enhancement"); },
callfunction_Forms:       function () { return this._makeApiCall("forms",       "forms");       },
callfunction_Fiori_App:   function () { return this._makeApiCall("fiori",       "fiori");       },
callfunction_workflow:    function () { return this._makeApiCall("workflow",    "workflow");    },
callfunction_Report:      function () { return this._makeApiCall("report",      "report");      },

        // CREATE TSD ORCHESTRATOR
        createTSD: function () {
            var that = this;
            var documentChildren = [];

            // standard_Response is now always a parsed object/array
            var stdResp = that.standard_Response;

            documentChildren = that.createHeadingTSD(documentChildren, stdResp);
            documentChildren = that.createTOCTSD(documentChildren, that.firstResponseData);
            documentChildren = that.createInitialCommonSectionTSD(documentChildren, stdResp);

            var promises = [];
            var dataItem = (that.firstResponseData && that.firstResponseData[0]) ? that.firstResponseData[0] : "";

            var isEnhancement = dataItem.includes("Enhancement") || dataItem.includes("ENHANCEMENT") || dataItem.includes("enhancement");
            var isForm = dataItem.includes("FORM") || dataItem.includes("Form");
            var isFiori = dataItem.includes("Fiori App") || dataItem.includes("FIORI APP");
            var isWorkflow = dataItem.includes("Workflow") || dataItem.includes("WORKFLOW");
            var isReport = dataItem.includes("EMBEDDED ANALYTICS") || dataItem.includes("Report");

            // Also detect type from RESPONSE object keys (new response format from API)
            if (!isEnhancement && !isForm && !isFiori && !isWorkflow && !isReport) {
                var r3Keys = that.RESPONSE && that.RESPONSE.response3 ? Object.keys(that.RESPONSE.response3) : [];
                // Check Workflow FIRST — its keys are highly specific; Fiori check comes after
                if (r3Keys.indexOf("Workflowdetails") >= 0 ||
                    r3Keys.indexOf("Additional_Information") >= 0 ||
                    r3Keys.indexOf("FM_Details") >= 0 ||
                    r3Keys.indexOf("BADI_Implementation_Details") >= 0 ||
                    r3Keys.indexOf("Deadline_Monitoring") >= 0) {
                    isWorkflow = true;
                } else if (r3Keys.indexOf("Enhancement") >= 0) {
                    isEnhancement = true;
                } else if (r3Keys.indexOf("Forms") >= 0) {
                    isForm = true;
                } else if (r3Keys.indexOf("Fiori_App") >= 0 || r3Keys.indexOf("Fiori App") >= 0) {
                    isFiori = true;
                }
            }

            // secondResponseData is the parsed cleancore object
            var cleanCoreData = that.secondResponseData;

            if (that.followUpPromptingFlag) {
                // Data already available from previous API call stored in developmetTypeResponses
                var d = safeParse(that.developmetTypeResponses, {});
                that.sanitizeObj(d);
                if (isEnhancement) { that.custdeliverytype = "Enhancement"; documentChildren = that.createEnhancementTSD(documentChildren, d, cleanCoreData); }
                else if (isForm) { that.custdeliverytype = "FORM"; documentChildren = that.createFormTSD(documentChildren, d, cleanCoreData); }
                else if (isFiori) { that.custdeliverytype = "Fiori App"; documentChildren = that.createFioriTSD(documentChildren, d, cleanCoreData); }
                else if (isWorkflow) { that.custdeliverytype = "Workflow"; documentChildren = that.createWorkflowTSD(documentChildren, d, cleanCoreData); }
                else if (isReport) { that.custdeliverytype = "Reports"; documentChildren = that.createReportTSD(documentChildren, d, cleanCoreData); }

                documentChildren = that.createLastCommonSectionTSD(documentChildren, stdResp);
                that._finalizeTSD(documentChildren);
            } else {
                if (isEnhancement) {
                    that.custdeliverytype = "Enhancement";
                    promises.push(that.callfunction_enhancement().then(function (res) {
                        documentChildren = that.createEnhancementTSD(documentChildren, res, cleanCoreData);
                    }).catch(function () {
                        // Fallback: use already-available RESPONSE data
                        if (that.RESPONSE) {
                            documentChildren = that.createEnhancementTSD(documentChildren, that.RESPONSE, cleanCoreData);
                        }
                    }));
                } else if (isForm) {
                    that.custdeliverytype = "FORM";
                    promises.push(that.callfunction_Forms().then(function (res) {
                        documentChildren = that.createFormTSD(documentChildren, res, cleanCoreData);
                    }).catch(function () {
                        // Fallback: use already-available RESPONSE data from developmentType call
                        var fallbackRes = that.RESPONSE || {};
                        documentChildren = that.createFormTSD(documentChildren, fallbackRes, cleanCoreData);
                    }));
                } else if (isFiori) {
                    that.custdeliverytype = "Fiori App";
                    promises.push(that.callfunction_Fiori_App().then(function (res) {
                        documentChildren = that.createFioriTSD(documentChildren, res, cleanCoreData);
                    }).catch(function () {
                        // Fallback: use already-available RESPONSE data from developmentType call
                        var fallbackRes = that.RESPONSE || {};
                        documentChildren = that.createFioriTSD(documentChildren, fallbackRes, cleanCoreData);
                    }));
                } else if (isWorkflow) {
                    that.custdeliverytype = "Workflow";
                    promises.push(that.callfunction_workflow().then(function (res) {
                        documentChildren = that.createWorkflowTSD(documentChildren, res, cleanCoreData);
                    }).catch(function () {
                        if (that.RESPONSE) {
                            documentChildren = that.createWorkflowTSD(documentChildren, that.RESPONSE, cleanCoreData);
                        }
                    }));
                } else if (isReport) {
                    that.custdeliverytype = "Report";
                    promises.push(that.callfunction_Report().then(function (res) {
                        documentChildren = that.createReportTSD(documentChildren, res, cleanCoreData);
                    }).catch(function () {}));
                }

                Promise.all(promises).then(function () {
                    documentChildren = that.createLastCommonSectionTSD(documentChildren, stdResp);
                    that._finalizeTSD(documentChildren);
                }).catch(function () {});
            }
        },

        _finalizeTSD: function (documentChildren) {
            var that = this;

            var doc = new window.docx.Document({
                sections: [{ children: documentChildren }]
            });

            window.docx.Packer.toBlob(doc)
                .then(function (blob) {

                    if (window.saveAs) {

                        var tittle = that.Tittle || "Technical Specification Document";

                        // Clean file name
                        var sanitizedTitle = tittle.replace(/[^a-zA-Z0-9\s]/g, "");
                        sanitizedTitle = sanitizedTitle.length > 60
                            ? sanitizedTitle.slice(0, 57).trim() + "..."
                            : sanitizedTitle;

                        // Save file
                        window.saveAs(blob, sanitizedTitle + ".docx");



                        // ✅ Always show prompt row
                        var oPromptRow = that.getView().byId("id_TSDPromptRow");
                        if (oPromptRow) {
                            oPromptRow.setVisible(true);
                        }
                    }

                })
                .catch(function () {
                });
        },

        // createHeadingTSD  — accepts parsed object directly
        createHeadingTSD: function (documentChildren, response) {
            var D = window.docx;
            var { Paragraph, TextRun, Table, TableRow, TableCell,
                AlignmentType, VerticalAlign } = D;

            // response is now always a parsed array
            var response2 = safeArr(response[1] && response[1]["response2"]);
            var title = (response2[1] && response2[1].Title) ? response2[1].Title : "Technical Specification Document";

            documentChildren.push(new Paragraph({
                children: [new TextRun({ text: "TSD- " + title, bold: true, size: 32, font: "Helvetica" })],
                spacing: { before: 0, after: 250 }
            }));

            documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Properties", bold: true, size: 20, font: "Helvetica" })], spacing: { after: 250 } }));

            var today2 = new Date();
            var currentDate2 = today2.toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });

            var propRows = [
                ["Deliverable ID", "<Deliverable ID>"],
                ["Document Type", "Technical Specification Document (TSD)"],
                ["Document Purpose", "To capture the technical development objects which need to be built to implement or were built as a result of implementing the referenced functional specification"],
                ["Document Status", ""],
                ["Created By", "GenAI"],
                ["Created On", currentDate2],
                ["Contributors", "GenAI"]
            ].map(function (pair) {
                return new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: pair[0], bold: true, size: 20, font: "Helvetica" })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
                            width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: pair[1], size: 20, font: "Helvetica" })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
                            width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER
                        })
                    ],
                    height: { value: 500, rule: "atLeast" }
                });
            });
            documentChildren.push(new Table({ rows: propRows }));

            // Reference Items
            documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Reference Items", bold: true, size: 24, font: "Helvetica" })], alignment: AlignmentType.LEFT, spacing: { after: 250, before: 150 } }));
            var refItems2 = (response2[1] && response2[1].ReferenceItems) ? response2[1].ReferenceItems : [{}];
            var refItems = safeObj(safeArr(refItems2)[0]);
            var refData = [
                { "Reference Type": "Incoming References", "Value": refItems["Incoming References"] || "" },
                { "Reference Type": "Reference Applications & IT Components LeanIX", "Value": refItems["Reference Applications & IT Components LeanIX"] || "" },
                { "Reference Type": "Reference Business Capabilities LeanIX", "Value": refItems["Reference Business Capabilities LeanIX"] || "" },
                { "Reference Type": "Reference Data Objects Precisely", "Value": refItems["Reference Data Objects Precisely"] || "" },
                { "Reference Type": "Reference Functional Specification", "Value": refItems["Reference Functional Specification"] || "" },
                { "Reference Type": "Reference Integrated Business Scenarios Signavio", "Value": refItems["Reference Integrated Business Scenarios Signavio"] || "" },
                { "Reference Type": "Reference Process Signavio", "Value": refItems["Reference Process Signavio"] || "" }
            ];
            documentChildren.push(makeColumnTable(D, ["Reference Type", "Value"], refData));

            // Related Items
            documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Related Items", bold: true, size: 24, font: "Helvetica" })], alignment: AlignmentType.LEFT, spacing: { after: 250, before: 150 } }));
            var relItemsData = safeObj(safeArr((response2[1] && response2[1].RelatedItems) ? response2[1].RelatedItems : [{}])[0]);
            var relData = [
                { "Related Type": "Dependent Documents", "Value": relItemsData["Dependent Documents"] || "" },
                { "Related Type": "Other Related Documents", "Value": relItemsData["Other Related Documents"] || "" },
                { "Related Type": "Related JIRA Issues", "Value": relItemsData["Related JIRA Issues"] || "" },
                { "Related Type": "Related Key Design Decisions", "Value": relItemsData["Related Key Design Decisions"] || "" },
                { "Related Type": "Related Requirements", "Value": relItemsData["Related Requirements"] || "" }
            ];
            documentChildren.push(makeColumnTable(D, ["Related Type", "Value"], relData));

            // Classification
            documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Classification", bold: true, size: 24, font: "Helvetica" })], alignment: AlignmentType.LEFT, spacing: { after: 250, before: 150 } }));
            var cls = safeObj(safeArr((response2[1] && response2[1].classification) ? response2[1].classification : [{}])[0]);
            var clsKVs = [
                ["Category", cls["Category"] || ""],
                ["Type", cls["Type"] || ""],
                ["Global/Local", cls["Global/Local"] || ""],
                ["Business Units Countries", cls["Business Units/Countries"] || ""],
                ["Localization Reason", cls["Localization Reason"] || ""],
                ["Target/Intermediate State", cls["Target/Intermediate State"] || ""],
                ["Complexity", cls["Complexity"] || ""],
                ["Release", cls["Release"] || ""],
                ["Group", cls["Group"] || ""],
                ["Sub Group", cls["Sub Group"] || ""]
            ];
            var clsRows = clsKVs.map(function (pair) { return makeKVRow(D, pair[0], pair[1]); });
            documentChildren.push(new Table({ rows: clsRows }));

            documentChildren.push(new Paragraph({ children: [], pageBreakBefore: true }));
            return documentChildren;
        },

        // createTOCTSD
        createTOCTSD: function (documentChildren, responseTOC) {
            var D = window.docx;
            var { Paragraph, TextRun } = D;

            function tocLine(text) {
                return new Paragraph({ children: [new TextRun({ text: text, size: 16, font: "Helvetica", color: "0000FF" })], spacing: { after: 100 } });
            }

            var commonLines = [
                "Table of Contents", " ● Brief Description", " ● Approved Versions, Change Control",
                " ● Principles, Policies, Standards & Guidelines", " ● Bespoke Development Details",
                " \t○ Transport requests & Rev-Trac", " \t\t▪ Transport request details:", " \t\t▪ List of impacted objects:",
                " \t○ Data descriptions", " \t\t▪ New domains", "\t\t\t▪ New data elements",
                "\t\t\t▪ Existing tables and structures", " \t\t▪ New tables and structures", " \t\t▪ Fields",
                " \t\t▪ New indexes", " \t\t▪ Index fields", " \t\t▪ Foreign keys", " \t\t▪ Update method for tables",
                " \t○ Data Dictionary Objects", " \t○ Non-Data Dictionary Objects"
            ];
            commonLines.forEach(function (l) { documentChildren.push(tocLine(l)); });

            var dataItem = (responseTOC && responseTOC[0]) ? responseTOC[0] : "";
            var typeLines = [];
            if (dataItem.includes("Enhancement") || dataItem.includes("ENHANCEMENT") || dataItem.includes("enhancement")) {
                typeLines = ["● Enhancement", "\t○ Data Dictionary Objects", "\t○ Table details", "\t○ Structure details", "\t○ Non-Data Dictionary Objects", "\t○ Custom business object", "\t\t▪ Key User Extension", "\t\t▪ Process in which BADI/exit/enhancement triggers", "\t\t▪ Enhancement control framework:", "\t\t▪ BADI Implementation Details", "\t\t▪ Exit"];
            } else if (dataItem.includes("FORM") || dataItem.includes("Form")) {
                typeLines = ["● Form", "\t○ Functional Requirements", "\t○ Triggers", "\t○ Data Selection, Validation & Processing", "\t○ Expected Output, Layout", "\t○ Hardcopy, Stationary & Printer Requirements", "\t○ Email Functionality"];
            } else if (dataItem.includes("Fiori App") || dataItem.includes("FIORI APP")) {
                typeLines = ["● Fiori App", "\t○ Technical details", "\t○ Fiori app Name & Type", "\t○ Selection screen", "\t○ Data Selection", "\t\t▪ Special Processing", "\t\t▪ Variants, Variant Data", "\t\t▪ Screens, Screen Flow", "\t\t▪ Program Flow & Pseudo-code", "\t\t\t▪ Process flow", "\t\t\t▪ Pseudo code", "\t○ Security Information", "\t\t▪ Exception Handling", "\t\t▪ Additional Information", "\t○ OData", "\t\t▪ Entity Diagram", "\t\t▪ Service Details", "\t\t▪ Entity 1", "\t○ UI5", "\t\t▪ UI5 App Details", "\t\t▪ UI5 App - Screen Flow"];
            } else if (dataItem.includes("EMBEDDED ANALYTICS") || dataItem.includes("Report")) {
                typeLines = ["● Report", "\t○ Technical details", "\t○ Report Name & Type", "\t○ Selection screen", "\t○ Data Selection", "\t\t▪ Special Processing", "\t\t▪ Variants, Variant Data", "\t\t▪ Screens, Screen Flow", "\t\t▪ Program Flow & Pseudo-code", "\t○ Security Information", "\t○ OData", "\t○ UI5"];
            } else if (dataItem.includes("Workflow") || dataItem.includes("WORKFLOW")) {
                typeLines = [
                    "● Workflow","\t○ Workflow General Information","\t○ Workflow Template Diagram",
                    "\t○ Workflow Scenario Definition",
                    "\t○ Workflow Container",
                    "\t○ Binding Workflow Container to Event Container",
                    "\t○ Steps Details:",
                    "\t○ Activity details",
                    "\t○ Agent Rules",
                    "\t○ Binding Workflow Container to Rule Container",
                    "\t○ Value Helps",
                    "\t○ Email Templates Details",
                    "\t○ Post Approval Process",
                    "\t○ Rejection Scenarios",
                    "\t○ Deadline Monitoring",
                    "\t○ External Connectivity",
                    "\t○ Workflow In App",
                    "\t○ Workflow Steps",
                    "\t○ Responsibility Management Details",
                    "\t○ Enhancement",
                    "\t○ Non-Data Dictionary Objects (NA)",
                    "\t○ Process in which BADI/exit/enhancement triggers",
                    "\t○ Enhancement control framework:",
                    "\t\t▪ Table: ZAXP_STP_ENH_HDR",
                    "\t\t▪ Table: ZAXP_STP_ENH_ITM",
                    "\t○ BADI Implementation Details"
                ];
            }
            typeLines.forEach(function (l) { documentChildren.push(tocLine(l)); });

            var bottomLines = [
                "● Technical unit test details", "\t○ Testing approach", "\t○ Completion criteria",
                "\t○ Test data and other needs", "\t○ Test tools and environment", "\t○ Testing scope",
                "\t○ Testing dependencies", "\t○ Technical unit test design", "\t○ Test methods (for reference)",
                "\t○ Technical unit test executors", "\t○ Technical unit test case scenarios",
                "● Additional Information", "\t○ Object Oriented Details:", "\t○ FM details:", "\t○ Background job",
                "● Code review", "\t○ ATC / Onapsis code review report", "\t○ ESLint / Onapsis code review report",
                "\t○ Peer review code", "\t○ Checklist", "● Assumptions", "\t○ Authorization objects", "\t○ Dependencies"
            ];
            bottomLines.forEach(function (l) { documentChildren.push(tocLine(l)); });

            documentChildren.push(new Paragraph({ children: [], pageBreakBefore: true }));
            return documentChildren;
        },

        // createInitialCommonSectionTSD  — accepts parsed object directly
        createInitialCommonSectionTSD: function (documentChildren, response) {
            var D = window.docx;
            var { Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, VerticalAlign, WidthType, Bookmark, ExternalHyperlink } = D;

            // response is the parsed standard_Response array
            var response2 = safeArr(response[1] && response[1]["response2"]);
            var response3 = safeObj(response[1] && response[1]["response3"]);

            // Brief Description
            safeSection(D, documentChildren, "Brief Description", function () {
                documentChildren.push(new Paragraph({
                    children: [new Bookmark({ id: "BriefDescription", children: [new TextRun({ text: "Brief Description", bold: true, size: 32, font: "Helvetica" })] })],
                    spacing: { after: 100 }
                }));
                var briefDescription = (response2[0] && response2[0]["brief_description"]) ? response2[0]["brief_description"] : "";
                if (briefDescription) {
                    briefDescription.split(/(?<=\.)\s+/).map(function (l) { return l.trim(); }).filter(function (l) { return l.length > 0; }).forEach(function (line) {
                        documentChildren.push(new Paragraph({ children: [new TextRun({ text: line, size: 20, font: "Helvetica" })], spacing: { after: 100 }, alignment: AlignmentType.JUSTIFIED }));
                    });
                }
            });

            // Approved Versions
            safeSection(D, documentChildren, "Approved Versions", function () {
                documentChildren.push(new Paragraph({
                    children: [new Bookmark({ id: "ApprovedVersionsChangeControl", children: [new TextRun({ text: "Approved Versions, Change Control", bold: true, size: 32, font: "Helvetica" })] })],
                    spacing: { before: 250, after: 250 }
                }));
                var avRows = [
                    new TableRow({
                        children: ["Approved Version", "Status", "Description, Reason for Change", "Additional Information"].map(function (h) {
                            return new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 20, font: "Helvetica" })] })], shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER });
                        }),
                        height: { value: 500, rule: "atLeast" }
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "1.00", bold: true, size: 20, font: "Helvetica" })] })], verticalAlign: VerticalAlign.CENTER }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Draft", color: "228B22", size: 20, font: "Helvetica" })] })], verticalAlign: VerticalAlign.CENTER }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Initial Version", size: 20, font: "Helvetica" })] })], verticalAlign: VerticalAlign.CENTER }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Created by GenAI", size: 20, font: "Helvetica" })] })], verticalAlign: VerticalAlign.CENTER })
                        ],
                        height: { value: 500, rule: "atLeast" }
                    })
                ];
                documentChildren.push(new Table({ rows: avRows }));
            });

            // Principles, Policies
            safeSection(D, documentChildren, "Principles Policies", function () {
                documentChildren.push(new Paragraph({
                    children: [new Bookmark({ id: "PrinciplePoliciesStandardsGuidelines", children: [new TextRun({ text: "Principles, Policies, Standards & Guidelines", bold: true, size: 32, font: "Helvetica" })] })],
                    spacing: { before: 250, after: 250 }
                }));
                var ppRows = [
                    new TableRow({
                        children: ["Documents", "Description", "Additional Information"].map(function (h) {
                            return new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 20, font: "Helvetica" })] })], shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER });
                        }),
                        height: { value: 500, rule: "atLeast" }
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new ExternalHyperlink({ link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-principles", children: [new TextRun({ text: "CCEP Architecture Design Principles", style: "Hyperlink", size: 20, font: "Helvetica" })] })] })] }), // eslint-disable-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint reference documentation
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Enterprise, Data & Integration Architecture Principles", size: 20, font: "Helvetica" })] })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 20, font: "Helvetica" })] })] })
                        ],
                        height: { value: 500, rule: "atLeast" }
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new ExternalHyperlink({ link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-policies-and-standards", children: [new TextRun({ text: "CCEP Enterprise Architecture Policies & Standards", style: "Hyperlink", size: 20, font: "Helvetica" })] })] })] }), // eslint-disable-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint reference documentation
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Data management, Integration & Non-platform Application, Infrastructure & Security Policies & Standards", size: 20, font: "Helvetica" })] })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 20, font: "Helvetica" })] })] })
                        ],
                        height: { value: 500, rule: "atLeast" }
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new ExternalHyperlink({ link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-patterns", children: [new TextRun({ text: "CCEP Enterprise Architecture Patterns", style: "Hyperlink", size: 20, font: "Helvetica" })] })] })] }), // eslint-disable-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint reference documentation
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Integration Architecture Patterns", size: 20, font: "Helvetica" })] })] }),
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "", size: 20, font: "Helvetica" })] })] })
                        ],
                        height: { value: 500, rule: "atLeast" }
                    })
                ];
                documentChildren.push(new Table({ rows: ppRows }));
            });

            // Bespoke Development Details
            safeSection(D, documentChildren, "Bespoke Development Details", function () {
                documentChildren.push(new Paragraph({
                    children: [new Bookmark({ id: "BespokeDevelopmentDetails", children: [new TextRun({ text: "Bespoke Development Details", bold: true, size: 32, font: "Helvetica" })] })],
                    spacing: { before: 250, after: 250 }
                }));
                var bespoke = safeObj(safeArr((response2[0] && response2[0].BespokeDevelopmentDetails) ? response2[0].BespokeDevelopmentDetails : [{}])[0]);
                var bdKeys = ["Application/Platform", "Object Name", "Object Type", "Object Description", "Step By Step Instructions", "Global/Local", "Business Unit/Country", "Localization Reason", "Target/Intermediate State", "Related JIRA Issues", "Additional Information"];
                var bdRows = bdKeys.map(function (k) {
                    var val = bespoke[k] || "";
                    var cellParagraph;
                    if (k === "Step By Step Instructions" && val) {
                        var steps = val.split(/\n/).filter(function (s) { return s.trim(); });
                        cellParagraph = new Paragraph({ children: steps.map(function (s) { return new TextRun({ text: s.trim(), size: 20, font: "Helvetica", break: 1 }); }) });
                    } else {
                        cellParagraph = new Paragraph({ children: [new TextRun({ text: val, size: 20, font: "Helvetica" })] });
                    }
                    return new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, size: 20, font: "Helvetica" })] })], shading: { fill: "#e5e7eb" }, width: { size: 3500, type: WidthType.DXA } }),
                            new TableCell({ children: [cellParagraph], width: { size: 6500, type: WidthType.DXA } })
                        ]
                    });
                });
                documentChildren.push(new Table({ rows: bdRows }));
            });

            // Transport Requests
            safeSection(D, documentChildren, "Transport Requests", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TransportRequestsRevTrac", children: [new TextRun({ text: "Transport requests & Rev-Trac", bold: true, size: 32, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TransportRequestDetails", children: [new TextRun({ text: "Transport request details:", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var trDetails = safeArr(response3.TR_details);
                var trData = trDetails.length ? trDetails : [{ "Object Name": "", "Transport Number": "", "Type": "" }];
                documentChildren.push(makeColumnTable(D, ["Object Name", "Transport Number", "Type"], trData));
            });

            // Impacted Objects
            safeSection(D, documentChildren, "Impacted Objects", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "ImpactedObjects", children: [new TextRun({ text: "List of impacted objects:", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var ioDetails = safeArr(response3.impacted_objects_details);
                var ioData = ioDetails.length ? ioDetails : [{ "Object Impacted": "", "Type": "", "Additional Information": "" }];
                documentChildren.push(makeColumnTable(D, ["Object Impacted", "Type", "Additional Information"], ioData));
            });

            // Data Descriptions
            safeSection(D, documentChildren, "Data Descriptions", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "DataDescriptions", children: [new TextRun({ text: "Data descriptions", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));

                // response2[2] may be { Data_description: [ {...} ] }
                var dataDescWrapper = response2[2] && response2[2]["Data_description"]
                    ? safeArr(response2[2]["Data_description"])
                    : [];
                var dataDesc = safeObj(dataDescWrapper[0]);

                function buildDataSection(heading, id, obj, fieldMap) {
                    documentChildren.push(new Paragraph({ children: [new Bookmark({ id: id, children: [new TextRun({ text: heading, bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                    var safeO = safeObj(obj);
                    var rows = Object.keys(fieldMap).map(function (displayKey) {
                        return makeKVRow(D, displayKey, safeO[fieldMap[displayKey]]);
                    });
                    if (rows.length) documentChildren.push(new Table({ rows: rows }));
                }

                buildDataSection("New domains", "NewDomains", dataDesc.NewDomains, {
                    "Domain Name": "Domain Name", "Description": "Description", "DataType/Length": "DataType/Length",
                    "Output Length": "OutputLength", "Lower Case": "LowerCase", "Check Table": "CheckTable"
                });
                buildDataSection("New data elements", "NewDataElements", dataDesc.NewDataElements, {
                    "DataElementName": "DataElementName", "Description": "Description", "Domain": "Domain",
                    "FieldLabels": "FieldLabels", "ParameterID": "ParameterID"
                });
                buildDataSection("Existing tables and structures", "ExistingTablesStructures", dataDesc.ExistingTablesAndStructures, {
                    "DataElement": "DataElement", "Description": "Description", "Domain": "Domain", "FieldName": "FieldName"
                });
                buildDataSection("New tables and structures", "NewTablesStructures", dataDesc.Newtablesandstructures, {
                    "Table Structure Name": "TableStructureName", "Type": "Type", "Description": "Description",
                    "Initial Size In Rows": "InitialSizeInRows", "Maximum Size In Rows": "MaximumSizeInRows",
                    "Data Class": "DataClass", "Buffering": "Buffering", "Buffering Type": "BufferingType", "Log Data Changes": "LogDataChanges"
                });
                buildDataSection("Fields", "Fields", dataDesc.Fields, {
                    "DataElement": "DataElement", "Description": "Description", "Domain": "Domain", "FieldName": "FieldName", "KeyFields": "KeyFields"
                });
                buildDataSection("New indexes", "NewIndexes", dataDesc.Newindexes, {
                    "Buffering": "Buffering", "Buffering Type": "BufferingType", "Data Class": "DataClass",
                    "Index Name": "IndexName", "Index Type": "IndexType", "Initial Size In Rows": "InitialSizeInRows",
                    "Log Data Changes": "LogDataChanges", "Maximum Size In Rows": "MaximumSizeInRows",
                    "Short Description": "ShortDescription", "Table Name": "TableName"
                });

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "IndexFields", children: [new TextRun({ text: "Index fields", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var idxF = safeObj(dataDesc.Indexfields);
                documentChildren.push(makeColumnTable(D, ["Field Description", "Field Name"], [{ "Field Description": idxF["Field Description"] || "", "Field Name": idxF["Field Name"] || "" }]));

                buildDataSection("Foreign keys", "ForeignKeys", dataDesc.ForeignKey, {
                    "Cardinality Right Side": "Cardinality Right Side", "Cardinality Left Side": "CardinalityLeftSide",
                    "Check Table": "CheckTable", "Error Message": "ErrorMessage", "Foreign Key Field": "ForeignKeyField",
                    "Foreign Key Field Type": "ForeignKeyFieldType", "Foreign Key Table": "ForeignKeyTable",
                    "Screen Check Required": "ScreenCheckRequired", "Short Text": "ShortText"
                });

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "UpdateMethodTables", children: [new TextRun({ text: "Update method for tables", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 100 } }));
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: dataDesc["Update method for tables"] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
            });

            return documentChildren;
        },

        // createLastCommonSectionTSD  — accepts parsed object directly
        createLastCommonSectionTSD: function (documentChildren, response) {
            var D = window.docx;
            var { Paragraph, TextRun, Table, Bookmark } = D;

            var TECHNICALTESTING = safeObj(safeArr(response[0] && response[0]["TechnicalUnitTestDetails"])[0]);
            var response2 = safeArr(response[1] && response[1]["response2"]);
            var response3 = safeObj(response[1] && response[1]["response3"]);

            // Technical Unit Test Details
            safeSection(D, documentChildren, "Technical Unit Test Details", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TechnicalUnitTestDetails", children: [new TextRun({ text: "Technical unit test details", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));

                var simpleFields = [
                    ["Testing Approach", "Testing approach"],
                    ["Completion Criteria", "Completion criteria"],
                    ["Test data and other needs", "Test data and other needs"],
                    ["Testing dependencies", "Testing dependencies"],
                    ["Technical unit test design", "Technical unit test design"],
                    ["Test methods (for reference)", "Test methods (for reference)"]
                ];
                simpleFields.forEach(function (pair) {
                    documentChildren.push(new Paragraph({ children: [new Bookmark({ id: pair[0].replace(/\s/g, ""), children: [new TextRun({ text: pair[0], bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 100 } }));
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: TECHNICALTESTING[pair[1]] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
                });

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TestToolsEnv", children: [new TextRun({ text: "Test Tools and Environment", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var ttRows = [
                    makeKVRow(D, "Test System", ""),
                    makeKVRow(D, "Client", ""),
                    makeKVRow(D, "Tool", "( ) CATT\n( ) Workbench Debugger\n( ) Other"),
                    makeKVRow(D, "Test program(s)", "")
                ];
                documentChildren.push(new Table({ rows: ttRows }));

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TechnicalUnitTestExecutorsHeader", children: [new TextRun({ text: "Technical Unit Test Executors", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var execData = safeObj(TECHNICALTESTING["Technical unit test executors"]);
                var execCols = ["Email", "Program Under Test", "Programmer Name", "Test Date", "Tester Name (if different)"];
                documentChildren.push(makeColumnTable(D, execCols, [execData]));

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "TestCaseScenarios", children: [new TextRun({ text: "Technical unit test case scenarios", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var scenarios = safeArr(TECHNICALTESTING["Technical unit test case scenarios"]);
                if (!scenarios.length) scenarios = [{ "ID": "", "Test Objective": "", "Test Steps/Procedures": "", "Test Data /Condition": "", "Expected Results": "", "Pass": "", "Fail": "", "Priority": "", "Comments": "" }];
                documentChildren.push(makeColumnTable(D, ["ID", "Test Objective", "Test Steps/Procedures", "Test Data /Condition", "Expected Results", "Pass", "Fail", "Priority", "Comments"], scenarios));
            });

            // Additional Information
            safeSection(D, documentChildren, "Additional Information", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Additional Information", bold: true, size: 32, font: "Helvetica" })], spacing: { before: 300, after: 200 } }));

                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Object-Oriented Class Details", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 300, after: 200 } }));
                var ooDetails = safeArr(response3.Object_oriented_details);
                if (ooDetails.length) {
                    documentChildren.push(makeColumnTable(D, ["Method Name", "Description", "Method Type", "Method Details", "Import/Export/Change/Return Parameters"], ooDetails));
                } else {
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No Object Oriented Details available.", size: 20, font: "Helvetica" })] }));
                }

                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "FM DETAILS", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 300, after: 200 } }));
                var fmDetails = safeArr(response3.FM_details);
                if (fmDetails.length) {
                    documentChildren.push(makeColumnTable(D, ["Function Group", "Function Module Name", "Description"], fmDetails));
                } else {
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No FM Details available.", size: 20, font: "Helvetica" })] }));
                }

                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "BackgroundJob", children: [new TextRun({ text: "Background job", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var bgDetails = safeArr(response3.Background_job_details);
                if (bgDetails.length) {
                    var bgCols = ["Job Catalog", "Job Catalog Entry", "Job Template", "Description", "System", "Class", "Execute Method", "Global/Local", "Business Unit/Country", "Localization Reason", "Target/Intermediate State", "Related JIRA Issues", "Additional Information"];
                    var bgMap = { "Job Catalog": "JobCatalog", "Job Catalog Entry": "JobCatalogEntry", "Job Template": "JobTemplate", "Description": "Description", "System": "System", "Class": "Class", "Execute Method": "ExecuteMethod", "Global/Local": "GlobalLocal", "Business Unit/Country": "BusinessUnitCountry", "Localization Reason": "LocalizationReason", "Target/Intermediate State": "TargetIntermediateState", "Related JIRA Issues": "RelatedJIRAIssues", "Additional Information": "AdditionalInformation" };
                    documentChildren.push(makeColumnTable(D, bgCols, bgDetails, bgMap));
                } else {
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No Background Job Details available.", size: 20, font: "Helvetica" })] }));
                }
            });

            // Code Review
            safeSection(D, documentChildren, "Code Review", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "CodeReview", children: [new TextRun({ text: "Code review", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var codeReview = safeObj(safeArr(response3.Code_review)[0]);
                var crFields = [
                    ["ATC / Onapsis code review report", "ATC / Onapsis code review report"],
                    ["ESLint / Onapsis code review report", "ESLint / Onapsis code review report"],
                    ["Peer review code", "Peer review code"],
                    ["Checklist", "Checklist"]
                ];
                crFields.forEach(function (pair) {
                    documentChildren.push(new Paragraph({ children: [new Bookmark({ id: pair[0].replace(/[^a-zA-Z]/g, ""), children: [new TextRun({ text: pair[0], bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: codeReview[pair[1]] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
                });
            });

            // Assumptions
            safeSection(D, documentChildren, "Assumptions", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "Assumptions_1", children: [new TextRun({ text: "Assumptions", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 200, after: 100 } }));
                var assumptions = (response2[1] && response2[1]["Assumptions"])
                    ? (Array.isArray(response2[1]["Assumptions"]) ? response2[1]["Assumptions"].join("\n") : response2[1]["Assumptions"])
                    : "";
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: assumptions, size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
            });

            // Authorization Objects
            safeSection(D, documentChildren, "Authorization Objects", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "AuthorizationObjects", children: [new TextRun({ text: "Authorization Objects", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var authObjs = safeArr(response3.authorization_objects_details);
                if (authObjs.length) {
                    var authCols = ["Authorization Object Name", "Authorization Restrictions", "Fields", "Roles"];
                    var authMap = { "Authorization Object Name": "Authorization Object", "Authorization Restrictions": "Restriction", "Fields": "Fields", "Roles": "Roles" };
                    documentChildren.push(makeColumnTable(D, authCols, authObjs, authMap));
                } else {
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No Authorization Objects available.", size: 20, font: "Helvetica" })] }));
                }
            });

            // Dependencies
            safeSection(D, documentChildren, "Dependencies", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "Dependencies", children: [new TextRun({ text: "Dependencies", bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var deps = (response2[0] && response2[0].Dependencies) ? response2[0].Dependencies : "";
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: deps, size: 20, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
            });

            return documentChildren;
        },

        // createEnhancementTSD  — response is already parsed object
        createEnhancementTSD: function (documentChildren, response, response_cleancore) {
            var D = window.docx;
            var { Paragraph, TextRun, Table, Bookmark } = D;
            var that = this;

            // Navigate safely: response.response3.Enhancement[0]  OR  response.response3[0]
            // New response structure has Enhancement array under response3.Enhancement
            var reportsData = safeObj(
                (response && response.response3 && Array.isArray(response.response3.Enhancement) && response.response3.Enhancement[0])
                    ? response.response3.Enhancement[0]
                    : (response && response.response3 && Array.isArray(response.response3) && response.response3[0])
                        ? response.response3[0]
                        : (response && response.response3 ? response.response3 : {})
            );

            // NEW: bind ProcessFlow string (new response has ProcessFlow as a plain string, not object)
            if (typeof reportsData["ProcessFlow"] === "string") {
                var pfStr = reportsData["ProcessFlow"];
                reportsData["ProcessFlow"] = { ProcessFlowDescription: pfStr, Steps: [] };
            }

            // NEW: normalise PseudoCode — new response has PseudoCode as object with
            //   CreateScenario, ChangeScenario, Overview etc. — map to expected fields
            var pcRaw = safeObj(reportsData["PseudoCode"]);
            if (!pcRaw["Overview"] && (pcRaw["CreateScenario"] || pcRaw["ChangeScenario"])) {
                pcRaw["Overview"] = pcRaw["Overview"] || pcRaw["CreateScenario"] || "";
                pcRaw["ProcessingLogic"] = pcRaw["ProcessingLogic"] || pcRaw["ChangeScenario"] || "";
                pcRaw["InputParameters"] = pcRaw["InputParameters"] || pcRaw["InputParameters"] || "";
                reportsData["PseudoCode"] = pcRaw;
            }

            // NEW: normalise ExceptionHandling — new response has different field names
            var ehRaw = safeObj(reportsData["ExceptionHandling"]);
            if (!ehRaw["SystemNotifications"] && ehRaw["ErrorMessages"]) {
                ehRaw["SystemNotifications"] = ehRaw["SystemNotifications"] || ehRaw["ErrorMessages"] || "";
                reportsData["ExceptionHandling"] = ehRaw;
            }

            // NEW: normalise Key_User_Extension — new response uses different field casing
            var kueRaw = safeObj(reportsData["Key_User_Extension"]);
            if (kueRaw["Business_Context"] === undefined && kueRaw["Business Context"]) {
                kueRaw["Business_Context"] = kueRaw["Business Context"];
            }
            if (kueRaw["Custom_field_name"] === undefined && kueRaw["Custom field name"]) {
                kueRaw["Custom_field_name"] = kueRaw["Custom field name"];
            }

            documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "EnhancementBookmark", children: [new TextRun({ text: "Enhancement", bold: true, size: 32, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));

            safeSection(D, documentChildren, "Enhancement - Data Dictionary Objects", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "DataDictionaryBookmark", children: [new TextRun({ text: "Data Dictionary Objects", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var ddo = safeObj(reportsData["Data_Dictionary_Objects"]);
                var ddFields = {
                    "Application/Platform": ddo["Application"],
                    "Object Name": ddo["Object_Name"],
                    "Object Type": ddo["Object_Type"],
                    "Object Description": ddo["Object_Description"],
                    "Step By Step Instructions": ddo["Step by step Instruction"],
                    "Global/Local": ddo["Global/Local"],
                    "System (S4/EWM/TM/etc)": ddo["System"],
                    "Business Unit/Country": ddo["Business Unit/Country"],
                    "Localization Reason": ddo["Localization Reason"],
                    "Target/Intermediate State": ddo["Target/Intermediate State"],
                    "Related JIRA Issues": ddo["Related JIRA Issue"],
                    "Additional Information": ddo["Additional Information"]
                };
                documentChildren.push(new Table({ rows: Object.keys(ddFields).map(function (k) { return makeKVRow(D, k, ddFields[k]); }) }));
            });

            safeSection(D, documentChildren, "Enhancement Descriptions", function () {
                var ed = safeObj(reportsData["Enhancement_descriptions"]);
                [["CDS View", "CDS View"], ["GLOBAL", "GLOBAL"], ["All Business Units / Countries", "All Business UnitsCountries"], ["TARGET STATE", "TARGET STATE"]].forEach(function (pair) {
                    documentChildren.push(new Paragraph({ children: [new Bookmark({ id: pair[0].replace(/\s/g, ""), children: [new TextRun({ text: pair[0], bold: true, size: 24, font: "Helvetica" })] })], spacing: { before: 250, after: 100 } }));
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: ed[pair[1]] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
                });
            });

            safeSection(D, documentChildren, "Enhancement - Table Details", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "Table1Details", children: [new TextRun({ text: "Table 1 details:", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var td = safeObj(reportsData["Table_Details"]);
                documentChildren.push(makeColumnTable(D, ["Field Name", "Field Type", "Field Length", "Key Information", "Field Description", "Check Table"],
                    [{ "Field Name": td["Field_Name"] || "", "Field Type": td["Field_Type"] || "", "Field Length": td["Field_Length"] || "", "Key Information": td["Key_Information"] || "", "Field Description": td["Field_Description"] || "", "Check Table": td["Check_Table"] || "" }]
                ));
            });

            safeSection(D, documentChildren, "Enhancement - Structure Details", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "StructureDetails", children: [new TextRun({ text: "Structure details:", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var sd = safeObj(reportsData["Structure_Details"]);
                documentChildren.push(makeColumnTable(D, ["Field Name", "Field Type", "Field Length", "Field Description"],
                    [{ "Field Name": sd["Field_Name"] || "", "Field Type": sd["Field_Type"] || "", "Field Length": sd["Field_Length"] || "", "Field Description": sd["Field_Description"] || "" }]
                ));
            });

            safeSection(D, documentChildren, "Enhancement - Non-Data Dictionary Objects", function () {
                documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "NonDataDictionaryObjects", children: [new TextRun({ text: "Non-Data Dictionary Objects", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));
                var ndo = safeObj(reportsData["Non_Data_Dictionary_Objects"]);
                var ndoCols = ["Object Name", "Object Type", "Object Description", "Step By Step Instructions", "Global/Local", "System", "Business Unit/Country", "Localization Reason", "Target/Intermediate State", "Related JIRA Issues", "Additional Information"];
                var ndoMap = { "Object Name": "Object_Name", "Object Type": "Object_Type", "Object Description": "Object_Description", "Step By Step Instructions": "Step by step Instruction", "Global/Local": "Global/Local", "System": "System (S4/EWM/TM/etc)", "Business Unit/Country": "Business Unit/Country", "Localization Reason": "Localization Reason", "Target/Intermediate State": "Target/Intermediate State", "Related JIRA Issues": "Related JIRA Issue", "Additional Information": "Additional Information" };
                documentChildren.push(makeColumnTable(D, ndoCols, [ndo], ndoMap));
            });

            safeSection(D, documentChildren, "Enhancement - Key User Extension", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Custom Business Object", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Key User Extension", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var kue = safeObj(reportsData["Key_User_Extension"]);
                var kueCols = ["Business_Context", "Description", "Custom_field_name", "Data_element", "Data_type", "Length", "Global/Local", "System", "Business Unit/Country", "Localization Reason", "Target/Intermediate State", "Related JIRA Issue", "Step-by-step logic key user extension", "Additional Information"];
                documentChildren.push(makeColumnTable(D, kueCols, [kue]));
            });

            safeSection(D, documentChildren, "Enhancement - Control Framework", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Process in which BADI/exit/enhancement triggers", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));


function fallback(val) {
    return (val && String(val).trim()) ? val : "N/A";
}

// PROCESS FLOW
documentChildren.push(new Paragraph({
    children: [new TextRun({
        text: "Process Flow",
        bold: true,
        size: 28,
        font: "Helvetica"
    })],
    spacing: { before: 300, after: 150 }
}));

var pf = safeObj(reportsData["ProcessFlow"]);

// Description
documentChildren.push(new Paragraph({
    children: [new TextRun({
        text: fallback(pf.ProcessFlowDescription),
        size: 20
    })],
    spacing: { after: 100 }
}));

// Steps (no numbering)
(pf.Steps && pf.Steps.length ? pf.Steps : [{}]).forEach(function (step) {
    documentChildren.push(new Paragraph({
        children: [new TextRun({
            text: fallback(step.StepDescription),
            size: 20
        })]
    }));
});


// PSEUDO CODE
documentChildren.push(new Paragraph({
    children: [new TextRun({
        text: "Pseudo Code",
        bold: true,
        size: 24,
        font: "Helvetica"
    })],
    spacing: { before: 200, after: 100 }
}));

var pc = safeObj(reportsData["PseudoCode"]);

var pseudoText = [
    pc.Overview,
    pc.InputParameters,
    pc.ProcessingLogic,
    pc.Conditions,
    pc.LoopsAndCalculations,
    pc.OutputDescription,
    pc.ErrorHandlingLogic
].map(fallback).join("\n");

pseudoText.split("\n").forEach(function (line) {
    documentChildren.push(new Paragraph({
        children: [new TextRun({
            text: line,
            size: 20,
            font: "Helvetica"
        })]
    }));
});


// EXCEPTION HANDLING
documentChildren.push(new Paragraph({
    children: [new TextRun({
        text: "Exception Handling",
        bold: true,
        size: 24,
        font: "Helvetica"
    })],
    spacing: { before: 200, after: 100 }
}));

var eh = safeObj(reportsData["ExceptionHandling"]);

// System Notifications
documentChildren.push(new Paragraph({
    children: [new TextRun({
        text: fallback(eh.SystemNotifications),
        size: 20
    })],
    spacing: { after: 100 }
}));

// Exception Scenarios
(eh.ExceptionScenarios && eh.ExceptionScenarios.length ? eh.ExceptionScenarios : [{}]).forEach(function (ex) {

    documentChildren.push(new Paragraph({
        children: [new TextRun({
            text: fallback(ex.Description),
            size: 20
        })]
    }));

    documentChildren.push(new Paragraph({
        children: [new TextRun({
            text: fallback(ex.HandlingApproach),
            size: 20
        })]
    }));

});

                // 🔼 ADD TILL HERE
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Enhancement control framework", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ecf = safeObj(reportsData["Enhancement_Control_Framework"]);
                var ecfMap = { "Module": "Module", "Header Table": "Header_Table", "Item Table": "Item_Table", "Enhancement Spot": "Enhancement_spot", "BADI Definition": "BADI_definition", "BADI Implementation": "BADI_Implementation", "Interface": "Interface", "Implementing Class": "Implementation_Class", "Global/Local": "Global/Local", "System": "System", "Business Unit/Country": "Business Unit/Country", "Localization Reason": "Localization Reason", "Target/Intermediate State": "Target/Intermediate State", "Related JIRA Issues": "Related JIRA Issue", "Additional Information": "Additional Information" };
                documentChildren.push(makeColumnTable(D, Object.keys(ecfMap), [ecf], ecfMap));
            });

            safeSection(D, documentChildren, "Enhancement - BADI Implementation Details", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "BADI Implementation Details", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var bid = safeObj(reportsData["BADI_Implementation_Details"]);
                documentChildren.push(new Table({
                    rows: [
                        makeKVRow(D, "Name of Enhancement", bid["Name_of_Enhancement"]),
                        makeKVRow(D, "Enhancement Type", bid["Enhancement_Type"]),
                        makeKVRow(D, "Logical Database", bid["Logical_Database"]),
                        makeKVRow(D, "Called Transactions", bid["Called_Transactions"]),
                        makeKVRow(D, "Function Group", bid["Function_Group"]),
                        makeKVRow(D, "Function Calls", bid["Function_Calls"]),
                        makeKVRow(D, "Fixed Point Arithmetic", bid["Fixed_point_Arithmetic"]),
                        makeKVRow(D, "Attachment", bid["Attachment"]),
                        makeKVRow(D, "BADI Definition", bid["BADI_Definition"]),
                        makeKVRow(D, "Step-by-step logic", bid["Step-by-step logic BADI_Implementation_Details"])
                    ]
                }));
                documentChildren.push(makeColumnTable(D, ["Enhancement Spot", "BADI Definition", "Implementing Class"],
                    [{ "Enhancement Spot": bid["Enhancement_Spot"] || "", "BADI Definition": bid["BADI_Definition"] || "", "Implementing Class": bid["Implementing_Class"] || "" }]
                ));
            });

            safeSection(D, documentChildren, "Enhancement - Implementing Class Details", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Implementation class details:", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var icd = safeObj(reportsData["Implementing_Class_Details"]);
                documentChildren.push(makeColumnTable(D, ["Method Name", "Description", "Type", "Details"],
                    [{ "Method Name": icd["Method_Name"] || "", "Description": icd["Description"] || "", "Type": icd["Type"] || "", "Details": icd["Details"] || "" }]
                ));
            });

            safeSection(D, documentChildren, "Enhancement - Exit", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Exit", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var exitData = safeObj(reportsData["Exit (NA)"]);
                documentChildren.push(new Table({
                    rows: [
                        makeKVRow(D, "Enhancement Spot", exitData["Enhancement Spot"]),
                        makeKVRow(D, "Implicit Enhancement", exitData["Implicit Enhancement"]),
                        makeKVRow(D, "Explicit Enhancement", exitData["Explicit Enhancement"]),
                        makeKVRow(D, "User exit", exitData["User exit"]),
                        makeKVRow(D, "Customer exit", exitData["Customer exit"])
                    ]
                }));
            });

            if (!this.cleancorerecommendation) {
                safeSection(D, documentChildren, "CleanCore Recommendations", function () {
                    that._appendCleanCoreRecommendations(documentChildren, response_cleancore, D);
                    that.cleancorerecommendation = true;
                });
            }

            return documentChildren;
        },

        // CleanCore helper  — response_cleancore is the parsed Instructions object
        _appendCleanCoreRecommendations: function (documentChildren, response_cleancore, D) {
            var { Paragraph, TextRun, Table, TableRow, TableCell } = D;
            var recs = [];
            try {
                var cc = safeObj(response_cleancore);
                var r3 = cc.response3;

                // response3 can be an object OR an array of objects — check both shapes
                var r3Candidates = Array.isArray(r3) ? r3 : [r3];
                r3Candidates.forEach(function (r3Item) {
                    var r3Obj = safeObj(r3Item);
                    if (r3Obj.clean_core_recommendations) {
                        recs = recs.concat(safeArr(r3Obj.clean_core_recommendations));
                    }
                });

                if (!recs.length && cc.clean_core_recommendations) {
                    // Try top-level array in case structure differs
                    recs = recs.concat(safeArr(cc.clean_core_recommendations));
                }

                if (!recs.length) {
                    // Last resort: gather any sibling keys that look like recommendation
                    // categories (e.g. "Class Recommendations", "Data Model Recommendations")
                    // directly under response3, in case they aren't nested under
                    // clean_core_recommendations at all.
                    r3Candidates.forEach(function (r3Item) {
                        var r3Obj = safeObj(r3Item);
                        var found = {};
                        var hasAny = false;
                        Object.keys(r3Obj).forEach(function (key) {
                            if (/recommendation/i.test(key)) {
                                found[key] = r3Obj[key];
                                hasAny = true;
                            }
                        });
                        if (hasAny) { recs.push(found); }
                    });
                }
            } catch {
                recs = [];
            }

            if (!recs.length) {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No CleanCore Recommendations available.", size: 20, font: "Helvetica" })] }));
                return;
            }

            documentChildren.push(new Paragraph({ children: [new TextRun({ text: "CleanCore Recommendations", bold: true, size: 28, color: "0000FF", font: "Helvetica" })], spacing: { before: 250, after: 250 } }));

            var headerRow = new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Recommendation Type", bold: true, size: 24, font: "Helvetica" })] })], shading: { fill: "#e5e7eb" } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Recommendations", bold: true, size: 24, font: "Helvetica" })] })], shading: { fill: "#e5e7eb" } })
                ]
            });
            var dataRows = [];

            function pushRow(typeLabel, value) {
                var lines = Array.isArray(value) ? value : String(value || "").split("\n");
                lines = lines.filter(function (l) { return String(l || "").trim() !== ""; });
                if (!lines.length) { lines = [""]; }
                var paragraphs = lines.map(function (line) {
                    return new Paragraph({ children: [new TextRun({ text: String(line || ""), size: 20, font: "Helvetica" })], bullet: { level: 0 } });
                });
                dataRows.push(new TableRow({
                    children: [
                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: typeLabel, bold: true, size: 20, font: "Helvetica" })] })] }),
                        new TableCell({ children: paragraphs })
                    ]
                }));
            }

            recs.forEach(function (rec) {
                if (rec === null || rec === undefined || rec === "") {
                    return;
                }
                if (typeof rec === "string") {
                    // Plain-string entry, e.g. "Class Recommendations: Create a dedicated..."
                    var splitIdx = rec.indexOf(":");
                    var typeLabel = splitIdx > -1 ? rec.substring(0, splitIdx).trim() : "Recommendation";
                    var text = splitIdx > -1 ? rec.substring(splitIdx + 1).trim() : rec;
                    pushRow(typeLabel, text);
                    return;
                }
                var recObj = safeObj(rec);
                var keys = Object.keys(recObj);
                if (!keys.length) {
                    return;
                }
                keys.forEach(function (key) {
                    pushRow(key, recObj[key]);
                });
            });

            if (!dataRows.length) {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "No CleanCore Recommendations available.", size: 20, font: "Helvetica" })] }));
                return;
            }
            documentChildren.push(new Table({ rows: [headerRow].concat(dataRows) }));
        },

        // createFormTSD


createFormTSD: function (documentChildren, response, response_cleancore) {
    var D = window.docx;
    var { Paragraph, TextRun, Table, TableRow, TableCell, Bookmark,
          WidthType } = D;
    var that = this;

    // Root data
    var forms = safeObj(safeArr(response && response.response3 && response.response3.Forms)[0]);

    // NEW: normalise ProcessFlow — new response has it as a plain string
    if (typeof forms["ProcessFlow"] === "string") {
        var pfStr = forms["ProcessFlow"];
        forms["ProcessFlow"] = { ProcessFlowDescription: pfStr, Steps: [] };
    }

    // NEW: normalise LayoutGeneral.FontStyles — new response may be a string not array
    var lgRaw = safeObj(forms["LayoutGeneral"]);
    if (typeof lgRaw["FontStyles"] === "string") {
        // Parse "Font Style 1 is Arial 13 Bold. Font Style 2 is Arial 10 Bold..." into array
        var fsText = lgRaw["FontStyles"];
        var parsedStyles = [];
        var fsPattern = /Font Style (\d+) is ([^.]+)\./g;
        var fsMatch;
        while ((fsMatch = fsPattern.exec(fsText)) !== null) {
            parsedStyles.push({ StyleName: "Style " + fsMatch[1], FontDetails: fsMatch[2].trim() });
        }
        lgRaw["FontStyles"] = parsedStyles.length > 0 ? parsedStyles : [];
        forms["LayoutGeneral"] = lgRaw;
    }

    // NEW: normalise LayoutGeneral.SectionSequence — new response may be a string
    if (typeof lgRaw["SectionSequence"] === "string") {
        // Convert prose description to empty array (no structured data available)
        lgRaw["SectionSequence"] = [];
        forms["LayoutGeneral"] = lgRaw;
    }

    // NEW: normalise LayoutGeneral.FieldLabelsTable — new response has richer string
    // keep as-is (already a string, will render in KV table)

    // NEW: normalise CompletionConfirmationSlip — new response has strings not objects
    var ccsRaw = safeObj(forms["CompletionConfirmationSlip"]);
    if (typeof ccsRaw["OperationDetailSection"] === "string") {
        ccsRaw["OperationDetailSection"] = { Description: ccsRaw["OperationDetailSection"], Fields: [] };
    }
    if (typeof ccsRaw["NotesSection"] === "string") {
        ccsRaw["NotesSection"] = { Description: ccsRaw["NotesSection"], Columns: "", Depth: "" };
    }
    if (typeof ccsRaw["OrderFooter"] === "string") {
        ccsRaw["OrderFooter"] = { Description: ccsRaw["OrderFooter"], Fields: [] };
    }
    if (typeof ccsRaw["PageFooter"] === "string") {
        ccsRaw["PageFooter"] = { Description: ccsRaw["PageFooter"], Format: "" };
    }
    forms["CompletionConfirmationSlip"] = ccsRaw;

    // NEW: normalise PseudoCode — new response has richer ODataForm and StandardForm strings
    var pcRaw2 = safeObj(forms["PseudoCode"]);
    if (typeof pcRaw2["ODataForm"] === "string") {
        pcRaw2["ODataForm"] = { GatewayService: pcRaw2["ODataForm"], EntityNodes: [], Associations: [], Methods: [] };
    }
    if (typeof pcRaw2["StandardForm"] === "string") {
        pcRaw2["StandardForm"] = { FormName: safeObj(forms["AdobeForms"])["FormName"] || "", InterfaceName: safeObj(forms["AdobeForms"])["InterfaceName"] || "", CustomClass: safeObj(forms["AdobeForms"])["CustomClass"] || "", Methods: [] };
    }
    forms["PseudoCode"] = pcRaw2;

    // NEW: normalise Translations — new response has SupportedLanguages as string not array
    var transRaw = safeObj(forms["Translations"]);
    if (typeof transRaw["SupportedLanguages"] === "string") {
        transRaw["SupportedLanguages"] = [transRaw["SupportedLanguages"]];
        forms["Translations"] = transRaw;
    }

    // Sub-objects from API response
    var processFlow    = safeObj(forms["ProcessFlow"]);
    var pseudoCode     = safeObj(forms["PseudoCode"]);
    var pageHeader     = safeObj(forms["PageHeader"]);
    var orderHeader    = safeObj(forms["OrderHeader"]);
    var funcReqs       = safeObj(forms["Functional_Requirements"]);
    var triggers       = safeObj(forms["Triggers"]);
    var dataSelText    = forms["Data Selection, Validation & Processing"] || "";
    var expectedOutput = forms["Expected_Output"] || "";
    var ccs            = safeObj(forms["CompletionConfirmationSlip"]);
    var layoutGeneral  = safeObj(forms["LayoutGeneral"]);
    var adobeForms     = safeObj(forms["AdobeForms"]);
    var hardcopy       = safeObj(forms["Hardcopy, Stationary & Printer Requirements"]);
    var emailFunc      = safeObj(forms["Email_Functionality"]);
    var testDataText   = forms["Test Data, Pre-conditions"] || "";
    var testDataObj    = safeObj(forms["TestDataPreConditions"]);
    var translations   = safeObj(forms["Translations"]);
    var logo           = safeObj(forms["Logo"]);
    var sampleForms    = forms["Sample Forms"] || "";

    // Helper: section heading H2
    function pushH2(text, bookmarkId) {
        var children = bookmarkId
            ? [new Bookmark({ id: bookmarkId, children: [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })] })]
            : [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })];
        documentChildren.push(new Paragraph({ children: children, spacing: { before: 300, after: 150 } }));
    }

    function pushH3(text) {
        documentChildren.push(new Paragraph({
            children: [new TextRun({ text: text, bold: true, size: 24, font: "Helvetica" })],
            spacing: { before: 200, after: 100 }
        }));
    }

    function pushText(text) {
        var val = (text === null || text === undefined) ? "" : String(text);
        documentChildren.push(new Paragraph({
            children: [new TextRun({ text: val, size: 20, font: "Helvetica" })],
            spacing: { before: 80, after: 100 }
        }));
    }

    function pushSpacer() {
        documentChildren.push(new Paragraph({ children: [], spacing: { before: 100, after: 100 } }));
    }

    // KV table (label | value) — reuses module-level makeKVRow
    function makeLocalKVTable(pairs) {
        var rows = pairs.map(function (pair) {
            return makeKVRow(D, pair[0], pair[1] !== undefined && pair[1] !== null ? pair[1] : "");
        });
        return new Table({ rows: rows });
    }

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 1 : "Form" top heading
    // ══════════════════════════════════════════════════════════════════════
    documentChildren.push(new Paragraph({
        children: [new Bookmark({
            id: "Formsheader",
            children: [new TextRun({ text: "Form", bold: true, size: 32, font: "Helvetica" })]
        })],
        spacing: { before: 250, after: 200 }
    }));

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 2 : Process flow
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Process flow", function () {
        pushH2("Process flow", "FormProcessFlow");

        var pfDesc = processFlow["ProcessFlowDescription"] || "";
        if (pfDesc) { pushText(pfDesc); }

        var steps = safeArr(processFlow["Steps"]);
        if (steps.length) {
            // Numbered list using separate paragraph per step
            steps.forEach(function (step, idx) {
                var stepDesc = step["StepDescription"] || step["StepNumber"] + ". " || "";
                documentChildren.push(new Paragraph({
                    children: [new TextRun({
                        text: (idx + 1) + ". " + stepDesc,
                        size: 20, font: "Helvetica"
                    })],
                    spacing: { before: 60, after: 60 },
                    indent: { left: 360 }
                }));
            });
        } else {
            pushText("No process flow steps available.");
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 3 : Pseudo code
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Pseudo code", function () {
        pushH2("Pseudo code", "FormPseudoCode");

        // Overview
        var overview = pseudoCode["Overview"] || "";
        if (overview) { pushText(overview); }

        // OData Form section
        var odataForm = safeObj(pseudoCode["ODataForm"]);
        if (odataForm && odataForm["GatewayService"]) {
            pushH3("Odata Form:");
            pushText("Gateway Service: " + (odataForm["GatewayService"] || ""));

            var entityNodes = safeArr(odataForm["EntityNodes"]);
            if (entityNodes.length) {
                pushText("Entity Nodes:");
                entityNodes.forEach(function (node, idx) {
                    documentChildren.push(new Paragraph({
                        children: [new TextRun({ text: (idx + 1) + ". " + (typeof node === "string" ? node : JSON.stringify(node)), size: 20, font: "Helvetica" })],
                        indent: { left: 360 },
                        spacing: { before: 40, after: 40 }
                    }));
                });
            }

            var assocs = safeArr(odataForm["Associations"]);
            if (assocs.length) {
                pushText("Associations:");
                assocs.forEach(function (a, idx) {
                    documentChildren.push(new Paragraph({
                        children: [new TextRun({ text: (idx + 1) + ". " + (typeof a === "string" ? a : JSON.stringify(a)), size: 20, font: "Helvetica" })],
                        indent: { left: 360 },
                        spacing: { before: 40, after: 40 }
                    }));
                });
            }

            var methods = safeArr(odataForm["Methods"]);
            if (methods.length) {
                pushText("Methods:");
                methods.forEach(function (m, idx) {
                    documentChildren.push(new Paragraph({
                        children: [new TextRun({ text: (idx + 1) + ". " + (typeof m === "string" ? m : JSON.stringify(m)), size: 20, font: "Helvetica" })],
                        indent: { left: 360 },
                        spacing: { before: 40, after: 40 }
                    }));
                });
            }
        }

        // Standard Form section
        var standardForm = safeObj(pseudoCode["StandardForm"]);
        pushH3("Standard Form:");
        pushText("Form Name: " + (standardForm["FormName"] || adobeForms["FormName"] || ""));
        pushText("Interface Name: " + (standardForm["InterfaceName"] || adobeForms["InterfaceName"] || ""));
        pushText("Custom Class: " + (standardForm["CustomClass"] || adobeForms["CustomClass"] || ""));

        var sfMethods = safeArr(standardForm["Methods"]);
        if (sfMethods.length) {
            pushText("Methods:");
            sfMethods.forEach(function (m) {
                documentChildren.push(new Paragraph({
                    children: [new TextRun({ text: (m["MethodName"] || "") + " - " + (m["Description"] || ""), size: 20, font: "Helvetica" })],
                    indent: { left: 360 },
                    spacing: { before: 40, after: 40 }
                }));
            });
        }

        // Multi-Lingual Text Management
        var formLogic = pseudoCode["FormLogic"] || pseudoCode["MultiLingualTextManagement"] || "";
        if (formLogic) {
            pushH3("Multi-Lingual Text Management:");
            pushText(formLogic);
        }

        // Form Logic
        pushH3("Form Logic:");
        pushText(pseudoCode["FormLogic"] || "Retrieve user country and language from SAP system and form will be printed in user specific language.");

        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 4 : Page Header
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Page Header", function () {
        pushH2("Page Header:", "FormPageHeader");
        pushText(pageHeader["Description"] || pageHeader["Layout"] || "The Page Header section should be formatted as in the screenshot above.");

        var phFields = safeArr(pageHeader["Fields"]);
        if (phFields.length) {
            var phRows = phFields.map(function (f) {
                return new TableRow({
                    children: [
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: f["FieldName"] || "", bold: true, size: 20, font: "Helvetica" })] })],
                            shading: { fill: "#e5e7eb" },
                            width: { size: 3000, type: WidthType.DXA }
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: f["FontStyle"] || "", size: 20, font: "Helvetica" })] })],
                            width: { size: 2500, type: WidthType.DXA }
                        }),
                        new TableCell({
                            children: [new Paragraph({ children: [new TextRun({ text: f["Source"] || "", size: 20, font: "Helvetica" })] })],
                            width: { size: 3860, type: WidthType.DXA }
                        })
                    ]
                });
            });

            var headerRow = new TableRow({
                children: ["Field Name", "Font Style", "Source"].map(function (h, i) {
                    return new TableCell({
                        children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 20, font: "Helvetica" })] })],
                        shading: { fill: "#e5e7eb" },
                        width: { size: [3000, 2500, 3860][i], type: WidthType.DXA }
                    });
                })
            });
            documentChildren.push(new Table({ rows: [headerRow].concat(phRows) }));
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 5 : Order Header
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Order Header", function () {
        pushH2("Order Header:", "FormOrderHeader");

        // Description
        pushText(orderHeader["Description"] || "");

        // Key fields as KV rows
        var ohPairs = [
            ["Barcode",        orderHeader["Barcode"] || ""],
            ["Company Name",   orderHeader["CompanyName"] || ""],
            ["Order No. and Description", orderHeader["OrderNoAndDescription"] || ""],
            ["Status",         orderHeader["Status"] || ""],
            ["Details",        orderHeader["Details"] || ""],
            ["Scheduling Details", orderHeader["SchedulingDetails"] || ""],
            ["Location Details", orderHeader["LocationDetails"] || ""],
            ["Planning Details", orderHeader["PlanningDetails"] || ""],
            ["Reference Objects", orderHeader["ReferenceObjects"] || ""],
            ["Responsibilities", orderHeader["Responsibilities"] || ""]
        ];
        documentChildren.push(makeLocalKVTable(ohPairs));

        // Sections table (if present)
        var ohSections = safeArr(orderHeader["Sections"]);
        if (ohSections.length) {
            pushH3("Sections:");
            ohSections.forEach(function (section) {
                pushText("Section: " + (section["SectionName"] || ""));
                var sFields = safeArr(section["Fields"]);
                sFields.forEach(function (f) {
                    documentChildren.push(new Paragraph({
                        children: [new TextRun({
                            text: "  " + (f["FieldName"] || "") + " — " + (f["Description"] || "") + " [Source: " + (f["Source"] || "") + "] [Font: " + (f["FontStyle"] || "") + "]",
                            size: 20, font: "Helvetica"
                        })],
                        spacing: { before: 40, after: 40 },
                        indent: { left: 360 }
                    }));
                });
            });
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 6 : Functional Requirements
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Functional Requirements", function () {
        pushH2("Functional Requirements", "FormFunctionalRequirements");

        // Client PDF shows a wide table with columns:
        // Output type | Application | Transmission medium | Driver program | Form routine |
        // Form name | Global/Local | System | Business Unit/Country | Localization Reason |
        // Target/Intermediate State | Related JIRA Issues | Additional Information

        var frCols = [
            "Output type", "Application", "Transmission medium", "Driver program",
            "Form routine", "Form name", "Global/Local", "System (S4/EWM/TM/etc)",
            "Business Unit/Country", "Localization Reason", "Target/Intermediate State",
            "Related JIRA Issues", "Additional Information"
        ];
        var frRow = {
            "Output type":              "Print Output",
            "Application":              funcReqs["Application"] || "",
            "Transmission medium":      funcReqs["Transmission_Medium"] || "",
            "Driver program":           funcReqs["Driver_Program"] || "",
            "Form routine":             funcReqs["Form_Routine"] || "",
            "Form name":                funcReqs["Form_Name"] || adobeForms["FormName"] || "",
            "Global/Local":             funcReqs["Global/Local"] || "",
            "System (S4/EWM/TM/etc)":   funcReqs["System(S4/EWM/TM/etc)"] || "",
            "Business Unit/Country":    funcReqs["Business Unit/Country"] || "",
            "Localization Reason":      funcReqs["Localization Reason"] || "",
            "Target/Intermediate State": funcReqs["Target/Intermediate State"] || "",
            "Related JIRA Issues":      funcReqs["Related JIRA Isuue"] || "",
            "Additional Information":   funcReqs["Additional Information"] || ""
        };
        documentChildren.push(makeColumnTable(D, frCols, [frRow]));
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 7 : Triggers
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Triggers", function () {
        pushH2("Triggers", "FormTriggers");

        // Client PDF columns:
        // Output type | Output type description | Transmission medium | Partner function |
        // Language | Processing mode | Global/Local | System | Business Unit/Country |
        // Localization Reason | Target/Intermediate State | Related JIRA Issues | Additional Information

        var trigCols = [
            "Output type", "Output type description", "Transmission medium",
            "Partner function", "Language", "Processing mode",
            "Global/Local", "System (S4/EWM/TM/etc)",
            "Business Unit/Country", "Localization Reason",
            "Target/Intermediate State", "Related JIRA Issues", "Additional Information"
        ];
        var trigRow = {
            "Output type":              triggers["OutputType"] || "",
            "Output type description":  triggers["OutputType_Description"] || "",
            "Transmission medium":      triggers["Transmission_Medium"] || "",
            "Partner function":         triggers["Partner_Function"] || "",
            "Language":                 triggers["Language"] || "",
            "Processing mode":          triggers["Processing_Mode"] || "",
            "Global/Local":             triggers["Global/Local"] || "",
            "System (S4/EWM/TM/etc)":   triggers["System (S4/EWM/TM/etc)"] || "",
            "Business Unit/Country":    triggers["Business Unit/Country"] || "",
            "Localization Reason":      triggers["Localization Reason"] || "",
            "Target/Intermediate State": triggers["Target/Intermediate State"] || "",
            "Related JIRA Issues":      triggers["Related JIRA Isuue"] || "",
            "Additional Information":   triggers["Additional Information"] || ""
        };
        documentChildren.push(makeColumnTable(D, trigCols, [trigRow]));
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 8 : Data Selection, Validation & Processing
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Data Selection", function () {
        pushH2("Data Selection, Validation & Processing", "FormDataSelection");
        pushText(dataSelText);
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 9 : Expected Output, Layout
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Expected Output", function () {
        pushH2("Expected Output, Layout", "FormExpectedOutput");
        pushText(expectedOutput);
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 10 : Completion Confirmation Slip
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Completion Confirmation Slip", function () {
        pushH2("Completion Confirmation Slip", "FormCCS");
        pushText(ccs["Description"] || "");

        // Operation Detail Section
        var opDetail = safeObj(ccs["OperationDetailSection"]);
        if (opDetail["Description"]) {
            pushH3("Operation Detail Section");
            pushText(opDetail["Description"]);
            var opFields = safeArr(opDetail["Fields"]);
            if (opFields.length) {
                documentChildren.push(makeColumnTable(D,
                    ["Field Name", "Description", "Font Style", "Source"],
                    opFields.map(function (f) {
                        return {
                            "Field Name": f["FieldName"] || "",
                            "Description": f["Description"] || "",
                            "Font Style": f["FontStyle"] || "",
                            "Source": f["Source"] || ""
                        };
                    })
                ));
            }
        }

        // Notes Section
        var notesSection = safeObj(ccs["NotesSection"]);
        if (notesSection["Description"]) {
            pushH3("Notes Section");
            documentChildren.push(makeLocalKVTable([
                ["Description", notesSection["Description"] || ""],
                ["Columns",     notesSection["Columns"] || ""],
                ["Depth",       notesSection["Depth"] || ""]
            ]));
        }

        // Order Footer
        var orderFooter = safeObj(ccs["OrderFooter"]);
        if (orderFooter["Description"]) {
            pushH3("Order Footer");
            pushText(orderFooter["Description"]);
            var ofFields = safeArr(orderFooter["Fields"]);
            if (ofFields.length) {
                documentChildren.push(makeColumnTable(D,
                    ["Field Name", "Font Style"],
                    ofFields.map(function (f) {
                        return {
                            "Field Name": f["FieldName"] || "",
                            "Font Style": f["FontStyle"] || ""
                        };
                    })
                ));
            }
        }

        // Page Footer
        var pageFooter = safeObj(ccs["PageFooter"]);
        if (pageFooter["Description"] || pageFooter["Format"]) {
            pushH3("Page Footer");
            pushText(pageFooter["Description"] || "");
            pushText(pageFooter["Format"] || "");
        }

        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 11 : Layout General
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Layout General", function () {
        pushH2("Layout General", "FormLayoutGeneral");

        // General properties
        documentChildren.push(makeLocalKVTable([
            ["Orientation",           layoutGeneral["Orientation"] || ""],
            ["Paper Size",            layoutGeneral["PaperSize"] || ""],
            ["Font",                  layoutGeneral["Font"] || ""],
            ["Field Labels Table",    layoutGeneral["FieldLabelsTable"] || ""],
            ["Page Numbering Format", layoutGeneral["PageNumberingFormat"] || ""]
        ]));

        // Font Styles table
        var fontStyles = safeArr(layoutGeneral["FontStyles"]);
        if (fontStyles.length) {
            pushH3("Font Styles");
            documentChildren.push(makeColumnTable(D,
                ["Style Name", "Font Details"],
                fontStyles.map(function (fs) {
                    return {
                        "Style Name":   fs["StyleName"] || "",
                        "Font Details": fs["FontDetails"] || ""
                    };
                })
            ));
        }

        // Section Sequence table — matches client PDF table exactly
        var sections = safeArr(layoutGeneral["SectionSequence"]);
        if (sections.length) {
            pushH3("Section Sequence");
            documentChildren.push(makeColumnTable(D,
                ["Section", "Description", "First Page", "Intermediate Pages", "Last Page", "Split Across Pages"],
                sections.map(function (s) {
                    return {
                        "Section":            s["Section"] || "",
                        "Description":        s["Description"] || "",
                        "First Page":         s["FirstPage"] || "",
                        "Intermediate Pages": s["IntermediatePages"] || "",
                        "Last Page":          s["LastPage"] || "",
                        "Split Across Pages": s["SplitAcrossPages"] || ""
                    };
                })
            ));
        }

        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 12 : Data Dictionary Objects
    //   Client PDF: wide table with columns matching BespokeDevelopmentDetails style
    //   Data comes from forms object — AdobeForms + other DDO fields
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Data Dictionary Objects", function () {
        pushH2("Data Dictionary Objects", "FormDDO");

        var ddoCols = [
            "Application/Platform", "Object Name", "Object Type", "Object Description",
            "Step By Step Instructions", "Global/Local", "System (S4/EWM/TM/etc)",
            "Business Unit/Country", "Localization Reason", "Target/Intermediate State",
            "Related JIRA Issues", "Additional Information"
        ];

        // Build DDO rows from AdobeForms and any other top-level form DDO data
        // The client PDF shows structures like ZAMM_S_SHOP_PAPERS, CDS Views, etc.
        // API response provides AdobeForms at minimum; render it as one row.
        var ddoRows = [];

        // AdobeForms entry
        if (adobeForms["FormName"] || adobeForms["FormType"]) {
            ddoRows.push({
                "Application/Platform":      funcReqs["Application"] || "",
                "Object Name":               adobeForms["FormName"] || "",
                "Object Type":               adobeForms["FormType"] || "Adobe Form",
                "Object Description":        adobeForms["Description"] || "",
                "Step By Step Instructions": "",
                "Global/Local":              funcReqs["Global/Local"] || "",
                "System (S4/EWM/TM/etc)":    funcReqs["System(S4/EWM/TM/etc)"] || "",
                "Business Unit/Country":     funcReqs["Business Unit/Country"] || "",
                "Localization Reason":        funcReqs["Localization Reason"] || "",
                "Target/Intermediate State":  funcReqs["Target/Intermediate State"] || "",
                "Related JIRA Issues":        funcReqs["Related JIRA Isuue"] || "",
                "Additional Information":     adobeForms["LinkedGatewayService"] ? "Linked Gateway Service: " + adobeForms["LinkedGatewayService"] : ""
            });
        }

        // Interface
        if (adobeForms["InterfaceName"]) {
            ddoRows.push({
                "Application/Platform":      funcReqs["Application"] || "",
                "Object Name":               adobeForms["InterfaceName"] || "",
                "Object Type":               "Interface",
                "Object Description":        "Interface for " + (adobeForms["FormName"] || "Adobe Form"),
                "Step By Step Instructions": "",
                "Global/Local":              funcReqs["Global/Local"] || "",
                "System (S4/EWM/TM/etc)":    funcReqs["System(S4/EWM/TM/etc)"] || "",
                "Business Unit/Country":     "",
                "Localization Reason":        "",
                "Target/Intermediate State":  "",
                "Related JIRA Issues":        "",
                "Additional Information":     ""
            });
        }

        // Custom Class
        if (adobeForms["CustomClass"]) {
            ddoRows.push({
                "Application/Platform":      funcReqs["Application"] || "",
                "Object Name":               adobeForms["CustomClass"] || "",
                "Object Type":               "Class",
                "Object Description":        "Custom class for " + (adobeForms["FormName"] || "Adobe Form"),
                "Step By Step Instructions": "",
                "Global/Local":              funcReqs["Global/Local"] || "",
                "System (S4/EWM/TM/etc)":    funcReqs["System(S4/EWM/TM/etc)"] || "",
                "Business Unit/Country":     "",
                "Localization Reason":        "",
                "Target/Intermediate State":  "",
                "Related JIRA Issues":        "",
                "Additional Information":     ""
            });
        }

        if (ddoRows.length === 0) {
            // Provide blank row matching client PDF style
            var blankRow = {};
            ddoCols.forEach(function (c) { blankRow[c] = ""; });
            ddoRows.push(blankRow);
        }

        documentChildren.push(makeColumnTable(D, ddoCols, ddoRows));

        // Sample Form reference (client PDF shows this after DDO table)
        if (sampleForms) {
            pushH3("Sample Form");
            pushText(sampleForms);
        }

        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 13 : Hardcopy, Stationary & Printer Requirements
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Hardcopy", function () {
        pushH2("Hardcopy, Stationary & Printer Requirements", "FormHardcopy");

        // Client PDF: wide multi-column table
        var hcCols = [
            "Printer type", "Printer name", "Page size", "Page layout", "Barcode",
            "Global/Local", "System (S4/EWM/TM/etc)", "Business Unit/Country",
            "Localization Reason", "Target/Intermediate State",
            "Related JIRA Issues", "Additional Information"
        ];
        var hcRow = {
            "Printer type":              hardcopy["PrinterType"] || "NA",
            "Printer name":              hardcopy["Printer_Name"] || "NA",
            "Page size":                 hardcopy["Page_size"] || "",
            "Page layout":               hardcopy["Page_Layout"] || "",
            "Barcode":                   hardcopy["Barcode"] || "",
            "Global/Local":              hardcopy["Global/Local"] || "",
            "System (S4/EWM/TM/etc)":    hardcopy["System (S4/EWM/TM/etc)"] || "",
            "Business Unit/Country":     hardcopy["Business Unit/Country"] || "",
            "Localization Reason":        hardcopy["Localization Reason"] || "",
            "Target/Intermediate State":  hardcopy["Target/Intermediate State"] || "",
            "Related JIRA Issues":        hardcopy["Related JIRA Isuue"] || "",
            "Additional Information":     hardcopy["Additional Information"] || ""
        };
        documentChildren.push(makeColumnTable(D, hcCols, [hcRow]));
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 14 : Email Functionality
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Email Functionality", function () {
        pushH2("Email Functionality", "FormEmail");

        // Client PDF: wide multi-column table
        var efCols = [
            "Recipient", "Recipient type", "Subject", "Body", "Signature", "Attachment",
            "Global/Local", "System (S4/EWM/TM/etc)", "Business Unit/Country",
            "Localization Reason", "Target/Intermediate State",
            "Related JIRA Issues", "Additional Information"
        ];
        var efRow = {
            "Recipient":                 emailFunc["Recipient"] || "",
            "Recipient type":            emailFunc["Recipient_type"] || "",
            "Subject":                   emailFunc["Subject"] || "",
            "Body":                      emailFunc["Body"] || "",
            "Signature":                 emailFunc["Signature"] || "",
            "Attachment":                emailFunc["Attachment"] || "",
            "Global/Local":              emailFunc["Global/Local"] || "",
            "System (S4/EWM/TM/etc)":    emailFunc["System"] || "",
            "Business Unit/Country":     emailFunc["Business Unit/Country"] || "",
            "Localization Reason":        emailFunc["Localization Reason"] || "",
            "Target/Intermediate State":  emailFunc["Target/Intermediate State"] || "",
            "Related JIRA Issues":        emailFunc["Related JIRA Isuue"] || "",
            "Additional Information":     emailFunc["Additional Information"] || ""
        };
        documentChildren.push(makeColumnTable(D, efCols, [efRow]));
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 15 : Test Data, Pre-conditions
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Test Data", function () {
        pushH2("Test Data, Pre-conditions", "FormTestData");

        // Free text from API
        if (testDataText) { pushText(testDataText); }

        // Structured data from TestDataPreConditions
        if (testDataObj && Object.keys(testDataObj).length) {
            documentChildren.push(makeLocalKVTable([
                ["Test System",   testDataObj["TestSystem"] || ""],
                ["Client",        testDataObj["Client"] || ""],
                ["Pre-Conditions", testDataObj["PreConditions"] || ""],
                ["Test Data",     testDataObj["TestData"] || ""]
            ]));
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 16 : Adobe Forms
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Adobe Forms", function () {
        pushH2("Adobe Forms", "FormAdobeForms");

        documentChildren.push(makeLocalKVTable([
            ["Form Name",              adobeForms["FormName"] || ""],
            ["Form Type",              adobeForms["FormType"] || ""],
            ["Interface Name",         adobeForms["InterfaceName"] || ""],
            ["Custom Class",           adobeForms["CustomClass"] || ""],
            ["Description",            adobeForms["Description"] || ""],
            ["Linked Gateway Service", adobeForms["LinkedGatewayService"] || ""]
        ]));
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 17 : Translations
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Translations", function () {
        pushH2("Translations", "FormTranslations");

        if (translations && Object.keys(translations).length) {
            var supportedLangs = safeArr(translations["SupportedLanguages"]);
            documentChildren.push(makeLocalKVTable([
                ["Translation Approach", translations["TranslationApproach"] || ""],
                ["Supported Languages",  supportedLangs.join(", ") || ""],
                ["Multilingual Details", translations["MultiLingualDetails"] || ""]
            ]));
        } else {
            pushText("NA");
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 18 : Logo
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Form - Logo", function () {
        pushH2("Logo", "FormLogo");

        if (logo && Object.keys(logo).length) {
            documentChildren.push(makeLocalKVTable([
                ["Logo Details",      logo["LogoDetails"] || ""],
                ["Logo Source",       logo["LogoSource"] || ""],
                ["Placement In Form", logo["PlacementInForm"] || ""]
            ]));
        } else {
            pushText("NA");
        }
        pushSpacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // SECTION 19 : CleanCore Recommendations (existing — unchanged)
    // ══════════════════════════════════════════════════════════════════════
    if (!this.cleancorerecommendation) {
        safeSection(D, documentChildren, "CleanCore Recommendations", function () {
            that._appendCleanCoreRecommendations(documentChildren, response_cleancore, D);
            that.cleancorerecommendation = true;
        });
    }

    return documentChildren;
},

        

createFioriTSD: function (documentChildren, response, response_cleancore) {
    var D = window.docx;
    var { Paragraph, TextRun, Table,
          Bookmark } = D;
    var that = this;

    // ROOT
    var fiori = safeObj(
        (response && response.response3 && Array.isArray(response.response3.Fiori_App) && response.response3.Fiori_App[0])
            ? response.response3.Fiori_App[0]
            : (response && response.response3 && Array.isArray(response.response3["Fiori App"]) && response.response3["Fiori App"][0])
                ? response.response3["Fiori App"][0]
                : {}
    );

    // DEBUG — remove after confirming data flows

    // fatd
    var fatd = safeObj(fiori["Fiori App Technical Details"]);

    // sub-objects — all paths verified
    var appNameType  = safeObj(fatd["Fiori App Name & Type"]);
    var selScreen    = safeArr(fatd["Selection Screen"] || []);
    var dataSel      = safeObj(fatd["Data Selection"] || {});
    var processFlow  = safeObj(fatd["ProcessFlow"] || {});
    var pseudoCode   = safeObj(fatd["Pseudo Code"] || {});
    // "ExceptionHandling" (camelCase) — the real key in new response
    var exHandling   = safeObj(fatd["ExceptionHandling"] || fatd["Exception Handling"] || {});
    var exAddInfo    = safeObj(fatd["Exception Handling Additional Information"] || {});
    var actionLogic  = safeObj(fatd["Action Logic"] || {});
    var validation   = fatd["Validation"] || "";
    var mockupScreen = fatd["Mock-up Screen"] || "";

    // Security — MERGE top-level (full) + fatd inner (L1-L5/App Name/Job Role)
    var secInfo = safeObj(fiori["Security Information"] || {});
    var fatdSec = safeObj(fatd["Security Information"] || {});
    // fatd inner has App Name, Job Role, L1-L5 — merge if missing at top level
    ["App Name","Job Role","L1","L2","L3","L4","L5"].forEach(function(k){
        if (!secInfo[k] && fatdSec[k]) secInfo[k] = fatdSec[k];
    });
    var secAddInfo = safeArr(secInfo["Additional Information"] || []);

    // OData
    var odata          = safeObj(fiori["OData"] || {});
    var entityDiagram  = odata["Entity Diagram"] || "";
    var serviceDetails = safeArr(odata["Service Details"] || []);
    var svcObjTable    = safeArr(odata["Service Details Object Table"] || []);
    var entity1        = safeObj(odata["Entity 1"] || {});

    // UI5
    var ui5           = safeObj(fiori["UI5"] || {});
    var ui5Details    = safeArr(ui5["UI5 App Details"] || []);
    var ui5ScreenFlow = safeArr(ui5["UI5 App - Screen Flow"] || []);

    // AI cols — reused for Additional Information tables
    var aiCols = [
        "Application/Platform","Object Name","Object Type","Object Description",
        "Step By Step Instructions","System (S4/EWM/TM/etc)","Global/Local",
        "Business Unit/Country","Localization Reason","Target/Intermediate State",
        "Related JIRA Issues","Additional Information"
    ];
    function mapAIRow(ai) {
        var r = {};
        r["Application/Platform"]      = ai["Application / Platform"]       || ai["Application/Platform"]      || "";
        r["Object Name"]               = ai["Object Name"]                  || "";
        r["Object Type"]               = ai["Object Type"]                  || "";
        r["Object Description"]        = ai["Object Description"]           || "";
        r["Step By Step Instructions"] = ai["Step By Step Instructions"]    || "";
        r["System (S4/EWM/TM/etc)"]    = ai["System (S4 / EWM / TM / etc)"]|| ai["System (S4/EWM/TM/etc)"]    || "";
        r["Global/Local"]              = ai["Global / Local"]               || ai["Global/Local"]              || "";
        r["Business Unit/Country"]     = ai["Business Unit / Country"]      || ai["Business Unit/Country"]     || "";
        r["Localization Reason"]       = ai["Localization Reason"]          || "";
        r["Target/Intermediate State"] = ai["Target / Intermediate State"]  || ai["Target/Intermediate State"] || "";
        r["Related JIRA Issues"]       = ai["Related JIRA Issues"]          || "";
        r["Additional Information"]    = ai["Additional Information"]        || "";
        return r;
    }

    // LOCAL HELPERS
    function ph2(text, bmId) {
        var ch = bmId
            ? [new Bookmark({ id: bmId, children: [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })] })]
            : [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })];
        documentChildren.push(new Paragraph({ children: ch, spacing: { before: 250, after: 150 } }));
    }
    function ph3(text) {
        documentChildren.push(new Paragraph({
            children: [new TextRun({ text: text, bold: true, size: 24, font: "Helvetica" })],
            spacing: { before: 200, after: 100 }
        }));
    }
    function pt(text) {
        String(text === null || text === undefined ? "" : text).split("\n").forEach(function(line){
            documentChildren.push(new Paragraph({
                children: [new TextRun({ text: line, size: 20, font: "Helvetica" })],
                spacing: { before: 60, after: 80 }
            }));
        });
    }
    function spacer() {
        documentChildren.push(new Paragraph({ children: [], spacing: { before: 80, after: 80 } }));
    }
    function kvTable(pairs) {
        return new Table({ rows: pairs.map(function(p){ return makeKVRow(D, p[0], p[1] != null ? p[1] : ""); }) });
    }

    // ══════════════════════════════════════════════════════════════════════
    // TOP HEADING
    // ══════════════════════════════════════════════════════════════════════
    documentChildren.push(new Paragraph({
        children: [new Bookmark({ id: "FioriApp", children: [new TextRun({ text: "Fiori App", bold: true, size: 32, font: "Helvetica" })] })],
        spacing: { before: 250, after: 250 }
    }));

    // ══════════════════════════════════════════════════════════════════════
    // 1. Technical details → Fiori App Name & Type
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - Technical Details", function () {
        ph2("Technical details", "FioriTechnicalDetails");
        ph3("Fiori app Name & Type");
        documentChildren.push(makeColumnTable(D,
            ["Fiori app", "App Title"],
            [{ "Fiori app": appNameType["Fiori App"] || "", "App Title": appNameType["App Title"] || "" }]
        ));
        spacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // 2. Selection screen
    //    KEY NOTE: actual key is "Default Values (From–To)" with U+2013 en-dash
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - Selection Screen", function () {
        ph2("Selection screen", "FioriSelectionScreen");

        var EN_DASH = "\u2013"; // –
        var colKey  = "Default Values (From" + EN_DASH + "To)";
        var ssCols  = [
            "Select Options / Parameters / Radio Buttons / Check Boxes",
            "Field Name",
            colKey,
            "Validation / Required / Optional"
        ];

        var ssRows = selScreen.length > 0
            ? selScreen.map(function(row){
                var r = {};
                r["Select Options / Parameters / Radio Buttons / Check Boxes"] =
                    row["Select Options / Parameters / Radio Buttons / Check Boxes"] || "";
                r["Field Name"] = row["Field Name"] || "";
                // Try en-dash key first (confirmed from new API), then fallbacks
                r[colKey] = row[colKey] ||
                             row["Default Values (From-To)"] ||
                             row["Default Values (From\u2014To)"] || "";
                r["Validation / Required / Optional"] =
                    row["Validation / Required / Optional"] || "";
                return r;
            })
            : [{ "Select Options / Parameters / Radio Buttons / Check Boxes": "",
                 "Field Name": "", [colKey]: "",
                 "Validation / Required / Optional": "" }];

        documentChildren.push(makeColumnTable(D, ssCols, ssRows));
        spacer();
    });

    // ══════════════════════════════════════════════════════════════════════
    // 3. Data Selection  (Special Processing / Variants / Screens / Program Flow)
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - Data Selection", function () {
        ph2("Data Selection", "FioriDataSelection");

        ph3("Special Processing");
        pt(dataSel["Special Processing"] || "");
        spacer();

        ph3("Variants, Variant Data");
        pt(dataSel["Variants, Variant Data"] || "");
        spacer();

        ph3("Screens, Screen Flow");
        pt(dataSel["Screens, Screen Flow"] || "");
        spacer();

        // Program Flow & Pseudo-code
        ph2("Program Flow & Pseudo-code", "FioriProgramFlow");

        // Process flow — from fatd["ProcessFlow"]
        ph3("Process flow");
        var pfDesc = processFlow["ProcessFlowDescription"] || "";
        if (pfDesc) pt(pfDesc);

        var pfSteps = safeArr(processFlow["Steps"] || []);
        if (pfSteps.length > 0) {
            pfSteps.forEach(function(step){
                var line = (step["StepNumber"] || "") + ". " + (step["StepDescription"] || "");
                if (step["ExpectedOutcome"]) line += " [Expected: " + step["ExpectedOutcome"] + "]";
                documentChildren.push(new Paragraph({
                    children: [new TextRun({ text: line, size: 20, font: "Helvetica" })],
                    spacing: { before: 60, after: 60 }, indent: { left: 360 }
                }));
            });
        } else {
            // Fallback to dataSel["Program Flow & Pseudo-code"]
            var pfAlt = dataSel["Program Flow & Pseudo-code"] || dataSel["Process Flow"] || "";
            if (pfAlt) pt(pfAlt); else pt("N/A");
        }
        spacer();

        // Pseudo code — from fatd["Pseudo Code"]
        ph3("Pseudo code");
        var pcFields = [
            ["Overview",             pseudoCode["Overview"]],
            ["Input Parameters",     pseudoCode["InputParameters"]],
            ["Processing Logic",     pseudoCode["ProcessingLogic"]],
            ["Conditions",           pseudoCode["Conditions"]],
            ["Loops & Calculations", pseudoCode["LoopsAndCalculations"]],
            ["Output Description",   pseudoCode["OutputDescription"]],
            ["Error Handling Logic", pseudoCode["ErrorHandlingLogic"]],
            ["Update Logic",         pseudoCode["UpdateLogic"]]
        ];
        pcFields.forEach(function(pair){
            if (pair[1] && String(pair[1]).trim() && String(pair[1]).trim() !== "N/A") {
                documentChildren.push(new Paragraph({
                    children: [new TextRun({ text: pair[0] + ":", bold: true, size: 20, font: "Helvetica" })],
                    spacing: { before: 120, after: 40 }
                }));
                pt(pair[1]);
            }
        });

        // Tables & Fields
        var pcTables = safeArr(pseudoCode["TablesAndFields"] || []);
        if (pcTables.length > 0) {
            documentChildren.push(new Paragraph({
                children: [new TextRun({ text: "Tables & Fields:", bold: true, size: 20, font: "Helvetica" })],
                spacing: { before: 120, after: 60 }
            }));
            documentChildren.push(makeColumnTable(D, ["Table Name","Field Name","Description","Usage"],
                pcTables.map(function(t){ return {
                    "Table Name": t["TableName"]||"", "Field Name": t["FieldName"]||"",
                    "Description": t["Description"]||"", "Usage": t["Usage"]||""
                }; })
            ));
        }
        spacer();

        // Validation
        if (validation && validation !== "N/A") {
            ph3("Validation:");
            pt(validation);
            spacer();
        }

        // Mock-up Screen
        if (mockupScreen && mockupScreen !== "N/A") {
            ph3("Mock-up Screen:");
            pt(mockupScreen);
            spacer();
        }
    });

    // ══════════════════════════════════════════════════════════════════════
    // 4. Security Information
    //    Source: fiori["Security Information"] (merged with fatd["Security Information"])
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - Security Information", function () {
        ph2("Security Information", "FioriSecurityInformation");

        // KV table — matches client PDF row order exactly
        documentChildren.push(kvTable([
            ["L1",               secInfo["L1"] || ""],
            ["L2",               secInfo["L2"] || ""],
            ["L3",               secInfo["L3"] || ""],
            ["L4",               secInfo["L4"] || ""],
            ["L5",               secInfo["L5"] || ""],
            ["Job Role",         secInfo["Job Role"] || secInfo["Jobposition"] || ""],
            ["Service Name",     secInfo["Service Name"] || ""],
            ["App Name",         secInfo["App Name"] || ""],
            ["Business Catalog", secInfo["Business Catalog"] || ""],
            ["Template Role",    secInfo["Template Role"] || ""]
        ]));
        spacer();

        if (secInfo["Narrative"]) {
            pt(secInfo["Narrative"]);
            spacer();
        }

        // Exception Handling — fatd["ExceptionHandling"]
        ph3("Exception Handling");
        // SystemNotifications is the main text; ErrorMessages is secondary
        var exText = exHandling["SystemNotifications"] ||
                     exHandling["ErrorMessages"]       ||
                     fatd["Exception Handling"]        ||
                     fiori["Exception Handling"]       || "";
        pt(exText || "N/A");
        spacer();

        // Exception Scenarios table
        var exScenarios = safeArr(exHandling["ExceptionScenarios"] || []);
        if (exScenarios.length > 0) {
            documentChildren.push(makeColumnTable(D,
                ["Exception Type","Description","Handling Approach","Severity"],
                exScenarios.map(function(sc){ return {
                    "Exception Type":    sc["ExceptionType"]    || "",
                    "Description":       sc["Description"]      || "",
                    "Handling Approach": sc["HandlingApproach"] || "",
                    "Severity":          sc["Severity"]         || ""
                }; })
            ));
            spacer();
        }

        // Additional Information (from fiori["Security Information"]["Additional Information"][])
        ph3("Additional Information");
        var aiRows = secAddInfo.length > 0
            ? secAddInfo.map(mapAIRow)
            : [(function(){ var r={}; aiCols.forEach(function(c){ r[c]=""; }); return r; }())];
        documentChildren.push(makeColumnTable(D, aiCols, aiRows));
        spacer();

        // Exception Handling Additional Information (fatd["Exception Handling Additional Information"])
        // Keys: "Application / Platform", "Object Name", etc. — same aiCols pattern
        if (exAddInfo && Object.keys(exAddInfo).length > 0) {
            var ehRow = mapAIRow(exAddInfo);
            var hasData = Object.values(ehRow).some(function(v){ return v && v !== "N/A"; });
            if (hasData) {
                ph3("Exception Handling — Additional Information");
                documentChildren.push(makeColumnTable(D, aiCols, [ehRow]));
                spacer();
            }
        }
    });

    // ══════════════════════════════════════════════════════════════════════
    // 5. OData
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - OData", function () {
        ph2("OData", "FioriOData");

        ph3("Entity Diagram");
        pt(entityDiagram || "N/A");
        spacer();

        // Service Details — odata["Service Details"][0]
        ph3("Service Details");
        var sd = serviceDetails.length > 0 ? safeObj(serviceDetails[0]) : {};
        documentChildren.push(kvTable([
            ["Package Name",       sd["Package Name"]       || ""],
            ["Service Definition", sd["Service Definition"] || ""],
            ["Service Binding",    sd["Service Binding"]    || ""],
            ["Service Name",       sd["Service Name"]       || ""]
        ]));
        spacer();

        // Service Details Object Table — wide column table
        if (svcObjTable.length > 0) {
            var sobjCols = [
                "Object Type","Relation","CDS View","SQL View","Interface",
                "Consumption View","Behavior Interface","Behavior Definition",
                "Metadata Extension","Access Control","Class","System (S4/EWM/TM/etc)"
            ];
            documentChildren.push(makeColumnTable(D, sobjCols,
                svcObjTable.map(function(obj){ return {
                    "Object Type":         obj["Object Type"]         || obj["Object"] || "",
                    "Relation":            obj["Relation"]            || "",
                    "CDS View":            obj["CDS View"]            || "",
                    "SQL View":            obj["SQL View"]            || "",
                    "Interface":           obj["Interface"]           || "",
                    "Consumption View":    obj["Consumption View"]    || "",
                    "Behavior Interface":  obj["Behavior Interface"]  || "",
                    "Behavior Definition": obj["Behavior Definition"] || "",
                    "Metadata Extension":  obj["Metadata Extension"]  || "",
                    "Access Control":      obj["Access Control"]      || "",
                    "Class":               obj["Class"]               || "",
                    "System (S4/EWM/TM/etc)": obj["System (S4/EWM/TM/etc)"] || obj["System (S4 / EWM / TM / etc)"] || ""
                }; })
            ));
            spacer();
        }

        // Entity 1
        ph3("Entity 1");

        var customTable = safeArr(entity1["Custom Table"] || []);
        if (customTable.length > 0) {
            ph3("Custom Table");
            documentChildren.push(makeColumnTable(D, ["Field","Description","Key","Type"],
                customTable.map(function(f){ return {
                    "Field": f["Field"]||"", "Description": f["Description"]||"",
                    "Key": f["Key"]||"", "Type": f["Type"]||""
                }; })
            ));
            spacer();
        }

        var entityDetails = safeArr(entity1["Entity Details"] || []);
        if (entityDetails.length > 0) {
            ph3("Entity Details");
            documentChildren.push(makeColumnTable(D, ["Properties","Technical Details","Description"],
                entityDetails.map(function(e){ return {
                    "Properties": e["Properties"]||"",
                    "Technical Details": e["Technical Details"]||"",
                    "Description": e["Description"]||""
                }; })
            ));
            spacer();
        }

        var opsAction = safeArr(entity1["Operations Action"] || entity1["OperationsAction"] || []);
        if (opsAction.length > 0) {
            ph3("Operations / Action");
            documentChildren.push(makeColumnTable(D, ["Operation/Action","Technical Details","Description"],
                opsAction.map(function(op){ return {
                    "Operation/Action":  op["Operation/Action"]  || "",
                    "Technical Details": op["Technical Details"] || "",
                    "Description":       op["Description"]       || ""
                }; })
            ));
            spacer();
        }

        var cdsJoin = entity1["CDS View Technical Details & Joins"] || "";
        if (cdsJoin) {
            ph3("CDS View Technical Details & Joins");
            pt(String(cdsJoin));
            spacer();
        }
    });

    // ══════════════════════════════════════════════════════════════════════
    // 6. UI5
    // ══════════════════════════════════════════════════════════════════════
    safeSection(D, documentChildren, "Fiori - UI5", function () {
        ph2("UI5", "FioriUI5");

        // UI5 App Details — fiori["UI5"]["UI5 App Details"][0]
        ph3("UI5 App Details");
        var ud = ui5Details.length > 0 ? safeObj(ui5Details[0]) : {};
        documentChildren.push(kvTable([
            ["Launchpad URL",    ud["Launchpad URL"]    || ""],
            ["App ID",           ud["App ID"]           || ud["APP ID"] || ""],
            ["Semantic Details", ud["Semantic Details"] || ""],
            ["Services",         ud["Services"]         || ""],
            ["GIT URL",          ud["GIT URL"]          || ""],
            ["MTA Details",      ud["MTA Details"]      || ""],
            ["Keywords",         ud["Keywords"]         || ""],
            ["Icons",            ud["Icons"]            || ""],
            ["Query Parameters", ud["Query Parameters"] || ""],
            ["App Type",         ud["App Type"]         || ""]
        ]));
        spacer();

        // UI5 App - Screen Flow — fiori["UI5"]["UI5 App - Screen Flow"][]
        ph3("UI5 App - Screen Flow");
        if (ui5ScreenFlow.length > 0) {
            documentChildren.push(makeColumnTable(D,
                ["View ID","View Layout","Controller ID","Controller Logic"],
                ui5ScreenFlow.map(function(sf){ return {
                    "View ID":          sf["View ID"]          || "",
                    "View Layout":      sf["View Layout"]      || "",
                    "Controller ID":    sf["Controller ID"]    || "",
                    "Controller Logic": sf["Controller Logic"] || ""
                }; })
            ));
        } else {
            documentChildren.push(makeColumnTable(D,
                ["View ID","View Layout","Controller ID","Controller Logic"],
                [{"View ID":"","View Layout":"","Controller ID":"","Controller Logic":""}]
            ));
        }
        spacer();

        // Action Logic — fatd["Action Logic"] = {Scan, Post, Complete}
        if (Object.keys(actionLogic).length > 0) {
            ph3("Action Logic");
            documentChildren.push(kvTable(
                Object.keys(actionLogic).map(function(k){ return [k, actionLogic[k] || ""]; })
            ));
            spacer();
        }
    });

    // ══════════════════════════════════════════════════════════════════════
    // 7. CleanCore Recommendations
    // ══════════════════════════════════════════════════════════════════════
    if (!this.cleancorerecommendation) {
        safeSection(D, documentChildren, "CleanCore Recommendations", function () {
            that._appendCleanCoreRecommendations(documentChildren, response_cleancore, D);
            that.cleancorerecommendation = true;
        });
    }

    return documentChildren;
},

        // createWorkflowTSD
        



createWorkflowTSD: function (documentChildren, response, response_cleancore) {

    // INTERNAL HELPERS

    var D = window.docx;
    var that = this;
    var { Paragraph, TextRun, Bookmark, Table, TableRow, TableCell,
        WidthType, BorderStyle, AlignmentType, TableLayoutType } = D;

    var PAGE_WIDTH = 9360;

    function safeArr(val) {
        if (Array.isArray(val)) return val;
        if (val !== null && val !== undefined) return [val];
        return [];
    }

    function safeObj(val) {
        return (val && typeof val === "object" && !Array.isArray(val)) ? val : {};
    }

    function safeStr(val) {
        if (val === null || val === undefined) return "";
        if (Array.isArray(val)) return val.join(", ");
        if (typeof val === "object") {
            try { return JSON.stringify(val); } catch { return ""; }
        }
        return String(val);
    }

    function safeSection(label, fn) {
        try { fn(); } catch { /* section failed silently; skip it */ }
    }

    function emptyRow(headers) {
        var row = {};
        headers.forEach(function (h, i) { row[h] = i === 0 ? "" : ""; });
        return [row];
    }

    function makeBorders() {
        var b = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
        return { top: b, bottom: b, left: b, right: b };
    }

    // Standard multi-column table (horizontal headers)
    function makeColumnTable(headers, rows, colWidths) {
        var borders = makeBorders();
        var widths;
        if (colWidths && colWidths.length === headers.length) {
            widths = colWidths;
        } else {
            var base = Math.floor(PAGE_WIDTH / headers.length);
            widths = headers.map(function () { return base; });
        }

        var headerRow = new TableRow({
            children: headers.map(function (h, i) {
                return new TableCell({
                    width: { size: widths[i], type: WidthType.DXA },
                    borders: borders,
                    shading: { fill: "E5E7EB" },
                    children: [new Paragraph({
                        alignment: AlignmentType.LEFT,
                        children: [new TextRun({ text: h || "", bold: true, color: "000000", size: 18, font: "Helvetica" })]
                    })]
                });
            })
        });

        var dataRows = (rows && rows.length > 0 ? rows : emptyRow(headers)).map(function (rowObj) {
            return new TableRow({
                children: headers.map(function (h, i) {
                    var val = (rowObj && rowObj[h] !== undefined && rowObj[h] !== null) ? safeStr(rowObj[h]) : "";
                    return new TableCell({
                        width: { size: widths[i], type: WidthType.DXA },
                        borders: borders,
                        children: [new Paragraph({
                            children: [new TextRun({ text: val, size: 18, font: "Helvetica" })]
                        })]
                    });
                })
            });
        });

        return new Table({
            layout: TableLayoutType.FIXED,
            width: { size: PAGE_WIDTH, type: WidthType.DXA },
            rows: [headerRow].concat(dataRows)
        });
    }

    // Activity side-by-side table (rows = fields, columns = activities)
    function makeActivityTable(adList) {
        if (!adList || adList.length === 0) {
            return makeColumnTable(["Activity"], [{ "Activity": "No activity details available." }]);
        }
        var borders = makeBorders();
        var labelW = 1800;
        var valW = Math.floor((PAGE_WIDTH - labelW) / adList.length);

        function makeActRow(labelText, valueFn, isGroupHeader) {
            var labelCell = new TableCell({
                width: { size: labelW, type: WidthType.DXA },
                borders: borders,
                shading: { fill: "E5E7EB" },
                children: [new Paragraph({
                    children: [new TextRun({ text: labelText, bold: true, size: 18, font: "Helvetica" })]
                })]
            });
            var valueCells = adList.map(function (ad) {
                var val = valueFn(ad);
                return new TableCell({
                    width: { size: valW, type: WidthType.DXA },
                    borders: borders,
                    shading: isGroupHeader ? { fill: "E5E7EB" } : undefined,
                    children: [new Paragraph({
                        children: [new TextRun({ text: safeStr(val), size: 18, font: "Helvetica", bold: !!isGroupHeader })]
                    })]
                });
            });
            return new TableRow({ children: [labelCell].concat(valueCells) });
        }

        var tableRows = [
            // Step / Type header rows
            makeActRow("", function (ad) {
                return "Step: " + safeStr(ad["step"] || ad["Step"] || ad.stepId || "");
            }, true),
            makeActRow("", function (ad) {
                return "Type: " + safeStr(ad["type"] || ad["Step Type"] || "Activity");
            }, true),
            makeActRow("Activity:", function (ad) { return ad["activity"] || ad["Activity"] || ""; }, false),
            makeActRow("Step Description:", function (ad) { return ad["stepDescription"] || ad["Step Description"] || ""; }, false),
            makeActRow("Responsibility:", function (ad) { return ad["responsibility"] || ad["Responsibility"] || "-"; }, false),
            makeActRow("Deadlines:", function (ad) { return ad["deadlines"] || ad["Deadlines"] || "-"; }, false),
            makeActRow("Exception Handling:", function (ad) { return ad["exceptionHandling"] || ad["Exception Handling"] || "-"; }, false),
            makeActRow("Task Abbreviation:", function (ad) { return ad["taskAbbreviation"] || ad["Task Abbreviation"] || ""; }, false),
            makeActRow("Task Number:", function (ad) { return ad["taskNumber"] || ad["Task Number"] || ""; }, false),
            makeActRow("Task Description:", function (ad) { return ad["taskDescription"] || ad["Task Description"] || ""; }, false),
            makeActRow("Object Type:", function (ad) { return ad["objectType"] || ad["Object Type"] || ""; }, false),
            makeActRow("Method:", function (ad) { return ad["method"] || ad["Method"] || ""; }, false),
            makeActRow("Work Item Text:", function (ad) { return ad["workItemText"] || ad["Work Item Text"] || ""; }, false),
            makeActRow("Agent Assignment:", function (ad) { return ad["agentAssignment"] || ad["Agent Assignment"] || "-"; }, false),
            makeActRow("Execution:", function (ad) { return ad["execution"] || ad["Execution"] || "-"; }, false)
        ];

        return new Table({
            layout: TableLayoutType.FIXED,
            width: { size: PAGE_WIDTH, type: WidthType.DXA },
            rows: tableRows
        });
    }

    // Heading helpers
    function pushH2(text, bookmarkId) {
        var children = bookmarkId
            ? [new Bookmark({ id: bookmarkId, children: [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })] })]
            : [new TextRun({ text: text, bold: true, size: 28, font: "Helvetica" })];
        documentChildren.push(new Paragraph({ children: children, spacing: { before: 250, after: 150 } }));
    }

    function pushH3(text) {
        documentChildren.push(new Paragraph({
            children: [new TextRun({ text: text, bold: true, size: 24, font: "Helvetica" })],
            spacing: { before: 200, after: 100 }
        }));
    }

    function pushText(text) {
        var lines = safeStr(text).split("\n");
        lines.forEach(function (line) {
            documentChildren.push(new Paragraph({
                children: [new TextRun({ text: line || "", size: 20, font: "Helvetica" })],
                spacing: { before: 60, after: 80 }
            }));
        });
    }

    function pushSpacer() {
        documentChildren.push(new Paragraph({ children: [], spacing: { before: 100, after: 100 } }));
    }

    // RESOLVE DATA ROOTS

    var r3 = safeObj(response && response.response3);

    // Workflowdetails[0] — primary data container
    var wd = safeObj(safeArr(r3.Workflowdetails)[0]);

    // NEW: pull Workflow Definition sub-object (new response nests data here)
    var wfDef = safeObj(wd["Workflow Definition"] || wd["Workflow definition"] || {});

    // NEW: if wd is empty but r3 has flat keys (old format), use r3 directly as wd
    if (Object.keys(wd).length === 0 && Object.keys(r3).length > 0) {
        wd = r3;
    }

    // Sub-objects from wd (camelCase — old nested format)
    // NEW: also try from wfDef for new nested format
    var wgi  = safeObj(wd.workflowGeneralInformation || wfDef["Workflow General Information"] || wfDef.workflowGeneralInformation || {});
    // NEW: also map from new response's Workflow_General_Information object
    if (Object.keys(wgi).length === 0 && wfDef["Workflow General Information"]) {
        var wgiNew = safeObj(wfDef["Workflow General Information"]);
        wgi = {
            objectName: wgiNew["Object Name"] || "",
            businessObjectType: wgiNew["Business Object Type"] || "",
            standardOrCustomWorkflowTemplate: wgiNew["Standard / Custom Workflow Template"] || "",
            comments: wgiNew["Comments"] || ""
        };
    }

    var wsd  = safeObj(wd.workflowScenarioDefinition || {});
    // NEW: map from wfDef.Workflow Template Diagram.Workflow Scenario Definition
    if (Object.keys(wsd).length === 0) {
        var wsdParent = safeObj(wfDef["Workflow Template Diagram"] || {});
        var wsdNew    = safeObj(wsdParent["Workflow Scenario Definition"] || {});
        if (Object.keys(wsdNew).length > 0) {
            wsd = {
                scenarioID: wsdNew["Scenario ID"] || "",
                abbreviation: wsdNew["Abbreviation"] || "",
                triggeringEvents: wsdNew["Triggering Events"] || "",
                objectType: wsdNew["Object Type"] || "",
                startConditions: wsdNew["Start Conditions"] || []
            };
        }
    }

    var td   = safeObj(wd.taskDetails || {});
    var fwsd = safeObj(wd.flexibleWorkflowScenarioDetails || {});
    var fbt  = safeObj(fwsd.flexibleBlockTabs || {});
    var rmOld = safeObj(wd.responsibilityManagementDetails || {});
    var extraOld = safeObj(wd.extra || {});

    // Helper: pick first truthy value from a list of sources
    function pick() {
        for (var i = 0; i < arguments.length; i++) {
            var v = arguments[i];
            if (v !== null && v !== undefined && v !== "") return v;
        }
        return "";
    }

    // Flat top-level r3 keys (new response format) with old fallbacks
    // NEW: also pull from r3 top-level flat keys and wd["Program Flow & Pseudo Code"]
    var pfPseudo = safeObj(wd["Program Flow & Pseudo Code"] || wd["Program Flow & Pseudo-code"] || {});
    var r3PostApproval    = pick(r3["Post_Approval_Process"], r3["Post Approval Process"],    wfDef["Post Approval Process"],    pfPseudo["Post Approval Process"],    extraOld["Post Approval Process"],    wd["Post Approval Process"]);
    var r3Rejection       = pick(r3["Rejection_Scenarios"],   r3["Rejection Scenarios"],       wfDef["Rejection Scenarios"],       pfPseudo["Rejection Scenarios"],       extraOld["Rejection Scenarios"],       wd["Rejection Scenarios"]);
    var r3Deadline        = pick(r3["Deadline_Monitoring"],   r3["Deadline Monitoring"],       wfDef["Deadline Monitoring"],       pfPseudo["Deadline Monitoring"],       extraOld["Deadline Monitoring"],       wd["Deadline Monitoring"]);
    var r3External        = pick(r3["External_Connectivity"], r3["External Connectivity"],     wfDef["External Connectivity"],     pfPseudo["External Connectivity"],     extraOld["External Connectivity"],     wd["External Connectivity"]);

    // NEW: post approval / rejection also inside wfDef
    if (!r3PostApproval) r3PostApproval = safeStr(wfDef["Post Approval Process"] || "");
    if (!r3Rejection)    r3Rejection    = safeStr(wfDef["Rejection Scenarios"]   || "");
    if (!r3Deadline)     r3Deadline     = safeStr(wfDef["Deadline Monitoring"]   || "");

    var r3ValueHelps      = safeArr(r3["Value_Helps"]      || r3["Value Helps"]      || wfDef["Value Helps"]      || wd["Value Helps"]      || fbt.valueHelps);
    var r3EmailTemplates  = safeArr(r3["Email_Templates"]  || r3["Email Templates"]  || wfDef["Email Templates"]  || wd["Email Templates"]  || fbt.emailTemplates);
    var r3EmailTplDetails = safeArr(r3["Email_Template_Details"] || r3["Email Template Details"] || wfDef["Email Template Details"] || wd["Email Templates Details"] || wd["Email Template Details"]);
    var r3WorkflowInApp   = safeArr(r3["Workflow_In_App"]  || r3["Workflow In App"]  || wfDef["Workflow In App"]  || wd["Workflow In App"]);
    var r3WorkflowSteps   = safeArr(r3["Workflow_Steps"]   || r3["Workflow Steps"]   || wfDef["Workflow Steps"]   || wd["Workflow Steps"]);
    var r3RespMgmt        = safeArr(r3["Responsibility_Management_Details"] || r3["Responsibility Management Details"] || wfDef["Responsibility Management Details"] || wd["Responsibility Management Details"]);
    var r3NonDDO          = safeArr(r3["Non_Data_Dictionary_Objects"] || r3["Non Data Dictionary Objects"] || wfDef["Non Data Dictionary Objects"] || wd["Non Data Dictionary Objects"]);
    var r3DisplayAgentRules = safeArr(
        r3["Display_Assigned_Agent_Rules_Relevant_for"] ||
        r3["Display Assigned Agent Rules Relevant for"] ||
        wfDef["Display Assigned Agent Rules Relevant for"] ||
        wd["Display Assigned Agent Rules Relevant for"] ||
        fbt.assignedAgentRules
    );
    var r3OpenPoints      = safeArr(r3["Open_Points"] || r3["Open Points"] || wfDef["Open Points"] || wd["Open Points"]);

    // BADI Implementation Details — new response has it at r3 top level as array
    var r3BADIImpl = safeArr(r3["BADI_Implementation_Details"] || r3["BADI Implementation Details"] || wfDef["BADI Implementation Details"] || wd["BADI Implementation Details"]);

    // Enhancement Control Framework — new response has it nested under Process_in_which_BADI_Triggers
    var badiTriggers = safeObj(
        r3["Process_in_which_BADI_Triggers"] ||
        r3["Process in which BADI Triggers"] ||
        wd["Process in which BADI Triggers"] ||
        wd["Process in which BADI/exit/enhancement triggers"] ||
        {}
    );
    var ecfRaw = badiTriggers["Enhancement Control Framework"] || badiTriggers["Enhancement control framework"] || [];
    var r3ECF = Array.isArray(ecfRaw) ? ecfRaw : (ecfRaw && typeof ecfRaw === "object" ? [ecfRaw] : []);

    // ZAXP tables
    var r3ENH_HDR = safeArr(wd["ZAXP_STP_ENH_HDR"] || r3["ZAXP_STP_ENH_HDR"]);
    var r3ENH_ITM = safeArr(wd["ZAXP_STP_ENH_ITM"] || r3["ZAXP_STP_ENH_ITM"]);

    // Steps Details — new response has it inside wfDef
    var r3Steps = safeArr(wd["Steps Details"] || wd["Steps Details:"] || wfDef["Steps Details"] || safeArr(wfDef["Steps Details:"])[0] || td.stepsDetails);
    // NEW: also try wfDef array directly
    if (r3Steps.length === 0) {
        var sds = wfDef["Steps Details:"] || wfDef["Steps Details"] || [];
        r3Steps = safeArr(sds);
    }

    // Activity Details — new response has it in wfDef["Activity Details"] array
    var adRaw = wd["Activity Details"] || wd["Activity details"] || wfDef["Activity Details"] || wfDef["Activity details"] || td.activityDetails;
    var r3Activities = Array.isArray(adRaw) ? adRaw
        : (adRaw && typeof adRaw === "object" ? [adRaw] : []);

    // Workflow Container — new response in wfDef
    var r3WFContainer = safeArr(wd["Workflow Container"] || wfDef["Workflow Container"] || wsd.workflowContainer);

    // Binding WF Container to Event Container
    var r3Binding = safeArr(
        wd["Binding Workflow Container to Event Container"] ||
        wfDef["Binding Workflow Container to Event Container"] ||
        wsd.bindingWorkflowContainerEventContainer
    );

    // Agent Rules — new response in wfDef
    var agentRulesRaw = wd["Agent Rules"] || wd["Agent Rules -"] || wfDef["Agent Rules"] || fbt.agentRules;
    var r3AgentRules = Array.isArray(agentRulesRaw) ? agentRulesRaw
        : (agentRulesRaw && typeof agentRulesRaw === "object" ? [agentRulesRaw] : []);

    var ardRaw = wd["Agent Rule Details"] || wd["Agent Rule details"] || wfDef["Agent Rule Details"] || wfDef["Agent Rule details"];
    var r3AgentRuleDetails = Array.isArray(ardRaw) ? ardRaw
        : (ardRaw && typeof ardRaw === "object" ? [ardRaw] : []);

    var bindingRuleRaw = wd["Binding Workflow Container to Rule Container"] || wfDef["Binding Workflow Container to Rule Container"];
    var r3BindingRule = safeArr(bindingRuleRaw);

    // Program Flow / Pseudo Code — new response inside pfPseudo or wfDef
    var r3ProcessFlow  = safeStr(pfPseudo["Process flow"] || pfPseudo["Process Flow"] || wd["Process flow"] || wd["Process Flow"] || "");
    var r3PseudoCode   = safeStr(
        (safeObj(pfPseudo["Pseudo Code"])["Class"] ? "Class: " + safeObj(pfPseudo["Pseudo Code"])["Class"] + "\nEvent: " + safeObj(pfPseudo["Pseudo Code"])["Event"] + "\nStandard flexible workflow: " + safeObj(pfPseudo["Pseudo Code"])["Standard flexible workflow"] : "") ||
        pfPseudo["Pseudo code"] || pfPseudo["Pseudo Code"] ||
        wd["Pseudo code"] || wd["Pseudo Code"] || ""
    );

    var r3POApprovalTypes = safeStr(
        (wd["Types of Purchase Order Approval and the sequence to determine the approval agents"] ||
        wd["Types of Purchase Order Approval"] ||
        r3["Types of Purchase Order Approval and the sequence to determine the approval agents"] ||
        safeArr(safeObj(pfPseudo["Workflow Approval Types"]))[0]) ? JSON.stringify(pfPseudo["Workflow Approval Types"]) : ""
    );

    // Security / Exception
    var r3ExceptionHandling = safeStr(wd["Exception Handling"] || pfPseudo["Exception Handling"] || r3["Exception Handling"] || "NA");
    var r3AdditionalInfoSec = safeObj(wd["Additional Information (NA)"] || wd["Additional Information"] || {});

    // Tables to be created — new response has it inside pfPseudo
    var tablesToCreateRaw = pfPseudo["Tables to be created"] || wd["Tables to be created"] || [];
    var r3TablesToCreate = Array.isArray(tablesToCreateRaw)
        ? tablesToCreateRaw.map(function(t) {
            var to = safeObj(t);
            return (to["Table Name"] || to.tableName || "") + (to["Purpose"] ? " — " + to["Purpose"] : "") + (to["Key Fields"] ? " (Key: " + to["Key Fields"] + ")" : "");
          }).join("\n")
        : safeStr(tablesToCreateRaw);

    // NEW: also pull Workflow General Info from new response's Workflow General Information sub-key
    var wgiFromWfDef = safeObj(wfDef["Workflow General Information"] || {});
    if (Object.keys(wgi).length === 0 && Object.keys(wgiFromWfDef).length > 0) {
        wgi = {
            objectName: wgiFromWfDef["Object Name"] || "",
            businessObjectType: wgiFromWfDef["Business Object Type"] || "",
            "Standard / Custom workflow Template": wgiFromWfDef["Standard / Custom Workflow Template"] || "",
            comments: wgiFromWfDef["Comments"] || ""
        };
    }

    // NEW: normalise Responsibility Management — pull from r3 top level if available
    if (r3RespMgmt.length === 0 && safeArr(r3["Responsibility_Management_Details"]).length > 0) {
        r3RespMgmt = safeArr(r3["Responsibility_Management_Details"]);
    }

    // NEW: normalise Email Template Details — r3.Email_Template_Details
    if (r3EmailTplDetails.length === 0 && safeArr(r3["Email_Template_Details"]).length > 0) {
        r3EmailTplDetails = safeArr(r3["Email_Template_Details"]).map(function(etd) {
            return {
                "Template Name": etd["Template Name"] || etd.templateName || "",
                "CDS View Name": etd["CDS View Name"] || etd.cdsViewName || "",
                "Usage": etd["Usage"] || etd.usage || ""
            };
        });
    }

    // NEW: normalise Display Agent Rules — r3.Display_Assigned_Agent_Rules_Relevant_for
    if (r3DisplayAgentRules.length === 0 && safeArr(r3["Display_Assigned_Agent_Rules_Relevant_for"]).length > 0) {
        r3DisplayAgentRules = safeArr(r3["Display_Assigned_Agent_Rules_Relevant_for"]);
    }

    // TOP-LEVEL BOOKMARK
    documentChildren.push(new Paragraph({
        children: [new Bookmark({
            id: "Workflow",
            children: [new TextRun({ text: "Workflow", bold: true, size: 32, font: "Helvetica" })]
        })],
        spacing: { before: 250, after: 250 }
    }));

    // SECTION 1: Program Flow & Pseudo-code
   
    safeSection("Program Flow & Pseudo-code", function () {
        pushH2("Program Flow & Pseudo-code", "ProgramFlowPseudoCode");

        pushH3("Process flow");
        if (r3ProcessFlow) {
            pushText(r3ProcessFlow);
        } else {
            // Placeholder matching client PDF note
            pushText("[Process flow diagram — see FSD / attached diagram]");
        }
        pushSpacer();

        pushH3("Pseudo code");
        if (r3PseudoCode) {
            pushText(r3PseudoCode);
        } else {
            // Populate from wd Workflow General Info if API didn't return it
            var wgiObjName = wgi.objectName || wd["Workflow General Information"] && safeObj(wd["Workflow General Information"]).objectName || "";
            var wsdTrigger = wsd.triggeringEvents || "";
            var stdWFTemplate = wgi.standardOrCustomWorkflowTemplate || wgi["Standard / Custom workflow Template"] || "";
            if (wgiObjName || stdWFTemplate || wsdTrigger) {
                pushText("Class: CL_MM_PUR_WF_OBJECT_PO");
                pushText("Event: " + (wsdTrigger || "SUBMITTED_FOR_APPROVAL"));
                pushText("Standard flexible workflow: " + (stdWFTemplate || "WS00800238"));
            }
        }
        pushSpacer();

        // Types of PO Approval
        if (r3POApprovalTypes) {
            pushH3("Types of Purchase Order Approval and the sequence to determine the approval agents:");
            pushText(r3POApprovalTypes);
            pushSpacer();
        }

        // Tables to be created
        if (r3TablesToCreate) {
            pushH3("Tables to be created:");
            pushText(r3TablesToCreate);
            pushSpacer();
        }
    });

    // SECTION 2: Security Information
   
    safeSection("Security Information", function () {
        pushH2("Security Information", "SecurityInformation");
        pushH3("Exception Handling");
        pushText(r3ExceptionHandling);
        pushSpacer();

        pushH3("Additional Information (NA)");
        var secAICols = [
            "Application/Platform", "Object Name", "Object Type", "Object Description",
            "Step By Step Instructions", "Global/Local", "System (S4/EWM/TM/etc)",
            "Business Unit/Country", "Localization Reason", "Target/Intermediate State",
            "Related JIRA Issues", "Additional Information"
        ];
        var secAIRow = {};
        secAICols.forEach(function (c) { secAIRow[c] = safeStr(r3AdditionalInfoSec[c] || ""); });
        documentChildren.push(makeColumnTable(secAICols, [secAIRow]));
        pushSpacer();
    });

    // SECTION 3: Workflow General Information
   
    safeSection("Workflow General Information", function () {
        pushH2("Workflow General Information", "WorkflowGeneralInformation");

        var wgiHeaders = ["Object Name", "Business Object Type", "Standard / Custom Workflow Template.", "Comments"];
        var wgiWidths  = [2000, 2500, 2860, 2000];
        documentChildren.push(makeColumnTable(wgiHeaders, [{
            "Object Name":                             safeStr(wgi.objectName || wd["Object Name"] || ""),
            "Business Object Type":                    safeStr(wgi.businessObjectType || wd["Business Object Type"] || ""),
            "Standard / Custom Workflow Template.":    safeStr(wgi.standardOrCustomWorkflowTemplate || wgi["Standard / Custom workflow Template"] || ""),
            "Comments":                                safeStr(Array.isArray(wgi.comments) ? wgi.comments.join("\n") : (wgi.comments || wgi["Comments"] || ""))
        }], wgiWidths));
        pushSpacer();
    });

    // SECTION 4: Workflow Template Diagram
    
    safeSection("Workflow Template Diagram", function () {
        pushH2("Workflow Template Diagram", "WorkflowTemplateDiagram");

        // 4a — Scenario Definition
        pushH3("Workflow Scenario Definition");
        var sdHeaders = ["Scenario ID", "Abbreviation:", "Triggering Events:", "Object Type:", "Start Conditions:"];
        var sdWidths  = [1200, 1200, 1800, 2000, 3160];

        // Start conditions — may be a string or array in API response
        var startCond = wsd.startConditions || wsd["Start Condition"] || wsd["Start Conditions"] || wd["Start Conditions"] || "";
        if (Array.isArray(startCond)) startCond = startCond.join("\n");

        documentChildren.push(makeColumnTable(sdHeaders, [{
            "Scenario ID":          safeStr(wsd.scenarioID || wsd["Scenario ID"] || ""),
            "Abbreviation:":        safeStr(wsd.abbreviation || wsd["Abbreviation"] || ""),
            "Triggering Events:":   safeStr(wsd.triggeringEvents || wsd["Triggering Events"] || ""),
            "Object Type:":         safeStr(wsd.objectType || wsd["Object Type"] || ""),
            "Start Conditions:":    safeStr(startCond)
        }], sdWidths));
        pushSpacer();

        // 4b — Workflow Container
        pushH3("Workflow Container");
        var wcH = ["Element", "Description", "Import\nY/N", "Export\nY/N", "Multiple\nValues Y/N", "Mandatory\nY/N", "Dictionary\nField", "Object\nType"];
        var wcW = [1400, 1800, 700, 700, 900, 800, 1360, 1700];

        documentChildren.push(makeColumnTable(wcH,
            r3WFContainer.length > 0
                ? r3WFContainer.map(function (c) {
                    return {
                        "Element":          safeStr(c.element || c["Element"] || ""),
                        "Description":      safeStr(c.description || c["Description"] || ""),
                        "Import\nY/N":      safeStr(c.import || c["Import"] || c["Import Y/N"] || ""),
                        "Export\nY/N":      safeStr(c.export || c["Export"] || c["Export Y/N"] || ""),
                        "Multiple\nValues Y/N": safeStr(c.multipleValues || c["Multiple Values"] || c["Multiple Values Y/N"] || ""),
                        "Mandatory\nY/N":   safeStr(c.mandatory || c["Mandatory"] || c["Mandatory Y/N"] || ""),
                        "Dictionary\nField": safeStr(c.dictionaryField || c["Dictionary Field"] || ""),
                        "Object\nType":     safeStr(c.objectType || c["Object Type"] || "")
                    };
                })
                : emptyRow(wcH),
            wcW
        ));
        pushSpacer();

        // 4c — Binding Workflow Container to Event Container
        pushH3("Binding Workflow Container to Event Container");
        var bwcH = [" Event Container Element", "Direction", "Workflow Container Element"];
        var bwcW = [3500, 1360, 4500];
        documentChildren.push(makeColumnTable(bwcH,
            r3Binding.length > 0
                ? r3Binding.map(function (b) {
                    return {
                        " Event Container Element":   safeStr(b.eventContainerElement || b["Event Container Element"] || b["&IV_OBJECT_ID&"] || ""),
                        "Direction":                  safeStr(b.direction || b["Direction"] || ""),
                        "Workflow Container Element": safeStr(b.workflowContainerElement || b["Workflow Container Element"] || "")
                    };
                })
                : emptyRow(bwcH),
            bwcW
        ));
        pushSpacer();
    });

    // SECTION 5: Steps Details & Activity Details
    
    safeSection("Steps Details", function () {
        pushH2("Steps Details:", "StepsDetails");

        var sdH = ["Step", "Step Type", "Step Description"];
        var sdW = [1560, 1800, 6000];
        documentChildren.push(makeColumnTable(sdH,
            r3Steps.length > 0
                ? r3Steps.map(function (s) {
                    return {
                        "Step":             safeStr(s.step || s["Step"] || ""),
                        "Step Type":        safeStr(s.stepType || s["Step Type"] || ""),
                        "Step Description": safeStr(s.stepDescription || s["Step Description"] || "")
                    };
                })
                : emptyRow(sdH),
            sdW
        ));
        pushSpacer();

        pushH2("Activity details", "ActivityDetails");
        documentChildren.push(makeActivityTable(r3Activities));
        pushSpacer();
    });

    // SECTION 6: Agent Rules
   
    safeSection("Agent Rules", function () {
        pushH2("Agent Rules -", "AgentRules");

        // 6a Agent Rules
        var arH = ["Unique Name", "Description", "Rule Category", "Agent Rule"];
        var arW = [2500, 3000, 1860, 2000];
        documentChildren.push(makeColumnTable(arH,
            r3AgentRules.length > 0
                ? r3AgentRules.map(function (ar) {
                    return {
                        "Unique Name":   safeStr(ar.uniqueName || ar["Unique Name"] || ar["Agent Rule"] || ""),
                        "Description":   safeStr(ar.description || ar["Description"] || ""),
                        "Rule Category": safeStr(ar.ruleCategory || ar["Rule Category"] || ""),
                        "Agent Rule":    safeStr(ar.agentRule || ar["Agent Rule"] || ar.ruleId || "")
                    };
                })
                : emptyRow(arH),
            arW
        ));
        pushSpacer();

        // 6b Agent Rule Details
        pushH3("Agent Rule details");
        var ardH = ["Rule ID", "Rule Abbreviation", "Rule Category", "Function Module"];
        var ardW = [1800, 2500, 2000, 3060];
        documentChildren.push(makeColumnTable(ardH,
            r3AgentRuleDetails.length > 0
                ? r3AgentRuleDetails.map(function (ard) {
                    return {
                        "Rule ID":           safeStr(ard["Rule ID"] || ard.ruleId || ""),
                        "Rule Abbreviation": safeStr(ard["Rule Abbreviation"] || ard.ruleAbbreviation || ""),
                        "Rule Category":     safeStr(ard["Rule Category"] || ard.ruleCategory || ""),
                        "Function Module":   safeStr(ard["Function Module"] || ard.functionModule || "")
                    };
                })
                : emptyRow(ardH),
            ardW
        ));
        pushSpacer();

        // 6c Binding WF Container to Rule Container
        pushH3("Binding Workflow Container to Rule Container");
        var brcH = ["Workflow Container Element", "Direction", "Rule Container Element"];
        var brcW = [3500, 1360, 4500];
        documentChildren.push(makeColumnTable(brcH,
            r3BindingRule.length > 0
                ? r3BindingRule.map(function (b) {
                    return {
                        "Workflow Container Element": safeStr(b["Workflow Container Element"] || b.workflowContainerElement || ""),
                        "Direction":                  safeStr(b["Direction"] || b.direction || ""),
                        "Rule Container Element":     safeStr(b["Rule Container Element"] || b.ruleContainerElement || "")
                    };
                })
                : emptyRow(brcH),
            brcW
        ));
        pushSpacer();

        // 6d Display Assigned Agent Rules Relevant for
        pushH3("Display Assigned Agent Rules Relevant for");
        var darH = ["ID", "Type", "Task", "Email\nNotify", "Processing", "Review", "Exceptional\nHandling", "Description"];
        var darW = [1200, 900, 900, 800, 900, 800, 900, 2960];
        documentChildren.push(makeColumnTable(darH,
            r3DisplayAgentRules.length > 0
                ? r3DisplayAgentRules.map(function (d) {
                    var desc = safeStr(
                        d["Description"] || d.description ||
                        d["relevantFor(Task Processing/Email Notify/Review/Exceptional Handling)Description"] || ""
                    );
                    return {
                        "ID":                 safeStr(d["ID"] || d.id || ""),
                        "Type":               safeStr(d["Type"] || d.type || ""),
                        "Task":               safeStr(d["Task"] || d.task || ""),
                        "Email\nNotify":      safeStr(d["Email Notify"] || d.emailNotify || ""),
                        "Processing":         safeStr(d["Processing"] || d.processing || ""),
                        "Review":             safeStr(d["Review"] || d.review || ""),
                        "Exceptional\nHandling": safeStr(d["Exception Handling"] || d["Exceptional Handling"] || d.exceptionHandling || d.exceptionalHandling || ""),
                        "Description":        desc
                    };
                })
                : emptyRow(darH),
            darW
        ));
        pushSpacer();
    });

    // SECTION 7: Value Helps -
   
    safeSection("Value Helps", function () {
        pushH2("Value Helps -", "ValueHelps");

        var vhH = ["TypeName", "Service Path", "Entity", "Property"];
        var vhW = [2000, 3360, 2000, 2000];
        documentChildren.push(makeColumnTable(vhH,
            r3ValueHelps.length > 0
                ? r3ValueHelps.map(function (vh) {
                    return {
                        "TypeName":     safeStr(vh.typeName || vh["Type Name"] || vh["TypeName"] || ""),
                        "Service Path": safeStr(vh.servicePath || vh["Service Path"] || ""),
                        "Entity":       safeStr(vh.entity || vh["Entity"] || ""),
                        "Property":     safeStr(vh.property || vh["Property"] || "")
                    };
                })
                : emptyRow(vhH),
            vhW
        ));
        pushSpacer();
    });

    // SECTION 8: Email Templates Details -
   
    safeSection("Email Templates Details", function () {
        pushH2("Email Templates Details -", "EmailTemplatesDetails");

        // Table 1 — Email Templates summary
        var etH = ["Unique Name", "Description", "Use Case"];
        var etW = [3000, 3360, 3000];
        documentChildren.push(makeColumnTable(etH,
            r3EmailTemplates.length > 0
                ? r3EmailTemplates.map(function (et) {
                    return {
                        "Unique Name": safeStr(et.uniqueName || et["Unique Name"] || ""),
                        "Description": safeStr(et.description || et["Description"] || et.shortText || et["Short Text"] || ""),
                        "Use Case":    safeStr(et.useCase || et["Use Case"] || et.purpose || "")
                    };
                })
                : emptyRow(etH),
            etW
        ));
        pushSpacer();

        // Table 2 — Email Templates Details (CDS view mapping)
        // Client PDF columns: Email Template Name | CDS View Name
        var etdH = ["Email Template Name", "CDS View Name"];
        var etdW = [4680, 4680];

        // Build rows — handle both array and object formats from API
        var etdRows = [];
        if (r3EmailTplDetails.length > 0) {
            etdRows = r3EmailTplDetails.map(function (etd) {
                // Object format: { "Email Template": "...", "CDS View": "..." }
                // Array format:  { emailTemplateName: "...", cdsViewName: "..." }
                var tplName = safeStr(etd["Email Template"] || etd["Template Name"] || etd.emailTemplateName || etd.name || "");
                var cdsView = safeStr(etd["CDS View"] || etd["CDS View Name"] || etd.cdsViewName || etd.cdsView || "");
                return { "Email Template Name": tplName, "CDS View Name": cdsView };
            });
        }
        // If it came back as a single object {Email Template, CDS View, ...}
        if (etdRows.length === 0) {
            var etdObj = safeObj(wd["Email Templates Details"]);
            if (etdObj["Email Template"]) {
                etdRows = [{ "Email Template Name": safeStr(etdObj["Email Template"]), "CDS View Name": safeStr(etdObj["CDS View"] || "") }];
            }
        }
        documentChildren.push(makeColumnTable(etdH, etdRows.length > 0 ? etdRows : emptyRow(etdH), etdW));
        pushSpacer();
    });

    // SECTION 9: Post Approval / Rejection / Deadline / External
    
    safeSection("Post Approval Process", function () {
        pushH2("Post Approval Process", "PostApprovalProcess");
        pushText(r3PostApproval || "");
        pushSpacer();

        pushH2("Rejection Scenarios", "RejectionScenarios");
        pushText(r3Rejection || "");
        pushSpacer();

        pushH2("Deadline Monitoring", "DeadlineMonitoring");
        pushText(r3Deadline || "");
        pushSpacer();

        pushH2("External Connectivity", "ExternalConnectivity");
        pushText(r3External || "");
        pushSpacer();
    });

    // SECTION 10: Workflow In App
    
    safeSection("Workflow In App", function () {
        pushH2("Workflow In App", "WorkflowInApp");

        if (r3WorkflowInApp.length > 0) {
            var borders = makeBorders();
            var wiaW = [1800, 1800, 2800, 2960];

            // Header row
            var wiaHeaderRow = new TableRow({
                children: ["Workflow Name", "Description", "Start Conditions", "Steps"].map(function (h, i) {
                    return new TableCell({
                        width: { size: wiaW[i], type: WidthType.DXA },
                        borders: borders,
                        shading: { fill: "E5E7EB" },
                        children: [new Paragraph({
                            children: [new TextRun({ text: h, bold: true, size: 18, font: "Helvetica" })]
                        })]
                    });
                })
            });

            var wiaDataRows = r3WorkflowInApp.map(function (w) {
                var stepsArr = Array.isArray(w["Steps"]) ? w["Steps"] : (w["Steps"] ? [w["Steps"]] : []);
                var startCond = safeStr(w["Start Conditions"] || w.startConditions || "");

                var stepsParagraphs = stepsArr.length > 0
                    ? stepsArr.map(function (s, idx) {
                        return new Paragraph({
                            children: [new TextRun({ text: (idx + 1) + ". " + safeStr(s), size: 18, font: "Helvetica" })]
                        });
                    })
                    : [new Paragraph({ children: [new TextRun({ text: "", size: 18, font: "Helvetica" })] })];

                return new TableRow({
                    children: [
                        new TableCell({
                            width: { size: wiaW[0], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(w["Workflow Name"] || w.workflowName || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wiaW[1], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(w["Description"] || w.description || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wiaW[2], type: WidthType.DXA }, borders: borders,
                            children: startCond.split("\n").map(function (line) {
                                return new Paragraph({ children: [new TextRun({ text: line, size: 18, font: "Helvetica" })] });
                            })
                        }),
                        new TableCell({
                            width: { size: wiaW[3], type: WidthType.DXA }, borders: borders,
                            children: stepsParagraphs
                        })
                    ]
                });
            });

            documentChildren.push(new Table({
                layout: TableLayoutType.FIXED,
                width: { size: PAGE_WIDTH, type: WidthType.DXA },
                rows: [wiaHeaderRow].concat(wiaDataRows)
            }));
        } else {
            documentChildren.push(makeColumnTable(
                ["Workflow Name", "Description", "Start Conditions", "Steps"],
                emptyRow(["Workflow Name", "Description", "Start Conditions", "Steps"])
            ));
        }
        pushSpacer();
    });

    // SECTION 11: Workflow Steps
    
    safeSection("Workflow Steps", function () {
        pushH2("Workflow Steps", "WorkflowSteps");

        if (r3WorkflowSteps.length > 0) {
            var borders = makeBorders();
            var wsW = [1500, 1200, 1200, 2360, 1500, 1600];
            var wsHeaders = ["Workflow Name", "Step\nName", "Recipients", "Step Conditions", "Deadlines", "Exception\nHandling"];

            var wsHeaderRow = new TableRow({
                children: wsHeaders.map(function (h, i) {
                    return new TableCell({
                        width: { size: wsW[i], type: WidthType.DXA },
                        borders: borders,
                        shading: { fill: "E5E7EB" },
                        children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 18, font: "Helvetica" })] })]
                    });
                })
            });

            var wsDataRows = r3WorkflowSteps.map(function (ws) {
                var stepCondText = safeStr(ws["Conditions"] || ws["Step Conditions"] || ws.stepConditions || "");
                var condLines = stepCondText.split("\n");

                return new TableRow({
                    children: [
                        new TableCell({
                            width: { size: wsW[0], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(ws["Workflow Name"] || ws.workflowName || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wsW[1], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(ws["Step Name"] || ws.stepName || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wsW[2], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(ws["Recipients"] || ws.recipients || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wsW[3], type: WidthType.DXA }, borders: borders,
                            children: condLines.map(function (line) {
                                return new Paragraph({ children: [new TextRun({ text: line, size: 18, font: "Helvetica" })] });
                            })
                        }),
                        new TableCell({
                            width: { size: wsW[4], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(ws["Deadlines"] || ws.deadlines || ""), size: 18, font: "Helvetica" })] })]
                        }),
                        new TableCell({
                            width: { size: wsW[5], type: WidthType.DXA }, borders: borders,
                            children: [new Paragraph({ children: [new TextRun({ text: safeStr(ws["Exception Handling"] || ws.exceptionHandling || ""), size: 18, font: "Helvetica" })] })]
                        })
                    ]
                });
            });

            documentChildren.push(new Table({
                layout: TableLayoutType.FIXED,
                width: { size: PAGE_WIDTH, type: WidthType.DXA },
                rows: [wsHeaderRow].concat(wsDataRows)
            }));
        } else {
            documentChildren.push(makeColumnTable(
                ["Workflow Name", "Step Name", "Recipients", "Step Conditions", "Deadlines", "Exception Handling"],
                emptyRow(["Workflow Name", "Step Name", "Recipients", "Step Conditions", "Deadlines", "Exception Handling"])
            ));
        }
        pushSpacer();
    });

    // SECTION 12: Responsibility Management Details
    
    safeSection("Responsibility Management Details", function () {
        pushH2("Responsibility Management Details", "ResponsibilityManagementDetails");

        var rmH = ["Fiori\nApp", "Team Name", "Team Type", "Responsibility Definition", "Direct SubTeam", "Direct SuperTeam"];
        var rmW = [1200, 1600, 1600, 2360, 1300, 1300];

        var rmData;
        if (r3RespMgmt.length > 0) {
            rmData = r3RespMgmt.map(function (rm) {
                return {
                    "Fiori\nApp":                safeStr(rm["Fiori App"] || rm.fioriApp || ""),
                    "Team Name":                 safeStr(rm["Team Name"] || rm.teamName || ""),
                    "Team Type":                 safeStr(rm["Team Type"] || rm.teamType || ""),
                    "Responsibility Definition": safeStr(rm["Responsibility Definition"] || rm.responsibilityDefinition || ""),
                    "Direct SubTeam":            safeStr(rm["Direct SubTeam"] || rm["Direct Sub Team"] || rm.directSubTeam || ""),
                    "Direct SuperTeam":          safeStr(rm["Direct SuperTeam"] || rm["Direct Super Team"] || rm.directSuperTeam || "")
                };
            });
        } else if (rmOld && Object.keys(rmOld).length > 0) {
            rmData = [{
                "Fiori\nApp":                safeStr(rmOld.fioriApp || ""),
                "Team Name":                 safeStr(rmOld.teamName || ""),
                "Team Type":                 safeStr(rmOld.teamType || ""),
                "Responsibility Definition": safeStr(rmOld.responsibilityDefinition || ""),
                "Direct SubTeam":            safeStr(rmOld.directSubTeam || ""),
                "Direct SuperTeam":          safeStr(rmOld.directSuperTeam || "")
            }];
        } else {
            rmData = emptyRow(rmH);
        }
        documentChildren.push(makeColumnTable(rmH, rmData, rmW));
        pushSpacer();

        // Custom Responsibility Definitions
        var crdArr = safeArr(rmOld.customResponsibilityDefinitions || []);
        if (crdArr.length > 0) {
            pushH3("Custom Responsibility Definitions");
            var crdH = ["Resp Def Name", "Resp Def Ext Name", "CDS View", "CDS Field", "Resp Def Entity", "Resp Def Property", "Service Namespace", "Resp Def Ext Service", "Resp Def Serv Ver", "Resp Def Desc"];
            var crdW = crdH.map(function () { return Math.floor(PAGE_WIDTH / crdH.length); });
            documentChildren.push(makeColumnTable(crdH,
                crdArr.map(function (c) {
                    return {
                        "Resp Def Name":        safeStr(c.respDefName || ""),
                        "Resp Def Ext Name":    safeStr(c.respDefExtName || ""),
                        "CDS View":             safeStr(c.cdsView || ""),
                        "CDS Field":            safeStr(c.cdsField || ""),
                        "Resp Def Entity":      safeStr(c.respDefEntity || ""),
                        "Resp Def Property":    safeStr(c.respDefProperty || ""),
                        "Service Namespace":    safeStr(c.serviceNamespace || ""),
                        "Resp Def Ext Service": safeStr(c.respDefExtSrvc || ""),
                        "Resp Def Serv Ver":    safeStr(c.respDefServVer || ""),
                        "Resp Def Desc":        safeStr(c.respDefDesc || "")
                    };
                }),
                crdW
            ));
            pushSpacer();
        }

        // Assign Custom Resp Def to CDS View
        var acdArr = safeArr(rmOld.assignCustomResponsibilityDefinitionToCDSView || []);
        if (acdArr.length > 0) {
            pushH3("Assign Custom Responsibility Definition to CDS View");
            var acdH = ["CDS View Name", "CDS View Fields", "Custom Responsibility Definition Names"];
            documentChildren.push(makeColumnTable(acdH,
                acdArr.map(function (a) {
                    return {
                        "CDS View Name":                         safeStr(a.cdsViewName || ""),
                        "CDS View Fields":                       safeStr(a.cdsViewFields || ""),
                        "Custom Responsibility Definition Names": safeStr(a.customResponsibilityDefinitionNames || "")
                    };
                })
            ));
            pushSpacer();
        }

        // BADIs for Enhancements in Standard Workflow
        var badiEnh = safeArr(rmOld.badiEnhancements || []);
        if (badiEnh.length > 0) {
            pushH3("BADIs for Enhancements in Standard Workflow");
            var beH = Object.keys(safeObj(badiEnh[0]));
            if (beH.length > 0) {
                documentChildren.push(makeColumnTable(beH, badiEnh));
            }
            pushSpacer();
        }
    });

    // SECTION 13: Enhancement
    
    safeSection("Enhancement", function () {
        pushH2("Enhancement", "Enhancement");

        // 13a Non-DDO
        pushH3("Non-Data Dictionary Objects (NA)");
        var enhH = [
            "Application\n/Platform", "Object\nName", "Object\nType", "Object\nDescription",
            "Released\nfor cloud", "Step By Step\nInstructions", "Global\n/Local",
            "System\n(S4/EWM/TM/etc)", "Business Unit\n/Country", "Localization\nReason",
            "Target\n/Intermediate State", "Related\nJIRA Issues", "Additional\nInformation"
        ];
        var enhW = enhH.map(function () { return Math.floor(PAGE_WIDTH / enhH.length); });
        documentChildren.push(makeColumnTable(enhH,
            r3NonDDO.length > 0
                ? r3NonDDO.map(function (e) {
                    return {
                        "Application\n/Platform":      safeStr(e["Application/Platform"] || e.applicationPlatform || ""),
                        "Object\nName":                safeStr(e["Object Name"] || e["Object\nName"] || e.objectName || ""),
                        "Object\nType":                safeStr(e["Object Type"] || e["Object\nType"] || e.objectType || ""),
                        "Object\nDescription":         safeStr(e["Object Description"] || e["Object\nDescription"] || e.objectDescription || ""),
                        "Released\nfor cloud":         safeStr(e["Released for cloud"] || e["Released for Cloud"] || e.releasedForCloud || ""),
                        "Step By Step\nInstructions":  safeStr(e["Step By Step Instructions"] || e.stepByStepInstructions || ""),
                        "Global\n/Local":              safeStr(e["Global/Local"] || e.globalLocal || ""),
                        "System\n(S4/EWM/TM/etc)":     safeStr(e["System (S4/EWM/TM/etc)"] || e.system || ""),
                        "Business Unit\n/Country":     safeStr(e["Business Unit/Country"] || e.businessUnitCountry || ""),
                        "Localization\nReason":        safeStr(e["Localization Reason"] || e.localizationReason || ""),
                        "Target\n/Intermediate State": safeStr(e["Target/Intermediate State"] || e.targetIntermediateState || ""),
                        "Related\nJIRA Issues":        safeStr(e["Related JIRA Issues"] || e.relatedJiraIssues || ""),
                        "Additional\nInformation":     safeStr(e["Additional Information"] || e.additionalInformation || "")
                    };
                })
                : emptyRow(enhH),
            enhW
        ));
        pushSpacer();

        // 13b Process in which BADI triggers
        pushH3("Process in which BADI/exit/enhancement triggers");
        pushSpacer();

        // 13c Enhancement control framework
        pushH3("Enhancement control framework:");
        var ecfH = [
            "Module", "Header\ntable", "Item\ntable", "Enhancement\nspot", "BADI\ndefinition",
            "BADI\nimplementation", "Interface", "Implementing\nclass", "Global\n/Local",
            "System\n(S4/EWM/TM/etc)", "Business Unit\n/Country", "Localization\nReason", "Target\n/Intermediate State"
        ];
        var ecfW = ecfH.map(function () { return Math.floor(PAGE_WIDTH / ecfH.length); });
        documentChildren.push(makeColumnTable(ecfH,
            r3ECF.length > 0
                ? r3ECF.map(function (ec) {
                    return {
                        "Module":                    safeStr(ec["Module"] || ec.module || ""),
                        "Header\ntable":             safeStr(ec["Header Table"] || ec["Header table"] || ec.headerTable || ""),
                        "Item\ntable":               safeStr(ec["Item Table"] || ec["Item table"] || ec.itemTable || ""),
                        "Enhancement\nspot":         safeStr(ec["Enhancement Spot"] || ec["Enhancement spot"] || ec.enhancementSpot || ""),
                        "BADI\ndefinition":          safeStr(ec["BADI Definition"] || ec["BADI definition"] || ec.badiDefinition || ""),
                        "BADI\nimplementation":      safeStr(ec["BADI Implementation"] || ec["BADI implementation"] || ec.badiImplementation || ""),
                        "Interface":                 safeStr(ec["Interface"] || ec.interface || ""),
                        "Implementing\nclass":       safeStr(ec["Implementing Class"] || ec["Implementing class"] || ec.implementingClass || ""),
                        "Global\n/Local":            safeStr(ec["Global/Local"] || ec.globalLocal || ""),
                        "System\n(S4/EWM/TM/etc)":   safeStr(ec["System (S4/EWM/TM/etc)"] || ec["System"] || ec.system || ""),
                        "Business Unit\n/Country":   safeStr(ec["Business Unit/Country"] || ec.businessUnitCountry || ""),
                        "Localization\nReason":      safeStr(ec["Localization Reason"] || ec.localizationReason || ""),
                        "Target\n/Intermediate State": safeStr(ec["Target/Intermediate State"] || ec.targetIntermediateState || "")
                    };
                })
                : emptyRow(ecfH),
            ecfW
        ));
        pushSpacer();

        // 13d ZAXP_STP_ENH_HDR
        pushH3("Table: ZAXP_STP_ENH_HDR");
        var zHDRcols = [
            "Purchasing\nOrganization", "Purchasing\nGroup", "Plant", "Storage\nLocation",
            "Company\nCode", "Country", "Free\nKey", "Enhancement Spot", "Name of a BADI", "Active", "Description", "Last\nChange Date"
        ];
        var zHDRW = zHDRcols.map(function () { return Math.floor(PAGE_WIDTH / zHDRcols.length); });

        documentChildren.push(makeColumnTable(zHDRcols,
            r3ENH_HDR.length > 0
                ? r3ENH_HDR.map(function (row) {
                    return {
                        "Purchasing\nOrganization": safeStr(row["Purchasing Organization"] || row.purchasingOrganization || ""),
                        "Purchasing\nGroup":        safeStr(row["Purchasing Group"] || row.purchasingGroup || ""),
                        "Plant":                    safeStr(row["Plant"] || row.plant || ""),
                        "Storage\nLocation":        safeStr(row["Storage Location"] || row.storageLocation || ""),
                        "Company\nCode":            safeStr(row["Company Code"] || row.companyCode || ""),
                        "Country":                  safeStr(row["Country"] || row.country || ""),
                        "Free\nKey":                safeStr(row["Free Key"] || row.freeKey || ""),
                        "Enhancement Spot":         safeStr(row["Enhancement Spot"] || row.enhancementSpot || ""),
                        "Name of a BADI":           safeStr(row["Name of a BADI"] || row.badiName || ""),
                        "Active":                   safeStr(row["Active"] || row.active || ""),
                        "Description":              safeStr(row["Description"] || row.description || ""),
                        "Last\nChange Date":        safeStr(row["Last Change Date"] || row.lastChangeDate || "")
                    };
                })
                : emptyRow(zHDRcols),
            zHDRW
        ));
        pushSpacer();

        // 13e ZAXP_STP_ENH_ITM
        pushH3("Table: ZAXP_STP_ENH_ITM");
        var zITMcols = [
            "Purchasing\nOrganization", "Purchasing\nGroup", "Plant", "Storage\nLocation",
            "Company\nCode", "Country", "Free\nKey", "Enhancement Spot", "Name of a BADI", "Reference type", "Active", "Description", "Last\nChange Date"
        ];
        var zITMW = zITMcols.map(function () { return Math.floor(PAGE_WIDTH / zITMcols.length); });

        documentChildren.push(makeColumnTable(zITMcols,
            r3ENH_ITM.length > 0
                ? r3ENH_ITM.map(function (row) {
                    return {
                        "Purchasing\nOrganization": safeStr(row["Purchasing Organization"] || row.purchasingOrganization || ""),
                        "Purchasing\nGroup":        safeStr(row["Purchasing Group"] || row.purchasingGroup || ""),
                        "Plant":                    safeStr(row["Plant"] || row.plant || ""),
                        "Storage\nLocation":        safeStr(row["Storage Location"] || row.storageLocation || ""),
                        "Company\nCode":            safeStr(row["Company Code"] || row.companyCode || ""),
                        "Country":                  safeStr(row["Country"] || row.country || ""),
                        "Free\nKey":                safeStr(row["Free Key"] || row.freeKey || ""),
                        "Enhancement Spot":         safeStr(row["Enhancement Spot"] || row.enhancementSpot || ""),
                        "Name of a BADI":           safeStr(row["Name of a BADI"] || row.badiName || ""),
                        "Reference type":           safeStr(row["Reference type"] || row["Reference Type"] || row.referenceType || ""),
                        "Active":                   safeStr(row["Active"] || row.active || ""),
                        "Description":              safeStr(row["Description"] || row.description || ""),
                        "Last\nChange Date":        safeStr(row["Last Change Date"] || row.lastChangeDate || "")
                    };
                })
                : emptyRow(zITMcols),
            zITMW
        ));
        pushSpacer();
    });

    // SECTION 14: BADI Implementation Details
    
    safeSection("BADI Implementation Details", function () {
        pushH2("BADI Implementation Details", "BADIImplementationDetails");

        var borders = makeBorders();
        var labelW = 1800;
        var badiCount = r3BADIImpl.length > 0 ? r3BADIImpl.length : 1;
        var valW = Math.floor((PAGE_WIDTH - labelW) / badiCount);

        function makeBADIRow(label, valueFn) {
            var labelCell = new TableCell({
                width: { size: labelW, type: WidthType.DXA },
                borders: borders,
                shading: { fill: "E5E7EB" },
                children: [new Paragraph({
                    children: [new TextRun({ text: label, bold: true, size: 18, font: "Helvetica" })]
                })]
            });
            var dataCells = (r3BADIImpl.length > 0 ? r3BADIImpl : [{}]).map(function (badi) {
                return new TableCell({
                    width: { size: valW, type: WidthType.DXA },
                    borders: borders,
                    children: [new Paragraph({
                        children: [new TextRun({ text: safeStr(valueFn(badi)), size: 18, font: "Helvetica" })]
                    })]
                });
            });
            return new TableRow({ children: [labelCell].concat(dataCells) });
        }

        documentChildren.push(new Table({
            layout: TableLayoutType.FIXED,
            width: { size: PAGE_WIDTH, type: WidthType.DXA },
            rows: [
                makeBADIRow("Name of\nEnhancement",      function (b) { return b["Name of Enhancement"] || b.nameOfEnhancement || ""; }),
                makeBADIRow("Enhancement\nType",          function (b) {
                    var et = b["Enhancement Type"] || b.enhancementType || "";
                    if (Array.isArray(et)) return et.join(", ");
                    return et;
                }),
                makeBADIRow("Released for\ncloud",        function (b) { return b["Released for Cloud"] || b["Released for cloud"] || b.releasedForCloud || ""; }),
                makeBADIRow("Logical\nDatabase",          function (b) { return b["Logical Database"] || b.logicalDatabase || ""; }),
                makeBADIRow("Called\nTransactions",       function (b) { return b["Called Transactions"] || b.calledTransactions || ""; }),
                makeBADIRow("Function\nGroup",            function (b) { return b["Function Group"] || b.functionGroup || ""; }),
                makeBADIRow("Function\nCalls",            function (b) { return b["Function Calls"] || b.functionCalls || ""; }),
                makeBADIRow("Fixed Point\nArithmetic",    function (b) { return b["Fixed Point Arithmetic"] || b.fixedPointArithmetic || ""; }),
                makeBADIRow("Attachment",                 function (b) { return b["Attachment"] || b.attachment || ""; })
            ]
        }));
        pushSpacer();

        // Technical mapping table: Enhancement Spot | BADI Definition | Implementing Class
        var t2H = ["Enhancement spot", "BADI definition", "Implementing class"];
        var t2W = [3120, 3120, 3120];
        documentChildren.push(makeColumnTable(t2H,
            r3BADIImpl.length > 0
                ? r3BADIImpl.map(function (b) {
                    return {
                        "Enhancement spot": safeStr(b["Enhancement Spot"] || b.enhancementSpot || ""),
                        "BADI definition":  safeStr(b["BADI Definition"] || b["BADI definition"] || b.badiDefinition || ""),
                        "Implementing class": safeStr(b["Implementing Class"] || b["Implementing class"] || b.implementingClass || "")
                    };
                })
                : emptyRow(t2H),
            t2W
        ));
        pushSpacer();

        // Implementation class details table
        pushText("Implementation class details:");
        var t3H = ["Method Name", "Description", "Type", "Details"];
        var t3W = [2500, 3000, 1360, 2500];
        var allMethods = [];
        r3BADIImpl.forEach(function (b) {
            safeArr(b["Implementation Class Details"] || b["Implementation class details"] || b.implementationClassDetails || []).forEach(function (m) {
                allMethods.push(m);
            });
        });
        documentChildren.push(makeColumnTable(t3H,
            allMethods.length > 0
                ? allMethods.map(function (m) {
                    return {
                        "Method Name": safeStr(m["Method Name"] || m["Implementing Method Name"] || m.methodName || ""),
                        "Description": safeStr(m["Description"] || m.description || ""),
                        "Type":        safeStr(m["Type"] || m.type || ""),
                        "Details":     safeStr(m["Details"] || m.details || "")
                    };
                })
                : emptyRow(t3H),
            t3W
        ));
        pushSpacer();
    });

    // SECTION 15: Open Points  (NEW — was completely missing)
   
    safeSection("Open Points", function () {
        pushH2("Open Points:", "OpenPoints");

        var opH = ["No.", "Open point details", "Action item with Status"];
        var opW = [600, 4880, 3880];
        documentChildren.push(makeColumnTable(opH,
            r3OpenPoints.length > 0
                ? r3OpenPoints.map(function (op, idx) {
                    return {
                        "No.":                  safeStr(op["No."] || op["No"] || op.number || (idx + 1)),
                        "Open point details":   safeStr(op["Open point details"] || op["Open Point Details"] || op.details || ""),
                        "Action item with Status": safeStr(op["Action item with Status"] || op["Action Item with Status"] || op.actionItem || "")
                    };
                })
                : emptyRow(opH),
            opW
        ));
        pushSpacer();
    });

    // SECTION 16: CleanCore Recommendations
    if (!this.cleancorerecommendation) {
        safeSection("CleanCore Recommendations", function () {
            that._appendCleanCoreRecommendations(documentChildren, response_cleancore, D);
            that.cleancorerecommendation = true;
        });
    }

    return documentChildren;
},

        _onTSDGeneratedSuccess: function () {
            this.byId("tsdNudgeBox").setVisible(true);

            this.byId("id_TSDStep1").setValidated(true);

            this.byId("btnTSD1").setVisible(false);
            this.byId("btnTSD2").setVisible(true);
        },

        createReportTSD: function (documentChildren, responseRaw, response_cleancore) {
            var D = window.docx;
            var { Paragraph, TextRun, Bookmark } = D;
            var that = this;

            var response = safeObj(safeArr(responseRaw && responseRaw.response3 && responseRaw.response3.Reports)[0]);

            documentChildren.push(new Paragraph({ children: [new Bookmark({ id: "ReportNameType", children: [new TextRun({ text: "Report Name & Type", bold: true, size: 28, font: "Helvetica" })] })], spacing: { before: 250, after: 250 } }));

            safeSection(D, documentChildren, "Report - Name & Type", function () {
                var rnt = safeObj(response["Report Name & Type"]);
                var rntRows = ["Report Type", "Logical Database", "External Form Calls", "Called Transactions", "Function Group", "Function Calls", "Fixed Point Arithmetic", "System (S4/EWM/TM/etc.)"].map(function (k) { return makeKVRow(D, k, rnt[k]); });
                documentChildren.push(new D.Table({ rows: rntRows }));
            });

            safeSection(D, documentChildren, "Report - Selection Screen", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Selection Screen", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ss = safeObj(response["Selection screen"]);
                documentChildren.push(new D.Table({
                    rows: [
                        makeKVRow(D, "Select Options / Parameters / Radio Buttons / Check Boxes", ss["Select Options / Parameters / Radio Buttons / Check Boxes"]),
                        makeKVRow(D, "Field Name", ss["Field Name"]),
                        makeKVRow(D, "Default Values (From\u2013To)", ss["Default Values (From\u2013To)"]),
                        makeKVRow(D, "Validation / Required / Optional", ss["Validation / Required / Optional"])
                    ]
                }));
            });

            safeSection(D, documentChildren, "Report - Data Selection", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Data Selection", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ds = safeObj(response["Data Selection"]);
                [["Special Processing", "Special Processing"], ["Variants, Variant Data", "Variants, Variant Data"], ["Screens, Screen Flow", "Screens, Screen Flow"], ["Pseudo-code", "Pseudo code"]].forEach(function (pair) {
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: pair[0], bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 100 } }));
                    documentChildren.push(new Paragraph({ children: [new TextRun({ text: ds[pair[1]] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
                });
            });

            safeSection(D, documentChildren, "Report - Security Information", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Security Information", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var si = safeObj(response["Security Information"]);
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Exception Handling", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 100 } }));
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: si["Exception Handling"] || "", size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Additional Information", bold: true, size: 28, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ai = safeObj(si["Additional Information"]);
                documentChildren.push(new D.Table({ rows: ["Application/Platform", "Object Name", "Object Type", "Object Description", "Step By Step Instructions", "System (S4/EWM/TM/etc)", "Global/Local", "Business Unit/Country", "Localization Reason"].map(function (k) { return makeKVRow(D, k, ai[k]); }) }));
            });

            safeSection(D, documentChildren, "Report - OData & Service Details", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "ODATA", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var sd = safeObj(response["ServiceDetails"]);
                documentChildren.push(new D.Table({ rows: [makeKVRow(D, "Package Name", sd["PackageName"]), makeKVRow(D, "Service Definition", sd["ServiceDefinition"]), makeKVRow(D, "Service Binding", sd["ServiceBinding"]), makeKVRow(D, "Service Name", sd["ServiceName"])] }));
                var obj = safeObj(response["Object"]);
                documentChildren.push(makeColumnTable(D, ["Object Type", "Relation", "CDS View", "SQL View", "Interface", "Consumption View", "Behavior Interface", "Behavior Definition", "Metadata Extension", "Access Control", "Class Name"],
                    [{ "Object Type": obj["ObjectType"] || "", "Relation": obj["Relation"] || "", "CDS View": obj["CDSView"] || "", "SQL View": obj["SQLView"] || "", "Interface": obj["Interface"] || "", "Consumption View": obj["ConsumptionView"] || "", "Behavior Interface": obj["BehaviorInterface"] || "", "Behavior Definition": obj["BehaviorDefinition"] || "", "Metadata Extension": obj["MetadataExtension"] || "", "Access Control": obj["AccessControl"] || "", "Class Name": obj["ClassName"] || "" }]
                ));
            });


            safeSection(D, documentChildren, "Report - Entity Details", function () {
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Entity 1", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Custom Table", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ct = safeArr(response["Entity 1"] && response["Entity 1"]["Custom Table"]);
                if (ct.length) documentChildren.push(makeColumnTable(D, ["Field", "Description", "Key", "Type"], ct));

                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Entity Details", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var ed = Array.isArray(response["Entity Details"]) ? response["Entity Details"] : (response["Entity Details"] ? [response["Entity Details"]] : []);
                if (ed.length) documentChildren.push(makeColumnTable(D, ["Properties", "Technical Details", "Description"], ed));

                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "Operations / Action", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 250 } }));
                var oa = Array.isArray(response["OperationsAction"]) ? response["OperationsAction"] : (response["OperationsAction"] ? [response["OperationsAction"]] : []);
                if (oa.length) documentChildren.push(makeColumnTable(D, ["Operation/Action", "Technical Details", "Description"], oa));

                documentChildren.push(new Paragraph({ children: [new TextRun({ text: "CDS View Technical Details & Joins", bold: true, size: 24, font: "Helvetica" })], spacing: { before: 250, after: 100 } }));
                var cdsJoin = (response["CDSViewTechnicalDetailsAndJoins"] && response["CDSViewTechnicalDetailsAndJoins"]["CDSViewTechnicalDetailsAndJoins"])
                    ? response["CDSViewTechnicalDetailsAndJoins"]["CDSViewTechnicalDetailsAndJoins"]
                    : (response["CDSViewTechnicalDetailsAndJoins"] || "");
                documentChildren.push(new Paragraph({ children: [new TextRun({ text: String(cdsJoin), size: 20, font: "Helvetica" })], spacing: { before: 100, after: 200 } }));
            });

            if (!this.cleancorerecommendation) {
                safeSection(D, documentChildren, "CleanCore Recommendations", function () {
                    that._appendCleanCoreRecommendations(documentChildren, response_cleancore, D);
                    that.cleancorerecommendation = true;
                });
            }
            return documentChildren;
        }
    };

    return TSDGenerator;
});