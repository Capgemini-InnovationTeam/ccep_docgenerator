/* global $:readonly */
/* eslint-disable @sap-ux/fiori-tools/sap-no-dom-insertion, @sap-ux/fiori-tools/sap-timeout-usage */
sap.ui.define([], function () {
	"use strict";

	return {

	
onGenerateFSDPress: function () {
    var selectedFiles = this._aSelectedFiles;

    // Validate selection
    if (!selectedFiles || selectedFiles.length === 0) {
        sap.m.MessageToast.show("Please select a file first.");
        return;
    }

    this.oBusyDialog.open();
    var that = this;
    var currentPath = window.location.pathname;
    var appModulePath = currentPath.substring(0, currentPath.lastIndexOf("/"));
    var sApiUrl = appModulePath + "/fsd_gen_api/drdtofsdconversion";

    // Helper: build a fresh FormData each attempt (FormData can only be sent once)
    function buildFormData() {
        var oFD = new FormData();
        for (var i = 0; i < selectedFiles.length; i++) {
            oFD.append("files", selectedFiles[i]);
        }
        return oFD;
    }

    // Helper: fetch a fresh CSRF token, then POST; retries up to maxRetries times on 403
    // Fetch CSRF token from app base path (the API endpoint only supports POST)
    var sTokenUrl = appModulePath + "/";
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
                    url: sApiUrl,
                    method: "POST",
                    headers: { "X-CSRF-Token": csrfToken || "" },
                    processData: false,
                    contentType: false,
                    timeout: 0,
                    data: buildFormData(),
                    success: function (response) {
    that.oBusyDialog.close();
    var stringData = (typeof response === "object") ? JSON.stringify(response) : response;
    that.response = stringData;
    that.responseapi1Global = stringData;

    // Toggle UI state FIRST, so it always happens even if document build fails
    if (typeof that._onFSDGeneratedSuccess === "function") {
        that._onFSDGeneratedSuccess();
    }

    // Wrap createFSD so a bad response shape can't kill the whole success handler
    try {
        if (typeof that.createFSD === "function") {
            that.createFSD(stringData);
        }
    } catch (e) {
        sap.m.MessageBox.error("FSD document could not be built: " + e.message);
    }

    sap.m.MessageToast.show("FSD Generated Successfully");
},
                    error: function (jqXHR, textStatus) {
                        // 403 = stale/missing CSRF token — re-fetch token and retry
                        if (jqXHR.status === 403 && attempt < maxRetries) {
                            setTimeout(function () { fetchTokenAndPost(attempt + 1, maxRetries); }, 1500);
                        } else {
                            that.oBusyDialog.close();
                            var sError = "FSD Generation failed (attempt " + attempt + "): " + textStatus + " (" + jqXHR.status + ")";
                            sap.m.MessageBox.error(sError);
                        }
                    }
                });
            }
        });
    }

    fetchTokenAndPost(1, 3); // up to 3 attempts
},


		// MAIN DOCUMENT BUILDER

			createFSD: function (responseData) {
			if (!this.fsdGenerated) {
				this.fsdGenerated = true;
				this.getView().byId("id_FSDPromptRow").setVisible(true);
				this.getView().byId("id_FSDPromptLabel").setVisible(true);
			}
			const {
				Document,
				Packer,
				Paragraph,
				TextRun,
			} = window.docx;

			var that = this;
			var documentChildrenFSD = [];
			const response = responseData;

			// ✅ Wrap entire function in try-catch for initial parsing
			try {
				var parsedRespone = JSON.parse(response);

				// ✅ Sanitize all null/undefined values recursively
				const sanitizeObj = (obj) => {
					if (Array.isArray(obj)) {
						obj.forEach((item, i) => {
							if (item === null || item === undefined) {
								obj[i] = "";
							} else if (typeof item === 'object') {
								sanitizeObj(item);
							}
						});
					} else if (typeof obj === 'object' && obj !== null) {
						Object.keys(obj).forEach(k => {
							if (obj[k] === null || obj[k] === undefined) {
								obj[k] = "";
							} else if (typeof obj[k] === 'object') {
								sanitizeObj(obj[k]);
							}
						});
					}
					return obj;
				};
				sanitizeObj(parsedRespone);

				
				var normalized = parsedRespone;
				if (!normalized.response_set_1) {
					// Try one level deeper for each key until we find response_set_1
					var keys = Object.keys(normalized);
					for (var ki = 0; ki < keys.length; ki++) {
						var candidate = normalized[keys[ki]];
						if (candidate && typeof candidate === "object" && candidate.response_set_1) {
							normalized = candidate;
							break;
						}
					}
				}

				// ✅ If response_set_1 still missing, create empty structure instead of failing
				if (!normalized.response_set_1) {
					normalized.response_set_1 = { title: "Untitled FSD", brief_description: "", solution: "" };
				}
				if (!normalized.response_set_2) {
					normalized.response_set_2 = {};
				}
				if (!normalized.response) {
					normalized.response = {};
				}

				// From here on, use normalized (correctly shaped) object
				// Re-stringify so all downstream functions (which call JSON.parse internally) keep working
				var normalizedString = JSON.stringify(normalized);

				var title = normalized.response_set_1.title || "Untitled FSD";

				var parsedData = normalized;
				// Re-sanitize parsedData as well
				sanitizeObj(parsedData);

				// parsedData.response holds the doc-type section (Enhancement/Forms/Workflow/etc.)
				// Guard against it being undefined after unwrapping
				var docTypeResponse = parsedData.response || {};

				//Promis
				var sectionPromises = [];

				// Section: Heading
				sectionPromises.push(new Promise(function (resolve) {
					try {
						var sectionChildren = [];
						sectionChildren = that.createFSDHeading(normalizedString, sectionChildren);
						resolve(sectionChildren);
					} catch {
						resolve([]);
					}
				}));

				// Section: TOC
				sectionPromises.push(new Promise(function (resolve) {
					try {
						var sectionChildren = [];
						sectionChildren = that.createFSDTOC(normalizedString, sectionChildren);
						resolve(sectionChildren);
					} catch {
						resolve([]);
					}
				}));

				// Section: Common Contents
				sectionPromises.push(new Promise(function (resolve) {
					try {
						var sectionChildren = [];
						sectionChildren = that.createFSDCommonContents(normalizedString, sectionChildren);
						resolve(sectionChildren);
					} catch {
						resolve([]);
					}
				}));

				// Section: Fiori App (conditional)
				if (docTypeResponse["fiori app"]) {
					sectionPromises.push(new Promise(function (resolve) {
						try {
							var sectionChildren = [];
							sectionChildren = that.createFSDFiori(normalizedString, sectionChildren);
							resolve(sectionChildren);
						} catch {
							resolve([]);
						}
					}));
				}

				// Section: Report / Embedded (conditional)
				if (docTypeResponse["embedded"] || docTypeResponse["report"]) {
					sectionPromises.push(new Promise(function (resolve) {
						try {
							var sectionChildren = [];
							sectionChildren = that.createFSDReport(normalizedString, sectionChildren);
							resolve(sectionChildren);
						} catch {
							resolve([]);
						}
					}));
				}

				// Section: Workflow (conditional)
				if (docTypeResponse["Workflow"]) {
					sectionPromises.push(new Promise(function (resolve) {
						try {
							var sectionChildren = [];
							sectionChildren = that.createFSDWorkflow(normalizedString, sectionChildren);
							resolve(sectionChildren);
						} catch {
							resolve([]);
						}
					}));
				}

				// Section: Enhancement (conditional)
				if (docTypeResponse["Enhancement"]) {
					sectionPromises.push(new Promise(function (resolve) {
						try {
							var sectionChildren = [];
							sectionChildren = that.createFSDEnhancement(normalizedString, sectionChildren);
							resolve(sectionChildren);
						} catch {
							resolve([]);
						}
					}));
				}

				// Section: Form (conditional)
				if (docTypeResponse["Forms"]) {
					sectionPromises.push(new Promise(function (resolve) {
						try {
							var sectionChildren = [];
							sectionChildren = that.createFSDForm(normalizedString, sectionChildren);
							resolve(sectionChildren);
						} catch {
							resolve([]);
						}
					}));
				}

				// Section: Last Contents
				sectionPromises.push(new Promise(function (resolve) {
					try {
						var sectionChildren = [];
						sectionChildren = that.createFSDLastContents(normalizedString, sectionChildren);
						resolve(sectionChildren);
					} catch {
						resolve([]);
					}
				}));

				// ✅ Use Promise.allSettled to wait for all sections, ignoring any failures
				Promise.allSettled(sectionPromises).then(function (results) {
					// Collect all successful section children into documentChildrenFSD
					results.forEach(function (result) {
						if (result.status === "fulfilled" && Array.isArray(result.value) && result.value.length > 0) {
							documentChildrenFSD = documentChildrenFSD.concat(result.value);
						}
					});

					// If no content was generated at all, add a placeholder paragraph
					if (documentChildrenFSD.length === 0) {
						documentChildrenFSD.push(new Paragraph({
							children: [new TextRun({
								text: "FSD Document - Some sections could not be generated due to missing data in the response.",
								size: 8 * 2,
								font: "Helvetica"
							})]
						}));
					}

					// Build the DOCX document
					var doc = new Document({
						sections: [{
							children: documentChildrenFSD
						}]
					});

					Packer.toBlob(doc)
						.then(function (blob) {
							if (window.saveAs) {
								if (title) {
									var sanitizedTitle = title.replace(/[^a-zA-Z0-9\s]/g, '');
									var maxLength = 60;
									var title1 = sanitizedTitle.length > maxLength
										? sanitizedTitle.slice(0, maxLength - 3).trim() + "..."
										: sanitizedTitle;
									var Title = "FSD - " + title1;
									window.saveAs(blob, Title + ".docx");
								} else {
									window.saveAs(blob, "FSD.docx");
								}
							} else {
								sap.m.MessageBox.error("saveAs function is not available.");
							}
						})
						.catch(function (err) {
							sap.m.MessageBox.error("Document generation failed: " + err.message);
						});
				}).catch(function (err) {
					sap.m.MessageBox.error("Document generation failed unexpectedly: " + err.message);
				});

			} catch (e) {
				sap.m.MessageBox.error("Error creating FSD: " + e.message);
			}
		},

		// TABLE OF CONTENTS

		createFSDTOC: function (response, documentChildrenFSD) {
			const { Paragraph, TextRun } = window.docx;

			const tocTextRunsCommon = [
				new TextRun({ text: "Table of Contents", size: 12 * 2, font: "Helvetica", bold: true, color: "0000FF" }),
				new TextRun({ text: " ● Brief Description", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Approved Versions, Change Control", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Principles, Policies, Standards & Guidelines", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Key Design Decisions", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Open Design Topics", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Requirements", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Assumptions", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Dependencies", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Out Of Scope", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Abbreviations", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Business Process Context", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Processes Group, Processes, Process Integration", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Integrated Business Scenarios", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Architectural Context", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Current, Target & Intermediate States", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Satellite Applications & Integration", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " ● Further Considerations", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Site, Country & Business Unit Localizations", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Organization", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Data, Master Data", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Security, Internal Controls", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Audit, Compliance & Legal", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Environment, Social & Governance (ESG)", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Country Specifics, Language", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Data Migration & Cutover", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Non-functional Considerations", size: 8 * 2, font: "Helvetica", color: "0000FF" }),
				new TextRun({ text: " \t○ Other Considerations", size: 8 * 2, font: "Helvetica", color: "0000FF" })
			];

			const tocParagraphsCommon = tocTextRunsCommon.map((textRun) =>
				new Paragraph({ children: [textRun], spacing: { after: 100 } })
			);
			documentChildrenFSD.push(...tocParagraphsCommon);

			let parsedData;
			try {
				parsedData = JSON.parse(response);
			} catch {
				return documentChildrenFSD;
			}

			if (parsedData && parsedData.response && parsedData.response["fiori app"] &&
				Object.keys(parsedData.response["fiori app"]).length > 0) {
				const tocTextRunsFioriApp = [
					new TextRun({ text: " • Specification Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Functional Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t\t▪ Report Selection fields", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " • Technical Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Security & Role Requirements, Sensitive Data", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Data Volumes & Performance Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Synchronous Processing, Response Times", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Background Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Test Scenarios, Test Data, Pre-conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Samples", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " • Additional Information", size: 16, font: "Helvetica", color: "0000FF" })
				];
				documentChildrenFSD.push(...tocTextRunsFioriApp.map((tr) =>
					new Paragraph({ children: [tr], spacing: { after: 100 } })
				));
			}

			if (parsedData && parsedData.response && parsedData.response["Workflow"] &&
				Object.keys(parsedData.response["Workflow"]).length > 0) {
				const tocTextRunsWorkflow = [
					new TextRun({ text: "● Workflow Specification Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t○ Functional Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Triggers, Start Conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Data Selection, Validation & Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Dialog Work Items, Agent Determination", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Delegation, Substitution", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Deadline Monitoring, Escalation Rules & Reminders", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Work Item Termination", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Expected Outcomes", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Notifications & Error Handling", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Further Constraints & Special Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Security & Role Requirements, Sensitive Data", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Data Volumes & Performance Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Synchronous Processing, Response Times", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Background Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Test Scenarios, Test Data, Pre-conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " ○ Sample Workflows", size: 16, font: "Helvetica", color: "0000FF" })
				];
				documentChildrenFSD.push(...tocTextRunsWorkflow.map((tr) =>
					new Paragraph({ children: [tr], spacing: { after: 100 } })
				));
			}

			if (parsedData && parsedData.response && parsedData.response["Forms"] &&
				Object.keys(parsedData.response["Forms"]).length > 0) {
				const tocTextRunsForm = [
					new TextRun({ text: "● Form Specification Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Functional Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t\t▪ Output Content", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Triggers", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Data Selection, Validation & Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Expected Output, Layout", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Hardcopy, Stationary & Printer Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Notifications & Error Handling", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Further Constraints & Special Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Security & Role Requirements, Sensitive Data", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Data Volumes & Performance Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Synchronous Processing, Response Times", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Background Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Test Data, Pre-conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Sample Forms", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "● Additional Information", size: 16, font: "Helvetica", color: "0000FF" })
				];
				documentChildrenFSD.push(...tocTextRunsForm.map((tr) =>
					new Paragraph({ children: [tr], spacing: { after: 100 } })
				));
			}

			if (parsedData && parsedData.response && parsedData.response["Enhancement"] &&
				Object.keys(parsedData.response["Enhancement"]).length > 0) {
				const tocTextRunsEnhancement = [
					new TextRun({ text: "● Enhancement Specification Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Functional Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Triggers", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Data Selection, Validation & Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Expected Outcome", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Notifications & Error Handling", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Further Constraints & Special Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Security & Role Requirements, Sensitive Data", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Data Volumes & Performance Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Synchronous Processing, Response Times", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Background Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Test Scenarios, Test Data, Pre-conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: " \t○ Sample Enhancement", size: 16, font: "Helvetica", color: "0000FF" })
				];
				documentChildrenFSD.push(...tocTextRunsEnhancement.map((tr) =>
					new Paragraph({ children: [tr], spacing: { after: 100 } })
				));
			}

			if (
				(parsedData && parsedData.response && parsedData.response["embedded"] && Object.keys(parsedData.response["embedded"]).length > 0) ||
				(parsedData && parsedData.response && parsedData.response["report"] && Object.keys(parsedData.response["report"]).length > 0)
			) {
				const tocTextRunsEmbedded = [
					new TextRun({ text: "●  Embedded Analytics Specification Details", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Functional Requirements", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Background", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Report Selections", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Expected Output & Layout", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Drill-down and Follow-on Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Security & Role Requirements, Sensitive Data", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Data Volumes & Performance Considerations", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Synchronous Processing, Response Times", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Background Processing", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Test Scenarios, Test Data, Pre-conditions", size: 16, font: "Helvetica", color: "0000FF" }),
					new TextRun({ text: "\t● Samples", size: 16, font: "Helvetica", color: "0000FF" })
				];
				documentChildrenFSD.push(...tocTextRunsEmbedded.map((tr) =>
					new Paragraph({ children: [tr], spacing: { after: 100 } })
				));
			}

			documentChildrenFSD.push(new Paragraph({ children: [], pageBreakBefore: true }));
			return documentChildrenFSD;
		},

		// HEADING
	createFSDHeading: function (response, documentChildrenFSD) {
			const {
				Paragraph,
				TextRun,
				Table,
				TableRow,
				TableCell,
				AlignmentType,
				VerticalAlign,
			} = window.docx;

			// ✅ SAFE PARSE with sanitizer
			var parsedData2 = response;
			var parsedData = JSON.parse(parsedData2);

			// ✅ Helper function - safe string value
			const safeStr = (val) => {
				if (val === null || val === undefined) return "";
				if (typeof val === 'object') return JSON.stringify(val);
				return String(val);
			};

			var TITLE = safeStr(parsedData.response_set_1.title);

			var FSDTitle = new Paragraph({
				children: [
					new TextRun({
						text: "FSD - " + TITLE,
						bold: true,
						size: 16 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					before: 0,
					after: 250
				}
			});

			var today = new Date();
			var options = {
				year: 'numeric',
				month: 'long',
				day: 'numeric'
			};
			var currentDate = today.toLocaleDateString("en-US", options);

			var Properties = new Paragraph({
				children: [
					new TextRun({
						text: "Properties",
						bold: true,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 250
				}
			});

			var Properties_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Deliverable ID", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "<Deliverable ID>", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Document Type", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Functional Specification Document (FSD)", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Document Purpose", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "To support the facilitation of the fit-to-standard workshops which will be used to validate the reference process against the demonstrator solution and capture the resulting delta requirements", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Document Status", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Created By", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: " GenAI", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Created On", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: currentDate, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Contributors", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: " GenAI", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
				]
			});

			documentChildrenFSD.push(FSDTitle);
			documentChildrenFSD.push(Properties);
			documentChildrenFSD.push(Properties_Table);

			var ReferenceItems = new Paragraph({
				children: [
					new TextRun({
						text: "Reference Items",
						bold: true,
						font: "Helvetica",
						size: 8 * 2
					})
				],
				spacing: { after: 250, before: 250 }
			});

			// ✅ FIX: Safe access for all reference items
			var refItems = parsedData.response_set_2 && parsedData.response_set_2.reference_items
				? parsedData.response_set_2.reference_items
				: {};

			var IncomingReferences = safeStr(refItems["Incoming References"]);
			var ReferenceApplications_ITComponents = safeStr(refItems["Reference Applications & IT Components (LeanIX)"]);
			var ReferenceBusinessCapabilities = safeStr(refItems["Reference Business Capabilities (LeanIX)"]);
			var ReferenceDataObjects = safeStr(refItems["Reference Data Objects (Precisely)"]);
			var ReferenceIntegratedBusinessScenarios = safeStr(refItems["Reference Integrated Business Scenarios (Signavio)"]);
			var ReferenceProcess = safeStr(refItems["Reference Process (Signavio)"]);
			var ReferenceSolutionDesignDocument = safeStr(refItems["Reference Solution Design Document"]);

			var ReferenceItems_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Process (Signavio)", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 17, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceProcess, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Integrate Business Scenarios (Signavio)", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceIntegratedBusinessScenarios, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Business Capabilities(Lean IX)", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceBusinessCapabilities, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Applications & IT Components (Lean IX)", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceApplications_ITComponents, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Data Objects (Precisely)", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceDataObjects, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Incoming References", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: IncomingReferences, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Reference Solution Design Document", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: ReferenceSolutionDesignDocument, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					})
				]
			});

			documentChildrenFSD.push(ReferenceItems);
			documentChildrenFSD.push(ReferenceItems_Table);

			var RelatedItems = new Paragraph({
				children: [
					new TextRun({
						text: "Related Items",
						bold: true,
						size: 8 * 2,
						font: "Helvetica",
					})
				],
				spacing: { after: 250, before: 250 }
			});

			// ✅ FIX: Safe access for ALL related items - this was the main bug
			var relatedItemsObj = parsedData.response_set_2 && parsedData.response_set_2.related_items
				? parsedData.response_set_2.related_items
				: {};

			var DependentDocuments = safeStr(relatedItemsObj["Dependent Documents"]);
			var OtherRelatedDocuments = safeStr(relatedItemsObj["Other Related Documents"]);
			var RelatedJIRAIssues = safeStr(relatedItemsObj["Related JIRA Issues"]);
			// var RelatedKeyDesignDecisions = safeStr(relatedItemsObj["Related Key Design Decisions"]);
			var RelatedRequirements = safeStr(relatedItemsObj["Related Requirements"]);

			var RelatedItems_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Dependent Documents", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 8, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: DependentDocuments, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Other Related Documents", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: OtherRelatedDocuments, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Related Requirements", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: RelatedRequirements, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: "Related JIRA Issues", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 2, type: "pct" },
								shading: { fill: "#e5e7eb" },
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: RelatedJIRAIssues, style: "Hyperlink", font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })],
								width: { size: 700, type: "pct" },
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: { value: 500, rule: "atLeast" }
					})
				]
			});

			documentChildrenFSD.push(RelatedItems);
			documentChildrenFSD.push(RelatedItems_Table);

			var Classification = new Paragraph({
				children: [
					new TextRun({
						text: "Classification",
						bold: true,
						size: 8 * 2,
						font: "Helvetica",
					})
				],
				spacing: { after: 250, before: 250 }
			});
			var classObj = parsedData.response_set_2 && parsedData.response_set_2.classification
				? parsedData.response_set_2.classification
				: {};

			var ClassificationBusinessUnitsCountries = safeStr(classObj["Business Units/Countries"]);
			var ClassificationCategory = safeStr(classObj["Category"]);
			var ClassificationComplexity = safeStr(classObj["Complexity"]);
			var ClassificationGlobalLocal = safeStr(classObj["Global/Local"]);
			var ClassificationGroup = safeStr(classObj["Group"]);
			var ClassificationLocalizationReason = safeStr(classObj["Localization Reason"]);
			var ClassificationRelease = safeStr(classObj["Release"]);
			var ClassificationSubGroup = safeStr(classObj["Sub Group"]);
			var ClassificationTargetIntermediateState = safeStr(classObj["Target/Intermediate State"]);
			var ClassificationType = safeStr(classObj["Type"]);

			var Classification_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Category", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationCategory, font: "Helvetica", color: "FFFFFF", bold: true, size: 8 * 2, shading: { fill: "FF0000" } })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Type", font: "Helvetica", bold: true, size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationType, font: "Helvetica", bold: true, size: 8 * 2, shading: { fill: "#FFA500" } })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Global/Local", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationGlobalLocal, font: "Helvetica", color: "FFFFFF", bold: true, size: 8 * 2, shading: { fill: "#008000" } })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Business Units Countries", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationBusinessUnitsCountries, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Localization Reason", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationLocalizationReason, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Target/Intermediate State", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationTargetIntermediateState, font: "Helvetica", color: "FFFFFF", bold: true, size: 8 * 2, shading: { fill: "#008000" } })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Complexity", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationComplexity, font: "Helvetica", bold: true, size: 8 * 2, shading: { fill: "#FFA500" } })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Release", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationRelease, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Group", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationGroup, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					}),
					new TableRow({
						children: [
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Sub Group", bold: true, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 2, type: "pct" }, shading: { fill: "#e5e7eb" }, verticalAlign: VerticalAlign.CENTER }),
							new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ClassificationSubGroup, font: "Helvetica", size: 8 * 2 })], alignment: AlignmentType.LEFT, indent: { left: 100 } })], width: { size: 700, type: "pct" }, verticalAlign: VerticalAlign.CENTER })
						],
						height: { value: 500, rule: "atLeast" }
					})
				]
			});

			documentChildrenFSD.push(Classification);
			documentChildrenFSD.push(Classification_Table);

			var pageBreak = new Paragraph({
				children: [],
				pageBreakBefore: true
			});
			documentChildrenFSD.push(pageBreak);

			return documentChildrenFSD;
		},


		createFSDCommonContents: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
				Table,
				ExternalHyperlink,
				TableRow,
				TableCell,
				AlignmentType,
				VerticalAlign,
			} = window.docx;
			var parsedData2 = response;

			var parsedData = JSON.parse(parsedData2);

			// Guard: any field the follow-up API omits will be undefined — treat as empty string
			const safeStr = (val) => {
				if (val === null || val === undefined) return "";
				if (typeof val === "object") return JSON.stringify(val);
				return String(val);
			};
			// Guard: any array field that may be missing — treat as empty array
			const safeArr = (val) => {
				if (Array.isArray(val)) return val;
				if (val === null || val === undefined || val === "") return [];
				return [val];
			};

			var BriefDescription1 = new Paragraph({
				children: [
					new Bookmark({
						id: "brief_description", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Brief Description",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					after: 100
				}
			});
			var brief_description1 = safeStr(parsedData.response_set_1.brief_description);
		

			var BriefDescription = brief_description1
				.split(/(?<=\.)\s+/)
				.map(line => line.replace(/\d+/g, '').trim())
				.filter(line => line.length > 0)
				.map(line => line.replace(/-\s*$/, "").trim());

			var BriefDescriptionParagraphs = BriefDescription.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100
				},
				alignment: AlignmentType.JUSTIFIED
			}));

			



			documentChildrenFSD.push(BriefDescription1);
			documentChildrenFSD.push(...BriefDescriptionParagraphs);

			var TITLE = safeStr(parsedData.response_set_1.title);

			var BriefDescriprion_Table_Solution = safeStr(parsedData.response_set_1.solution);
			var BriefDescriprion_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "DRD",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
												// Make the text bold
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 6,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),

							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Solution",
												size: 8 * 2,
												font: "Helvetica",
												bold: true
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: TITLE,
												size: 8 * 2,
												font: "Helvetica",
												bold: true, // Make the text bold
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 6,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: BriefDescriprion_Table_Solution,
												size: 8 * 2,
												font: "Helvetica"
												// bold: true 
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),

				],
				spacing: {
					before: 350,
					after: 200

				}
			});

			documentChildrenFSD.push(BriefDescriprion_Table);

			var ApprovedVersionsChangeControl = new Paragraph({
				children: [
					new Bookmark({
						id: "approved_versions", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Approved Versions, Change Control",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var approvedVersion = safeStr(parsedData.response_set_1.approved_version_change_control && parsedData.response_set_1.approved_version_change_control["Approved Version"]);
			var descriptionReasonForChange = safeStr(parsedData.response_set_1.approved_version_change_control && parsedData.response_set_1.approved_version_change_control["Description, Reason for Change"]);
			var additionalInformation = safeStr(parsedData.response_set_1.approved_version_change_control && parsedData.response_set_1.approved_version_change_control["Additional Information"]);

			var ApprovedVersionsChangeControl_Table = new Table({
				rows: [
					// Header row
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Approved Version",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 15,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Status",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description, Reason for Change",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 30,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 35,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					// Data row
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: approvedVersion,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 15,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												// text: status,
												text: "Initial Draft",

												color: "228B22",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: descriptionReasonForChange,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 30,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: additionalInformation,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 35,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});
			documentChildrenFSD.push(ApprovedVersionsChangeControl);
			documentChildrenFSD.push(ApprovedVersionsChangeControl_Table);
			var PrinciplesPoliciesStandardsGuidelines = new Paragraph({
				children: [
					new Bookmark({
						id: "principles_policies", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Principles, Policies, Standards & Guidelines",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});
			var paragraph12 = new Paragraph({
				children: [
					new TextRun({
						text: "Please refer to the following principles, policies, standards & guidelines when completing this enhancement functional specification",
						// bold: true,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});
			var PrinciplesPoliciesStandardsGuidelines_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Documents",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 12,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},

								verticalAlign: VerticalAlign.CENTER
							}),

							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),

						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new ExternalHyperlink({
												// eslint-disable-next-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint/Signavio reference documentation
												link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-principles", // Your hyperlink URL
												children: [
													new TextRun({
														text: "CCEP Architecture Design Principles", // The text that will be clickable
														style: "Hyperlink",
														size: 8 * 2,
														font: "Helvetica"
													})
												]
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Enterprise, Data & Integration Architecture Principles",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new ExternalHyperlink({
												// eslint-disable-next-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint/Signavio reference documentation
												link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-policies-and-standards", // Your hyperlink URL
												children: [
													new TextRun({
														text: "CCEP Enterprise Architecture Policies & Standards",
														style: "Hyperlink",
														size: 8 * 2,
														font: "Helvetica"
													})
												]
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Data management, Integration & Non-platform Application, Infrastructure & Security Policies & Standards",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new ExternalHyperlink({
												// eslint-disable-next-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint/Signavio reference documentation
												link: "https://ccep.sharepoint.com/sites/comms39_EnterpriseArchitectureManagement/SitePages/Governance.aspx#enterprise-architecture-patterns", // Your hyperlink URL
												children: [
													new TextRun({
														text: "CCEP Enterprise Architecture Patterns",
														style: "Hyperlink",
														size: 8 * 2,
														font: "Helvetica"
													})
												]
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Integration Architecture Patterns",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 700,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "",
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 2,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
				]
			});
			documentChildrenFSD.push(PrinciplesPoliciesStandardsGuidelines);
			documentChildrenFSD.push(paragraph12);
			documentChildrenFSD.push(PrinciplesPoliciesStandardsGuidelines_Table);

			var KeyDesignDecisions = new Paragraph({
				children: [
					new Bookmark({
						id: "key_design_decisions", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Key Design Decisions",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"

							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});

			var keyDesignDecisions = safeArr(parsedData.response_set_1.key_design_decisions);

			var kddkeyDesignDecision_0 = keyDesignDecisions[0]["Key Design Decision"];
			var kdddescription_0 = keyDesignDecisions[0].Description;
			var kddadditionalInfo_0 = keyDesignDecisions[0]["Additional Information"];

			var KeyDesignDecisions_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Key Design Decision",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},

								verticalAlign: VerticalAlign.CENTER
							}),

							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Status",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 15,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),

						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: kddkeyDesignDecision_0,
												// bold: true, // Make the text bold
												size: 8 * 2,
												font: "Helvetica"
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 10,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												// text: kddstatus_0,
												text: "Initial Draft",
												color: "228B22", // Set the text color to forest green

												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: kdddescription_0,
												// bold: true, // Make the text bold
												size: 8 * 2,
												font: "Helvetica"
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 10,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: kddadditionalInfo_0,
												// bold: true 
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),

				]
			});
			documentChildrenFSD.push(KeyDesignDecisions);
			documentChildrenFSD.push(KeyDesignDecisions_Table);

			var openDesignTopics = safeArr(parsedData.response_set_1.open_design_topics);

			if (openDesignTopics && openDesignTopics.length > 0) {

				if (openDesignTopics[0] && openDesignTopics[0]["Design Topic"]) {

					var odtdesignTopic_0 = openDesignTopics[0]["Design Topic"];
					var odtdescription_0 = openDesignTopics[0].Description;
					var odtkeyDesignDecision_0 = openDesignTopics[0]["Key Design Decision"];
					var odtresolved_0 = openDesignTopics[0].Resolved;
					var odtresolution_0 = openDesignTopics[0].Resolution;
					var odtadditionalInfo_0 = openDesignTopics[0]["Additional Information"];

					var OpenDesignTopics = new Paragraph({
						children: [
							new Bookmark({
								id: "opendesigntopics",
								children: [
									new TextRun({
										text: "Open Design Topics",
										bold: true,
										size: 16 * 2,
										font: "Helvetica"
									})
								]
							})
						],
						spacing: {
							before: 250,
							after: 100
						}
					});

					var OpenDesignTopics_Table = new Table({
						rows: [

							new TableRow({
								children: [
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Design Topic",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100 // Very small left padding
												}
											})
										],
										width: {
											size: 15,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Description",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 20,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Key Design Decision",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Received?",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Resolution",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: "Additional Information",
														bold: true,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 15,
											type: "pct"
										},
										shading: {
											fill: "#e5e7eb"
										},
										verticalAlign: VerticalAlign.CENTER
									})
								],
								height: {
									value: 500,
									rule: "atLeast"
								}
							}),
							// Table data row
							new TableRow({
								children: [
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtdesignTopic_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtdescription_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtkeyDesignDecision_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtresolved_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtresolution_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									}),
									new TableCell({
										children: [
											new Paragraph({
												children: [
													new TextRun({
														text: odtadditionalInfo_0,
														size: 8 * 2,
														font: "Helvetica"
													})
												],
												alignment: AlignmentType.LEFT,
												indent: {
													left: 100
												}
											})
										],
										width: {
											size: 10,
											type: "pct"
										},
										verticalAlign: VerticalAlign.CENTER
									})
								],
								height: {
									value: 500,
									rule: "atLeast"
								}
							})
						]
					});
					documentChildrenFSD.push(OpenDesignTopics);

					documentChildrenFSD.push(OpenDesignTopics_Table);
				} else {
					// No open design topics to render for this response
				}
			} else {
				// No open design topics data available for this response
			}

			var Requirements = new Paragraph({
				children: [
					new Bookmark({
						id: "requirements",
						children: [
							new TextRun({
								text: "Requirements",
								bold: true,
								size: 16 * 2,

								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});

			var requirementsData = safeArr(parsedData.response_set_1.requirements);

			var reqrequirement_0 = requirementsData[0]["Requirement"];
			var reqstatus_0 = requirementsData[0].Status;
			var reqdescription_0 = requirementsData[0].Description;
			var reqadditionalInfo_0 = requirementsData[0]["Additional Information"];
			var Requirements_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Requirement",
												bold: true,
												font: "Helvetica",
												size: 8 * 2
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Status",
												bold: true,
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description",
												bold: true,
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),

						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: reqrequirement_0,
												font: "Helvetica",
												size: 8 * 2
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 10,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: reqstatus_0,
												color: "228B22",
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: reqdescription_0,
												font: "Helvetica",
												size: 8 * 2
												// bold: true, // Make the text bold
											})
										],

										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}

									})
								],
								width: {
									size: 10,
									type: "pct"
								},

								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: reqadditionalInfo_0,
												font: "Helvetica",
												size: 8 * 2
												// bold: true 
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),

						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),

				]
			});
			documentChildrenFSD.push(Requirements);
			documentChildrenFSD.push(Requirements_Table);

			var Assumptions = new Paragraph({
				children: [
					new Bookmark({
						id: "assumptions",
						children: [
							new TextRun({
								text: "Assumptions",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});
			documentChildrenFSD.push(Assumptions);
			const assumptionsData = safeArr(parsedData.response_set_1.assumptions);
			const normalizedAssumptionsData = Array.isArray(assumptionsData)
				? assumptionsData
				: [assumptionsData];
			const objectDetails_Assumptions = normalizedAssumptionsData.map(assumption => ({
				"Assumptions": assumption["Assumptions"] || "N/A",
				"Description": assumption["Description"] || "N/A",
				"Additional Information": assumption["Additional Information"] || "N/A"
			}));
			const columnTitles_Assumptions = [
				"Assumptions",
				"Description",
				"Additional Information"
			];
			const headerRow_Assumptions = new TableRow({
				children: columnTitles_Assumptions.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 16
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_Assumptions = objectDetails_Assumptions.map(field => {
				return new TableRow({
					children: columnTitles_Assumptions.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "N/A", // Handle missing values gracefully
											size: 8 * 2, // Font size: 16
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const AssumptionsTable = new Table({
				rows: [headerRow_Assumptions, ...dataRows_Assumptions]
			});
			documentChildrenFSD.push(AssumptionsTable);



			var Dependencies = new Paragraph({
				children: [
					new Bookmark({
						id: "dependencies",
						children: [
							new TextRun({
								text: "Dependencies",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});
			var dependencies = safeArr(parsedData.response_set_2 && parsedData.response_set_2.dependencies);
			var tableRows = [];
			tableRows.push(
				new TableRow({
					children: [
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "E2E Process Area/ Workstream",
											bold: true,
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100 // Very small left padding
									}
								})
							],
							width: {
								size: 20,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Status",
											bold: true,
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 10,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Description",
											bold: true,
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 40,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Outcome",
											bold: true,
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 20,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Additional Information",
											bold: true,
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 10,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						})
					],
					height: {
						value: 500,
						rule: "atLeast"
					}
				})
			);
			for (var i = 0; i < dependencies.length; i++) {
				tableRows.push(
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: dependencies[i]["E2E Process Area / Workstream"], // Dynamically set text
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							// new TableCell({
							// 	children: [
							// 		new Paragraph({
							// 			children: [
							// 				new TextRun({
							// 					text: dependencies[i].Status, // Dynamically set status
							// 					size: 8 * 2,
							// 					font: "Helvetica"
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: dependencies[i].Status, // Dynamically set status
												size: 8 * 2,
												font: "Helvetica",
												shading: {
														fill: dependencies[i].Status === "OPEN" ? "FFA500" : undefined // Orange shading for OPEN status, undefined for others
													},
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER,
								// shading: dependencies[i].Status === "OPEN" ? {
								// 	fill: "FFA500", // Orange shading
								// 	// type: ShadingType.SOLID
								
							}),
							
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: dependencies[i].Description, // Dynamically set description
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: dependencies[i].Outcome, // Dynamically set outcome
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: dependencies[i].Additional_Information, // Dynamically set additional info
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 10,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				);
			}
			var Dependencies_Table = new Table({
				rows: tableRows
			});
			documentChildrenFSD.push(Dependencies);
			documentChildrenFSD.push(Dependencies_Table);


			var OutOfScope = new Paragraph({
				children: [
					new Bookmark({
						id: "out_of_scope", // Bookmark ID
						children: [
							new TextRun({
								text: "Out of Scope Items",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			documentChildrenFSD.push(OutOfScope);
			var outOfScopeData = safeArr(parsedData.response_set_1.outofscope);
			const normalizedOutOfScopeData = Array.isArray(outOfScopeData)
				? outOfScopeData
				: [outOfScopeData];
			const objectDetails_OutOfScope = normalizedOutOfScopeData.map(item => ({
				"Scope Item": item["Scope Item"] || "N/A",
				"Description": item["Description"] || "N/A",
				"Additional Information": item["Additional Information"] || "N/A"
			}));
			const columnTitles_OutOfScope = [
				"Scope Item",
				"Description",
				"Additional Information"
			];
			const headerRow_OutOfScope = new TableRow({
				children: columnTitles_OutOfScope.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 16
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_OutOfScope = objectDetails_OutOfScope.map(field => {
				return new TableRow({
					children: columnTitles_OutOfScope.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "N/A", // Handle missing values gracefully
											size: 8 * 2, // Font size: 16
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const OutOfScopeTable = new Table({
				rows: [headerRow_OutOfScope, ...dataRows_OutOfScope]
			});
			documentChildrenFSD.push(OutOfScopeTable);




			var abbreviationsData = safeArr(parsedData.response_set_2 && parsedData.response_set_2.abbreviations);
			if (abbreviationsData && abbreviationsData.length > 0) {
				var Abbrevations = new Paragraph({
					children: [
						new Bookmark({
							id: "abbreviations",
							children: [
								new TextRun({
									text: "Abbreviations",
									font: "Helvetica",
									bold: true,
									size: 16 * 2
								})
							]
						})
					],
					spacing: {
						before: 250,
						after: 100
					}
				});
				var headerRow = new TableRow({
					children: [
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Abbreviation",
											font: "Helvetica",
											size: 8 * 2,
											bold: true
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 6,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Meaning",
											font: "Helvetica",
											size: 8 * 2,
											bold: true
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 15,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: "Additional Information",
											font: "Helvetica",
											size: 8 * 2,
											bold: true
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 15,
								type: "pct"
							},
							shading: {
								fill: "#e5e7eb"
							},
							verticalAlign: VerticalAlign.CENTER
						}),
					],
					height: {
						value: 500,
						rule: "atLeast"
					}
				});
				var abbreviationsTableRows = [headerRow];
				abbreviationsData.forEach(function (item) {
					var row = new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: item.Abbreviation,
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 6,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: item.Meaning,
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 15,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: item["Additional Information"],
												font: "Helvetica",
												size: 8 * 2
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 15,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					});

					// Add each dynamically created row to the abbreviationsTableRows array
					abbreviationsTableRows.push(row);
				});
				var Abbrevations_Table = new Table({
					rows: abbreviationsTableRows
				});
				documentChildrenFSD.push(Abbrevations);
				documentChildrenFSD.push(Abbrevations_Table);

			}
			var BusinessProcessContext1 = new Paragraph({
				children: [
					new Bookmark({
						id: "business_process_context",
						children: [
							new TextRun({
								text: "Business Process Context",
								bold: true,
								size: 16 * 2,
								font: "Helvetica",
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 250

				}
			});

			var BusinessProcessContext2 = safeStr(parsedData.response_set_2 && parsedData.response_set_2.business_process_context);

			var BusinessProcessContext = BusinessProcessContext2
				.split(/\n(?=\d+\.\s)/)
				.map(line => line.trim())
				.filter(line => line.length > 0);
			var BusinessProcessContext_data = BusinessProcessContext.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100 // Space between paragraphs
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(BusinessProcessContext1);
			documentChildrenFSD.push(...BusinessProcessContext_data);
			var ProcessesGroup_Processes_ProcessIntegration = new Paragraph({
				children: [
					new Bookmark({
						id: "processgroup",
						children: [
							new TextRun({
								text: "Processes Group, Processes, Process Integration",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100

				}
			});
			var paragraph22 = new Paragraph({
				children: [
					new TextRun({
						text: "Please see ",
						font: "Helvetica",
						size: 8 * 2, // Font size (adjust as needed)
						bold: true,
					}),
					new ExternalHyperlink({
						// eslint-disable-next-line @sap-ux/fiori-tools/sap-no-hardcoded-url -- intentional link to internal SharePoint/Signavio reference documentation
						link: "https://editor.signavio.com/p/hub-preview/model/86e6ee4ebdd04a078594556f99d1e8db",
						children: [
							new TextRun({
								text: "PTR-200-02-10 Create Fixed Asset Record - Indirect Capitalization",
								font: "Helvetica",
								style: "Hyperlink", // Optional if you want to apply hyperlink style
								size: 8 * 2, // Font size (adjust as needed)
								bold: true,
							})
						]
					}),
					new TextRun({
						text: " in Signavio for more details.",
						font: "Helvetica",
						size: 8 * 2, // Font size (adjust as needed)
						bold: true,
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			documentChildrenFSD.push(ProcessesGroup_Processes_ProcessIntegration);
			documentChildrenFSD.push(paragraph22);
			var IntegratedBusinessScenarios = new Paragraph({
				children: [
					new Bookmark({
						id: "integrated_business",
						children: [
							new TextRun({
								text: "Integrated Business Scenarios",
								font: "Helvetica",
								bold: true,
								size: 12 * 2 // Font size: 24
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 200 // Space after paragraph
				}
			});

			var integrated_business_scenarios2 = safeStr(parsedData.response_set_2 && parsedData.response_set_2.integrated_business_scenarios);

			var IntegratedBusinessScenarios_Data;

			if (integrated_business_scenarios2.trim().toUpperCase() === "NA") {
				IntegratedBusinessScenarios_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var IntegratedBusinessScenarios1 = integrated_business_scenarios2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				IntegratedBusinessScenarios_Data = IntegratedBusinessScenarios1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(IntegratedBusinessScenarios);
			documentChildrenFSD.push(...IntegratedBusinessScenarios_Data);
			var ArchitecturalContext = new Paragraph({
				children: [
					new Bookmark({
						id: "architectural_context",
						children: [
							new TextRun({
								text: "Architectural Context",
								font: "Helvetica",
								bold: true,
								size: 16 * 2 // Font size: 32
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 200 // Space after paragraph
				}
			});

			var ArchitecturalContex2 = safeStr(parsedData.response_set_2 && parsedData.response_set_2.architectural_context);

			var ArchitecturalContext_Data;

			if (ArchitecturalContex2.trim().toUpperCase() === "NA") {
				ArchitecturalContext_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2 // Font size: 16
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var ArchitecturalContext1 = ArchitecturalContex2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				ArchitecturalContext_Data = ArchitecturalContext1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(ArchitecturalContext);
			documentChildrenFSD.push(...ArchitecturalContext_Data);

			var Current_Target_IntermediateStates = new Paragraph({
				children: [
					new Bookmark({
						id: "Current_Target_Intermediate_States",
						children: [
							new TextRun({
								text: "Current, Target & Intermediate States",
								font: "Helvetica",
								bold: true,
								size: 12 * 2 // Font size: 24
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 200 // Space after paragraph
				}
			});

			var Current_Target_IntermediateStates2 = parsedData.response_set_3.current_target_and_intermidiate_states;

			var Current_Target_IntermediateStates_Data;

			if (Current_Target_IntermediateStates2.trim().toUpperCase() === "NA") {
				Current_Target_IntermediateStates_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2 // Font size: 16
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var Current_Target_IntermediateStates1 = Current_Target_IntermediateStates2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Current_Target_IntermediateStates_Data = Current_Target_IntermediateStates1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(Current_Target_IntermediateStates);
			documentChildrenFSD.push(...Current_Target_IntermediateStates_Data);

			var SatelliteApplications_Integration = new Paragraph({
				children: [
					new Bookmark({
						id: "Satellite_Applications_Integration",
						children: [
							new TextRun({
								text: "Satellite Applications & Integration",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200

				}
			});
			documentChildrenFSD.push(SatelliteApplications_Integration);
			var SatelliteApplications_Integration2 = parsedData.response_set_3.satellite_applications_and_integration;

			var SatelliteApplications_Integration_Data;

			if (SatelliteApplications_Integration2.trim().toUpperCase() === "NA") {
				// If the value is "NA", just add a paragraph with "NA"
				SatelliteApplications_Integration_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2, // Font size: 16
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100 // Space between paragraphs
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				// Process the data as normal if it's not "NA"
				var SatelliteApplications_Integration1 = SatelliteApplications_Integration2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SatelliteApplications_Integration_Data = SatelliteApplications_Integration1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(...SatelliteApplications_Integration_Data);

			var FurtherConsiderations = new Paragraph({
				children: [
					new Bookmark({
						id: "further_considerations",
						children: [
							new TextRun({
								text: "Further Considerations",
								bold: true,
								font: "Helvetica",
								size: 16 * 2 // Font size: 32
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 200 // Space after the paragraph
				}
			});

			var FurtherConsiderations2 = parsedData.response_set_3.further_considerations;

			var FurtherConsiderations_Data;

			if (FurtherConsiderations2.trim().toUpperCase() === "NA") {
				// If the value is "NA", just add a paragraph with "NA"
				FurtherConsiderations_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2, // Font size: 16
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100 // Space after the paragraph
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				// Process the data as normal if it's not "NA"
				var FurtherConsiderations1 = FurtherConsiderations2
					// .split(/\n(?=\d+\.\s)/) // Split based on numbered lines
					// .split(/\n\d+\.\s*/)
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering
					.map(line => line.trim()) // Trim whitespace
					.filter(line => line.length > 0); // Remove empty lines

				FurtherConsiderations_Data = FurtherConsiderations1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space after the paragraph
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(FurtherConsiderations);
			documentChildrenFSD.push(...FurtherConsiderations_Data);

			var Site_Country_BusinessUnitLocalizations = new Paragraph({
				children: [
					new Bookmark({
						id: "site_country",
						children: [
							new TextRun({
								text: "Site, Country & Business Unit Localizations",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var Site_Country_BusinessUnitLocalizations2 = parsedData.response_set_3.site_country_business_unit_localizations;

			var Site_Country_BusinessUnitLocalizations_Data;

			if (Site_Country_BusinessUnitLocalizations2.trim().toUpperCase() === "NA") {
				Site_Country_BusinessUnitLocalizations_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2,
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				var Site_Country_BusinessUnitLocalizations1 = Site_Country_BusinessUnitLocalizations2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Site_Country_BusinessUnitLocalizations_Data = Site_Country_BusinessUnitLocalizations1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(Site_Country_BusinessUnitLocalizations);
			documentChildrenFSD.push(...Site_Country_BusinessUnitLocalizations_Data);

			var Organization = new Paragraph({
				children: [
					new Bookmark({
						id: "organisation",
						children: [
							new TextRun({
								text: "Organization",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var Organization2 = parsedData.response_set_3.organization;

			var Organization_Data;

			if (Organization2.trim().toUpperCase() === "NA") {
				Organization_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2,
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				var Organization1 = Organization2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Organization_Data = Organization1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(Organization);
			documentChildrenFSD.push(...Organization_Data);
			var Data_MasterData = new Paragraph({
				children: [
					new Bookmark({
						id: "data_masterdata",
						children: [
							new TextRun({
								text: "Data, Master Data",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var Data_MasterData2 = parsedData.response_set_3.data_master_data;

			var Data_MasterData_Data;

			if (Data_MasterData2.trim().toUpperCase() === "NA") {
				Data_MasterData_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2,
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				var Data_MasterData1 = Data_MasterData2
					// .split(/\n(?=\d+\.\s)/)
					// .split(/\n\d+\.\s*/)
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Data_MasterData_Data = Data_MasterData1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(Data_MasterData);
			documentChildrenFSD.push(...Data_MasterData_Data);
			var Security_InternalControls = new Paragraph({
				children: [
					new Bookmark({
						id: "security_internalcontrols",
						children: [
							new TextRun({
								text: "Security, Internal Controls",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var Security_InternalControls2 = parsedData.response_set_3.security_internal_controls;

			var Security_InternalControls_Data;

			if (Security_InternalControls2.trim().toUpperCase() === "NA") {
				Security_InternalControls_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								size: 8 * 2,
								font: "Helvetica"
							})
						],
						spacing: {
							after: 100
						},
						alignment: AlignmentType.JUSTIFIED
					})
				];
			} else {
				var Security_InternalControls1 = Security_InternalControls2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Security_InternalControls_Data = Security_InternalControls1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					},
					alignment: AlignmentType.JUSTIFIED
				}));
			}

			documentChildrenFSD.push(Security_InternalControls);
			documentChildrenFSD.push(...Security_InternalControls_Data);
			var Audit_Compliance_Legal = new Paragraph({
				children: [
					new Bookmark({
						id: "audit_compliance_legal",
						children: [
							new TextRun({
								text: "Audit, Compliance & Legal",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var Audit_Compliance_Legal2 = parsedData.response_set_3.audit_compliance_legal;

			var Audit_Compliance_Legal_Data;

			if (Audit_Compliance_Legal2.trim().toUpperCase() === "NA") {
				Audit_Compliance_Legal_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var Audit_Compliance_Legal1 = Audit_Compliance_Legal2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Audit_Compliance_Legal_Data = Audit_Compliance_Legal1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							font: "Helvetica",
							size: 8 * 2
						})
					],
					spacing: {
						after: 100
					}
				}));
			}

			documentChildrenFSD.push(Audit_Compliance_Legal);
			documentChildrenFSD.push(...Audit_Compliance_Legal_Data);
			var EnvironmentSocialGovernance_ESG = new Paragraph({
				children: [
					new Bookmark({
						id: "esg",
						children: [
							new TextRun({
								text: "Environment, Social & Governance (ESG)",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var EnvironmentSocialGovernance_ESG2 = parsedData.response_set_3.environment_social_governance;

			var EnvironmentSocialGovernance_ESG_Data;

			if (EnvironmentSocialGovernance_ESG2.trim().toUpperCase() === "NA") {
				EnvironmentSocialGovernance_ESG_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var EnvironmentSocialGovernance_ESG1 = EnvironmentSocialGovernance_ESG2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				EnvironmentSocialGovernance_ESG_Data = EnvironmentSocialGovernance_ESG1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							font: "Helvetica",
							size: 8 * 2
						})
					],
					spacing: {
						after: 100
					}
				}));
			}

			documentChildrenFSD.push(EnvironmentSocialGovernance_ESG);
			documentChildrenFSD.push(...EnvironmentSocialGovernance_ESG_Data);
			var CountrySpecifics_Language = new Paragraph({
				children: [
					new Bookmark({
						id: "CountrySpecifics_Language",
						children: [
							new TextRun({
								text: "Country Specifics, Language",
								bold: true,
								font: "Helvetica",
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var CountrySpecifics_Language2 = parsedData.response_set_3.country_specifics_language;

			var CountrySpecifics_Language_Data;

			if (CountrySpecifics_Language2.trim().toUpperCase() === "NA") {
				CountrySpecifics_Language_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var CountrySpecifics_Language1 = CountrySpecifics_Language2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				CountrySpecifics_Language_Data = CountrySpecifics_Language1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}

			documentChildrenFSD.push(CountrySpecifics_Language);
			documentChildrenFSD.push(...CountrySpecifics_Language_Data);

			var DataMigration_Cutover = new Paragraph({
				children: [
					new Bookmark({
						id: "DataMigration_Cutover",
						children: [
							new TextRun({
								text: "Data Migration & Cutover",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var DataMigration_Cutover2 = parsedData.response_set_3.data_migration_cutover;

			var DataMigration_Cutover_Data;

			if (DataMigration_Cutover2.trim().toUpperCase() === "NA") {
				DataMigration_Cutover_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var DataMigration_Cutover1 = DataMigration_Cutover2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				DataMigration_Cutover_Data = DataMigration_Cutover1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}

			documentChildrenFSD.push(DataMigration_Cutover);
			documentChildrenFSD.push(...DataMigration_Cutover_Data);

			var NonfunctionalConsiderations = new Paragraph({
				children: [
					new Bookmark({
						id: "NonfunctionalConsiderations",
						children: [
							new TextRun({
								text: "Non-functional Considerations",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var NonfunctionalConsiderations2 = parsedData.response_set_3.non_functional_considerations;

			var NonfunctionalConsiderations_Data;

			if (NonfunctionalConsiderations2.trim().toUpperCase() === "NA") {
				NonfunctionalConsiderations_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var NonfunctionalConsiderations1 = NonfunctionalConsiderations2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				NonfunctionalConsiderations_Data = NonfunctionalConsiderations1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}

			documentChildrenFSD.push(NonfunctionalConsiderations);
			documentChildrenFSD.push(...NonfunctionalConsiderations_Data);
			var OtherConsiderations = new Paragraph({
				children: [
					new Bookmark({
						id: "OtherConsiderations",
						children: [
							new TextRun({
								text: "Other Considerations",
								font: "Helvetica",
								bold: true,
								size: 12 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});

			var OtherConsiderations2 = parsedData.response_set_3.other_considerations;

			var OtherConsiderations_Data;

			if (OtherConsiderations2.trim().toUpperCase() === "NA") {
				OtherConsiderations_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var OtherConsiderations1 = OtherConsiderations2
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				OtherConsiderations_Data = OtherConsiderations1.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}

			documentChildrenFSD.push(OtherConsiderations);
			documentChildrenFSD.push(...OtherConsiderations_Data);

			return documentChildrenFSD;

		},

		createFSDLastContents: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
			} = window.docx;
			var parsedData3 = response;
			 
			var parsedData = JSON.parse(parsedData3);

			var AdditionalInformationSection = new Paragraph({
				children: [
					new Bookmark({
						id: "AdditionalInfoID",
						children: [
							new TextRun({
								text: "Additional Information ",
								bold: true,
								size: 16 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					after: 100
				}
			});
			documentChildrenFSD.push(AdditionalInformationSection);
			var AdditionalInformationData = parsedData.response_set_3.additional_information
			var AdditionalInformationParagraphs;

			if (AdditionalInformationData.trim().toUpperCase() === "N/A") {
				AdditionalInformationParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var AdditionalInformationList = AdditionalInformationData
					
					// .split(/\n(?=\d+\.\s)/)
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering
					.map(line => line.trim())
					.filter(line => line.length > 0);

				AdditionalInformationParagraphs = AdditionalInformationList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}


			documentChildrenFSD.push(...AdditionalInformationParagraphs);



			return documentChildrenFSD;

		},
	createFSDFiori: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
				Table,
				TableRow,
				TableCell,
				AlignmentType,
				VerticalAlign,
			} = window.docx;

			var parsedFioriData = JSON.parse(response);
			var fiori = parsedFioriData.response['fiori app'].Fiori_App.Fiori_App[0];

			var fiori2 = parsedFioriData.response['fiori app']['Fiori_App2']['Functional_Requirements'][0];
			var fioriSpecificationDetails = fiori['Specification Details'];

			var FioriSpecificationDetailsArray = fioriSpecificationDetails
				.split(/(?<=\.)\s+|\n(?=\d+\.\s)/)
				.map(line => line.trim())
				.filter(line => line.length > 0);
			var Fiori_Specification_Details = new Paragraph({
				children: [
					new TextRun({
						text: "Fiori Specification Details",
						bold: true,
						size: 16 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var FioriSpecificationDetailsParagraphs = FioriSpecificationDetailsArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Fiori_Specification_Details);
			documentChildrenFSD.push(...FioriSpecificationDetailsParagraphs);

			var Fiori_Functional_Requirements = new Paragraph({
				children: [
					new TextRun({
						text: "Functional Requirements",
						bold: true,
						size: 14 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			documentChildrenFSD.push(Fiori_Functional_Requirements);


			var purpose = fiori2['Purpose'];
			var PurposeArray = purpose
				
				.replace(/^\d+\.\s*/, '')
				.split(/\s*\d+\.\s*/)
				.map(line => line.trim().replace(/\.+$/, ''))
				.filter(line => line.length > 0);
			var PurposeParagraphs = PurposeArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line, // Append period to each line if necessary
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			// Add Purpose Heading and Paragraphs to document children
			// documentChildrenFSD.push(Purpose_Heading);
			documentChildrenFSD.push(...PurposeParagraphs);

			var fioriCustomDevelopment = fiori2['Custom Development'];
			var FioriCustomDevelopmentArray = fioriCustomDevelopment
				
				.replace(/^\d+\.\s*/, '')
				.split(/\s*\d+\.\s*/)
				.map(line => line.trim().replace(/\.+$/, ''))  // Remove any trailing period(s)
				.filter(line => line.length > 0);
			var Fiori_Custom_Development = new Paragraph({
				children: [
					new TextRun({
						text: "Custom Development",
						bold: true,
						size: 12 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var FioriCustomDevelopmentParagraphs = FioriCustomDevelopmentArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line + ".",
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Fiori_Custom_Development);
			documentChildrenFSD.push(...FioriCustomDevelopmentParagraphs);

			var fieldsAndInputsData = fiori2["Details of Fields and Inputs"];
			var columnTitles_FieldsAndInputs = [
				"Field",
				"Field Name",
				"Optional/Mandatory Input",
				"Single/Multiple input/Selection"
			];
			var headerRow_FieldsAndInputs = new TableRow({
				children: columnTitles_FieldsAndInputs.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb",
						},
					})
				),
			});
			var dataRows_FieldsAndInputs = fieldsAndInputsData.map(field => {
				return new TableRow({
					children: columnTitles_FieldsAndInputs.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2,
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			var fieldsAndInputsDetailsTable = new Table({
				rows: [headerRow_FieldsAndInputs, ...dataRows_FieldsAndInputs]
			});
			documentChildrenFSD.push(fieldsAndInputsDetailsTable);

			var logicProcessFlow = fiori2["Logic/Process Flow"];
			var LogicProcessFlowArray = logicProcessFlow
				// .split(/\n(?=\d+\.\s)/)
				// .map(line => line.trim())
				.replace(/^\d+\.\s*/, '')
				.split(/\s*\d+\.\s*/)
				.map(line => line.trim().replace(/\.+$/, ''))
				.filter(line => line.length > 0);
			var Logic_Process_Flow_Heading = new Paragraph({
				children: [
					new TextRun({
						text: "Logic/Process Flow",
						bold: true,
						size: 12 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var LogicProcessFlowParagraphs = LogicProcessFlowArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line, // No period added as the original data already includes it
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Logic_Process_Flow_Heading);
			documentChildrenFSD.push(...LogicProcessFlowParagraphs);

			var keyFeatures = fiori2['Key Features'];
			var KeyFeaturesArray = keyFeatures
				// .split(/\n(?=\d+\.\s)/)
				// .map(line => line.trim())
				.replace(/^\d+\.\s*/, '')
				.split(/\s*\d+\.\s*/)
				.map(line => line.trim().replace(/\.+$/, ''))
				.filter(line => line.length > 0);
			var Key_Features_Heading = new Paragraph({
				children: [
					new TextRun({
						text: "Key Features",
						bold: true,
						size: 12 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var KeyFeaturesParagraphs = KeyFeaturesArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line, // Append period to each line
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Key_Features_Heading);
			documentChildrenFSD.push(...KeyFeaturesParagraphs);

			var useCases = fiori2['Use Cases'];
			var UseCasesArray = useCases
				// .split(/\n(?=\d+\.\s)/)
				// .map(line => line.trim())
				.replace(/^\d+\.\s*/, '')
				.split(/\s*\d+\.\s*/)
				.map(line => line.trim().replace(/\.+$/, ''))
				.filter(line => line.length > 0);
			var Use_Cases_Heading = new Paragraph({
				children: [
					new TextRun({
						text: "Use Cases",
						bold: true,
						size: 12 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var UseCasesParagraphs = UseCasesArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line, // Append period to each line if necessary
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Use_Cases_Heading);
			documentChildrenFSD.push(...UseCasesParagraphs);

			var fioriReportSelectionFields = fiori['Report Selection fields'];
			var FioriReportSelectionFieldsArray = fioriReportSelectionFields
				.split("\n")
				.map(line => line.trim())
				.filter(line => line.length > 0);
			var Fiori_Report_Selection_Fields = new Paragraph({
				children: [
					new TextRun({
						text: "Report Selection fields",
						bold: true,
						size: 14 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var FioriReportSelectionFieldsParagraphs = FioriReportSelectionFieldsArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Fiori_Report_Selection_Fields);
			documentChildrenFSD.push(...FioriReportSelectionFieldsParagraphs);

			var fioriTechnicalDetails = fiori['Technical Details'];
			var FioriTechnicalDetailsArray = fioriTechnicalDetails
				.split(". ")
				.map(line => line.trim())
				.filter(line => line.length > 0);
			var Fiori_Technical_Details = new Paragraph({
				children: [
					new TextRun({
						text: "Technical Details",
						bold: true,
						size: 16 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var FioriTechnicalDetailsParagraphs = FioriTechnicalDetailsArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line + ".",
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Fiori_Technical_Details);
			documentChildrenFSD.push(...FioriTechnicalDetailsParagraphs);

			var FioriSecurityRequirements = new Paragraph({
				children: [
					new Bookmark({
						id: "FioriSecurityRequirements",
						children: [
							new TextRun({
								text: "Security & Role Requirements, Sensitive Data", // Heading text
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var FioriSecurityRequirementsData = fiori["Security & Role Requirements, Sensitive Data"];
			var FioriSecurityRequirementsParagraphs;
			if (FioriSecurityRequirementsData.trim().toUpperCase() === "N/A") {
				FioriSecurityRequirementsParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var FioriSecurityRequirementsList = FioriSecurityRequirementsData
					.split(/\n(?=\d+\.\s)/) // Split by numbers followed by a period and newline
					.map(line => line.trim())
					.filter(line => line.length > 0);

				FioriSecurityRequirementsParagraphs = FioriSecurityRequirementsList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(FioriSecurityRequirements);
			documentChildrenFSD.push(...FioriSecurityRequirementsParagraphs);

			var PerformanceConsiderations = new Paragraph({
				children: [
					new Bookmark({
						id: "performance_considerations",
						children: [
							new TextRun({
								text: "Data Volumes & Performance Considerations",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var performanceData = fiori["Data Volumes & Performance Considerations"];
			var messageSizeData = performanceData["Message Size"];
			var initialLoadData = performanceData["Initial Load Volumne"];
			var avgVolumeData = performanceData["Average Volume"];
			var peakVolumeData = performanceData["Peak Volume"];
			var metricsData = performanceData["Performance Metrics"];
			var additionalDetails = performanceData["Additional Information"];
			var PerformanceConsiderations_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Message Size",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Initial Load Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Average Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Peak Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Performance Metrics",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: messageSizeData,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: initialLoadData,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: avgVolumeData,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: peakVolumeData,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: metricsData,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: additionalDetails,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});
			documentChildrenFSD.push(PerformanceConsiderations);
			documentChildrenFSD.push(PerformanceConsiderations_Table);

			var FioriProcessingData_heading = new Paragraph({
				children: [
					new Bookmark({
						id: "Fiori Processing, Response Times", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Synchronous Processing, Response Times",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			documentChildrenFSD.push(FioriProcessingData_heading);
			var FioriProcessingData = fiori["Synchronous Processing, Response Times"];
			var FioriProcessingParagraphs;
			if (FioriProcessingData.trim().toUpperCase() === "N/A") {
				FioriProcessingParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var FioriProcessingList = FioriProcessingData
					.split(/(?<=\.)\s+|\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				FioriProcessingParagraphs = FioriProcessingList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(...FioriProcessingParagraphs);

			var FioriBackgroundProcessing = new Paragraph({
				children: [
					new Bookmark({
						id: "fiori_background_processing", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Background Processing",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var fioriBackgroundProcessingData = fiori["Background Processing"];
			var additionalInformation = fioriBackgroundProcessingData["Additional Information"];
			var backgroundJob = fioriBackgroundProcessingData["Background Job"];
			var criticallyErrorHandling = fioriBackgroundProcessingData["Critically, Error Handling"];
			var description = fioriBackgroundProcessingData["Description"];
			var notificationDistributions = fioriBackgroundProcessingData["Notification, Distributions"];
			var recurrence = fioriBackgroundProcessingData["Recurrence"];
			var startConditionsTiming = fioriBackgroundProcessingData["Start Conditions, Timing"];
			var FioriBackgroundProcessing_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Background Job",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Start Conditions, Timing",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Recurrence",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Criticality, Error Handling",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Notification, Distribution",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: backgroundJob,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: description,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: startConditionsTiming,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: recurrence,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: criticallyErrorHandling,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: notificationDistributions,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: additionalInformation,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});
			documentChildrenFSD.push(FioriBackgroundProcessing);
			documentChildrenFSD.push(FioriBackgroundProcessing_Table);

			var TestScenarioSection = new Paragraph({
				children: [
					new Bookmark({
						id: "test_scenarios_section", // Updated bookmark id
						children: [
							new TextRun({
								text: "Test Scenarios, Test Data, Pre-conditions",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var testScenarioData = parsedFioriData.response['fiori app'].Fiori_App.TestData;
			var testScenariosArray = Array.isArray(testScenarioData) ? testScenarioData : [testScenarioData];
			var columnTitles_TestScenarios = [
				"SINO",
				"Scenario",
				"Outcome"
			];
			var headerRow_TestScenarios = new TableRow({
				children: columnTitles_TestScenarios.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb",
						},
					})
				),
			});
			var dataRows_TestScenarios = testScenariosArray.map(field => {
				return new TableRow({
					children: columnTitles_TestScenarios.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2,
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			var testScenariosDetailsTable = new Table({
				rows: [headerRow_TestScenarios, ...dataRows_TestScenarios]
			});
			documentChildrenFSD.push(TestScenarioSection);
			documentChildrenFSD.push(testScenariosDetailsTable);

			var sampleData_heading = new Paragraph({
				children: [
					new Bookmark({
						id: "sampleData_heading", // Updated bookmark id
						children: [
							new TextRun({
								text: "Samples",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var sampleData = fiori["Samples"];
			var sampleParagraphs;
			if (sampleData.trim().toUpperCase() === "NA") {
				sampleParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var sampleList = sampleData
					.split(/\n(?=\d+\.\s)/) // Split by numbers followed by a period and newline
					.map(line => line.trim())
					.filter(line => line.length > 0);

				sampleParagraphs = sampleList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(sampleData_heading);
			documentChildrenFSD.push(...sampleParagraphs);

			
			return documentChildrenFSD;
		},

		createFSDReport: function (response, documentChildrenFSD) {
			const {
				Paragraph, Bookmark, TextRun, Table,
				TableRow, TableCell, AlignmentType
			} = window.docx;

			var responseReport = response;
			var parsedResponse;
			try {
				parsedResponse = JSON.parse(responseReport);
			} catch {
				sap.m.MessageToast.show("Report section: Response is not valid JSON. Skipping section.");
				documentChildrenFSD.push(new Paragraph({
					children: [new TextRun({ text: "[Report section skipped: Invalid JSON response]", color: "FF0000", size: 16, font: "Helvetica" })]
				}));
				return documentChildrenFSD;
			}
			if (!parsedResponse || !parsedResponse.response) {
				sap.m.MessageToast.show("Warning: 'response' key is missing. Report section skipped.");
				documentChildrenFSD.push(new Paragraph({
					children: [new TextRun({ text: "[Report section skipped: 'response' key not found]", color: "FF0000", size: 16, font: "Helvetica" })]
				}));
				return documentChildrenFSD;
			}
			if (!parsedResponse.response.embedded) {
				sap.m.MessageToast.show("Warning: 'embedded' section is missing from response. Report section skipped.");
				documentChildrenFSD.push(new Paragraph({
					children: [new TextRun({ text: "[Report section skipped: 'embedded' not found in response]", color: "FF0000", size: 16, font: "Helvetica" })]
				}));
				return documentChildrenFSD;
			}
			var embeddedData;
			if (
				parsedResponse.response.embedded.embedded &&
				Array.isArray(parsedResponse.response.embedded.embedded) &&
				parsedResponse.response.embedded.embedded.length > 0
			) {

				embeddedData = parsedResponse.response.embedded.embedded;
			} else {
				sap.m.MessageToast.show("Warning: 'embedded.embedded' array is missing or empty. Some fields will be skipped.");
				embeddedData = null;
			}
			var embeddeddata2;
			if (parsedResponse.response.embedded.embedded2) {
				embeddeddata2 = parsedResponse.response.embedded.embedded2;
			} else {
				sap.m.MessageToast.show("Warning: 'embedded.embedded2' is missing. Functional Requirements and Test Data will be skipped.");
				embeddeddata2 = null;
			}
			var safeStr = function (val) {
				if (val === null || val === undefined) return "N/A";
				if (typeof val === "object") return "N/A";
				return String(val).trim() || "N/A";
			};
			var errPara = function (fieldName) {
				return new Paragraph({
					children: [new TextRun({
						text: "[" + fieldName + ": field not found in response]",
						color: "FF0000", size: 16, font: "Helvetica"
					})],
					spacing: { after: 100 }
				});
			};
			var isNA = function (val) {
				if (val === null || val === undefined || val === "") return true;
				return String(val).trim().toUpperCase() === "NA" ||
					String(val).trim().toUpperCase() === "N/A";
			};
			var makeTextParagraphs = function (text) {
				return String(text).split(/\n(?=\d+\.\s)/)
					.map(function (line) { return line.trim(); })
					.filter(function (line) { return line.length > 0; })
					.map(function (line) {
						return new Paragraph({
							children: [new TextRun({ text: line, size: 16, font: "Helvetica" })],
							spacing: { after: 100 },
							alignment: AlignmentType.JUSTIFIED
						});
					});
			};
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "SpecificationDetails",
					children: [new TextRun({ text: "Specification Details", size: 32, bold: true, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Specification Details"] !== undefined) {
				var specVal = embeddedData[0]["Specification Details"];
				if (isNA(specVal)) {
					documentChildrenFSD.push(new Paragraph({
						children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })],
						spacing: { after: 100 }
					}));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(specVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Specification Details' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Specification Details"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "FunctionalRequirementsFSD",
					children: [new TextRun({ text: "Functional Requirements", size: 28, bold: true, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddeddata2 && embeddeddata2.Functional_Requirements !== undefined) {
				try {
					var funcText = safeStr(embeddeddata2.Functional_Requirements);
					var headers = ["Purpose:", "Key Features:", "Use Cases:",
						"Details of Fields and Inputs:", "Logic/Process Flow:",
						"Custom Development:", "Special Note:"];
					var paragraphs_report = [];
					var currentParagraph = "";

					funcText.split(/\n+/).forEach(function (sentence) {
						var trimmed = sentence.trim();
						var isHeader = headers.some(function (h) { return trimmed.startsWith(h); });
						if (isHeader) {
							if (currentParagraph.trim()) {
								paragraphs_report.push({ text: currentParagraph.trim(), isHeader: false });
							}
							var hdr = headers.find(function (h) { return trimmed.startsWith(h); });
							paragraphs_report.push({ text: hdr, isHeader: true });
							currentParagraph = trimmed.replace(hdr, "").trim() + " ";
						} else {
							currentParagraph += trimmed + " ";
						}
					});
					if (currentParagraph.trim()) {
						paragraphs_report.push({ text: currentParagraph.trim(), isHeader: false });
					}

					paragraphs_report.forEach(function (item) {
						documentChildrenFSD.push(new Paragraph({
							children: item.isHeader
								? [new TextRun({ text: item.text, bold: true, size: 16, font: "Helvetica" }),
								new TextRun({ text: " ", size: 22 })]
								: [new TextRun({ text: item.text, bold: false, size: 16, font: "Helvetica" })],
							spacing: { after: 150 },
							alignment: AlignmentType.LEFT
						}));
					});
				} catch (e) {
					sap.m.MessageToast.show("Warning: Error rendering 'Functional Requirements': " + e.message);
					documentChildrenFSD.push(errPara("Functional Requirements (render error: " + e.message + ")"));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Functional_Requirements' is missing in embedded2.");
				documentChildrenFSD.push(errPara("Functional Requirements"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "BackgroundFSD",
					children: [new TextRun({ text: "Background", bold: true, size: 24, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Background"] !== undefined) {
				var bgVal = embeddedData[0]["Background"];
				if (isNA(bgVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(bgVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Background' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Background"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "ReportSelectionID",
					children: [new TextRun({ text: "Report Selection", bold: true, size: 24, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Report Selections"] !== undefined) {
				var rsVal = embeddedData[0]["Report Selections"];
				if (isNA(rsVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(rsVal));
				}
			} else {
				// ❌ Field missing
				sap.m.MessageToast.show("Warning: 'Report Selections' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Report Selections"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "ExpectedOutputLayoutID",
					children: [new TextRun({ text: "Expected Output & Layout", bold: true, size: 24, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Expected Output & Layout"] !== undefined) {
				var eoVal = embeddedData[0]["Expected Output & Layout"];
				if (isNA(eoVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(eoVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Expected Output & Layout' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Expected Output & Layout"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "DrillDownFollowOnProcessingID",
					children: [new TextRun({ text: "Drill-down and Follow-on Processing", size: 24, bold: true, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Drill-down and Follow-on Processing"] !== undefined) {
				var ddVal = embeddedData[0]["Drill-down and Follow-on Processing"];
				if (isNA(ddVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(ddVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Drill-down and Follow-on Processing' field is missing.");
				documentChildrenFSD.push(errPara("Drill-down and Follow-on Processing"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "SecurityRoleRequirementsID",
					children: [new TextRun({ text: "Security & Role Requirements, Sensitive Data", bold: true, size: 28, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Security & Role Requirements, Sensitive Data"] !== undefined) {
				var srVal = embeddedData[0]["Security & Role Requirements, Sensitive Data"];
				if (isNA(srVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(srVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Security & Role Requirements, Sensitive Data' field is missing.");
				documentChildrenFSD.push(errPara("Security & Role Requirements, Sensitive Data"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "DataVolumesPerformanceID",
					children: [new TextRun({ text: "Data Volumes & Performance Considerations", bold: true, size: 28, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Data Volumes & Performance Considerations"] !== undefined) {
				var dvVal = embeddedData[0]["Data Volumes & Performance Considerations"];
				if (isNA(dvVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(dvVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Data Volumes & Performance Considerations' field is missing.");
				documentChildrenFSD.push(errPara("Data Volumes & Performance Considerations"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "SynchronousProcessingID",
					children: [new TextRun({ text: "Synchronous Processing, Response Times", size: 28, bold: true, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Synchronous Processing, Response Times"] !== undefined) {
				var spVal = embeddedData[0]["Synchronous Processing, Response Times"];
				if (isNA(spVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(spVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Synchronous Processing, Response Times' field is missing.");
				documentChildrenFSD.push(errPara("Synchronous Processing, Response Times"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "BackgroundProcessingID",
					children: [new TextRun({ text: "Background Processing", bold: true, size: 28, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Background Processing"] !== undefined) {
				var bpVal = embeddedData[0]["Background Processing"];
				if (isNA(bpVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(bpVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Background Processing' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Background Processing"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "TestScenariosID",
					children: [new TextRun({ text: "Test Scenarios, Test Data, Pre-conditions", bold: true, size: 28, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddeddata2 && embeddeddata2.TestData !== undefined) {
				try {
					var testData = Array.isArray(embeddeddata2.TestData)
						? embeddeddata2.TestData : [embeddeddata2.TestData];

					var tsCols = ["SINO", "Scenario", "Outcome"];
					var tsHeaderRow = new TableRow({
						children: tsCols.map(function (t) {
							return new TableCell({
								children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 16, font: "Helvetica" })] })],
								shading: { fill: "#e5e7eb" }
							});
						})
					});
					var tsDataRows = testData.map(function (item) {
						return new TableRow({
							children: tsCols.map(function (t) {
								return new TableCell({
									children: [new Paragraph({
										children: [new TextRun({
											text: safeStr(item ? item[t] : null),
											size: 16, font: "Helvetica"
										})]
									})]
								});
							})
						});
					});
					documentChildrenFSD.push(new Table({ rows: [tsHeaderRow, ...tsDataRows] }));
				} catch (e) {
					sap.m.MessageToast.show("Warning: Error rendering Test Scenarios table: " + e.message);
					documentChildrenFSD.push(errPara("Test Scenarios table (render error: " + e.message + ")"));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'TestData' is missing in embedded2.");
				documentChildrenFSD.push(errPara("Test Scenarios (TestData)"));
			}
			documentChildrenFSD.push(new Paragraph({
				children: [new Bookmark({
					id: "SamplesID",
					children: [new TextRun({ text: "Samples", bold: true, size: 28, font: "Helvetica" })]
				})],
				spacing: { before: 250, after: 250 }
			}));

			if (embeddedData && embeddedData[0] && embeddedData[0]["Samples"] !== undefined) {
				var samVal = embeddedData[0]["Samples"];
				if (isNA(samVal)) {
					documentChildrenFSD.push(new Paragraph({ children: [new TextRun({ text: "N/A", font: "Helvetica", size: 16 })], spacing: { after: 100 } }));
				} else {
					documentChildrenFSD.push(...makeTextParagraphs(samVal));
				}
			} else {
				sap.m.MessageToast.show("Warning: 'Samples' field is missing in embedded data.");
				documentChildrenFSD.push(errPara("Samples"));
			}

			return documentChildrenFSD;
		},

	createFSDWorkflow: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
				Table,
				TableRow,
				TableCell,
				AlignmentType,
			} = window.docx;

			var response1 = response;
			var parsedWorkflowResponse = JSON.parse(response1);

			var WorkflowFSD = new Paragraph({
				children: [
					new Bookmark({
						id: "WorkflowID",
						children: [
							new TextRun({
								text: "Workflow",
								bold: true,
								size: 16 * 2, // Font size: 28
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 250   // Space after paragraph
				}
			});
			documentChildrenFSD.push(WorkflowFSD);


			var functionalRequirementsTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "FunctionalRequirements",
						children: [
							new TextRun({
								text: "Functional Requirements",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(functionalRequirementsTitle);


			
			var FunctionalRequirementsArray_Workflow = parsedWorkflowResponse.response.Workflow.workflow2.Functional_Requirements;

const headers_workflow = [
    "Purpose:",
    "Key Features:",
    "Use Cases:",
    "Details of Fields and Inputs:",
    "Logic/Process Flow:",
    "Custom Development:",
    "Special Note:"
];

let paragraphs_workflow = [];
let currentParagraph = "";
let lastHeader = null;

// Split the text by one or more newlines and process each sentence
FunctionalRequirementsArray_Workflow.split(/\n+/).forEach(sentence => {
    const trimmedSentence = sentence.trim();
    const isHeader = headers_workflow.some(header => trimmedSentence.startsWith(header));

    if (isHeader) {
        // Push the accumulated content before starting a new section
        if (currentParagraph.trim()) {
            paragraphs_workflow.push({ text: currentParagraph.trim(), isHeader: false });
        }
        // Find the header that this sentence starts with
        lastHeader = headers_workflow.find(header => trimmedSentence.startsWith(header));
        // Remove the header text from the sentence to get only the content
        let content = trimmedSentence.replace(lastHeader, "").trim();
        // Add the header as its own paragraph
        paragraphs_workflow.push({ text: lastHeader, isHeader: true });
        // Reset currentParagraph with any content that was on the header line
        currentParagraph = content ? content + " " : "";
    } else {
        // Continue adding content to the current paragraph
        currentParagraph += trimmedSentence + " ";
    }
});

// Push any remaining text as a final paragraph
if (currentParagraph.trim()) {
    paragraphs_workflow.push({ text: currentParagraph.trim(), isHeader: false });
}

// Map each section to a formatted Paragraph (assuming you're using a document-generation library)
const functionalRequirementParagraphs_Workflow = paragraphs_workflow.map(item => {
    return new Paragraph({
        children: item.isHeader
            ? [
                  new TextRun({
                      text: item.text,
                      bold: true, // Bold headers
                      size: 8 * 2,
                      font: "Helvetica"
                  }),
                  new TextRun({ text: " ", size: 22 }) // Add spacing after header
              ]
            : [
                  new TextRun({
                      text: item.text,
                      bold: false,
                      size: 8 * 2,
                      font: "Helvetica"
                  })
              ],
        spacing: {
            after: 150
        },
        alignment: AlignmentType.LEFT
    });
});

// Finally, add the formatted paragraphs to your document
documentChildrenFSD.push(...functionalRequirementParagraphs_Workflow);



			var triggersTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "Triggers",
						children: [
							new TextRun({
								text: "Triggers",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(triggersTitle);
			var triggersData = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Triggers'];
			var Triggers_Data;
			if (triggersData.trim().toUpperCase() === "NA") {
				Triggers_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var TriggersArray = triggersData

					.split(/(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Triggers_Data = TriggersArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...Triggers_Data);

			var dataSelectionValidationProcessingTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "DataSelectionValidationProcessing",
						children: [
							new TextRun({
								text: "Data Selection, Validation & Processing",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(dataSelectionValidationProcessingTitle);
			var dataSelectionValidationProcessingDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Data Selection, Validation & Processing'];
			var DataSelectionValidationProcessing_Data;
			if (dataSelectionValidationProcessingDetails.trim().toUpperCase() === "NA") {
				// Handle case where data is "NA"
				DataSelectionValidationProcessing_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2, // Font size: 16
							}),
						],
						spacing: {
							after: 100, // Space after the paragraph
						},
					}),
				];
			} else {
				// Split and process the data into numbered sections
				var DataSelectionValidationProcessingArray = dataSelectionValidationProcessingDetails
					// .split(/\n(?=\d+\.\s)/) // Split on newline before a numbered point
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering
					.map(line => line.trim()) // Trim whitespace from each line
					.filter(line => line.length > 0); // Remove empty lines

				DataSelectionValidationProcessing_Data = DataSelectionValidationProcessingArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100, // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED, // Justified alignment
				}));
			}
			documentChildrenFSD.push(...DataSelectionValidationProcessing_Data);

			var dialogWorkItemsAgentDeterminationTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "DialogWorkItemsAgentDetermination",
						children: [
							new TextRun({
								text: "Dialog Work Items, Agent Determination",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(dialogWorkItemsAgentDeterminationTitle);
			var dialogWorkItemsAgentDeterminationDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Dialog Work Items, Agent Determination'];
			var DialogWorkItemsAgentDetermination_Data;
			if (dialogWorkItemsAgentDeterminationDetails.trim().toUpperCase() === "NA") {

				DialogWorkItemsAgentDetermination_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100, // Space after the paragraph
						},
					}),
				];
			} else {

				var DialogWorkItemsAgentDeterminationArray = dialogWorkItemsAgentDeterminationDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				DialogWorkItemsAgentDetermination_Data = DialogWorkItemsAgentDeterminationArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...DialogWorkItemsAgentDetermination_Data);


			var specificationDetailsTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "SpecificationDetails",
						children: [
							new TextRun({
								text: "Specification Details",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(specificationDetailsTitle);
			var specificationDetailsData = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Specification Details'];
			var SpecificationDetails_Data;
			if (specificationDetailsData.trim().toUpperCase() === "NA") {
				SpecificationDetails_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SpecificationDetailsArray = specificationDetailsData
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SpecificationDetails_Data = SpecificationDetailsArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SpecificationDetails_Data);

			var delegationSubstitutionTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "DelegationSubstitution",
						children: [
							new TextRun({
								text: "Delegation, Substitution",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(delegationSubstitutionTitle);
			var delegationSubstitutionDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Delegation, Substitution'];
			var DelegationSubstitution_Data;
			if (delegationSubstitutionDetails.trim().toUpperCase() === "NA") {
				DelegationSubstitution_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2, // Font size: 16
							}),
						],
						spacing: {
							after: 100, // Space after the paragraph
						},
					}),
				];
			} else {

				var DelegationSubstitutionArray = delegationSubstitutionDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				DelegationSubstitution_Data = DelegationSubstitutionArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100, // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED, // Justified alignment
				}));
			}
			documentChildrenFSD.push(...DelegationSubstitution_Data);

			var deadlineMonitoringTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "DeadlineMonitoring",
						children: [
							new TextRun({
								text: "Deadline Monitoring, Escalation Rules & Reminders",
								bold: true,
								size: 12 * 2,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(deadlineMonitoringTitle);
			var deadlineMonitoringDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Deadline Monitoring, Escalation Rules & Reminders'];
			var DeadlineMonitoring_Data;
			if (deadlineMonitoringDetails.trim().toUpperCase() === "NA") {

				DeadlineMonitoring_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {

				var DeadlineMonitoringArray = deadlineMonitoringDetails
					.split(/(?<=\.)\s+/) // 
					.map(line => line.trim())
					.filter(line => line.length > 0);

				DeadlineMonitoring_Data = DeadlineMonitoringArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...DeadlineMonitoring_Data);

			var workItemTerminationTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "WorkItemTermination",
						children: [
							new TextRun({
								text: "Work Item Termination",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(workItemTerminationTitle);
			var workItemTerminationData = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Work Item Termination'];
			var WorkItemTermination_Data;
			if (workItemTerminationData.trim().toUpperCase() === "NA") {
				WorkItemTermination_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var WorkItemTerminationArray = workItemTerminationData
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				WorkItemTermination_Data = WorkItemTerminationArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...WorkItemTermination_Data);

			var expectedOutcomeTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "ExpectedOutcome",
						children: [
							new TextRun({
								text: "Expected Outcome",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(expectedOutcomeTitle);
			var expectedOutcomeDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Expected Outcome'];
			var ExpectedOutcome_Data;
			if (expectedOutcomeDetails.trim().toUpperCase() === "NA") {

				ExpectedOutcome_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								bold: true,
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {

				var ExpectedOutcomeArray = expectedOutcomeDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				ExpectedOutcome_Data = ExpectedOutcomeArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...ExpectedOutcome_Data);

			var notificationsErrorHandlingTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "NotificationsErrorHandling",
						children: [
							new TextRun({
								text: "Notifications & Error Handling",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(notificationsErrorHandlingTitle);
			const workflowsNotificationsDetails = parsedWorkflowResponse.response.Workflow.workflow2.NotificationsErrorHandling;
			const normalizedWorkflowsNotificationsDetails = Array.isArray(workflowsNotificationsDetails)
				? workflowsNotificationsDetails
				: [workflowsNotificationsDetails];
			const objectDetails_WorkflowsNotifications = normalizedWorkflowsNotificationsDetails.map(notification => ({
				"Error Code": notification["error_code"] || "N/A",
				"Message Type": notification["message_type"] || "N/A",
				"Description": notification["description"] || "N/A",
				"Trigger Event": notification["trigger_event"] || "N/A",
				"Display Location": notification["display_location"] || "N/A"
			}));
			const columnTitles_WorkflowsNotifications = [
				"Error Code",
				"Message Type",
				"Description",
				"Trigger Event",
				"Display Location"
			];
			const headerRow_WorkflowsNotifications = new TableRow({
				children: columnTitles_WorkflowsNotifications.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_WorkflowsNotifications = objectDetails_WorkflowsNotifications.map(field => {
				return new TableRow({
					children: columnTitles_WorkflowsNotifications.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const WorkflowsNotificationsDetailsTable = new Table({
				rows: [headerRow_WorkflowsNotifications, ...dataRows_WorkflowsNotifications]
			});
			documentChildrenFSD.push(WorkflowsNotificationsDetailsTable);




			var furtherConstraintsSpecialConsiderationsTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "FurtherConstraintsSpecialConsiderations",
						children: [
							new TextRun({
								text: "Further Constraints & Special Considerations",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(furtherConstraintsSpecialConsiderationsTitle);
			var furtherConstraintsSpecialConsiderationsDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Further Constraints & Special Considerations'];
			var FurtherConstraintsSpecialConsiderations_Data;
			if (furtherConstraintsSpecialConsiderationsDetails.trim().toUpperCase() === "NA") {

				FurtherConstraintsSpecialConsiderations_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {

				var FurtherConstraintsSpecialConsiderationsArray = furtherConstraintsSpecialConsiderationsDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				FurtherConstraintsSpecialConsiderations_Data = FurtherConstraintsSpecialConsiderationsArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100, // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED, // Justified alignment
				}));
			}
			documentChildrenFSD.push(...FurtherConstraintsSpecialConsiderations_Data);

			var securityAndRoleRequirementsTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "SecurityAndRoleRequirements",
						children: [
							new TextRun({
								text: "Security & Role Requirements, Sensitive Data",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(securityAndRoleRequirementsTitle);
			var securityAndRoleRequirementsDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Security & Role Requirements, Sensitive Data'];
			var SecurityAndRoleRequirements_Data;
			if (securityAndRoleRequirementsDetails.trim().toUpperCase() === "NA") {
				SecurityAndRoleRequirements_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SecurityAndRoleRequirementsArray = securityAndRoleRequirementsDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SecurityAndRoleRequirements_Data = SecurityAndRoleRequirementsArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SecurityAndRoleRequirements_Data);

			var DataVolumesPerformanceFSD = new Paragraph({
				children: [
					new Bookmark({
						id: "DataVolumesPerformanceID",
						children: [
							new TextRun({
								text: "Data Volumes & Performance Considerations",
								bold: true,
								size: 12 * 2, // Font size: 28
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 250   // Space after paragraph
				}
			});
			documentChildrenFSD.push(DataVolumesPerformanceFSD);
			const dataVolumesPerformanceDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]["Data Volumes & Performance Considerations"];
			const objectDetails_DataVolumesPerformance = [
				{
					"Additional Information": dataVolumesPerformanceDetails["Additional Information"] || "N/A",
					"Average Volume": dataVolumesPerformanceDetails["Average Volume"] || "N/A",
					"Initial Load Volume": dataVolumesPerformanceDetails["Initial Load Volumne"] || "N/A", // Note: Fix typo in "Volumne"
					"Message Size": dataVolumesPerformanceDetails["Message Size"] || "N/A",
					"Peak Volume": dataVolumesPerformanceDetails["Peak Volume"] || "N/A",
					"Performance Metrics": dataVolumesPerformanceDetails["Performance Metrics"] || "N/A"
				},
			];
			const columnTitles_DataVolumesPerformance = [
				"Additional Information",
				"Average Volume",
				"Initial Load Volume",
				"Message Size",
				"Peak Volume",
				"Performance Metrics"
			];
			const headerRow_DataVolumesPerformance = new TableRow({
				children: columnTitles_DataVolumesPerformance.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_DataVolumesPerformance = objectDetails_DataVolumesPerformance.map(field => {
				return new TableRow({
					children: columnTitles_DataVolumesPerformance.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const DataVolumesPerformanceDetailsTable = new Table({
				rows: [headerRow_DataVolumesPerformance, ...dataRows_DataVolumesPerformance],
			});
			documentChildrenFSD.push(DataVolumesPerformanceDetailsTable);

			var synchronousProcessingTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "SynchronousProcessing",
						children: [
							new TextRun({
								text: "Synchronous Processing, Response Times",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(synchronousProcessingTitle);
			var synchronousProcessingData = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Synchronous Processing, Response Times'];
			var SynchronousProcessing_Data;
			if (synchronousProcessingData.trim().toUpperCase() === "NA") {
				SynchronousProcessing_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SynchronousProcessingArray = synchronousProcessingData
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SynchronousProcessing_Data = SynchronousProcessingArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SynchronousProcessing_Data);


			var BackgroundProcessingFSD = new Paragraph({
				children: [
					new Bookmark({
						id: "BGID",
						children: [
							new TextRun({
								text: "Background Processing",
								bold: true,
								size: 12 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 250
				}
			});
			documentChildrenFSD.push(BackgroundProcessingFSD);
			const backgroundProcessingDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]["Background Processing"];
			const objectDetails_BackgroundProcessing = [
				{
					"Additional Information": backgroundProcessingDetails["Additional Information"] || "N/A",
					"Background Job": backgroundProcessingDetails["Background Job"] || "N/A",
					"Critically, Error Handling": backgroundProcessingDetails["Critically, Error Handling"] || "N/A",
					"Description": backgroundProcessingDetails["Description"] || "N/A",
					"Notification, Distributions": backgroundProcessingDetails["Notification, Distributions"] || "N/A",
					"Recurrence": backgroundProcessingDetails["Recurrence"] || "N/A",
					"Start Conditions, Timing": backgroundProcessingDetails["Start Conditions, Timing"] || "N/A"
				},
			];
			var columnTitles_BackgroundProcessing = [
				"Additional Information",
				"Background Job",
				"Critically, Error Handling",
				"Description",
				"Notification, Distributions",
				"Recurrence",
				"Start Conditions, Timing"
			];
			var headerRow_BackgroundProcessing = new TableRow({
				children: columnTitles_BackgroundProcessing.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb",
						},
					})
				),
			});
			var dataRows_BackgroundProcessing = objectDetails_BackgroundProcessing.map(field => {
				return new TableRow({
					children: columnTitles_BackgroundProcessing.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2,
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			var BackgroundProcessingDetailsTable = new Table({
				rows: [headerRow_BackgroundProcessing, ...dataRows_BackgroundProcessing]
			});
			documentChildrenFSD.push(BackgroundProcessingDetailsTable);

			var testScenariosTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "TestScenariosID",
						children: [
							new TextRun({
								text: "Test Scenarios, Test Data, Pre-conditions",
								bold: true,
								size: 12 * 2, // Font size: 28
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250, // Space before paragraph
					after: 250   // Space after paragraph
				}
			});
			documentChildrenFSD.push(testScenariosTitle);
			const testScenariosDetails = parsedWorkflowResponse.response.Workflow.workflow2.TestData;
			const normalizedTestScenariosDetails = Array.isArray(testScenariosDetails)
				? testScenariosDetails
				: [testScenariosDetails];
			const objectDetails_TestScenarios = normalizedTestScenariosDetails.map(item => ({
				"SINO": item["SINO"] || "N/A",
				"Scenario": item["Scenario"] || "N/A",
				"Outcome": item["Outcome"] || "N/A"
			}));
			const columnTitles_TestScenarios = [
				"SINO",
				"Scenario",
				"Outcome"
			];
			const headerRow_TestScenarios = new TableRow({
				children: columnTitles_TestScenarios.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_TestScenarios = objectDetails_TestScenarios.map(field => {
				return new TableRow({
					children: columnTitles_TestScenarios.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const TestScenariosDetailsTable = new Table({
				rows: [headerRow_TestScenarios, ...dataRows_TestScenarios]
			});
			documentChildrenFSD.push(TestScenariosDetailsTable);


			var samplesTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "Samples",
						children: [
							new TextRun({
								text: "Samples",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(samplesTitle);
			var samplesDetails = parsedWorkflowResponse.response.Workflow.workflow.Workflowdetails[0]['Samples'];
			var Samples_Data;
			if (samplesDetails.trim().toUpperCase() === "NA") {

				Samples_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {

				var SamplesArray = samplesDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				Samples_Data = SamplesArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100, // Space between paragraphs
					},
					alignment: AlignmentType.JUSTIFIED, // Justified alignment
				}));
			}
			documentChildrenFSD.push(...Samples_Data);



			return documentChildrenFSD;
		},

		createFSDEnhancement: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
				Table,
				TableRow,
				TableCell,
				AlignmentType,
			} = window.docx;

			var response_Enhancement = response;
			var parsedEnhancementResponse = JSON.parse(response_Enhancement);

			

			var specificationDetailsTitle_Enhancement = new Paragraph({
				children: [
					new Bookmark({
						id: "SpecificationDetailsEnhancement",
						children: [
							new TextRun({
								text: "Specification Details",
								size: 14 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(specificationDetailsTitle_Enhancement);
			 
			var specificationDetailsData_Enhancement = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Specification Details'];
			// parsedEnhancementResponse.response.Enhancement.Enhancement.Enhancement[0]['Specification Details']
			var SpecificationDetails_Data_Enhancement;
			if (specificationDetailsData_Enhancement.trim().toUpperCase() === "NA") {
				SpecificationDetails_Data_Enhancement = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SpecificationDetailsArray_Enhancement = specificationDetailsData_Enhancement
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SpecificationDetails_Data_Enhancement = SpecificationDetailsArray_Enhancement.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SpecificationDetails_Data_Enhancement);


			var functionalRequirementsEnhancementTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "FunctionalRequirementsEnhancement",
						children: [
							new TextRun({
								text: "Functional Requirements",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(functionalRequirementsEnhancementTitle);


		
		 
		



		
var FunctionalRequirementsArray_Enhancement = parsedEnhancementResponse.response.Enhancement.Enhancement2.Functional_Requirements;

const headers = [
    "Objective:",
    "Current State (As-Is):",
    "Future State (To-Be):",
    "Features and Functional Scope:",
    "Dependencies:",
    "Logic/Process Flow:",
    "Additional Considerations:"
];

let paragraphs_enhancement = [];
let currentParagraph = "";
let lastHeader = null;

// Split the text while focusing ous10n all defined headers
FunctionalRequirementsArray_Enhancement.split(/\n+/).forEach(sentence => {
    const trimmedSentence = sentence.trim();
    const isHeader = headers.some(header => trimmedSentence.startsWith(header));

    if (isHeader) {
        // Push the previous paragraph before starting a new section
        if (currentParagraph.trim()) {
            paragraphs_enhancement.push({ text: currentParagraph.trim(), isHeader: false });
        }
        lastHeader = headers.find(header => trimmedSentence.startsWith(header)); // Get the actual header
        let content = trimmedSentence.replace(lastHeader, "").trim(); // Remove header from the text
        paragraphs_enhancement.push({ text: lastHeader, isHeader: true });
        currentParagraph = content ? content + " " : "";
    } else {
        currentParagraph += trimmedSentence + " ";
    }
});

// Push the last paragraph if there is any remaining text
if (currentParagraph.trim()) {
    paragraphs_enhancement.push({ text: currentParagraph.trim(), isHeader: false });
}

// Now the extracted paragraphs_enhancement will contain all sections
const functionalRequirementParagraphs_Enhancement = paragraphs_enhancement.map(item => {
    return new Paragraph({
        children: item.isHeader
            ? [
                  new TextRun({
                      text: item.text,
                      bold: true, // Apply bold to headers only
                      size: 8 * 2,
                      font: "Helvetica"
                  }),
                  new TextRun({ text: " ", size: 22 }) // Space after header
              ]
            : [
                  new TextRun({
                      text: item.text,
                      bold: false, // Ensure normal text formatting for content
                      size: 8 * 2,
                      font: "Helvetica"
                  })
              ],
        spacing: {
            after: 150
        },
        alignment: AlignmentType.LEFT
    });
});

// Add the formatted paragraphs to the document
documentChildrenFSD.push(...functionalRequirementParagraphs_Enhancement);


			var triggersTitleEnhancement = new Paragraph({
				children: [
					new Bookmark({
						id: "Triggers",
						children: [
							new TextRun({
								text: "Triggers",
								bold: true,
								size: 12 * 2,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(triggersTitleEnhancement);
			var triggersData = parsedEnhancementResponse.response.Enhancement.Enhancement[0].Triggers;
			var Triggers_Data;
			if (triggersData.trim().toUpperCase() === "NA") {
				Triggers_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var TriggersArray = triggersData
					// .split(/(?=\d+\.\s)/)
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering

					.map(line => line.trim())
					.filter(line => line.length > 0);

				Triggers_Data = TriggersArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...Triggers_Data);

			var dataSelectionValidationProcessingParagraph = new Paragraph({
				children: [
					new Bookmark({
						id: "DataSelectionValidationProcessingBookmarkID",
						children: [
							new TextRun({
								text: "Data Selection, Validation & Processing",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(dataSelectionValidationProcessingParagraph);
			var dataSelectionValidationProcessingDetails = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Data Selection, Validation & Processing'];
			var processedDataSelectionValidationProcessingContent;
			if (dataSelectionValidationProcessingDetails.trim().toUpperCase() === "NA") {
				processedDataSelectionValidationProcessingContent = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 16,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var dataSelectionValidationProcessingArray = dataSelectionValidationProcessingDetails
					// .split(/\n(?=\d+\.\s)/)
					.replace(/^\d+\.\s*/, '') // Remove leading "1." or any "n." at the start
					.split(/\n\d+\.\s*/) // Split at numbered sections and remove numbering
					.map(line => line.trim())
					.filter(line => line.length > 0);

				processedDataSelectionValidationProcessingContent = dataSelectionValidationProcessingArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 16,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...processedDataSelectionValidationProcessingContent);

			var expectedOutcomeEnhancementTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "ExpectedOutcomeEnhancement",
						children: [
							new TextRun({
								text: "Expected Outcome",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(expectedOutcomeEnhancementTitle);
			var expectedOutcomeEnhancementDetails = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Expected Outcome'];
			var ExpectedOutcomeEnhancement_Data;
			if (expectedOutcomeEnhancementDetails.trim().toUpperCase() === "NA") {
				ExpectedOutcomeEnhancement_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								bold: true,
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var ExpectedOutcomeEnhancementArray = expectedOutcomeEnhancementDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				ExpectedOutcomeEnhancement_Data = ExpectedOutcomeEnhancementArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...ExpectedOutcomeEnhancement_Data);

			var notificationsErrorHandlingEnhancementTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "NotificationsErrorHandlingEnhancement",
						children: [
							new TextRun({
								text: "Notifications & Error Handling",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(notificationsErrorHandlingEnhancementTitle);
			const enhancedNotificationsDetails = parsedEnhancementResponse.response.Enhancement.Enhancement2.NotificationsErrorHandling;
			const normalizedEnhancedNotificationsDetails = Array.isArray(enhancedNotificationsDetails)
				? enhancedNotificationsDetails
				: [enhancedNotificationsDetails];
			const objectDetails_EnhancedNotifications = normalizedEnhancedNotificationsDetails.map(notification => ({
				"Error Code": notification["error_code"] || "N/A",
				"Message Type": notification["message_type"] || "N/A",
				"Description": notification["description"] || "N/A",
				"Trigger Event": notification["trigger_event"] || "N/A",
				"Display Location": notification["display_location"] || "N/A"
			}));
			const columnTitles_EnhancedNotifications = [
				"Error Code",
				"Message Type",
				"Description",
				"Trigger Event",
				"Display Location"
			];
			const headerRow_EnhancedNotifications = new TableRow({
				children: columnTitles_EnhancedNotifications.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_EnhancedNotifications = objectDetails_EnhancedNotifications.map(field => {
				return new TableRow({
					children: columnTitles_EnhancedNotifications.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const EnhancedNotificationsDetailsTable = new Table({
				rows: [headerRow_EnhancedNotifications, ...dataRows_EnhancedNotifications]
			});
			documentChildrenFSD.push(EnhancedNotificationsDetailsTable);



			var furtherConstraintsSpecialConsiderationsEnhancementTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "FurtherConstraintsSpecialConsiderationsEnhancement",
						children: [
							new TextRun({
								text: "Further Constraints & Special Considerations",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250, // Space before the paragraph
					after: 250,  // Space after the paragraph
				},
			});
			documentChildrenFSD.push(furtherConstraintsSpecialConsiderationsEnhancementTitle);
			var furtherConstraintsSpecialConsiderationsEnhancementDetails = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Further Constraints & Special Considerations'];
			var FurtherConstraintsSpecialConsiderationsEnhancement_Data;
			if (furtherConstraintsSpecialConsiderationsEnhancementDetails.trim().toUpperCase() === "NA") {
				FurtherConstraintsSpecialConsiderationsEnhancement_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var FurtherConstraintsSpecialConsiderationsEnhancementArray = furtherConstraintsSpecialConsiderationsEnhancementDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				FurtherConstraintsSpecialConsiderationsEnhancement_Data = FurtherConstraintsSpecialConsiderationsEnhancementArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...FurtherConstraintsSpecialConsiderationsEnhancement_Data);



			var securityAndRoleRequirementsTitle_Enhancement = new Paragraph({
				children: [
					new Bookmark({
						id: "SecurityAndRoleRequirementsEnhancement",
						children: [
							new TextRun({
								text: "Security & Role Requirements, Sensitive Data",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(securityAndRoleRequirementsTitle_Enhancement);
			var securityAndRoleRequirementsDetails_Enhancement = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Security & Role Requirements, Sensitive Data'];
			var SecurityAndRoleRequirements_Data_Enhancement;
			if (securityAndRoleRequirementsDetails_Enhancement.trim().toUpperCase() === "NA") {
				SecurityAndRoleRequirements_Data_Enhancement = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SecurityAndRoleRequirementsArray_Enhancement = securityAndRoleRequirementsDetails_Enhancement
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SecurityAndRoleRequirements_Data_Enhancement = SecurityAndRoleRequirementsArray_Enhancement.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SecurityAndRoleRequirements_Data_Enhancement);


			var DataVolumesPerformanceParagraph = new Paragraph({
				children: [
					new Bookmark({
						id: "DataVolumesPerformanceBookmarkID",
						children: [
							new TextRun({
								text: "Data Volumes & Performance Considerations",
								bold: true,
								size: 12 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 250
				}
			});
			documentChildrenFSD.push(DataVolumesPerformanceParagraph);
			const dataVolumesPerformanceData = parsedEnhancementResponse.response.Enhancement.Enhancement[0]["Data Volumes & Performance Considerations"];
			const objectDetails_DataVolumesPerformanceEnhanced = [
				{
					"Message Size": dataVolumesPerformanceData["Message Size"] || "N/A",
					"Initial Load Volume": dataVolumesPerformanceData["Initial Load Volumne"] || "N/A",
					"Average Volume": dataVolumesPerformanceData["Average Volume"] || "N/A",
					"Peak Volume": dataVolumesPerformanceData["Peak Volume"] || "N/A",
					"Performance Metrics": dataVolumesPerformanceData["Performance Metrics"] || "N/A",
					"Additional Information": dataVolumesPerformanceData["Additional Information"] || "N/A"
				},
			];
			const columnTitles_DataVolumesPerformanceEnhanced = [
				"Message Size",
				"Initial Load Volume",
				"Average Volume",
				"Peak Volume",
				"Performance Metrics",
				"Additional Information"
			];
			const headerRow_DataVolumesPerformanceEnhanced = new TableRow({
				children: columnTitles_DataVolumesPerformanceEnhanced.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb",
						},
					})
				),
			});
			const dataRows_DataVolumesPerformanceEnhanced = objectDetails_DataVolumesPerformanceEnhanced.map(field => {
				return new TableRow({
					children: columnTitles_DataVolumesPerformanceEnhanced.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "",
											size: 8 * 2,
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const DataVolumesPerformanceEnhancedTable = new Table({
				rows: [headerRow_DataVolumesPerformanceEnhanced, ...dataRows_DataVolumesPerformanceEnhanced]
			});
			documentChildrenFSD.push(DataVolumesPerformanceEnhancedTable);


			var synchronousProcessingTitle_Enhancement = new Paragraph({
				children: [
					new Bookmark({
						id: "SynchronousProcessingEnhancement",
						children: [
							new TextRun({
								text: "Synchronous Processing, Response Times",
								size: 12 * 2,
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(synchronousProcessingTitle_Enhancement);
			var synchronousProcessingData_Enhancement = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Synchronous Processing, Response Times'];
			var SynchronousProcessing_Data_Enhancement;
			if (synchronousProcessingData_Enhancement.trim().toUpperCase() === "NA") {
				SynchronousProcessing_Data_Enhancement = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SynchronousProcessingArray_Enhancement = synchronousProcessingData_Enhancement
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SynchronousProcessing_Data_Enhancement = SynchronousProcessingArray_Enhancement.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SynchronousProcessing_Data_Enhancement);

			var BackgroundProcessingFSDSection = new Paragraph({
				children: [
					new Bookmark({
						id: "BGID_FSD",
						children: [
							new TextRun({
								text: "Background Processing",
								bold: true,
								size: 12 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 250
				}
			});
			documentChildrenFSD.push(BackgroundProcessingFSDSection);
			const bgProcessingDetailsFSD = parsedEnhancementResponse.response.Enhancement.Enhancement[0]["Background Processing"];
			const bgProcessingObjectDetails = [
				{
					"Additional Information": bgProcessingDetailsFSD["Additional Information"] || "N/A",
					"Background Job": bgProcessingDetailsFSD["Background Job"] || "N/A",
					"Critically, Error Handling": bgProcessingDetailsFSD["Critically, Error Handling"] || "N/A",
					"Description": bgProcessingDetailsFSD["Description"] || "N/A",
					"Notification, Distributions": bgProcessingDetailsFSD["Notification, Distributions"] || "N/A",
					"Recurrence": bgProcessingDetailsFSD["Recurrence"] || "N/A",
					"Start Conditions, Timing": bgProcessingDetailsFSD["Start Conditions, Timing"] || "N/A"
				}
			];
			var bgProcessingColumnTitles = [
				"Additional Information",
				"Background Job",
				"Critically, Error Handling",
				"Description",
				"Notification, Distributions",
				"Recurrence",
				"Start Conditions, Timing"
			];
			var bgProcessingHeaderRow = new TableRow({
				children: bgProcessingColumnTitles.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb",
						},
					})
				),
			});
			var bgProcessingDataRows = bgProcessingObjectDetails.map(field => {
				return new TableRow({
					children: bgProcessingColumnTitles.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "",
											size: 8 * 2,
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			var bgProcessingDetailsTable = new Table({
				rows: [bgProcessingHeaderRow, ...bgProcessingDataRows]
			});
			documentChildrenFSD.push(bgProcessingDetailsTable);


			var testScenariosTitle_Enhancement = new Paragraph({
				children: [
					new Bookmark({
						id: "TestScenariosEnhancementID",
						children: [
							new TextRun({
								text: "Test Scenarios, Test Data, Pre-conditions",
								bold: true,
								size: 12 * 2,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(testScenariosTitle_Enhancement);
			const testScenariosDetails_Enhancement = parsedEnhancementResponse.response.Enhancement.Enhancement2.TestData;;
			const normalizedTestScenariosDetails_Enhancement = Array.isArray(testScenariosDetails_Enhancement)
				? testScenariosDetails_Enhancement
				: [testScenariosDetails_Enhancement];
			const objectDetails_TestScenarios_Enhancement = normalizedTestScenariosDetails_Enhancement.map(item => ({
				"SINO": item["SINO"] || "N/A",
				"Scenario": item["Scenario"] || "N/A",
				"Outcome": item["Outcome"] || "N/A"
			}));
			const columnTitles_TestScenarios_Enhancement = [
				"SINO",
				"Scenario",
				"Outcome"
			];
			const headerRow_TestScenarios_Enhancement = new TableRow({
				children: columnTitles_TestScenarios_Enhancement.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_TestScenarios_Enhancement = objectDetails_TestScenarios_Enhancement.map(field => {
				return new TableRow({
					children: columnTitles_TestScenarios_Enhancement.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const TestScenariosDetailsTable_Enhancement = new Table({
				rows: [headerRow_TestScenarios_Enhancement, ...dataRows_TestScenarios_Enhancement]
			});
			documentChildrenFSD.push(TestScenariosDetailsTable_Enhancement);


			var samplesEnhancementTitle = new Paragraph({
				children: [
					new Bookmark({
						id: "SamplesEnhancement",
						children: [
							new TextRun({
								text: "Samples",
								size: 12 * 2, // Font size: 28
								bold: true,
								font: "Helvetica",
							}),
						],
					}),
				],
				spacing: {
					before: 250,
					after: 250,
				},
			});
			documentChildrenFSD.push(samplesEnhancementTitle);
			var samplesEnhancementDetails = parsedEnhancementResponse.response.Enhancement.Enhancement[0]['Samples'];
			var SamplesEnhancement_Data;
			if (samplesEnhancementDetails.trim().toUpperCase() === "NA") {
				SamplesEnhancement_Data = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2,
							}),
						],
						spacing: {
							after: 100,
						},
					}),
				];
			} else {
				var SamplesEnhancementArray = samplesEnhancementDetails
					.split(/(?<=\.)\s+/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SamplesEnhancement_Data = SamplesEnhancementArray.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica",
						}),
					],
					spacing: {
						after: 100,
					},
					alignment: AlignmentType.JUSTIFIED,
				}));
			}
			documentChildrenFSD.push(...SamplesEnhancement_Data);


			return documentChildrenFSD;
		},

			createFSDForm: function (response, documentChildrenFSD) {
			 
			const {
				Paragraph,
				Bookmark,
				TextRun,
				Table,
				TableRow,
				TableCell,
				AlignmentType,
				VerticalAlign,
			} = window.docx;
			var parsedData2 = response;

			var parsedData = JSON.parse(parsedData2)
			var forms = parsedData.response.Forms.Forms.Forms[0];

			var specificationDetails = forms["Specification Details"];
			var SpecificationDetailsArray = specificationDetails
				.split(/(?<=\.)\s+|\n(?=\d+\.\s)/)
				.map(line => line.trim())
				.filter(line => line.length > 0);
			var Form_Specification_Details = new Paragraph({
				children: [
					new TextRun({
						text: "Form Specification Details",
						bold: true,
						size: 16 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});

			var SpecificationDetailsParagraphs = SpecificationDetailsArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100,
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(Form_Specification_Details);
			documentChildrenFSD.push(...SpecificationDetailsParagraphs);

			
			var FunctionalRequirementsTitle = new Paragraph({
				children: [
					new TextRun({
						text: "Functional Requirements",
						bold: true,
						size: 14 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			documentChildrenFSD.push(FunctionalRequirementsTitle);

			

			var FunctionalRequirementsArray = parsedData.response.Forms.Forms2.Functional_Requirements;

const headers = [
    "Purpose:",
    "Details to Capture:",
    "Output Transmission:",
    "Static Content:",
    "Logic/Process Flow:",
    "Output Content:",
    "Special Note:"
];

let paragraphs_forms = [];
let currentParagraph = "";
let lastHeader = null;

// Split the text while focusing on all defined headers
FunctionalRequirementsArray.split(/\n+/).forEach(sentence => {
    const trimmedSentence = sentence.trim();
    const isHeader = headers.some(header => trimmedSentence.startsWith(header));

    if (isHeader) {
        // Push the previous paragraph before starting a new section
        if (currentParagraph.trim()) {
            paragraphs_forms.push({ text: currentParagraph.trim(), isHeader: false });
        }
        lastHeader = headers.find(header => trimmedSentence.startsWith(header)); // Get the actual header
        let content = trimmedSentence.replace(lastHeader, "").trim(); // Remove header from the text
        paragraphs_forms.push({ text: lastHeader, isHeader: true });
        currentParagraph = content ? content + " " : "";
    } else {
        currentParagraph += trimmedSentence + " ";
    }
});

// Push the last paragraph if there is any remaining text
if (currentParagraph.trim()) {
    paragraphs_forms.push({ text: currentParagraph.trim(), isHeader: false });
}

// Now the extracted paragraphs_forms will contain all sections
const functionalRequirementParagraphs_Forms = paragraphs_forms.map(item => {
    return new Paragraph({
        children: item.isHeader
            ? [
                  new TextRun({
                      text: item.text,
                      bold: true, // Apply bold to headers only
                      size: 8 * 2,
                      font: "Helvetica"
                  }),
                  new TextRun({ text: " ", size: 22 }) // Space after header
              ]
            : [
                  new TextRun({
                      text: item.text,
                      bold: false, // Normal text formatting for content
                      size: 8 * 2,
                      font: "Helvetica"
                  })
              ],
        spacing: {
            after: 150
        },
        alignment: AlignmentType.LEFT
    });
});

// Add the formatted paragraphs to the document
documentChildrenFSD.push(...functionalRequirementParagraphs_Forms);



			var triggers = forms["Triggers"];
			var TriggersArray = triggers
				// .split(/\n(?=\d+\.\s)/)
				.split(/(?=\d+\.\s)/)

				.map(line => line.trim())
				.filter(line => line.length > 0);
			var TriggersTitle = new Paragraph({
				children: [
					new TextRun({
						text: "Triggers",
						bold: true,
						size: 14 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 200
				}
			});
			var TriggersParagraphs = TriggersArray.map(line => new Paragraph({
				children: [
					new TextRun({
						text: line,
						size: 8 * 2,
						font: "Helvetica"
					})
				],
				spacing: {
					after: 100
				},
				alignment: AlignmentType.JUSTIFIED
			}));
			documentChildrenFSD.push(TriggersTitle);
			documentChildrenFSD.push(...TriggersParagraphs);
			var DataSelectionValidationProcessing = new Paragraph({
				children: [
					new Bookmark({
						id: "data_selection_validation_processing", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Data Selection, Validation & Processing",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});

			var dataSelectionValidationProcessing = parsedData.response.Forms.Forms.Forms[0]["Data Selection, Validation & Processing"];

			var details_0 = dataSelectionValidationProcessing.Details;
			var formField_0 = dataSelectionValidationProcessing["Form Field"];
			var formSection_0 = dataSelectionValidationProcessing["Form Section"];
			var sourceField_0 = dataSelectionValidationProcessing["Source Field"];
			var sourceObject_0 = dataSelectionValidationProcessing["Source Object"];

			var DataSelectionValidationProcessing_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Details",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Form Field",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Form Section",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Source Field",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Source Object",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: details_0,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: formField_0,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: formSection_0,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: sourceField_0,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: sourceObject_0,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});

			documentChildrenFSD.push(DataSelectionValidationProcessing);
			documentChildrenFSD.push(DataSelectionValidationProcessing_Table);

			var ExpectedOutcome = new Paragraph({
				children: [
					new Bookmark({
						id: "ExpectedOutcome",
						children: [
							new TextRun({
								text: "Expected Outcome", // Heading text
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var ExpectedOutcomeData = parsedData.response.Forms.Forms.Forms[0]["Expected Outcome"];
			var ExpectedOutcomeParagraphs;
			if (ExpectedOutcomeData.trim().toUpperCase() === "N/A") {
				ExpectedOutcomeParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var ExpectedOutcomeList = ExpectedOutcomeData
					.split(/(?<=\.)\s+|\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				ExpectedOutcomeParagraphs = ExpectedOutcomeList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(ExpectedOutcome);
			documentChildrenFSD.push(...ExpectedOutcomeParagraphs);

			var HardcopyRequirements = new Paragraph({
				children: [
					new Bookmark({
						id: "HardcopyRequirements",
						children: [
							new TextRun({
								text: "Hardcopy, Stationary & Printer Requirements",
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var HardcopyRequirementsData = forms["Hardcopy, Stationary & Printer Requirements"];
			var HardcopyRequirementsParagraphs;
			if (HardcopyRequirementsData.trim().toUpperCase() === "N/A") {
				HardcopyRequirementsParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var HardcopyRequirementsList = HardcopyRequirementsData
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				HardcopyRequirementsParagraphs = HardcopyRequirementsList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					}
				}));
			}
			documentChildrenFSD.push(HardcopyRequirements);
			documentChildrenFSD.push(...HardcopyRequirementsParagraphs);

			var NotificationsErrorHandling = new Paragraph({
				children: [
					new Bookmark({
						id: "NotificationsErrorHandling",
						children: [
							new TextRun({
								text: "Notifications & Error Handling", // Title without "13.12"
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			documentChildrenFSD.push(NotificationsErrorHandling);
			const formsNotificationsDetails = parsedData.response.Forms.Forms2.NotificationsErrorHandling;
			const normalizedFormsNotificationsDetails = Array.isArray(formsNotificationsDetails)
				? formsNotificationsDetails
				: [formsNotificationsDetails];
			const objectDetails_FormsNotifications = normalizedFormsNotificationsDetails.map(notification => ({
				"Error Code": notification["error_code"] || "N/A",
				"Message Type": notification["message_type"] || "N/A",
				"Description": notification["description"] || "N/A",
				"Trigger Event": notification["trigger_event"] || "N/A",
				"Display Location": notification["display_location"] || "N/A"
			}));
			const columnTitles_FormsNotifications = [
				"Error Code",
				"Message Type",
				"Description",
				"Trigger Event",
				"Display Location"
			];
			const headerRow_FormsNotifications = new TableRow({
				children: columnTitles_FormsNotifications.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2, // Font size: 20
										font: "Helvetica",
									}),
								],
							}),
						],
						shading: {
							fill: "#e5e7eb", // Light gray background for the header
						},
					})
				),
			});
			const dataRows_FormsNotifications = objectDetails_FormsNotifications.map(field => {
				return new TableRow({
					children: columnTitles_FormsNotifications.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "", // Handle missing values gracefully
											size: 8 * 2, // Font size: 20
											font: "Helvetica",
										}),
									],
								}),
							],
						})
					),
				});
			});
			const FormsNotificationsDetailsTable = new Table({
				rows: [headerRow_FormsNotifications, ...dataRows_FormsNotifications]
			});
			documentChildrenFSD.push(FormsNotificationsDetailsTable);



			var FurtherConstraints = new Paragraph({
				children: [
					new Bookmark({
						id: "FurtherConstraints",
						children: [
							new TextRun({
								text: "Further Constraints & Special Considerations", // Removed "13.12"
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var FurtherConstraintsData = forms["Further Constraints & Special Considerations"];
			var FurtherConstraintsParagraphs;
			if (FurtherConstraintsData.trim().toUpperCase() === "N/A") {
				FurtherConstraintsParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var FurtherConstraintsList = FurtherConstraintsData
					.split(/\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				FurtherConstraintsParagraphs = FurtherConstraintsList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2,
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100
					}
				}));
			}
			documentChildrenFSD.push(FurtherConstraints);
			documentChildrenFSD.push(...FurtherConstraintsParagraphs);

			var SecurityRequirements = new Paragraph({
				children: [
					new Bookmark({
						id: "SecurityRequirements",
						children: [
							new TextRun({
								text: "Security & Role Requirements, Sensitive Data", // Heading text
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var SecurityRequirementsData = forms["Security & Role Requirements, Sensitive Data"];
			var SecurityRequirementsParagraphs;
			if (SecurityRequirementsData.trim().toUpperCase() === "N/A") {
				SecurityRequirementsParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var SecurityRequirementsList = SecurityRequirementsData
					.split(/\n(?=\d+\.\s)/) // Split by numbers followed by a period and newline
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SecurityRequirementsParagraphs = SecurityRequirementsList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(SecurityRequirements);
			documentChildrenFSD.push(...SecurityRequirementsParagraphs);

			var DataVolumes_PerformanceConsiderations = new Paragraph({
				children: [
					new Bookmark({
						id: "data_volumes_performance_considerations",
						children: [
							new TextRun({
								text: "Data Volumes & Performance Considerations",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var dataVolumesData = parsedData.response.Forms.Forms.Forms[0]["Data Volumes & Performance Considerations"];
			var messageSize = dataVolumesData["Message Size"];
			var initialLoadVolume = dataVolumesData["Initial Load Volume"];
			var averageVolume = dataVolumesData["Average Volume"];
			var peakVolume = dataVolumesData["Peak Volume"];
			var performanceMetrics = dataVolumesData["Performance Metrics"];
			var additionalInfo = dataVolumesData["Additional Information"];
			var DataVolumes_PerformanceConsiderations_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Message Size",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Initial Load Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Average Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Peak Volume",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Performance Metrics",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: messageSize,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: initialLoadVolume,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: averageVolume,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: peakVolume,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: performanceMetrics,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: additionalInfo,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});
			documentChildrenFSD.push(DataVolumes_PerformanceConsiderations);
			documentChildrenFSD.push(DataVolumes_PerformanceConsiderations_Table);

			var SynchronousProcessingData_heading = new Paragraph({
				children: [
					new Bookmark({
						id: "Synchronous Processing, Response Times", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Synchronous Processing, Response Times",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			documentChildrenFSD.push(SynchronousProcessingData_heading);
			var SynchronousProcessingData = parsedData.response.Forms.Forms.Forms[0]["Synchronous Processing, Response Times"];
			var SynchronousProcessingParagraphs;
			if (SynchronousProcessingData.trim().toUpperCase() === "N/A") {
				SynchronousProcessingParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "N/A",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var SynchronousProcessingList = SynchronousProcessingData
					.split(/(?<=\.)\s+|\n(?=\d+\.\s)/)
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SynchronousProcessingParagraphs = SynchronousProcessingList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(...SynchronousProcessingParagraphs);

			var BackgroundProcessing = new Paragraph({
				children: [
					new Bookmark({
						id: "background_processing", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Background Processing",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			var backgroundProcessingData = parsedData.response.Forms.Forms.Forms[0]["Background Processing"];
			var additionalInformation = backgroundProcessingData["Additional Information"];
			var backgroundJob = backgroundProcessingData["Background Job"];
			var criticallyErrorHandling = backgroundProcessingData["Critically, Error Handling"];
			var description = backgroundProcessingData["Description"];
			var notificationDistributions = backgroundProcessingData["Notification, Distributions"];
			var recurrence = backgroundProcessingData["Recurrence"];
			var startConditionsTiming = backgroundProcessingData["Start Conditions, Timing"];
			var BackgroundProcessing_Table = new Table({
				rows: [
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Background Job",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Description",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Start Conditions, Timing",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Recurrence",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Criticality, Error Handling",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Notification, Distribution",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: "Additional Information",
												bold: true,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								shading: {
									fill: "#e5e7eb"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					}),
					new TableRow({
						children: [
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: backgroundJob,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100 // Very small left padding
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: description,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: startConditionsTiming,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: recurrence,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: criticallyErrorHandling,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: notificationDistributions,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							}),
							new TableCell({
								children: [
									new Paragraph({
										children: [
											new TextRun({
												text: additionalInformation,
												size: 8 * 2,
												font: "Helvetica"
											})
										],
										alignment: AlignmentType.LEFT,
										indent: {
											left: 100
										}
									})
								],
								width: {
									size: 20,
									type: "pct"
								},
								verticalAlign: VerticalAlign.CENTER
							})
						],
						height: {
							value: 500,
							rule: "atLeast"
						}
					})
				]
			});
			documentChildrenFSD.push(BackgroundProcessing);
			documentChildrenFSD.push(BackgroundProcessing_Table);



			var TestScenarios = new Paragraph({
				children: [
					new Bookmark({
						id: "test_scenarios", // Create the bookmark here inside the paragraph
						children: [
							new TextRun({
								text: "Test Scenarios, Test Data, Pre-conditions",
								bold: true,
								size: 14 * 2,
								font: "Helvetica"
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 100
				}
			});
			documentChildrenFSD.push(TestScenarios);
			const testScenariosData = parsedData.response.Forms.Forms2.TestData;
			const normalizedTestScenariosData = Array.isArray(testScenariosData)
				? testScenariosData
				: [testScenariosData];
			const objectDetails_TestScenarios = normalizedTestScenariosData.map(item => ({

				"SINO": item["SINO"] || "N/A",
				"Scenario": item["Scenario"] || "N/A",
				"Outcome": item["Outcome"] || "N/A"
			}));
			const columnTitles_TestScenarios = ["SINO", "Scenario", "Outcome"];

			const headerRow_TestScenarios = new TableRow({
				children: columnTitles_TestScenarios.map(title =>
					new TableCell({
						children: [
							new Paragraph({
								children: [
									new TextRun({
										text: title,
										bold: true,
										size: 8 * 2,
										font: "Helvetica"
									})
								],
								alignment: AlignmentType.LEFT,
								indent: {
									left: 100 // Very small left padding
								}
							})
						],
						width: {
							size: 33,


							type: "pct"
						},
						shading: {
							fill: "#e5e7eb"
						},
						verticalAlign: VerticalAlign.CENTER
					})
				),
				height: {
					value: 500,
					rule: "atLeast"
				}
			});
			const dataRows_TestScenarios = objectDetails_TestScenarios.map(field => {
				return new TableRow({
					children: columnTitles_TestScenarios.map(title =>
						new TableCell({
							children: [
								new Paragraph({
									children: [
										new TextRun({
											text: field[title] || "N/A", // Handle missing values gracefully
											size: 8 * 2,
											font: "Helvetica"
										})
									],
									alignment: AlignmentType.LEFT,
									indent: {
										left: 100
									}
								})
							],
							width: {
								size: 33,

								type: "pct"
							},
							verticalAlign: VerticalAlign.CENTER
						})
					),
					height: {
						value: 500,
						rule: "atLeast"
					}
				});
			});
			const TestScenarios_Table = new Table({
				rows: [headerRow_TestScenarios, ...dataRows_TestScenarios]
			});
			documentChildrenFSD.push(TestScenarios_Table);


			var Samples = new Paragraph({
				children: [
					new Bookmark({
						id: "Samples",
						children: [
							new TextRun({
								text: "Samples Forms", // Heading text
								font: "Helvetica",
								bold: true,
								size: 14 * 2
							})
						]
					})
				],
				spacing: {
					before: 250,
					after: 200
				}
			});
			var SamplesData = forms["Samples"];
			var SamplesParagraphs;
			if (SamplesData.trim().toUpperCase() === "NA") {
				SamplesParagraphs = [
					new Paragraph({
						children: [
							new TextRun({
								text: "NA",
								font: "Helvetica",
								size: 8 * 2
							})
						],
						spacing: {
							after: 100
						}
					})
				];
			} else {
				var SamplesList = SamplesData
					.split(/\n(?=\d+\.\s)/) // Split by numbers followed by a period and newline
					.map(line => line.trim())
					.filter(line => line.length > 0);

				SamplesParagraphs = SamplesList.map(line => new Paragraph({
					children: [
						new TextRun({
							text: line,
							size: 8 * 2, // Font size: 16
							font: "Helvetica"
						})
					],
					spacing: {
						after: 100 // Space between paragraphs
					}
				}));
			}
			documentChildrenFSD.push(Samples);
			documentChildrenFSD.push(...SamplesParagraphs);


			// 	AdditionalInfoParagraphs = AdditionalInfoList.map(line => new Paragraph({
			// 		children: [
			// 			new TextRun({
			// 				text: line,
			// 				size: 8 * 2,
			// 				font: "Helvetica"

			return documentChildrenFSD;
		},
	};
});