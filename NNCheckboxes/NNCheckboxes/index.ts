import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { SpawnSyncOptionsWithBufferEncoding } from "child_process";
import { RelationshipInfo } from "./reldef";
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class NNCheckboxes implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Reference to the control container HTMLDivElement
	// This element contains all elements of our custom control example
	private _container: HTMLDivElement;
	private _divFlexBox: HTMLDivElement;
	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// Event Handler 'refreshData' reference
	private _parentRecordId: string;
	private _parentRecordType: string;
	private _childRecordType: string;
	private _labelAttributeName: string;
	private _backgroundColorAttributeName: string;
	private _backgroundColorIsFromOptionSet: boolean;
	private _foreColorAttributeName: string;
	private _foreColorIsFromOptionSet: boolean;
	private _numberOfColumns: number;
	private _categoryAttributeName: string;
	private _categoryUseDisplayName: boolean;
	private _useToggleSwitch: boolean;
	private _relationshipSchemaName: string | null;
	private _colors: any;
	private _relationshipInfo: RelationshipInfo;
	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		// Add control initialization code
		this._context = context;
		this._container = document.createElement("div");
		this._container.setAttribute("class", "nncb-main")
		this._useToggleSwitch = (context.parameters.useToggleSwitch && context.parameters.useToggleSwitch.raw && context.parameters.useToggleSwitch.raw.toLowerCase() === "true") ? true : false;
		// @ts-ignore
		this._parentRecordId = context.mode.contextInfo.entityId;
		// @ts-ignore
		this._parentRecordType = context.mode.contextInfo.entityTypeName;


		if (context.parameters.toggleDefaultBackgroundColorOn && context.parameters.toggleDefaultBackgroundColorOn.raw) {
			let styleSheet: any = this.GetStyleSheet();
			if (styleSheet != null) {
				let rules: any = styleSheet.rules;
				for (let i = 0; i < rules.length; i++) {
					let rule = rules[i];

					if (rule.selectorText === "input:checked + .nncb-slider") {
						styleSheet.deleteRule(i);
						styleSheet.insertRule("input:checked + .nncb-slider { background-color: " + context.parameters.toggleDefaultBackgroundColorOn.raw + ";}", rule.index)
					}
				}
			}
		}

		if (context.parameters.toggleDefaultBackgroundColorOff && context.parameters.toggleDefaultBackgroundColorOff.raw) {
			let styleSheet: any = this.GetStyleSheet();
			if (styleSheet != null) {
				let rules = styleSheet.rules;
				for (let i = 0; i < rules.length; i++) {
					let rule = rules[i];

					if (rule.selectorText === ".nncb-slider") {
						styleSheet.deleteRule(i);
						styleSheet.insertRule(".nncb-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: " + context.parameters.toggleDefaultBackgroundColorOff.raw + "; -webkit-transition: .4s; transition: .4s;", rule.index)
					}
				}
			}
		}

		container.appendChild(this._container);

		this._childRecordType = context.parameters.nnRelationshipDataSet.getTargetEntityType();
		this._numberOfColumns = context.parameters.columnsNumber ? context.parameters.columnsNumber.raw : 1;

		for (var i = 0; i < context.parameters.nnRelationshipDataSet.columns.length; i++) {
			var column = context.parameters.nnRelationshipDataSet.columns[i];
			if (column.alias === "displayAttribute") {
				this._labelAttributeName = column.name;
			}
			else if (column.alias === "backgroundColorAttribute") {
				this._backgroundColorAttributeName = column.name;

				if (column.dataType === "OptionSet" || column.dataType === "" && column.name === "statuscode") {
					if (!this._colors || !this._colors[column.name]) {
						await this.GetOptionSetColors(column.name);
					}

					this._backgroundColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === "foreColorAttribute") {
				this._foreColorAttributeName = column.name;

				if (column.dataType === "OptionSet" || column.dataType === "" && column.name === "statuscode") {
					if (!this._colors || !this._colors[column.name]) {
						await this.GetOptionSetColors(column.name);
					}

					this._foreColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === "categoryAttribute") {
				this._categoryAttributeName = column.name;
				this._categoryUseDisplayName = column.dataType === "Lookup.Simple" || column.dataType === "OptionSet" || column.dataType === "TwoOptions" || column.dataType === "" && column.name === "statuscode";
			}
		}

		// If no category grouping, then only one flexbox is needed
		if (!this._categoryAttributeName) {
			this._divFlexBox = document.createElement("div");
			this._divFlexBox.setAttribute("class", "nncb-flex");
			this._container.appendChild(this._divFlexBox);
		}

		try {
			this._relationshipSchemaName = await this.GetNNRelationshipNameByEntityNames();
		}
		catch (error) {
			this._context.navigation.openAlertDialog(
				{
					text: "NNCheckboxes " + error
				});
			return;
		}

		var thisCtrl = this;

		context.webAPI.retrieveRecord("savedquery", context.parameters.nnRelationshipDataSet.getViewId(), "?$select=fetchxml")
			.then(
				function (view) {
					context.webAPI.retrieveMultipleRecords(thisCtrl._childRecordType, "?fetchXml=" + encodeURIComponent(<string>view.fetchxml)
					)
						.then(function (result) {
							var category = "";
							var divFlexCtrl = document.createElement("div");

							for (var i = 0; i < result.entities.length; i++) {
								var record = result.entities[i];

								// If using category
								if (thisCtrl._categoryAttributeName) {
									// We need to display new category only if the category 
									// is different from the previous one
									if (category != record[thisCtrl._categoryAttributeName]) {
										category = record[thisCtrl._categoryAttributeName];

										let label = record[thisCtrl._categoryAttributeName + (thisCtrl._categoryUseDisplayName ? "@OData.Community.Display.V1.FormattedValue" : "")];
										if (!label) {
											label = context.resources.getString("No_Category");
										}

										// Add the category
										var categoryDiv = document.createElement("div");
										categoryDiv.setAttribute("style", "margin-bottom: 10px;border-bottom: solid 1px #828181;padding-bottom: 5px;");
										categoryDiv.innerHTML = label;

										thisCtrl._container.appendChild(categoryDiv);

										// Add a new flex box
										thisCtrl._divFlexBox = document.createElement("div");
										thisCtrl._divFlexBox.setAttribute("class", "nncb-flex");
										thisCtrl._container.appendChild(thisCtrl._divFlexBox);
									}
								}

								// Add flex content
								divFlexCtrl = document.createElement("div");
								divFlexCtrl.setAttribute("style", "flex: 0 " + (100 / thisCtrl._numberOfColumns) + "% !important");
								thisCtrl._divFlexBox.appendChild(divFlexCtrl);

								// With style if configured with colors
								var styles = new Array();
								if (thisCtrl._backgroundColorAttributeName) {
									if (thisCtrl._backgroundColorIsFromOptionSet) {
										if (thisCtrl._colors
											&& thisCtrl._colors[thisCtrl._backgroundColorAttributeName]
											&& thisCtrl._colors[thisCtrl._backgroundColorAttributeName][record[thisCtrl._backgroundColorAttributeName]]) {
											var color = thisCtrl._colors[thisCtrl._backgroundColorAttributeName][record[thisCtrl._backgroundColorAttributeName]];
											styles.push("background-color:" + color);
										}
									}
									else {
										styles.push("background-color:" + record[thisCtrl._backgroundColorAttributeName]);
									}
								}
								if (thisCtrl._foreColorAttributeName) {
									if (thisCtrl._foreColorIsFromOptionSet) {
										if (thisCtrl._colors
											&& thisCtrl._colors[thisCtrl._foreColorAttributeName]
											&& thisCtrl._colors[thisCtrl._foreColorAttributeName][record[thisCtrl._backgroundColorAttributeName]]) {
											var color = thisCtrl._colors[thisCtrl._foreColorAttributeName][record[thisCtrl._backgroundColorAttributeName]];
											styles.push("color:" + color);
										}
									}
									else {
										styles.push("color:" + record[thisCtrl._foreColorAttributeName]);
									}
								}

								var lblContainer = document.createElement("label");
								divFlexCtrl.appendChild(lblContainer);

								if (thisCtrl._useToggleSwitch) {
									lblContainer.setAttribute("class", "nncb-container-switch");

									var spanLabel = document.createElement("span");
									spanLabel.setAttribute("class", "nncb-switch-label");
									spanLabel.textContent = record[thisCtrl._labelAttributeName];
									divFlexCtrl.appendChild(spanLabel);
								}
								else {
									lblContainer.setAttribute("class", "nncb-container");
									lblContainer.setAttribute("style", styles.join(";"))
								}

								var chk = document.createElement("input");
								chk.setAttribute("type", "checkbox");
								chk.setAttribute("id", thisCtrl._relationshipInfo.Name + "|" + record[thisCtrl._childRecordType + "id"]);
								chk.setAttribute("value", thisCtrl._relationshipInfo.Name + "|" + record[thisCtrl._childRecordType + "id"]);
								chk.addEventListener("change", function () {
									let entity1name: string;
									let entity2name: string;
									let record1Id: string;
									let record2Id: string;
									if (thisCtrl._relationshipInfo.Entity1AttributeName === thisCtrl._childRecordType) {
										entity1name = thisCtrl._childRecordType;
										record1Id = this.id.split('|')[1];
										entity2name = thisCtrl._parentRecordType;
										record2Id = thisCtrl._parentRecordId;
									}
									else {
										entity1name = thisCtrl._parentRecordType;
										record1Id = thisCtrl._parentRecordId;
										entity2name = thisCtrl._childRecordType;
										record2Id = this.id.split('|')[1];
									}

									if (this.checked) {
										var associateRequest = new class {
											target = {
												id: record1Id,
												entityType: entity1name
											};
											relatedEntities = [
												{
													id: record2Id,
													entityType: entity2name
												}
											];
											relationship = thisCtrl._relationshipInfo.Name;
											getMetadata(): any {
												return {
													boundParameter: undefined,
													parameterTypes: {
														"target": {
															"typeName": "mscrm." + entity1name,
															"structuralProperty": 5
														},
														"relatedEntities": {
															"typeName": "mscrm." + entity2name,
															"structuralProperty": 4
														},
														"relationship": {
															"typeName": "Edm.String",
															"structuralProperty": 1
														}
													},
													operationType: 2,
													operationName: "Associate"
												};
											}
										}();

										// @ts-ignore
										thisCtrl._context.webAPI.execute(associateRequest)
											.then(
												// @ts-ignore
												function (result) {
													console.log("NNCheckboxes: records were successfully associated")
												},
												// @ts-ignore
												function (error) {
													thisCtrl._context.navigation.openAlertDialog({ text: "An error occured when associating records. Please check NNCheckboxes control configuration" });
												}
											);
									}
									else {
										var theRecordId = this.id;
										var disassociateRequest = new class {
											target = {
												id: record1Id,
												entityType: entity1name
											};
											relatedEntityId = record2Id;
											relationship = thisCtrl._relationshipInfo.Name;
											getMetadata(): any {
												return {
													boundParameter: undefined,
													parameterTypes: {
														"target": {
															"typeName": "mscrm." + entity1name,
															"structuralProperty": 5
														},
														"relationship": {
															"typeName": "Edm.String",
															"structuralProperty": 1
														}
													},
													operationType: 2,
													operationName: "Disassociate"
												};
											}
										}();

										// @ts-ignore
										thisCtrl._context.webAPI.execute(disassociateRequest)
											.then(
												// @ts-ignore
												function (result) {
													console.log("NNCheckboxes: records were successfully disassociated")
												},
												// @ts-ignore
												function (error) {
													thisCtrl._context.navigation.openAlertDialog({ text: "An error occured when disassociating records. Please check NNCheckboxes control configuration" });
												}
											);
									}
								});

								if (context.mode.isControlDisabled) {
									chk.setAttribute("disabled", "disabled");
								}

								if (thisCtrl._useToggleSwitch) {
									var toggle = document.createElement("span");
									toggle.setAttribute("class", "nncb-slider nncb-round");

									if (styles.length > 0)
										toggle.setAttribute("style", styles.join(";"))

									lblContainer.appendChild(chk);
									lblContainer.appendChild(toggle);
								}
								else {
									var mark = document.createElement("span");
									mark.setAttribute("class", "nncb-checkmark");

									lblContainer.innerHTML += record[thisCtrl._labelAttributeName];
									lblContainer.appendChild(chk);
									lblContainer.appendChild(mark);
								}
							}

							thisCtrl._context.parameters.nnRelationshipDataSet.paging.reset();
							thisCtrl._context.parameters.nnRelationshipDataSet.refresh();
						},
							function (error) {
								thisCtrl._context.navigation.openAlertDialog(
									{
										text: "NNCheckboxes: " + error
									});
							}
						);
				},
				function (error) {
					thisCtrl._context.navigation.openAlertDialog(
						{
							text: "NNCheckboxes: " + error
						});
				}
			);
	}
	public GetStyleSheet() {
		for (var i = 0; i < document.styleSheets.length; i++) {
			var sheet = document.styleSheets[i];
			if (sheet.href && sheet.href.endsWith("NNCheckboxes.css")) {
				return sheet;
			}
		}
		return null;
	}

	public async GetOptionSetColors(attribute: string) {
		let requestUrl =
			"/api/data/v9.0/EntityDefinitions(LogicalName='" + this._childRecordType + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$filter=LogicalName eq '" + attribute + "'&$expand=OptionSet";

		var thisCtrl = this;
		let request = new XMLHttpRequest();
		request.open("GET", requestUrl, true);
		request.setRequestHeader("OData-MaxVersion", "4.0");
		request.setRequestHeader("OData-Version", "4.0");
		request.setRequestHeader("Accept", "application/json");
		request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		request.onreadystatechange = () => {
			if (request.readyState === 4) {
				request.onreadystatechange = null;
				if (request.status === 200) {
					let entityMetadata = JSON.parse(request.response);
					let options = entityMetadata.value[0].OptionSet.Options;
					thisCtrl._colors = {};
					thisCtrl._colors[attribute] = {}
					for (var i = 0; i < options.length; i++) {
						thisCtrl._colors[attribute][options[i].Value] = options[i].Color;
					}
				} else {
					let errorText = request.responseText;
					console.log(errorText);
				}
			}
		};
		request.send();
	}

	public async GetNNRelationshipNameByEntityNames() {
		let schemaNameParameter = this._context.parameters.relationshipSchemaName;
		if (schemaNameParameter != undefined && schemaNameParameter.raw != null) {
			let entityMetadata = await this._context.utils.getEntityMetadata(this._parentRecordType);
			let nnRelationships = entityMetadata.ManyToManyRelationships.getAll();

			for (let i = 0; i < nnRelationships.length; i++) {
				if (nnRelationships[i].IntersectEntityName.toLowerCase() === this._context.parameters.relationshipSchemaName.raw.toLowerCase()) {
					this._relationshipInfo = new RelationshipInfo();
					this._relationshipInfo.Entity1LogicalName = nnRelationships[i].Entity1LogicalName;
					this._relationshipInfo.Entity1AttributeName = nnRelationships[i].Entity1IntersectAttribute;
					this._relationshipInfo.Entity2LogicalName = nnRelationships[i].Entity2LogicalName;
					this._relationshipInfo.Entity2AttributeName = nnRelationships[i].Entity2IntersectAttribute;
					this._relationshipInfo.Name = nnRelationships[i].SchemaName;
				}
			}

			return Promise.resolve(<string>schemaNameParameter.raw);
		}
		let entityMetadata = await this._context.utils.getEntityMetadata(this._parentRecordType);
		let nnRelationships = entityMetadata.ManyToManyRelationships.getAll();

		let count = 0;
		let foundSchemaName = "";

		for (let i = 0; i < nnRelationships.length; i++) {
			if ((nnRelationships[i].Entity1LogicalName == this._parentRecordType && nnRelationships[i].Entity2LogicalName == this._childRecordType) ||
				(nnRelationships[i].Entity1LogicalName == this._childRecordType && nnRelationships[i].Entity2LogicalName == this._parentRecordType)) {
				count++;
				foundSchemaName = nnRelationships[i].SchemaName;

				this._relationshipInfo = new RelationshipInfo();
				this._relationshipInfo.Entity1LogicalName = nnRelationships[i].Entity1LogicalName;
				this._relationshipInfo.Entity1AttributeName = nnRelationships[i].Entity1IntersectAttribute;
				this._relationshipInfo.Entity2LogicalName = nnRelationships[i].Entity2LogicalName;
				this._relationshipInfo.Entity2AttributeName = nnRelationships[i].Entity2IntersectAttribute;
				this._relationshipInfo.Name = nnRelationships[i].SchemaName;
			}
		}

		if (foundSchemaName.length === 0) {
			return Promise.reject(new Error(this._context.resources.getString("No_Relationship_Found")));
		}
		if (count > 1) {
			return Promise.reject(new Error(this._context.resources.getString("Multiple_Relationships_Found")));
		}

		return Promise.resolve(<string>foundSchemaName);
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		if (!context.updatedProperties.includes("nnRelationshipDataSet")) return;

		if (context.parameters.nnRelationshipDataSet.paging.hasNextPage) {
			context.parameters.nnRelationshipDataSet.paging.loadNextPage();
			return;
		}

		var selectedIds = context.parameters.nnRelationshipDataSet.sortedRecordIds;
		for (var j = 0; j < selectedIds.length; j++) {
			let chk = <HTMLInputElement>window.document.getElementById(this._relationshipInfo.Name + "|" + selectedIds[j]);
			if (chk) {
				chk.checked = true;
			}
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}

}