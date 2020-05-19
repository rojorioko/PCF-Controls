import { IDataSetRecord } from './ts/interface/datasetrecord.interface';
import { Constant } from './ts/constant/constant';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
// import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
// import { SpawnSyncOptionsWithBufferEncoding } from "child_process";
// import { stringify } from 'querystring';
import { IAttributeValue } from './ts/interface/attributevalue.interface';
import { IRelationDefinition } from './ts/interface/relationdefinition.interface';
import { INNRelationshipInfo } from './ts/interface/nnrelationshipinfo.interface';
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class ListCheckboxes implements ComponentFramework.StandardControl<IInputs, IOutputs> {

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
	private _colors: any;
	private _relationshipInfo: INNRelationshipInfo;
	private _useCustomRelationship: boolean;
	private _customRelationshipDefinitionChild: IRelationDefinition;
	private _customRelationshipDefinitionCurrent: IRelationDefinition;
	private _changedCheckboxes: HTMLInputElement[] = [];

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
		this._useCustomRelationship = (context.parameters.useCustomIntersect && context.parameters.useCustomIntersect.raw.toLowerCase() === 'true') ? true : false;

		if (!this._parametersAreValid()) return;

		if (this._useCustomRelationship) {
			this._customRelationshipDefinitionChild = await this._getRelationshipDefinitionByName(this._context.parameters.customIntersectChildRelationship.raw);
			this._customRelationshipDefinitionCurrent = await this._getRelationshipDefinitionByName(this._context.parameters.relationshipSchemaName.raw);
		}

		this._container = document.createElement("div");
		this._container.setAttribute("class", Constant.NncbMain);
		container.appendChild(this._container);

		this._setToggleDefaultBackgroud();
		this._extractDataSetParameters();

		// If no category grouping, then only one flexbox is needed
		if (!this._categoryAttributeName) {
			this._divFlexBox = document.createElement("div");
			this._divFlexBox.setAttribute("class", Constant.NncbFlex);
			this._container.appendChild(this._divFlexBox);
		}

		try {
			let saveQuery: ComponentFramework.WebApi.Entity = await this._getViewFetchXML();

			if (!this._useCustomRelationship)
				await this._setRelationshipInformation();

			let result: ComponentFramework.WebApi.RetrieveMultipleResponse = await context.webAPI.retrieveMultipleRecords(saveQuery.returnedtypecode, "?fetchXml=" + encodeURIComponent(saveQuery.fetchxml));

			if (this._useCustomRelationship)
				this._displayViewRecordsCustom(result)
			else
				this._displayViewRecordsDefault(result);
		}
		catch (error) {
			this._showAlertMessage("Checkboxes Control " + error);
			return;
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		if (!context.updatedProperties.includes(Constant.DatasetName)) return;

		if (context.parameters.nnRelationshipDataSet.paging.hasNextPage) {
			context.parameters.nnRelationshipDataSet.paging.loadNextPage();
			return;
		}

		if (this._changedCheckboxes.length) {
			this._changedCheckboxes.forEach(chx => {
				chx.disabled = false;
			});

			this._changedCheckboxes.length = 0;
		}

		let dataSet = context.parameters.nnRelationshipDataSet;

		for (var j = 0; j < dataSet.sortedRecordIds.length; j++) {
			let matchId: string | null = dataSet.sortedRecordIds[j];
			if (this._useCustomRelationship) {
				let records: IDataSetRecord[] = dataSet.sortedRecordIds.map(r => ({
					key: r,
					values: dataSet.columns.map(c => {
						let ikeyValue: IAttributeValue = {
							key: c.name,
							value: dataSet.records[r].getFormattedValue(c.name)
						}

						let fieldValue: any = dataSet.records[r].getValue(c.name);
						if (c.name === this._customRelationshipDefinitionChild.ReferencingAttribute && fieldValue) {
							ikeyValue.value = {
								id: fieldValue.id.guid,
								name: fieldValue.name,
								entityType: fieldValue.etn
							};
						}

						return ikeyValue;
					})
				}));
				let found = records.find(record => record.key === dataSet.sortedRecordIds[j] && record.values.find(data => data.key === this._customRelationshipDefinitionChild.ReferencingAttribute));
				let lookupValue = found ? <ComponentFramework.EntityReference>found.values.find(fieldValue => fieldValue.key === this._customRelationshipDefinitionChild.ReferencingAttribute)!.value : undefined;
				matchId = lookupValue ? lookupValue.id : null;
			}

			let chk = matchId ? <HTMLInputElement>window.document.getElementById(matchId) : null;

			if (chk) {
				chk.checked = true;
				chk.setAttribute(Constant.IntersectData, dataSet.sortedRecordIds[j]);
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

	//#region Private Functions
	private async _getRelationshipDefinitionByName(schemaName: string): Promise<any> {
		//@ts-ignore
		let requestUrl = `${this._context.page.getClientUrl()}/api/data/v9.1/RelationshipDefinitions(SchemaName='${schemaName}')`;
		console.log("SCHEMA " + requestUrl);
		let result: any = {};
		let request = new XMLHttpRequest();
		// Return it as a Promise
		return new Promise(function (resolve, reject) {
			// Setup our listener to process compeleted requests
			request.onreadystatechange = function () {
				if (request.readyState === 4) {
					request.onreadystatechange = null;
					if (request.status === 200) {
						result = JSON.parse(this.response);
					} else {
						let errorText = request.responseText;
						reject(new Error(errorText));
					}

					resolve(result);
				}
			};

			request.open("GET", requestUrl, true);
			request.setRequestHeader("OData-MaxVersion", "4.0");
			request.setRequestHeader("OData-Version", "4.0");
			request.setRequestHeader("Accept", "application/json");
			request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			// Send the request
			request.send();
		});
	}
	private async _getViewFetchXML(): Promise<ComponentFramework.WebApi.Entity> {
		let saveQuery: ComponentFramework.WebApi.Entity | null;

		if (this._useCustomRelationship) {
			if (this._context.parameters.customIntersectChildFetchXML.raw) {
				saveQuery = {
					"returnedtypecode": this._customRelationshipDefinitionChild.ReferencedEntity,
					"fetchxml": this._context.parameters.customIntersectChildFetchXML.raw
				};
			}
			else {
				let queryOption: string = `?$select=fetchxml,returnedtypecode&$filter=name eq '${this._context.parameters.customIntersectChildView.raw}'`;
				let result: ComponentFramework.WebApi.RetrieveMultipleResponse = await this._context.webAPI.retrieveMultipleRecords(Constant.SaveQuery, queryOption);
				saveQuery = result.entities && result.entities.length > 0 ? result.entities[0] : null;
			}
		} else {
			saveQuery = await this._context.webAPI.retrieveRecord(Constant.SaveQuery, this._context.parameters.nnRelationshipDataSet.getViewId(), "?$select=fetchxml,returnedtypecode");
		}

		if (saveQuery != null && saveQuery.fetchxml) {
			return Promise.resolve(saveQuery);
		}
		else {
			let viewNotFound = this._context.resources.getString(Constant.ViewNotFound);
			return Promise.reject(new Error(viewNotFound));
		}
	}

	private _displayViewRecordsDefault(result: ComponentFramework.WebApi.RetrieveMultipleResponse) {
		var category = "";
		var divFlexCtrl = document.createElement("div");

		for (var i = 0; i < result.entities.length; i++) {
			var record = result.entities[i];

			// If using category
			if (this._categoryAttributeName) {
				// We need to display new category only if the category 
				// is different from the previous one
				if (category != record[this._categoryAttributeName]) {
					category = record[this._categoryAttributeName];

					let label = record[this._categoryAttributeName + (this._categoryUseDisplayName ? "@OData.Community.Display.V1.FormattedValue" : "")];
					if (!label) {
						label = this._context.resources.getString("No_Category");
					}

					// Add the category
					var categoryDiv = document.createElement("div");
					categoryDiv.setAttribute("style", "margin-bottom: 10px;border-bottom: solid 1px #828181;padding-bottom: 5px;");
					categoryDiv.innerHTML = label;

					this._container.appendChild(categoryDiv);

					// Add a new flex box
					this._divFlexBox = document.createElement("div");
					this._divFlexBox.setAttribute("class", Constant.NncbFlex);
					this._container.appendChild(this._divFlexBox);
				}
			}

			// Add flex content
			divFlexCtrl = document.createElement("div");
			divFlexCtrl.setAttribute("style", "flex: 0 " + (100 / this._numberOfColumns) + "% !important");
			this._divFlexBox.appendChild(divFlexCtrl);

			// With style if configured with colors
			var styles = new Array();
			if (this._backgroundColorAttributeName) {
				if (this._backgroundColorIsFromOptionSet) {
					if (this._colors
						&& this._colors[this._backgroundColorAttributeName]
						&& this._colors[this._backgroundColorAttributeName][record[this._backgroundColorAttributeName]]) {
						var color = this._colors[this._backgroundColorAttributeName][record[this._backgroundColorAttributeName]];
						styles.push("background-color:" + color);
					}
				}
				else {
					styles.push("background-color:" + record[this._backgroundColorAttributeName]);
				}
			}
			if (this._foreColorAttributeName) {
				if (this._foreColorIsFromOptionSet) {
					if (this._colors
						&& this._colors[this._foreColorAttributeName]
						&& this._colors[this._foreColorAttributeName][record[this._backgroundColorAttributeName]]) {
						var color = this._colors[this._foreColorAttributeName][record[this._backgroundColorAttributeName]];
						styles.push("color:" + color);
					}
				}
				else {
					styles.push("color:" + record[this._foreColorAttributeName]);
				}
			}

			var lblContainer = document.createElement("label");
			divFlexCtrl.appendChild(lblContainer);

			if (this._useToggleSwitch) {
				lblContainer.setAttribute("class", "nncb-container-switch");

				var spanLabel = document.createElement("span");
				spanLabel.setAttribute("class", "nncb-switch-label");
				spanLabel.textContent = record[this._labelAttributeName];
				divFlexCtrl.appendChild(spanLabel);
			}
			else {
				lblContainer.setAttribute("class", "nncb-container");
				lblContainer.setAttribute("style", styles.join(";"))
			}

			var chk = document.createElement("input");
			chk.setAttribute("type", "checkbox");
			chk.setAttribute("id", record[this._childRecordType + "id"]);
			chk.setAttribute("value", record[this._childRecordType + "id"]);
			chk.addEventListener("change", this._onCheckboxChange.bind(this));

			if (this._context.mode.isControlDisabled) {
				chk.setAttribute("disabled", "disabled");
			}

			if (this._useToggleSwitch) {
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

				lblContainer.innerHTML += record[this._labelAttributeName];
				lblContainer.appendChild(chk);
				lblContainer.appendChild(mark);
			}
		}

		this._context.parameters.nnRelationshipDataSet.paging.reset();
		this._context.parameters.nnRelationshipDataSet.refresh();
	}

	private _displayViewRecordsCustom(result: ComponentFramework.WebApi.RetrieveMultipleResponse) {
		let divFlexCtrl = document.createElement("div");

		for (var i = 0; i < result.entities.length; i++) {
			let record = result.entities[i];

			// Add flex content
			divFlexCtrl = document.createElement("div");
			divFlexCtrl.setAttribute("style", "flex: 0 " + (100 / this._numberOfColumns) + "% !important");
			this._divFlexBox.appendChild(divFlexCtrl);

			let lblContainer = document.createElement("label");
			divFlexCtrl.appendChild(lblContainer);

			if (this._useToggleSwitch) {
				lblContainer.setAttribute("class", "nncb-container-switch");

				let spanLabel = document.createElement("span");
				spanLabel.setAttribute("class", "nncb-switch-label");
				spanLabel.textContent = record[this._labelAttributeName];
				divFlexCtrl.appendChild(spanLabel);
			}
			else {
				lblContainer.setAttribute("class", "nncb-container");
			}

			let chk = document.createElement("input");
			chk.setAttribute("type", "checkbox");
			chk.setAttribute("id", record[this._customRelationshipDefinitionChild.ReferencedEntity + "id"]);
			chk.setAttribute("value", record[this._customRelationshipDefinitionChild.ReferencedEntity + "id"]);
			chk.addEventListener("change", this._onCheckboxChangeCustom.bind(this));

			if (this._context.mode.isControlDisabled) {
				chk.setAttribute("disabled", "disabled");
			}

			if (this._useToggleSwitch) {
				let toggle = document.createElement("span");
				toggle.setAttribute("class", "nncb-slider nncb-round");

				lblContainer.appendChild(chk);
				lblContainer.appendChild(toggle);
			}
			else {
				var mark = document.createElement("span");
				mark.setAttribute("class", "nncb-checkmark");

				lblContainer.innerHTML += record[this._labelAttributeName];
				lblContainer.appendChild(chk);
				lblContainer.appendChild(mark);
			}
		}

		this._context.parameters.nnRelationshipDataSet.paging.reset();
		this._context.parameters.nnRelationshipDataSet.refresh();
	}

	private _onCheckboxChange(event: any) {
		let currentTarget = event.currentTarget;
		let entity1name: string;
		let entity2name: string;
		let record1Id: string;
		let record2Id: string;
		if (this._relationshipInfo.Entity1AttributeName === this._childRecordType) {
			entity1name = this._childRecordType;
			record1Id = currentTarget.id;
			entity2name = this._parentRecordType;
			record2Id = this._parentRecordId;
		}
		else {
			entity1name = this._parentRecordType;
			record1Id = this._parentRecordId;
			entity2name = this._childRecordType;
			record2Id = currentTarget.id;
		}

		let thisCtrl: any = this;
		if (currentTarget.checked) {
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
			var theRecordId = currentTarget.id;
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
						thisCtrl._showAlertMessage(thisCtrl._context.resources.getString("Error_Disassociate"));
					}
				);
		}
	}

	private async _onCheckboxChangeCustom(event: any) {
		let currentTarget = <HTMLInputElement>event.currentTarget;
		currentTarget.disabled = true;

		if (!this._changedCheckboxes.find(chx => chx.id === currentTarget.id)) {
			this._changedCheckboxes.push(currentTarget);
		}

		try {
			if (currentTarget.checked) {
				var newRecord: any = {};
				newRecord[`${this._customRelationshipDefinitionChild.ReferencingAttribute}@odata.bind`] = `/${this._customRelationshipDefinitionChild.ReferencedEntity}s(${currentTarget.id})`;
				newRecord[`${this._customRelationshipDefinitionCurrent.ReferencingAttribute}@odata.bind`] = `/${this._customRelationshipDefinitionCurrent.ReferencingEntity}s(${this._parentRecordId})`;

				let newRecordId = await this._context.webAPI.createRecord(this._customRelationshipDefinitionChild.ReferencingEntity, newRecord);
				currentTarget.setAttribute(Constant.IntersectData, newRecordId.id);
			}
			else {
				let intersectRecordId = currentTarget.getAttribute(Constant.IntersectData);
				if (intersectRecordId) {
					await this._context.webAPI.deleteRecord(this._customRelationshipDefinitionChild.ReferencingEntity, intersectRecordId);
					currentTarget.removeAttribute(Constant.IntersectData);
				}
				else {
					let message = this._context.resources.getString(Constant.RecordNotFound);
					this._showAlertMessage(message);
				}
			}
		}
		catch (error) {
			this._showAlertMessage(error.message);
		}
		finally {
			this._context.parameters.nnRelationshipDataSet.paging.reset();
			this._context.parameters.nnRelationshipDataSet.refresh();
			//currentTarget.disabled = false;
		}
	}

	private async _extractDataSetParameters(): Promise<void> {
		this._useToggleSwitch = (this._context.parameters.useToggleSwitch
			&& this._context.parameters.useToggleSwitch.raw
			&& this._context.parameters.useToggleSwitch.raw.toLowerCase() === 'true')
			? true : false;
		this._childRecordType = this._context.parameters.nnRelationshipDataSet.getTargetEntityType();
		this._numberOfColumns = this._context.parameters.columnsNumber ? this._context.parameters.columnsNumber.raw : 1;
		let mode: any = this._context.mode;
		this._parentRecordId = mode.contextInfo.entityId;
		this._parentRecordType = mode.contextInfo.entityTypeName;

		for (var i = 0; i < this._context.parameters.nnRelationshipDataSet.columns.length; i++) {
			var column = this._context.parameters.nnRelationshipDataSet.columns[i];
			if (column.alias === Constant.DisplayName) {
				this._labelAttributeName = column.name;
			}
			else if (column.alias === Constant.BackgoundColorAttribute) {
				this._backgroundColorAttributeName = column.name;

				if (column.dataType === "OptionSet" || column.dataType === "" && column.name === "statuscode") {
					if (!this._colors || !this._colors[column.name]) {
						this._setOptionSetColors(column.name);
					}

					this._backgroundColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === Constant.ForeColorAttribute) {
				this._foreColorAttributeName = column.name;

				if (column.dataType === "OptionSet" || column.dataType === "" && column.name === "statuscode") {
					if (!this._colors || !this._colors[column.name]) {
						this._setOptionSetColors(column.name);
					}

					this._foreColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === Constant.CategoryAttribute) {
				this._categoryAttributeName = column.name;
				this._categoryUseDisplayName = column.dataType === "Lookup.Simple" || column.dataType === "OptionSet" || column.dataType === "TwoOptions" || column.dataType === "" && column.name === "statuscode";
			}
		}
	}

	private _getStyleSheet() {
		for (var i = 0; i < document.styleSheets.length; i++) {
			var sheet = document.styleSheets[i];
			if (sheet.href && sheet.href.endsWith(Constant.ControlCSS)) {
				return sheet;
			}
		}
		return null;
	}

	private _setOptionSetColors(attribute: string) {
		let requestUrl =
			"/api/data/v9.0/EntityDefinitions(LogicalName='"
			+ this._childRecordType + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$filter=LogicalName eq '"
			+ attribute + "'&$expand=OptionSet";

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

	private async _setRelationshipInformation(): Promise<void> {
		let schemaNameParameter = this._context.parameters.relationshipSchemaName;
		if (schemaNameParameter != undefined && schemaNameParameter.raw != null) {
			let entityMetadata = await this._context.utils.getEntityMetadata(this._parentRecordType);
			let nnRelationships = entityMetadata.ManyToManyRelationships.getAll();

			for (let i = 0; i < nnRelationships.length; i++) {
				if (nnRelationships[i].IntersectEntityName.toLowerCase() === this._context.parameters.relationshipSchemaName.raw.toLowerCase()) {
					this._relationshipInfo = {
						Entity1LogicalName: nnRelationships[i].Entity1LogicalName,
						Entity1AttributeName: nnRelationships[i].Entity1IntersectAttribute,
						Entity2LogicalName: nnRelationships[i].Entity2LogicalName,
						Entity2AttributeName: nnRelationships[i].Entity2IntersectAttribute,
						Name: nnRelationships[i].SchemaName
					};

				}
			}
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

				this._relationshipInfo = {
					Entity1LogicalName: nnRelationships[i].Entity1LogicalName,
					Entity1AttributeName: nnRelationships[i].Entity1IntersectAttribute,
					Entity2LogicalName: nnRelationships[i].Entity2LogicalName,
					Entity2AttributeName: nnRelationships[i].Entity2IntersectAttribute,
					Name: nnRelationships[i].SchemaName
				};
			}
		}

		if (foundSchemaName.length === 0) {
			return Promise.reject(new Error(this._context.resources.getString("No_Relationship_Found")));
		}
		if (count > 1) {
			return Promise.reject(new Error(this._context.resources.getString("Multiple_Relationships_Found")));
		}
	}

	private _setToggleDefaultBackgroud() {
		if (this._context.parameters.toggleDefaultBackgroundColorOn && this._context.parameters.toggleDefaultBackgroundColorOn.raw) {
			// @ts-ignore
			let styleSheet: any = this._getStyleSheet();
			if (styleSheet != null) {
				// @ts-ignore
				let rules = styleSheet.rules;
				for (let i = 0; i < rules.length; i++) {
					let rule = rules[i];
					if (rule.selectorText === "input:checked + .nncb-slider") {
						// @ts-ignore
						styleSheet.deleteRule(i);
						// @ts-ignore
						styleSheet.insertRule("input:checked + .nncb-slider { background-color: " + this._context.parameters.toggleDefaultBackgroundColorOn.raw + ";}", rule.index)
					}
				}
			}
		}

		if (this._context.parameters.toggleDefaultBackgroundColorOff && this._context.parameters.toggleDefaultBackgroundColorOff.raw) {
			// @ts-ignore
			let styleSheet: any = this._getStyleSheet();
			if (styleSheet != null) {
				// @ts-ignore
				let rules = styleSheet.rules;
				for (let i = 0; i < rules.length; i++) {
					let rule = rules[i];
					if (rule.selectorText === ".nncb-slider") {
						// @ts-ignore
						styleSheet.deleteRule(i);
						// @ts-ignore
						styleSheet.insertRule(".nncb-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: " + this._context.parameters.toggleDefaultBackgroundColorOff.raw + "; -webkit-transition: .4s; transition: .4s;", rule.index)
					}
				}
			}
		}
	}

	private _parametersAreValid(): boolean {
		let isInputValid: boolean = true;

		if (this._useCustomRelationship) {
			isInputValid = this._context.parameters.customIntersectChildDisplayAttibute.raw
				&& this._context.parameters.customIntersectChildRelationship.raw
				&& (this._context.parameters.customIntersectChildFetchXML.raw || this._context.parameters.customIntersectChildView.raw)
				&& this._context.parameters.relationshipSchemaName
				? true : false;
		}
		else {
			let displayNameDataSetParam = this._context.parameters.nnRelationshipDataSet.columns.find(param => param.alias === Constant.DisplayName);
			isInputValid = displayNameDataSetParam ? true : false;
		}

		if (!isInputValid) {
			this._showAlertMessage(this._context.resources.getString(Constant.InvalidParameters));
		}

		return isInputValid;
	}

	private _showAlertMessage(message: string): void {
		this._context.navigation.openAlertDialog({ text: message });
	}
	//#endregion
}