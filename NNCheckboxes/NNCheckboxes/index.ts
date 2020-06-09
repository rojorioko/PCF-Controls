import { IDataSetRecord } from './ts/interface/datasetrecord.interface';
import { Constant } from './ts/constant/constant';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IAttributeValue } from './ts/interface/attributevalue.interface';
import { IRelationDefinition } from './ts/interface/relationdefinition.interface';
import { INNRelationshipInfo } from './ts/interface/nnrelationshipinfo.interface';

export class ListCheckboxes implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Reference to the control container HTMLDivElement
	// This element contains all elements of our custom control example
	private _container: HTMLDivElement;
	private _divFlexBox: HTMLDivElement;
	// Reference to ComponentFramework Context object
	private _context: ComponentFramework.Context<IInputs>;
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
	private _checkboxRecords: HTMLCollectionOf<Element> | null;
	private _customIntersectDisplayEntityFetchXML: string | null;
	private _hasValidDataSource: boolean;
	private _elementPreFix: string;

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
		this._container.setAttribute("class", Constant.NncbMain);
		container.appendChild(this._container);
		this._useCustomRelationship = (context.parameters.useCustomIntersect?.raw?.toLowerCase() === 'true');

		if (this._parametersAreNotValid()) {
			this._showAlertMessage(this._context.resources.getString(Constant.InvalidParameters));
			return;
		}

		try {
			this._setToggleDefaultBackgroud();
			this._extractDataSetParameters();

			await this._setRelationshipDetails();

			if (this._useCustomRelationship && !this._hasValidDataSource) {
				this._showAlertMessage(this._context.resources.getString(Constant.InvalidParameters));
				return;
			}

			let saveQuery: ComponentFramework.WebApi.Entity = await this._getViewFetchXML();

			let result: ComponentFramework.WebApi.RetrieveMultipleResponse = await context.webAPI.retrieveMultipleRecords(saveQuery.returnedtypecode, "?fetchXml=" + encodeURIComponent(saveQuery.fetchxml));

			this._displayViewRecords(result, this._useCustomRelationship);
		}
		catch (error) {
			this._showAlertMessage(error.message);
			return;
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public async updateView(context: ComponentFramework.Context<IInputs>) {
		if (!context.updatedProperties.includes(Constant.DatasetName) && !context.updatedProperties.includes(Constant.CustomIntersectDisplayEntityFetchXML))
			return;

		if (context.parameters.nnRelationshipDataSet.paging.hasNextPage) {
			context.parameters.nnRelationshipDataSet.paging.loadNextPage();
			return;
		}

		//Check if fetchXML has changed, then reset the display data
		if (context.updatedProperties.includes(Constant.CustomIntersectDisplayEntityFetchXML)
			&& context.parameters.fetchXmlData.raw !== this._customIntersectDisplayEntityFetchXML) {

			this._customIntersectDisplayEntityFetchXML = context.parameters.fetchXmlData.raw

			try {
				let saveQuery = await this._getViewFetchXML();
				let result: ComponentFramework.WebApi.RetrieveMultipleResponse = await context.webAPI.retrieveMultipleRecords(
					saveQuery.returnedtypecode,
					"?fetchXml=" + encodeURIComponent(saveQuery.fetchxml)
				);

				// reset the control display								
				this._divFlexBox.innerHTML = "";
				this._displayViewRecords(result, this._useCustomRelationship);

			} catch (error) {
				this._showAlertMessage(error.message);
			}
		}

		// Enable checkbox after reset.
		if (this._checkboxRecords && this._checkboxRecords.length > 0) {
			for (let index = 0; index < this._checkboxRecords.length; index++) {
				(<HTMLInputElement>this._checkboxRecords[index]).disabled = false;
			}

			this._checkboxRecords = null;
		}

		let dataSet = context.parameters.nnRelationshipDataSet;

		// Check all records that are already associated to current record
		for (let j = 0; j < dataSet.sortedRecordIds.length; j++) {
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

			let elementId = matchId ? `${matchId}|${this._elementPreFix}` : null;
			let chk = elementId ? <HTMLInputElement>window.document.getElementById(elementId) : null;

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

	/**
	 * Create main container for the control
	 */
	private _createMainContainer() {
		this._divFlexBox = document.createElement("div");
		this._divFlexBox.setAttribute("class", Constant.NncbFlex);
		this._container.appendChild(this._divFlexBox);
	}

	/**
	 * Set the relationship definition of the current and target entity
	 */
	private async _setRelationshipDetails() {
		if (this._useCustomRelationship) {
			this._customRelationshipDefinitionChild = await this._getRelationshipDefinitionByName(this._context.parameters.customIntersectDisplayEntityRelationship.raw);
			this._customRelationshipDefinitionCurrent = await this._getRelationshipDefinitionByName(this._context.parameters.relationshipSchemaName.raw);
			this._elementPreFix = this._customRelationshipDefinitionCurrent?.SchemaName;
		}
		else {
			await this._setRelationshipInformation();
			this._elementPreFix = this._relationshipInfo.Name;
		}
	}

	/**
	 * Get the relationship definition based on schema name. This will return Entities and Fields of both child and parent entity.
	 * @param schemaName Schema Name of the relation
	 */
	private async _getRelationshipDefinitionByName(schemaName: string): Promise<any> {
		//Why not use context and webapi feature for this?
		//this._context.webAPI.retrieveMultipleRecords();

		//@ts-ignore
		let requestUrl = `${this._context.page.getClientUrl()}/api/data/v9.1/RelationshipDefinitions(SchemaName='${schemaName}')`;
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

	/**
	 * Get the SaveQuery entity record. It contains the fetchxml data to be executed and it's target entity.
	 */
	private async _getViewFetchXML(): Promise<ComponentFramework.WebApi.Entity> {
		let saveQuery: ComponentFramework.WebApi.Entity | null = null;;

		if (this._context.parameters.fetchXmlData?.raw) {
			let targetDisplayEntity = this._useCustomRelationship ?
				this._customRelationshipDefinitionChild.ReferencedEntity : this._parentRecordType === this._relationshipInfo.Entity1LogicalName ?
					this._relationshipInfo.Entity2LogicalName : this._relationshipInfo.Entity1LogicalName;
			saveQuery = {
				"fetchxml": this._context.parameters.fetchXmlData.raw,
				"returnedtypecode": targetDisplayEntity
			};
		}

		if (saveQuery === null) {
			if (this._useCustomRelationship) {
				let queryOption: string = `?$select=fetchxml,returnedtypecode&$filter=name eq '${this._context.parameters.customIntersectDisplayEntityView?.raw}'`;
				let result: ComponentFramework.WebApi.RetrieveMultipleResponse = await this._context.webAPI.retrieveMultipleRecords(Constant.SaveQuery, queryOption);
				saveQuery = result.entities && result.entities.length > 0 ? result.entities[0] : null;
			} else {
				saveQuery = await this._context.webAPI.retrieveRecord(Constant.SaveQuery, this._context.parameters.nnRelationshipDataSet.getViewId(), "?$select=fetchxml,returnedtypecode");
			}
		}

		if (saveQuery !== null && saveQuery.fetchxml) {
			return Promise.resolve(saveQuery);
		}
		else {
			let viewNotFound = this._context.resources.getString(Constant.InvalidParameters);
			return Promise.reject(new Error(viewNotFound));
		}
	}

	/**
	 * Display entity records as chexbox
	 * @param result Entity collection to be displayed as checkbox
	 * @param isCustomRelationship Identify whether to display Custom NN relationship or default
	 */
	private _displayViewRecords(result: ComponentFramework.WebApi.RetrieveMultipleResponse, isCustomRelationship: boolean) {
		this._createMainContainer();
		let category = "";

		for (let i = 0; i < result.entities.length; i++) {
			let record = result.entities[i];

			// If using category
			if (this._categoryAttributeName && !isCustomRelationship) {
				// We need to display new category only if the category 
				// is different from the previous one
				if (category !== record[this._categoryAttributeName]) {
					category = record[this._categoryAttributeName];

					let label = record[`${this._categoryAttributeName + (this._categoryUseDisplayName ? "@OData.Community.Display.V1.FormattedValue" : "")}`];
					if (!label) {
						label = this._context.resources.getString(Constant.NoCategory);
					}

					// Add the category
					let categoryDiv = document.createElement("div");
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
			let divFlexCtrl = document.createElement("div");
			divFlexCtrl.setAttribute("style", `flex: 0 ${(100 / this._numberOfColumns)}% !important`);
			this._divFlexBox.appendChild(divFlexCtrl);

			// With style if configured with colors
			let styles = new Array();
			if (!isCustomRelationship) {
				if (this._backgroundColorAttributeName) {
					if (this._backgroundColorIsFromOptionSet) {
						if (this._colors
							&& this._colors[this._backgroundColorAttributeName]
							&& this._colors[this._backgroundColorAttributeName][record[this._backgroundColorAttributeName]]) {
							let color = this._colors[this._backgroundColorAttributeName][record[this._backgroundColorAttributeName]];
							styles.push(`background-color:${color}`);
						}
					}
					else {
						styles.push(`background-color: ${record[this._backgroundColorAttributeName]}`);
					}
				}
				if (this._foreColorAttributeName) {
					if (this._foreColorIsFromOptionSet) {
						if (this._colors
							&& this._colors[this._foreColorAttributeName]
							&& this._colors[this._foreColorAttributeName][record[this._backgroundColorAttributeName]]) {
							let color = this._colors[this._foreColorAttributeName][record[this._backgroundColorAttributeName]];
							styles.push(`color: ${color}`);
						}
					}
					else {
						styles.push(`color: ${record[this._foreColorAttributeName]}`);
					}
				}
			}

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
				lblContainer.setAttribute("style", styles.join(";"))
			}

			let inputId = isCustomRelationship ? record[`${this._customRelationshipDefinitionChild.ReferencedEntity}id`] + `|${this._elementPreFix}`
				: record[`${this._childRecordType}id`] + `|${this._elementPreFix}`;
			let inputValue = isCustomRelationship ? record[`${this._customRelationshipDefinitionChild.ReferencedEntity}id`] + `|${this._elementPreFix}`
				: record[`${this._childRecordType}id`] + `|${this._elementPreFix}`;
			let changeHandler = isCustomRelationship ? this._onCheckboxChangeCustom.bind(this) : this._onCheckboxChange.bind(this);
			let chk = document.createElement("input");

			chk.setAttribute("type", "checkbox");
			chk.setAttribute("id", inputId);
			chk.setAttribute("value", inputValue);

			if (isCustomRelationship)
				chk.setAttribute("class", this._elementPreFix);

			chk.addEventListener("change", changeHandler);

			if (this._context.mode.isControlDisabled)
				chk.setAttribute("disabled", "disabled");

			if (this._useToggleSwitch) {
				let toggle = document.createElement("span");
				toggle.setAttribute("class", "nncb-slider nncb-round");

				if (styles.length > 0)
					toggle.setAttribute("style", styles.join(";"))

				lblContainer.appendChild(chk);
				lblContainer.appendChild(toggle);
			}
			else {
				let mark = document.createElement("span");
				mark.setAttribute("class", "nncb-checkmark");

				lblContainer.innerHTML += record[this._labelAttributeName];
				lblContainer.appendChild(chk);
				lblContainer.appendChild(mark);
			}
		}

		this._context.parameters.nnRelationshipDataSet.paging.reset();
		this._context.parameters.nnRelationshipDataSet.refresh();
	}

	/**
	 * Main handler for checkboxes with default NN relationship. 
	 * It will associate selected record to current entity if checked and disassociate if unchecked.	 
	 * @param event Default event parameter on checkbox
	 */
	private async _onCheckboxChange(event: any) {
		let currentTarget = event.currentTarget;
		let entity1name: string;
		let entity2name: string;
		let record1Id: string;
		let record2Id: string;
		if (this._relationshipInfo.Entity1AttributeName === this._childRecordType) {
			entity1name = this._childRecordType;
			record1Id = currentTarget.id.split('|')[0];
			entity2name = this._parentRecordType;
			record2Id = this._parentRecordId;
		}
		else {
			entity1name = this._parentRecordType;
			record1Id = this._parentRecordId;
			entity2name = this._childRecordType;
			record2Id = currentTarget.id.split('|')[0];
		}

		let thisCtrl: any = this;
		let request: any;
		if (currentTarget.checked) {
			request = new class {
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
		}
		else {
			request = new class {
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
		}

		try {
			await thisCtrl._context.webAPI.execute(request);
		} catch (error) {
			this._showAlertMessage(error.message);
		}
	}

	/**
	 * Main handler for checkboxes with custom NN relationship. 
	 * It will create record for intersect entity if checked and delete instance of intersect entity if unchecked.
	 * @param event Default event parameter on checkbox
	 */
	private async _onCheckboxChangeCustom(event: any) {
		let currentTarget = <HTMLInputElement>event.currentTarget;
		currentTarget.disabled = true;

		this._checkboxRecords = document.getElementsByClassName(this._elementPreFix);
		if (this._checkboxRecords.length > 0) {
			for (let index = 0; index < this._checkboxRecords.length; index++) {
				(<HTMLInputElement>this._checkboxRecords[index]).disabled = true;
			}
		}

		try {
			if (currentTarget.checked) {
				let newRecord: any = {};
				newRecord[`${this._customRelationshipDefinitionChild.ReferencingAttribute}@odata.bind`] = `/${this._customRelationshipDefinitionChild.ReferencedEntity}s(${currentTarget.id.split('|')[0]})`;
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
		}
	}

	/**
	 * This will map the parameters from the form control to the appropriate attibutes or properties.
	 */
	private async _extractDataSetParameters(): Promise<void> {
		this._useToggleSwitch = (this._context.parameters.useToggleSwitch?.raw?.toLowerCase() === 'true');
		this._childRecordType = this._context.parameters.nnRelationshipDataSet.getTargetEntityType();
		this._numberOfColumns = this._context.parameters.columnsNumber?.raw ?? 1;
		let mode: any = this._context.mode;
		this._parentRecordId = mode.contextInfo.entityId;
		this._parentRecordType = mode.contextInfo.entityTypeName;
		this._customIntersectDisplayEntityFetchXML = this._context.parameters.fetchXmlData?.raw ?? null;

		for (let i = 0; i < this._context.parameters.nnRelationshipDataSet.columns.length; i++) {
			let column = this._context.parameters.nnRelationshipDataSet.columns[i];
			if (column.alias === Constant.DisplayName) {
				this._labelAttributeName = column.name;
			}
			else if (column.alias === Constant.BackgoundColorAttribute) {
				this._backgroundColorAttributeName = column.name;

				if (column.dataType === Constant.OptionSet || column.dataType === "" && column.name === Constant.StatusCode) {
					if (!this._colors || !this._colors[column.name]) {
						this._setOptionSetColors(column.name);
					}

					this._backgroundColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === Constant.ForeColorAttribute) {
				this._foreColorAttributeName = column.name;

				if (column.dataType === Constant.OptionSet || column.dataType === "" && column.name === Constant.StatusCode) {
					if (!this._colors || !this._colors[column.name]) {
						this._setOptionSetColors(column.name);
					}

					this._foreColorIsFromOptionSet = true;
				}
			}
			else if (column.alias === Constant.CategoryAttribute) {
				this._categoryAttributeName = column.name;
				this._categoryUseDisplayName = column.dataType === Constant.LookupSimple || column.dataType === Constant.OptionSet || column.dataType === Constant.TwoOptions || column.dataType === "" && column.name === Constant.StatusCode;
			}
		}
	}

	/**
	 * Retrieve the main CSS file of the Control
	 * @returns Return the main CSS file.
	 */
	private _getStyleSheet() {
		for (let i = 0; i < document.styleSheets.length; i++) {
			let sheet = document.styleSheets[i];
			if (sheet.href && sheet.href.endsWith(Constant.ControlCSS)) {
				return sheet;
			}
		}
		return null;
	}

	/**
	 * Set the backbround and fore color based on the Optionset value color coding.
	 * @param attribute Logical name of the optionset field.
	 */
	private _setOptionSetColors(attribute: string) {
		let requestUrl =
			"/api/data/v9.0/EntityDefinitions(LogicalName='"
			+ this._childRecordType + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$filter=LogicalName eq '"
			+ attribute + "'&$expand=OptionSet";

		let thisCtrl = this;
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
					for (let i = 0; i < options.length; i++) {
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

	/**
	 * Get NN relationship definition of the current entity and target entity.
	 * If there are more than 1 NN relationship and the schema name of the relationship is not provided, the control will throw an error.
	 */
	private async _setRelationshipInformation(): Promise<void> {
		let schemaNameParameter = this._context.parameters.relationshipSchemaName;
		let entityMetadata = await this._context.utils.getEntityMetadata(this._parentRecordType);
		let nnRelationships = entityMetadata.ManyToManyRelationships.getAll();

		if (schemaNameParameter?.raw) {
			for (let i = 0; i < nnRelationships.length; i++) {
				if (nnRelationships[i].IntersectEntityName.toLowerCase() === this._context.parameters.relationshipSchemaName.raw.toLowerCase()) {
					this._relationshipInfo = {
						Entity1LogicalName: nnRelationships[i].Entity1LogicalName,
						Entity1AttributeName: nnRelationships[i].Entity1IntersectAttribute,
						Entity2LogicalName: nnRelationships[i].Entity2LogicalName,
						Entity2AttributeName: nnRelationships[i].Entity2IntersectAttribute,
						Name: nnRelationships[i].SchemaName
					};

					break;
				}
			}

			if (this._relationshipInfo) return;
		}

		let count = 0;
		let foundSchemaName = "";

		for (let i = 0; i < nnRelationships.length; i++) {
			if ((nnRelationships[i].Entity1LogicalName === this._parentRecordType && nnRelationships[i].Entity2LogicalName === this._childRecordType) ||
				(nnRelationships[i].Entity1LogicalName === this._childRecordType && nnRelationships[i].Entity2LogicalName === this._parentRecordType)) {
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
			return Promise.reject(new Error(this._context.resources.getString(Constant.NoRelationshipFound)));
		}
		if (count > 1) {
			return Promise.reject(new Error(this._context.resources.getString(Constant.MultipleRelationshipFound)));
		}
	}

	/**
	 * Set background color of the Switch.
	 * This is only applicable if the Toggle Switch is enabled instead of Checkbox.
	 */
	private _setToggleDefaultBackgroud() {
		let styleSheet: any = this._getStyleSheet();
		if (!styleSheet) return;

		let rules = styleSheet.rules;

		if (this._context.parameters.toggleDefaultBackgroundColorOn?.raw) {
			for (let i = 0; i < rules.length; i++) {
				let rule = rules[i];
				if (rule.selectorText === "input:checked + .nncb-slider") {
					styleSheet.deleteRule(i);
					styleSheet.insertRule(`input:checked + .nncb-slider { background-color: ${this._context.parameters.toggleDefaultBackgroundColorOn.raw};}`, rule.index)
				}
			}
		}

		if (this._context.parameters.toggleDefaultBackgroundColorOff?.raw) {
			for (let i = 0; i < rules.length; i++) {
				let rule = rules[i];
				if (rule.selectorText === ".nncb-slider") {
					styleSheet.deleteRule(i);
					styleSheet.insertRule(`.nncb-slider { position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0; background-color: ${this._context.parameters.toggleDefaultBackgroundColorOff.raw}; -webkit-transition: .4s; transition: .4s;`, rule.index);
				}
			}
		}
	}

	/**
	 * Validate the form parameters.
	 * @returns True if all parameters are valid, otherwise, returns false.
	 */
	private _parametersAreNotValid(): boolean {
		let isInputValid: boolean = true;

		if (this._useCustomRelationship) {
			isInputValid = this._context.parameters.customIntersectDisplayEntityDisplayAttibute?.raw
				&& this._context.parameters.customIntersectDisplayEntityRelationship?.raw
				&& this._context.parameters.relationshipSchemaName?.raw
				? true : false;
		}
		else {
			let displayNameDataSetParam = this._context.parameters.nnRelationshipDataSet.columns.find(param => param.alias === Constant.DisplayName);
			isInputValid = displayNameDataSetParam ? true : false;
		}

		this._hasValidDataSource = this._context.parameters.fetchXmlData?.raw
			|| this._context.parameters.customIntersectDisplayEntityView?.raw
			? true : false;

		return !isInputValid;
	}

	/**
	 * Main handler of the messages needed to be shown to the user.
	 * @param message Message to display on the Form
	 */
	private _showAlertMessage(message: string): void {
		this._context.navigation.openAlertDialog({ text: message });
	}

	//#endregion
}