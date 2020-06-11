import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class DisplayImageForOptionset implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _selectedOptionsetValue: string;
	private _optionsetText: string;
	// Power Apps component framework delegate which will be assigned to this object which would be called whenever any update happens.
	private _notifyOutputChanged: () => void;
	// label element created as part of this component
	private lblRemainingCount: HTMLLabelElement;
	private imageElement: HTMLImageElement;
	// reference to the component container HTMLDivElement
	// This element contains all elements of our code component example
	private _container: HTMLDivElement;
	// reference to Power Apps component framework Context object
	private _context: ComponentFramework.Context<IInputs>;
	// Event Handler 'refreshData' reference
	private _refreshData: EventListenerOrEventListenerObject;
	private options: ComponentFramework.PropertyHelper.OptionMetadata[];
	private _dropdown: HTMLSelectElement;
	private blankOptions: ComponentFramework.PropertyHelper.OptionMetadata[];
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
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		this._refreshData = this.refreshData.bind(this);

		let dropdown = document.createElement("select");
		dropdown.addEventListener("change", this._refreshData);
		this._dropdown = dropdown;
		this._dropdown.setAttribute("class", "dropDown");

		this._selectedOptionsetValue = context.parameters.optionSet.raw?.toString() || "0";

		this.options = context.parameters.optionSet.attributes?.Options || this.blankOptions;

		this.SetOptionsetValues();
		container.appendChild(dropdown);
		if (this._selectedOptionsetValue != null) {
			dropdown.value = this._selectedOptionsetValue;
		}

		this._container = document.createElement("div");
		// Add control initialization code
		this._context = context;
		this.imageElement = document.createElement("img");
		this.imageElement.setAttribute("id", "imgOptionset");
		this.imageElement.setAttribute("class", "image");
		this.imageElement.src = "https://learningjune2020.crm6.dynamics.com/WebResources/new_" + this._optionsetText.replace(" ", "");
		this.imageElement.height = 100;
		this.imageElement.width = 100;

		this._container.appendChild(this.imageElement);
		container.appendChild(this._container);

	}

	public SetOptionsetValues(): void {
		let optionTextSet = false;
		if (this.options != null && this.options != undefined) {
			if (this._dropdown.options == null || this._dropdown.options.length == 0) {
				this.options.forEach(element => {
					let option = document.createElement("option");
					option.value = element.Value.toString();
					option.text = element.Label;
					this._dropdown.appendChild(option);

				});
			}
			this.options.forEach(element => {
				if (!optionTextSet && element.Value.toString() == this._selectedOptionsetValue) {
					this._optionsetText = element.Label;
					optionTextSet = true;
				}
			});
		}

	}

	public refreshData(evt: Event): void {
		this._selectedOptionsetValue = this._dropdown.value;
		this.SetOptionsetValues();
		this.imageElement.src = "https://learningjune2020.crm6.dynamics.com/WebResources/new_" + this._optionsetText.replace(" ", "");

		this._notifyOutputChanged();
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view

		// storing the latest context from the control.
		this._selectedOptionsetValue = context.parameters.optionSet.raw ? context.parameters.optionSet.raw.toString() : "0";

		this._context = context;

	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			optionSet: Number(this._selectedOptionsetValue)
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		this._dropdown.removeEventListener("change", this.refreshData);
	}
}