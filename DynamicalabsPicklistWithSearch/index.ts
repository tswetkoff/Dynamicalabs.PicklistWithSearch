import {IInputs, IOutputs} from "./generated/ManifestTypes";
import {IOptionValue} from "./optionValue";

export class DynamicalabsPicklistWithSearch implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private options: Array<IOptionValue>;
    private selectedOption: IOptionValue;
    private notifyOutputChangedDelegate: () => void;
    private context: ComponentFramework.Context<IInputs>;
    private container: HTMLDivElement;
    private inputElement: HTMLInputElement;
    private selectedLabelElement: HTMLSpanElement;
    private searchButton: HTMLButtonElement;
    private dropDownContainer: HTMLDivElement;
    private isRequired: boolean;

	constructor()
	{

	}

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this.context = context;

        this.isRequired = [1, 2].indexOf(context.parameters.OptionValue.attributes?.RequiredLevel || 0) > -1;

        this.options = this.isRequired ? [] : [{
            Value: null,
            Label: "--Select--"
        }];

        this.options = this.options.concat(context.parameters.OptionValue.attributes?.Options.map((m: any) => {
            return {
                Label: m.Label,
                Value: m.Value
            }
        }) || []);

        this.selectedOption = {
            Value: context.parameters.OptionValue.raw,
            Label: context.parameters.OptionValue.raw !== null ? this.options.find(f => f.Value === context.parameters.OptionValue.raw)?.Label || `Unknown option:${context.parameters.OptionValue.raw} ` : "--Select--"
        };

        this.notifyOutputChangedDelegate = notifyOutputChanged;

        this.container = container;
        this.container.addEventListener("keydown", this.onKeyPressContainer.bind(this));
        this.container.classList.add("dnl-option-container");

        this.initSearchInput();
        this.initSelectedItem();
        this.initSearchButton();
        this.initDropDownContainer();
        this.renderSelected();

        window.addEventListener('resize', this.hideDropDownContainer.bind(this));
        window.addEventListener('click', this.hideDropDownContainer.bind(this));
	}

	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.selectedOption = this.selectedOption = {
            Value: context.parameters.OptionValue.raw,
            Label: context.parameters.OptionValue.raw !== null ? this.options.find(f => f.Value === context.parameters.OptionValue.raw)?.Label || `Unknown option:${context.parameters.OptionValue.raw} ` : "--Select--"
        };
        this.inputElement.value = "";
        this.renderSelected();
	}


	private controlValueChange() {
        this.notifyOutputChangedDelegate();
    }

    private initSearchInput() {
        this.inputElement = document.createElement("input");
        this.inputElement.setAttribute("placeholder", "--Select or search option--");
        this.inputElement.classList.add("dnl-input-element");
        this.inputElement.addEventListener("keyup", this.doSearch.bind(this));

        this.container.append(this.inputElement);
    }

    private initSelectedItem() {
        this.selectedLabelElement = document.createElement("span");
        this.selectedLabelElement.classList.add("selected-label");
        this.selectedLabelElement.addEventListener('click', this.onClickLabel.bind(this));

        this.container.append(this.selectedLabelElement);
    }

    private onClickLabel() {
        this.container.classList.remove("selected");
        this.inputElement.value = "";
        this.inputElement.focus();
    }

    private renderSelected() {
        if (this.selectedOption.Value !== null) {
            this.selectedLabelElement.innerHTML = this.selectedOption.Label;
            this.selectedLabelElement.setAttribute("title", this.selectedOption.Label);
            this.container.classList.add("selected");
        } else {
            this.container.classList.remove("selected");
            this.selectedLabelElement.innerHTML = "---";
        }
    }

    private initSearchButton() {
        this.searchButton = document.createElement("button");
        this.searchButton.classList.add("dnl-search-btn");
        this.searchButton.innerHTML = `<svg viewBox=\"0 0 16 16\" fill=\"currentColor\" xmlns=\"http://www.w3.org/2000/svg\">  <path fill-rule=\"evenodd\" d=\"M1.646 4.646a.5.5 0 0 1 .708 0L8 10.293l5.646-5.647a.5.5 0 0 1 .708.708l-6 6a.5.5 0 0 1-.708 0l-6-6a.5.5 0 0 1 0-.708z\"></path></svg>`;
        this.searchButton.addEventListener('click', this.onClickFilterButton.bind(this));
        this.container.append(this.searchButton);
    }

    private initDropDownContainer() {
        this.dropDownContainer = document.createElement("div");
        this.dropDownContainer.classList.add("custom-dropdown-container");
        this.container.append(this.dropDownContainer);
    }

    private initEvents() {
        const dropDownItems = this.dropDownContainer.querySelectorAll('.custom-drop-down-item');
        if (dropDownItems) {
            dropDownItems.forEach(dropDownItem => {
                dropDownItem.addEventListener("click", this.onChooseItem.bind(this, dropDownItem));
            });
        }
    }

    private onKeyPressContainer(event: KeyboardEvent) {
        const targetElement = event.target as HTMLElement;

        //enter
        if (event.keyCode == 13) {
            if (this.container.classList.contains("expanded") && this.dropDownContainer.querySelector('.custom-drop-down-item.selected')) {
                this.onChooseItem(this.dropDownContainer.querySelector('.custom-drop-down-item.selected'));
                event.preventDefault();
                return;
            }

            if (targetElement.tagName === "INPUT" || (targetElement.tagName === "BUTTON" && targetElement.classList.contains("dnl-search-btn"))) {
                this.doSearch(null);
                event.preventDefault();
                return;
            }
        }
        //Esc
        if (event.keyCode == 27) {
            if (this.container.classList.contains("expanded")) {
                this.onClickFilterButton(event);
            }
        }

        //down
        if (event.keyCode == 40) {
            this.selectNextItem();
            event.preventDefault();
        }

        //up
        if (event.keyCode == 38) {
            this.selectPreviousItem();
            event.preventDefault();
        }
    }

    private selectNextItem() {
        const selectItems = this.dropDownContainer.querySelectorAll('.custom-drop-down-item');
        const selectedItem = this.dropDownContainer.querySelector('.custom-drop-down-item.selected');
        if (!selectItems || !selectItems.length || !this.container.classList.contains("expanded")) {
            return;
        }
        if (selectedItem) {
            const currentSelectedIndex = Array.from(selectItems).indexOf(selectedItem);
            if (selectItems.length > (currentSelectedIndex + 1)) {
                selectItems.item(currentSelectedIndex).classList.remove("selected");
                selectItems.item(currentSelectedIndex + 1).classList.add("selected");
                this.scrollToItem(selectItems.item(currentSelectedIndex + 1));
            }
        } else {
            const firstElement = selectItems.item(0);
            firstElement.classList.add("selected");
            this.scrollToItem(firstElement);
        }
    }

    private selectPreviousItem() {
        const selectItems = this.dropDownContainer.querySelectorAll('.custom-drop-down-item');
        const selectedItem = this.dropDownContainer.querySelector('.custom-drop-down-item.selected');
        if (!selectItems || !selectItems.length || !this.container.classList.contains("expanded")) {
            this.onClickLabel();
            return;
        }
        if (selectedItem) {
            const currentSelectedIndex = Array.from(selectItems).indexOf(selectedItem);
            selectItems.item(currentSelectedIndex).classList.remove("selected");
            if (currentSelectedIndex > 0) {
                selectItems.item(currentSelectedIndex - 1).classList.add("selected");
                this.scrollToItem(selectItems.item(currentSelectedIndex - 1));
            } else {
                this.onClickLabel();
            }
        }
    }

    private scrollToItem(element: any) {
        const divScrollable = this.dropDownContainer.querySelector(".custom-drop-down-scrollable");
        if (divScrollable) {
            element.focus();
            divScrollable.scrollTop = element.offsetTop > 235 ? element.offsetTop : 0;
        }
    }

    private onChooseItem(selectedItem: any) {
        if (!selectedItem) {
            return;
        }

        selectedItem.classList.remove("selected");
        this.hideDropDownContainer();


        this.selectedOption = this.selectedOption = {
            Value: +selectedItem.getAttribute("data-id"),
            Label: +selectedItem.getAttribute("data-id")
                ? this.options.find(f => f.Value === +selectedItem.getAttribute("data-id"))?.Label || `Unknown option:${selectedItem.getAttribute("data-id")}`
                : ""
        };

        this.controlValueChange();
        this.renderSelected();
    }

    private onClickFilterButton(event: any) {
        if (this.container.classList.contains("expanded")) {
            this.hideDropDownContainer();
            this.renderSelected();
        } else {
            this.doSearch(null);
            event.stopPropagation()
        }
    }

    private doSearch(event: any) {
        if (event && [13, 27, 9, 40, 37, 38, 39].indexOf(event.keyCode) > -1) {
            return;
        }

        const filteredItems =
            this.inputElement.value.trim()
                ? this.options.filter(
                f => f.Label.toLowerCase().includes(this.inputElement.value.toLowerCase()))
                : this.options;

        this.renderDropDownContainer(filteredItems);
    }

    private renderDropDownContainer(optionValues: Array<IOptionValue>) {
        let options = "";
        for (let v = 0; v < optionValues.length; v++) {
            options +=
                `<div  data-id="${optionValues[v].Value}" class="custom-drop-down-item" >
                    <div class="option-text">
                        <div class="option-name" title="${optionValues[v].Label}">${optionValues[v].Label}</div>
                    </div>
                 </div>`
        }

        options = options.length === 0 ? "<div class='no-data'>No data</div>" : options;

        this.dropDownContainer.innerHTML =
            `<div class="custom-drop-down-scrollable">                 
                 ${options}                 
             </div>`;
        this.initEvents();
        this.showDropDownContainer();
    }

    private showDropDownContainer() {
        const docHeight = document.body.clientHeight;
        const currentPosition = this.container.getBoundingClientRect().y;
        if (currentPosition + 335 > docHeight) {
            this.container.classList.add("expanded-top");
        } else {
            this.container.classList.remove("expanded-top");
        }
        this.container.classList.add("expanded");
    }

    private hideDropDownContainer() {
        this.container.classList.remove("expanded");
        this.inputElement.value = "";
    }

	public getOutputs(): IOutputs
	{
		return {
            OptionValue: this.selectedOption.Value || undefined,
        }
	}

	public destroy(): void
	{}
}