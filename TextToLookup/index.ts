import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class TextToLookup implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _value: string;
    private _selectedRecord: { id: string; name: string } | null;
    private _fetchXml: string;
    private _notifyOutputChanged: () => void;
    private _context: ComponentFramework.Context<IInputs>;

    // UI Elements
    private _container: HTMLDivElement;
    private _inputWrapper: HTMLDivElement;
    private _textInput: HTMLInputElement;
    private _searchIcon: HTMLDivElement;
    private _clearButton: HTMLButtonElement;
    private _dropdownContainer: HTMLDivElement;
    private _dropdown: HTMLUListElement;
    private _loadingSpinner: HTMLDivElement;
    private _isDropdownOpen: boolean = false;
    private _isLoading: boolean = false;

    // Debounce timer
    private _searchTimer: number | null = null;

    constructor() {
        this._value = "";
        this._selectedRecord = null;
        this._fetchXml = "";
    }

    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;

        // Initialize values
        this._value = context.parameters.selectedValue.raw || "";
        this._fetchXml = context.parameters.fetchXml.raw || "";

        // Parse existing selected record if available
        if (this._value) {
            try {
                this._selectedRecord = JSON.parse(this._value);
            } catch {
                this._selectedRecord = null;
            }
        }

        this.createUI();
        this.updateDisplay();
    }

    private createUI(): void {
        // Add CSS class to container
        this._container.className = "text-lookup-container";

        // Input wrapper
        this._inputWrapper = document.createElement("div");
        this._inputWrapper.className = "text-lookup-input-wrapper";

        // Search icon
        this._searchIcon = document.createElement("div");
        this._searchIcon.className = "text-lookup-search-icon";
        this._searchIcon.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001c.03.04.062.078.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1.007 1.007 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0z" fill="#666"/>
            </svg>`;

        // Text input
        this._textInput = document.createElement("input");
        this._textInput.type = "text";
        this._textInput.className = "text-lookup-input";
        this._textInput.placeholder = "Search records...";

        // Clear button
        this._clearButton = document.createElement("button");
        this._clearButton.className = "text-lookup-clear-btn";
        this._clearButton.innerHTML = `
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                <path d="M9 3L3 9M3 3l6 6" stroke="#666" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
        `;

        // Loading spinner
        this._loadingSpinner = document.createElement("div");
        this._loadingSpinner.className = "text-lookup-spinner";
        this._loadingSpinner.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none">
                <circle cx="12" cy="12" r="10" stroke="#0078d4" stroke-width="2" stroke-linecap="round" stroke-dasharray="31.416" stroke-dashoffset="31.416">
                    <animate attributeName="stroke-dasharray" dur="2s" values="0 31.416;15.708 15.708;0 31.416;0 31.416" repeatCount="indefinite"/>
                    <animate attributeName="stroke-dashoffset" dur="2s" values="0;-15.708;-31.416;-31.416" repeatCount="indefinite"/>
                </circle>
            </svg>
        `;

        // Dropdown container
        this._dropdownContainer = document.createElement("div");
        this._dropdownContainer.className = "text-lookup-dropdown-container";

        // Dropdown list
        this._dropdown = document.createElement("ul");
        this._dropdown.className = "text-lookup-dropdown";

        // Event listeners
        this._textInput.addEventListener("input", this.onTextInputChange.bind(this));
        this._textInput.addEventListener("focus", this.onTextInputFocus.bind(this));
        this._textInput.addEventListener("blur", this.onTextInputBlur.bind(this));
        this._textInput.addEventListener("keydown", this.onTextInputKeyDown.bind(this));
        this._clearButton.addEventListener("click", this.onClearClick.bind(this));

        // Close dropdown when clicking outside
        document.addEventListener("click", (event) => {
            if (!this._container.contains(event.target as Node)) {
                this.closeDropdown();
            }
        });

        // Assemble UI
        this._inputWrapper.appendChild(this._searchIcon);
        this._inputWrapper.appendChild(this._textInput);
        this._inputWrapper.appendChild(this._loadingSpinner);
        this._inputWrapper.appendChild(this._clearButton);
        this._dropdownContainer.appendChild(this._dropdown);
        this._container.appendChild(this._inputWrapper);
        this._container.appendChild(this._dropdownContainer);
    }

    private updateDisplay(): void {
        if (this._selectedRecord) {
            this._textInput.value = this._selectedRecord.name;
            this._textInput.classList.add("has-value");
            this._clearButton.style.display = "flex";
        } else {
            this._textInput.value = "";
            this._textInput.classList.remove("has-value");
            this._clearButton.style.display = "none";
        }
    }

    private onTextInputChange(): void {
        const searchText = this._textInput.value.trim();

        // Clear previous timer
        if (this._searchTimer) {
            clearTimeout(this._searchTimer);
        }

        // If there's a selected record and user is typing, clear it
        if (this._selectedRecord && searchText !== this._selectedRecord.name) {
            this.clearSelection();
        }

        // Debounce search
        this._searchTimer = window.setTimeout(() => {
            if (searchText.length >= 2) {
                this.performSearch(searchText);
            } else {
                this.closeDropdown();
            }
        }, 300);
    }

    private onTextInputFocus(): void {
        const searchText = this._textInput.value.trim();
        if (searchText.length >= 2) {
            this.performSearch(searchText);
        }
    }

    private onTextInputBlur(): void {
        // Delay to allow click on dropdown items
        setTimeout(() => {
            if (!this._container.matches(':hover')) {
                this.closeDropdown();
            }
        }, 200);
    }

    private onTextInputKeyDown(event: KeyboardEvent): void {
        if (event.key === "Enter") {
            event.preventDefault();
            const searchText = this._textInput.value.trim();
            if (searchText.length >= 2) {
                this.performSearch(searchText);
            }
        } else if (event.key === "Escape") {
            this.closeDropdown();
        }
    }

    private onClearClick(): void {
        this.clearSelection();
        this._textInput.focus();
    }

    private clearSelection(): void {
        this._selectedRecord = null;
        this._value = "";
        this.updateDisplay();
        this.closeDropdown();
        this._notifyOutputChanged();
    }

    private async performSearch(searchText: string): Promise<void> {
        if (!this._fetchXml) {
            return;
        }

        this.showLoading(true);

        try {
            // Replace search parameter in FetchXML
            const searchFetchXml = this._fetchXml.replace(/\{searchText\}/g, searchText);

            // Execute FetchXML query
            const result = await this._context.webAPI.retrieveMultipleRecords(
                this.getEntityNameFromFetchXml(searchFetchXml),
                `?fetchXml=${encodeURIComponent(searchFetchXml)}`
            );

            this.displaySearchResults(result.entities);
        } catch (error) {
            console.error("Search error:", error);
            this.closeDropdown();
        } finally {
            this.showLoading(false);
        }
    }

    private getEntityNameFromFetchXml(fetchXml: string): string {
        const match = fetchXml.match(/<entity[^>]+name=['"]([^'"]+)['"]/i);
        return match ? match[1] : "";
    }

    private displaySearchResults(entities: any[]): void {
        // Clear previous results
        this._dropdown.innerHTML = "";

        if (entities.length === 0) {
            const noResultsItem = document.createElement("li");
            noResultsItem.className = "text-lookup-dropdown-item no-results";
            noResultsItem.innerHTML = `
                <div class="text-lookup-item-content">
                    <div class="text-lookup-item-icon">
                        <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <path d="M7 13A6 6 0 1 0 7 1a6 6 0 0 0 0 12ZM7 4v3M7 9.5h.01" stroke="#666" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
                        </svg>
                    </div>
                    <span>No results found</span>
                </div>
            `;
            this._dropdown.appendChild(noResultsItem);
        } else {
            entities.forEach(entity => {
                const listItem = document.createElement("li");
                listItem.className = "text-lookup-dropdown-item";

                listItem.addEventListener("click", () => {
                    this.selectRecord(entity);
                });

                const displayName = this.getDisplayName(entity);
                const recordId = this.getPrimaryIdField(entity);

                listItem.innerHTML = `
                    <div class="text-lookup-item-content">
                        <div class="text-lookup-item-icon">
                            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                                <path d="M2 3.5A1.5 1.5 0 0 1 3.5 2h7A1.5 1.5 0 0 1 12 3.5v7a1.5 1.5 0 0 1-1.5 1.5h-7A1.5 1.5 0 0 1 2 10.5v-7Z" stroke="#0078d4" stroke-width="1.2" fill="none"/>
                                <path d="M4.5 6.5h5M4.5 8.5h3" stroke="#0078d4" stroke-width="1" stroke-linecap="round"/>
                            </svg>
                        </div>
                        <span class="text-lookup-item-name">${displayName}</span>
                    </div>
                `;

                this._dropdown.appendChild(listItem);
            });
        }

        this.openDropdown();
    }

    private getDisplayName(entity: any): string {
        // Common display name fields - adjust based on your needs
        const nameFields = ['name', 'fullname', 'subject', 'title', 'displayname', 'firstname', 'lastname'];

        for (const field of nameFields) {
            if (entity[field]) {
                return entity[field];
            }
        }

        // For contacts, combine first and last name
        if (entity['firstname'] && entity['lastname']) {
            return `${entity['firstname']} ${entity['lastname']}`;
        }

        // Fallback to first string property
        const firstStringProperty = Object.keys(entity).find(key =>
            typeof entity[key] === 'string' && !key.startsWith('_') && !key.includes('@')
        );

        return firstStringProperty ? entity[firstStringProperty] : entity[Object.keys(entity)[0]];
    }

    private selectRecord(entity: any): void {
        // Get record details
        const recordId = entity[this.getPrimaryIdField(entity)];
        const recordName = this.getDisplayName(entity);

        // Set new selection (replaces any existing selection)
        this._selectedRecord = {
            id: recordId,
            name: recordName
        };

        this._value = JSON.stringify(this._selectedRecord);

        this.updateDisplay();
        this.closeDropdown();
        this._notifyOutputChanged();
    }

    private getPrimaryIdField(entity: any): string {
        // Find the primary ID field (usually ends with 'id')
        const idField = Object.keys(entity).find(key =>
            key.endsWith('id') && !key.startsWith('_') && !key.includes('@')
        );
        return idField || Object.keys(entity)[0];
    }

    private showLoading(show: boolean): void {
        this._isLoading = show;
        if (show) {
            this._loadingSpinner.style.display = "flex";
            this._searchIcon.style.display = "none";
        } else {
            this._loadingSpinner.style.display = "none";
            this._searchIcon.style.display = "flex";
        }
    }

    private openDropdown(): void {
        this._dropdownContainer.classList.add("open");
        this._isDropdownOpen = true;
    }

    private closeDropdown(): void {
        this._dropdownContainer.classList.remove("open");
        this._isDropdownOpen = false;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Update FetchXML if it changed
        const newFetchXml = context.parameters.fetchXml.raw || "";
        if (newFetchXml !== this._fetchXml) {
            this._fetchXml = newFetchXml;
            // Clear current selection if FetchXML changed
            if (this._selectedRecord) {
                this.clearSelection();
            }
        }

        // Update selected value if it changed externally
        const newSelectedValue = context.parameters.selectedValue.raw || "";
        if (newSelectedValue !== this._value) {
            this._value = newSelectedValue;

            // Parse the new selected record
            if (this._value) {
                try {
                    this._selectedRecord = JSON.parse(this._value);
                } catch {
                    this._selectedRecord = null;
                }
            } else {
                this._selectedRecord = null;
            }

            this.updateDisplay();
        }
    }

    public getOutputs(): IOutputs {
        return {
            selectedValue: this._value || ""
        };
    }

    public destroy(): void {
        if (this._searchTimer) {
            clearTimeout(this._searchTimer);
        }
        // Clean up event listeners if needed
    }
}