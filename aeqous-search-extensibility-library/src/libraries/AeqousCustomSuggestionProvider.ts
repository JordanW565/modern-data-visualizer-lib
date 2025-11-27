import { BaseSuggestionProvider, ISuggestion } from "@aequos/extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';

export interface IAeqousCustomSuggestionProviderProperties {
    customApiEndpoint?: string;
    numberOfSuggestions?: number;
}

/**
 * Custom Suggestion Provider for Aeqous Search Box
 * This provider demonstrates how to create custom search suggestions
 * for the Modern Data Visualizer search box component
 */
export class AeqousCustomSuggestionProvider extends BaseSuggestionProvider<IAeqousCustomSuggestionProviderProperties> {

    private _zeroTermSuggestions: ISuggestion[] = [];

    /**
     * Initialize the suggestion provider
     * This is called when the provider is first loaded
     */
    public async onInit(): Promise<void> {
        this._onSuggestionSelected = this._onSuggestionSelected.bind(this);

        // Define zero-term suggestions (shown when search box is empty)
        this._zeroTermSuggestions = [
            {
                displayText: 'Recent Documents',
                groupName: 'Quick Access',
                hoverText: 'View your recently accessed documents',
                description: 'Recently accessed documents',
                onSuggestionSelected: this._onSuggestionSelected,
            },
            {
                displayText: 'Popular Content',
                groupName: 'Quick Access',
                hoverText: 'View the most popular content in your organization',
                description: 'Most viewed content',
                onSuggestionSelected: this._onSuggestionSelected,
            },
            {
                displayText: 'My Files',
                groupName: 'Quick Access',
                hoverText: 'View your personal files',
                description: 'Your personal documents',
                onSuggestionSelected: this._onSuggestionSelected,
            }
        ];
    }

    /**
     * Enable or disable zero-term suggestions
     * Return true to show suggestions when the search box is empty
     */
    public get isZeroTermSuggestionsEnabled(): boolean {
        return true;
    }

    /**
     * Get suggestions based on the user's query
     * @param queryText The text entered by the user
     * @returns Array of suggestions matching the query
     */
    public async getSuggestions(queryText: string): Promise<ISuggestion[]> {
        return this._getCustomSuggestions(queryText);
    }

    /**
     * Get zero-term suggestions (shown when search box is empty)
     * @returns Array of zero-term suggestions
     */
    public async getZeroTermSuggestions(): Promise<ISuggestion[]> {
        return this._zeroTermSuggestions;
    }

    /**
     * Handle when a suggestion is selected by the user
     * @param suggestion The selected suggestion
     */
    private _onSuggestionSelected = (suggestion: ISuggestion): void => {
        console.log('Aeqous Custom Suggestion Selected:', suggestion);

        // You can add custom logic here, such as:
        // - Navigate to a specific URL
        // - Trigger a custom action
        // - Log analytics

        if (suggestion.targetUrl) {
            window.location.href = suggestion.targetUrl;
        }
    }

    /**
     * Generate custom suggestions based on the query text
     * In a real implementation, this would call an API or search service
     * @param queryText The user's search query
     * @returns Array of custom suggestions
     */
    private _getCustomSuggestions = async (queryText: string): Promise<ISuggestion[]> => {

        // Example: Filter suggestions based on query text
        const allSuggestions: ISuggestion[] = [
            {
                displayText: 'Project Documentation',
                groupName: 'Documents',
                description: 'Technical documentation for projects',
                hoverText: `Search for project documentation related to "${queryText}"`,
                onSuggestionSelected: this._onSuggestionSelected,
            },
            {
                displayText: 'Team Resources',
                groupName: 'Resources',
                description: 'Shared team resources and files',
                hoverText: `Find team resources related to "${queryText}"`,
                onSuggestionSelected: this._onSuggestionSelected,
            },
            {
                displayText: 'Knowledge Base Articles',
                groupName: 'Knowledge Base',
                description: 'Help articles and guides',
                hoverText: `Search knowledge base for "${queryText}"`,
                onSuggestionSelected: this._onSuggestionSelected,
            },
            {
                displayText: 'Training Materials',
                groupName: 'Training',
                description: 'Training videos and guides',
                hoverText: `Find training materials for "${queryText}"`,
                onSuggestionSelected: this._onSuggestionSelected,
            }
        ];

        // Filter suggestions based on query text
        if (!queryText || queryText.trim() === '') {
            return [];
        }

        const query = queryText.toLowerCase().trim();
        return allSuggestions.filter(suggestion =>
            suggestion.displayText.toLowerCase().includes(query) ||
            suggestion.description.toLowerCase().includes(query) ||
            suggestion.groupName.toLowerCase().includes(query)
        );

        /*
         * EXAMPLE: Call a custom API endpoint for suggestions
         *
         * if (this.properties.customApiEndpoint) {
         *     try {
         *         const response = await fetch(
         *             `${this.properties.customApiEndpoint}?query=${encodeURIComponent(queryText)}&count=${this.properties.numberOfSuggestions || 10}`
         *         );
         *         const data = await response.json();
         *
         *         return data.suggestions.map(item => ({
         *             displayText: item.title,
         *             groupName: item.category,
         *             description: item.description,
         *             hoverText: item.tooltip,
         *             targetUrl: item.url,
         *             iconSrc: item.iconUrl,
         *             onSuggestionSelected: this._onSuggestionSelected
         *         }));
         *     } catch (error) {
         *         console.error('Error fetching custom suggestions:', error);
         *         return [];
         *     }
         * }
         */
    }

    /**
     * Configure property pane fields for this suggestion provider
     * This allows administrators to configure the provider's behavior
     */
    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
        return [
            {
                groupName: 'Aeqous Custom Suggestion Settings',
                groupFields: [
                    PropertyPaneTextField('providerProperties.customApiEndpoint', {
                        label: 'Custom API Endpoint',
                        description: 'Optional: URL to your custom suggestion API',
                        placeholder: 'https://api.example.com/suggestions'
                    }),
                    PropertyPaneTextField('providerProperties.numberOfSuggestions', {
                        label: 'Number of Suggestions',
                        description: 'Maximum number of suggestions to display (default: 10)',
                        placeholder: '10'
                    })
                ]
            }
        ];
    }
}
