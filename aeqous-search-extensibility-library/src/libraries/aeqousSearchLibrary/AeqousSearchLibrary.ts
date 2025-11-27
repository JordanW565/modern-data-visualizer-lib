import { ServiceKey } from "@microsoft/sp-core-library";
import {
    IExtensibilityLibrary,
    IDataSourceDefinition,
    ILayoutDefinition,
    IComponentDefinition,
    ISuggestionProviderDefinition,
    ISuggestionProvider
} from "@aequos/extensibility";
import { AeqousCustomSuggestionProvider } from "../AeqousCustomSuggestionProvider";

/**
 * Aeqous Search Box Extensibility Library
 *
 * This library provides custom suggestion providers for the Aeqous Modern Data Visualizer
 * search box component (version 1.5.4).
 *
 * To use this library:
 * 1. Deploy the .sppkg file to your SharePoint app catalog
 * 2. In the Search Box web part, go to the last property pane page
 * 3. In "Extensibility configuration", add this library's manifest ID
 * 4. The custom suggestion provider will be available for selection
 */
export class AeqousSearchLibrary implements IExtensibilityLibrary {

    /**
     * Register custom suggestion providers
     * These will be available in the Search Box web part
     */
    public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
        return [
            {
                name: 'Aeqous Custom Suggestions',
                key: 'AeqousCustomSuggestionProvider',
                description: 'Custom suggestion provider for Aeqous search with configurable options',
                serviceKey: ServiceKey.create<ISuggestionProvider>(
                    'Aeqous:CustomSuggestionProvider',
                    AeqousCustomSuggestionProvider
                )
            }
        ];
    }

    /**
     * Custom layouts (not used for search box)
     * Return empty array as search box doesn't support custom layouts
     */
    public getCustomLayouts(): ILayoutDefinition[] {
        return [];
    }

    /**
     * Custom web components (not used for search box)
     * Return empty array - web components are loaded via Data Visualizer
     */
    public getCustomWebComponents(): IComponentDefinition<any>[] {
        return [];
    }

    /**
     * Custom data sources (not used for search box)
     * Return empty array as search box doesn't support custom data sources
     */
    public getCustomDataSources(): IDataSourceDefinition[] {
        return [];
    }

    /**
     * Register custom Handlebars helpers (not used for search box)
     * Search box doesn't use Handlebars templating
     */
    public registerHandlebarsCustomizations?(namespace: typeof Handlebars): void {
        // Not applicable for search box
    }
}
