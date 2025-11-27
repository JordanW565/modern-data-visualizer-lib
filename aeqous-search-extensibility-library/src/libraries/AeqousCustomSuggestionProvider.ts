import { BaseSuggestionProvider, ISuggestion } from "@aequos/extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField, PropertyPaneSlider, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IAeqousCustomSuggestionProviderProperties {
    numberOfSuggestions?: number;
    fuzzyMatchThreshold?: number;
    enableFuzzyMatch?: boolean;
    siteUrl?: string;
}

/**
 * SharePoint Page item interface
 */
interface ISharePointPage {
    Title: string;
    Path: string;
    Description?: string;
    PromotedState?: number;
}

/**
 * Suggestion with similarity score for sorting
 */
interface ISuggestionWithScore extends ISuggestion {
    score: number;
}

/**
 * Custom Suggestion Provider for Aeqous Search Box
 * Features:
 * - Fetches only published SharePoint pages (excludes news pages)
 * - Fuzzy string matching for typo tolerance
 * - Configurable similarity threshold
 */
export class AeqousCustomSuggestionProvider extends BaseSuggestionProvider<IAeqousCustomSuggestionProviderProperties> {

    private _zeroTermSuggestions: ISuggestion[] = [];
    private _cachedPages: ISharePointPage[] = [];
    private _cacheExpiry: number = 0;
    private readonly CACHE_DURATION_MS: number = 5 * 60 * 1000; // 5 minutes

    /**
     * Initialize the suggestion provider
     * This is called when the provider is first loaded
     */
    public async onInit(): Promise<void> {
        this._onSuggestionSelected = this._onSuggestionSelected.bind(this);

        // Load recent pages for zero-term suggestions
        await this._loadRecentPages();
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
        if (!queryText || queryText.trim() === '') {
            return [];
        }

        return this._getPageSuggestions(queryText);
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

        if (suggestion.targetUrl) {
            window.location.href = suggestion.targetUrl;
        }
    }

    /**
     * Load recent pages for zero-term suggestions
     */
    private async _loadRecentPages(): Promise<void> {
        try {
            const pages = await this._fetchPublishedPages('*', 5);

            this._zeroTermSuggestions = pages.map(page => ({
                displayText: page.Title,
                groupName: 'Recent Pages',
                description: page.Description || 'Recent page',
                hoverText: page.Title,
                targetUrl: page.Path,
                onSuggestionSelected: this._onSuggestionSelected
            }));
        } catch (error) {
            console.error('Error loading recent pages for zero-term suggestions:', error);
            // Fallback to default suggestions
            this._zeroTermSuggestions = [
                {
                    displayText: 'Browse All Pages',
                    groupName: 'Quick Access',
                    hoverText: 'View all pages',
                    description: 'Browse all available pages',
                    onSuggestionSelected: this._onSuggestionSelected,
                }
            ];
        }
    }

    /**
     * Get page suggestions with fuzzy matching support
     * @param queryText The user's search query
     * @returns Array of page suggestions sorted by relevance
     */
    private _getPageSuggestions = async (queryText: string): Promise<ISuggestion[]> => {
        try {
            // Fetch published pages matching the query
            const maxResults = this.properties.numberOfSuggestions || 10;
            const pages = await this._fetchPublishedPages(queryText, maxResults * 2);

            if (pages.length === 0) {
                return [];
            }

            const enableFuzzy = this.properties.enableFuzzyMatch !== false; // Default to true
            const threshold = this.properties.fuzzyMatchThreshold || 0.6;

            // Score and filter pages based on similarity
            const scoredSuggestions: ISuggestionWithScore[] = pages.map(page => {
                let score: number;

                if (enableFuzzy) {
                    // Calculate similarity score with fuzzy matching
                    score = this._calculateSimilarity(queryText.toLowerCase(), page.Title.toLowerCase());
                } else {
                    // Exact match scoring
                    const titleLower = page.Title.toLowerCase();
                    const queryLower = queryText.toLowerCase();
                    score = titleLower.includes(queryLower) ? 1 : 0;
                }

                return {
                    displayText: page.Title,
                    groupName: 'Pages',
                    description: page.Description || 'SharePoint page',
                    hoverText: page.Title,
                    targetUrl: page.Path,
                    onSuggestionSelected: this._onSuggestionSelected,
                    score: score
                };
            });

            // Filter by threshold and sort by score (descending)
            const filteredSuggestions = scoredSuggestions
                .filter(s => s.score >= threshold)
                .sort((a, b) => b.score - a.score)
                .slice(0, maxResults);

            // Remove score property before returning
            return filteredSuggestions.map(({ score, ...suggestion }) => suggestion);

        } catch (error) {
            console.error('Error getting page suggestions:', error);
            return [];
        }
    }

    /**
     * Fetch published SharePoint pages (excluding news pages)
     * @param queryText Search query text
     * @param rowLimit Maximum number of results
     * @returns Array of SharePoint pages
     */
    private async _fetchPublishedPages(queryText: string, rowLimit: number): Promise<ISharePointPage[]> {
        // Check cache first
        const now = Date.now();
        if (this._cachedPages.length > 0 && now < this._cacheExpiry) {
            return this._cachedPages;
        }

        try {
            const spHttpClient: SPHttpClient = this.serviceScope.consume(SPHttpClient.serviceKey);
            const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;

            // Build SharePoint Search query
            // ContentTypeId starts with 0x0101009D1CB255DA76424F860D91F20E6C4118 for Site Pages
            // PromotedState=0 means regular page (not news: PromotedState=2)
            // IsDocument=1 for pages
            const searchQuery = queryText === '*' ? '' : queryText;

            const queryTemplate = searchQuery
                ? `(Title:"${searchQuery}*" OR Title:"*${searchQuery}*") AND ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118* AND PromotedState:0 AND IsDocument:1`
                : `ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118* AND PromotedState:0 AND IsDocument:1`;

            const searchUrl = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(queryTemplate)}'` +
                `&rowlimit=${rowLimit}` +
                `&selectproperties='Title,Path,Description,PromotedState'` +
                `&sortlist='LastModifiedTime:descending'` +
                `&trimduplicates=false`;

            const response: SPHttpClientResponse = await spHttpClient.get(
                searchUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Search query failed: ${response.statusText}`);
            }

            const data = await response.json();
            const results = data.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

            const pages: ISharePointPage[] = results.map((row: any) => {
                const cells = row.Cells;
                return {
                    Title: this._getCellValue(cells, 'Title') || 'Untitled',
                    Path: this._getCellValue(cells, 'Path') || '',
                    Description: this._getCellValue(cells, 'Description') || '',
                    PromotedState: parseInt(this._getCellValue(cells, 'PromotedState') || '0', 10)
                };
            }).filter((page: ISharePointPage) =>
                // Additional filter to ensure we only get regular pages (not news)
                page.PromotedState === 0 && page.Path && page.Path.length > 0
            );

            // Update cache
            this._cachedPages = pages;
            this._cacheExpiry = now + this.CACHE_DURATION_MS;

            return pages;

        } catch (error) {
            console.error('Error fetching published pages:', error);
            return [];
        }
    }

    /**
     * Extract cell value from SharePoint search results
     * @param cells Array of cells from search result row
     * @param key The key to extract
     * @returns The cell value or null
     */
    private _getCellValue(cells: any[], key: string): string | null {
        const cell = cells.find((c: any) => c.Key === key);
        return cell ? cell.Value : null;
    }

    /**
     * Calculate similarity between two strings using Levenshtein distance
     * Returns a score between 0 and 1 (1 being most similar)
     * @param str1 First string
     * @param str2 Second string
     * @returns Similarity score (0-1)
     */
    private _calculateSimilarity(str1: string, str2: string): number {
        // Handle exact match or substring match with bonus
        if (str1 === str2) {
            return 1.0;
        }
        if (str2.includes(str1)) {
            return 0.95;
        }
        if (str2.startsWith(str1)) {
            return 0.9;
        }

        // Calculate Levenshtein distance
        const distance = this._levenshteinDistance(str1, str2);
        const maxLength = Math.max(str1.length, str2.length);

        if (maxLength === 0) {
            return 1.0;
        }

        // Convert distance to similarity score (0-1)
        return 1 - (distance / maxLength);
    }

    /**
     * Calculate Levenshtein distance between two strings
     * @param str1 First string
     * @param str2 Second string
     * @returns Edit distance
     */
    private _levenshteinDistance(str1: string, str2: string): number {
        const len1 = str1.length;
        const len2 = str2.length;
        const matrix: number[][] = [];

        // Initialize matrix
        for (let i = 0; i <= len1; i++) {
            matrix[i] = [i];
        }
        for (let j = 0; j <= len2; j++) {
            matrix[0][j] = j;
        }

        // Fill matrix
        for (let i = 1; i <= len1; i++) {
            for (let j = 1; j <= len2; j++) {
                const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
                matrix[i][j] = Math.min(
                    matrix[i - 1][j] + 1,      // deletion
                    matrix[i][j - 1] + 1,      // insertion
                    matrix[i - 1][j - 1] + cost // substitution
                );
            }
        }

        return matrix[len1][len2];
    }

    /**
     * Configure property pane fields for this suggestion provider
     * This allows administrators to configure the provider's behavior
     */
    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
        return [
            {
                groupName: 'Page Suggestion Settings',
                groupFields: [
                    PropertyPaneTextField('providerProperties.siteUrl', {
                        label: 'Site URL',
                        description: 'SharePoint site URL to search (leave empty for current site)',
                        placeholder: 'https://contoso.sharepoint.com/sites/yoursite'
                    }),
                    PropertyPaneSlider('providerProperties.numberOfSuggestions', {
                        label: 'Number of Suggestions',
                        min: 3,
                        max: 20,
                        value: 10,
                        showValue: true,
                        step: 1
                    })
                ]
            },
            {
                groupName: 'Fuzzy Match Settings',
                groupFields: [
                    PropertyPaneToggle('providerProperties.enableFuzzyMatch', {
                        label: 'Enable Fuzzy Matching',
                        onText: 'Enabled',
                        offText: 'Disabled',
                        checked: true
                    }),
                    PropertyPaneSlider('providerProperties.fuzzyMatchThreshold', {
                        label: 'Similarity Threshold',
                        min: 0.3,
                        max: 1.0,
                        value: 0.6,
                        showValue: true,
                        step: 0.05,
                        disabled: !(this.properties.enableFuzzyMatch !== false)
                    })
                ]
            }
        ];
    }
}
