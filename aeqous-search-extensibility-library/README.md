# Aeqous Search Box Extensibility Library

An extensibility library for the **Aeqous Modern Data Visualizer v1.5.4** that provides custom suggestion providers for the search box component.

## Overview

This SharePoint Framework (SPFx) library component extends the Aeqous Modern Data Visualizer search box with custom suggestion capabilities. It demonstrates how to create and implement custom search suggestions that integrate seamlessly with the Modern Data Visualizer solution.

## Features

- **Published Pages Only**: Automatically fetches and suggests only published SharePoint pages
- **News Page Filtering**: Excludes news pages (PromotedState=2) from suggestions, showing only regular site pages
- **Fuzzy String Matching**: Intelligent typo tolerance using Levenshtein distance algorithm
  - Finds similar matches even when users make spelling mistakes
  - Configurable similarity threshold (default: 0.6)
  - Prevents users from hitting dead ends with no results
- **Smart Result Ranking**: Suggestions sorted by relevance score
  - Exact matches scored highest (1.0)
  - Substring matches (0.95)
  - Prefix matches (0.9)
  - Fuzzy matches based on edit distance
- **Zero-Term Suggestions**: Display recent pages when search box is empty
- **Performance Optimized**: Built-in caching (5-minute cache) for faster response times
- **Configurable Options**: Admin-configurable settings through the property pane
  - Number of suggestions (3-20)
  - Enable/disable fuzzy matching
  - Adjust similarity threshold
  - Custom site URL for cross-site search
- **Extensible Architecture**: Built using the @aequos/extensibility framework (v1.5.0)

## Prerequisites

Before using this library, ensure you have:

- SharePoint Online environment
- Modern Data Visualizer v1.5.4 or compatible version installed
- SharePoint App Catalog (tenant or site collection level)
- Node.js v10.24.1 (as specified in `.nvmrc`)
- SPFx 1.12.1

## Installation

### Step 1: Build the Solution

```bash
# Install dependencies
npm install

# Build and package the solution
npm run build
```

This will create a `.sppkg` file in the `sharepoint/solution` folder.

### Step 2: Deploy to App Catalog

1. Navigate to your SharePoint App Catalog
2. Upload the `aeqous-search-extensibility-library.sppkg` file
3. When prompted, click **Deploy**

### Step 3: Register with Search Box Web Part

1. Add or edit a **Search Box** web part on your SharePoint page
2. Open the web part's property pane
3. Navigate to the last configuration page
4. In the **Extensibility configuration** section, click **Add**
5. Enter the library manifest ID: `8f7a9c2e-5b3d-4a1f-9e8c-7d6e5f4a3b2c`
6. Enable the library
7. Save your changes

### Step 4: Select the Custom Provider

1. In the Search Box web part settings
2. Go to the **Suggestions** section
3. Select **"Aeqous Custom Suggestions"** as your suggestion provider
4. Configure any additional options as needed
5. Save your changes

## Library Components

### AeqousSearchLibrary

The main library class that implements `IExtensibilityLibrary` from @aequos/extensibility. This class registers all custom extensions with the Modern Data Visualizer.

**Location**: `src/libraries/aeqousSearchLibrary/AeqousSearchLibrary.ts`

### AeqousCustomSuggestionProvider

A custom suggestion provider that extends `BaseSuggestionProvider`. This component provides:

- **SharePoint Search Integration**: Queries SharePoint Search API for published pages
- **Automatic Filtering**: Excludes news pages and unpublished content
- **Fuzzy Matching**: Levenshtein distance algorithm for typo tolerance
- **Dynamic Suggestions**: Suggestions that update based on user input with smart ranking
- **Zero-Term Suggestions**: Recent pages shown when search box is empty
- **Caching**: 5-minute cache for improved performance
- **Custom Actions**: Handle suggestion selection events

**Location**: `src/libraries/AeqousCustomSuggestionProvider.ts`

## Configuration Options

The custom suggestion provider includes configurable properties accessible through the property pane:

### Page Suggestion Settings
- **Site URL**: SharePoint site URL to search (leave empty to use current site)
- **Number of Suggestions**: Maximum number of suggestions to display (3-20, default: 10)

### Fuzzy Match Settings
- **Enable Fuzzy Matching**: Toggle fuzzy string matching on/off (default: enabled)
- **Similarity Threshold**: Minimum similarity score required (0.3-1.0, default: 0.6)
  - Lower values = more lenient matching (more results, may be less relevant)
  - Higher values = stricter matching (fewer results, more relevant)
  - Recommended: 0.6 for general use, 0.4 for very lenient matching

## How It Works

### Page Filtering

The library uses SharePoint Search API with the following filters:
```
ContentTypeId: 0x0101009D1CB255DA76424F860D91F20E6C4118* (Site Pages)
PromotedState: 0 (Regular pages only, excludes news where PromotedState=2)
IsDocument: 1 (Published pages only)
```

### Fuzzy Matching Algorithm

When users type search queries, the provider:
1. Fetches published pages from SharePoint
2. Calculates similarity score for each page title:
   - **Exact match**: Score = 1.0
   - **Contains substring**: Score = 0.95
   - **Starts with query**: Score = 0.9
   - **Fuzzy match**: Score = 1 - (edit_distance / max_length)
3. Filters results by similarity threshold
4. Sorts by score (highest first)
5. Returns top N results

**Example**: User types "Projct" (typo for "Project")
- "Project Overview" → Score: 0.93 (1 character edit distance)
- "Project Plan" → Score: 0.92
- "Team Project" → Score: 0.85
- All shown if threshold ≤ 0.85

## Customization

### Adjusting Search Scope

To search multiple site collections or modify the query:

Edit `_fetchPublishedPages()` in `AeqousCustomSuggestionProvider.ts`:
```typescript
// Current: searches current site or specified site
const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;

// To search entire tenant, modify the queryTemplate to remove site restriction
```

### Modifying Page Filters

To include different content types or adjust filters:

Edit the `queryTemplate` in `_fetchPublishedPages()`:
```typescript
// Current query excludes news (PromotedState:0)
// To include news pages, remove: AND PromotedState:0
// To include other content types, modify: ContentTypeId:...
```

### Adjusting Fuzzy Matching

To customize the similarity algorithm:

Edit `_calculateSimilarity()` in `AeqousCustomSuggestionProvider.ts`:
```typescript
// Adjust scoring bonuses:
if (str2.includes(str1)) {
    return 0.95; // Change this value (0-1)
}
```

### Adding Custom Properties

To add new configuration options:

1. Update the `IAeqousCustomSuggestionProviderProperties` interface
2. Add new property pane fields in `getPropertyPaneGroupsConfiguration()`
3. Use the properties in your suggestion logic

## Development

### Local Development

```bash
# Start the local development server
gulp serve
```

This will start the SharePoint Workbench where you can test your library.

### Debug in SharePoint

1. Run `gulp serve` to start the local server
2. Navigate to your SharePoint site's hosted workbench:
   ```
   https://your-tenant.sharepoint.com/_layouts/workbench.aspx
   ```
3. Add a Search Box web part and register your library's manifest ID
4. Set breakpoints in your TypeScript code

### Update Version

To update the library version:

```bash
gulp update-version --value 1.5.5
```

This will update version numbers in:
- `package-solution.json`
- `AeqousSearchLibrary.manifest.json`

## Project Structure

```
aeqous-search-extensibility-library/
├── config/                          # SPFx configuration files
│   ├── config.json                 # Bundle configuration
│   ├── package-solution.json       # Solution package settings
│   └── serve.json                  # Development server settings
├── src/
│   ├── libraries/
│   │   ├── aeqousSearchLibrary/
│   │   │   ├── AeqousSearchLibrary.ts          # Main library class
│   │   │   ├── AeqousSearchLibrary.manifest.json
│   │   │   └── loc/                             # Localization files
│   │   └── AeqousCustomSuggestionProvider.ts    # Custom provider implementation
│   └── index.ts                                 # Library entry point
├── package.json                     # Dependencies and scripts
├── tsconfig.json                   # TypeScript configuration
├── gulpfile.js                     # Build tasks
└── README.md                       # This file
```

## API Reference

### ISuggestion Interface

Suggestions returned by the provider should conform to:

```typescript
{
  displayText: string;           // Text shown to the user
  groupName: string;             // Group/category name
  description?: string;          // Additional description
  hoverText?: string;           // Tooltip text
  targetUrl?: string;           // URL to navigate to
  iconSrc?: string;             // Icon URL
  onSuggestionSelected: (suggestion: ISuggestion) => void;
}
```

### BaseSuggestionProvider Methods

Key methods to implement:

- `onInit()`: Initialize the provider
- `getSuggestions(queryText: string)`: Get suggestions for a query
- `getZeroTermSuggestions()`: Get suggestions when query is empty
- `isZeroTermSuggestionsEnabled`: Enable/disable zero-term suggestions
- `getPropertyPaneGroupsConfiguration()`: Define configuration UI

## Compatibility

- **Modern Data Visualizer**: v1.5.4 and compatible versions
- **@aequos/extensibility**: v1.5.0
- **SharePoint Framework**: v1.12.1
- **Node.js**: v10.24.1
- **TypeScript**: v3.3.x

## Troubleshooting

### Library Not Showing in Web Part

1. Verify the library is deployed in the app catalog
2. Ensure the manifest ID is correctly registered
3. Check browser console for any loading errors
4. Confirm the library is enabled in the extensibility configuration

### Suggestions Not Appearing

1. Verify the suggestion provider is selected in web part settings
2. Check browser console for errors in provider logic
3. Ensure `isZeroTermSuggestionsEnabled` returns `true` for zero-term suggestions
4. Verify suggestion objects have required properties

### Build Errors

1. Ensure Node.js version matches `.nvmrc` (v10.24.1)
2. Delete `node_modules` and run `npm install` again
3. Clear SPFx cache: `gulp clean`
4. Check TypeScript version compatibility

## Support

For issues related to:
- **This library**: Create an issue in your repository
- **Modern Data Visualizer**: Visit [Aeqous Documentation](https://www.aequos.ca)
- **@aequos/extensibility**: Check the [npm package](https://www.npmjs.com/package/@aequos/extensibility)

## License

Specify your license here.

## Additional Resources

- [Modern Data Visualizer Documentation](https://modern-data-visualizer.readthedocs.io/)
- [@aequos/extensibility Package](https://www.npmjs.com/package/@aequos/extensibility)
- [SharePoint Framework Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
- [SPFx Library Components](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/library-component-overview)

## Credits

Built for use with Aeqous Modern Data Visualizer by Franck Cornu and the Aeqous team.
