# Aeqous Search Box Extensibility Library

An extensibility library for the **Aeqous Modern Data Visualizer v1.5.4** that provides custom suggestion providers for the search box component.

## Overview

This SharePoint Framework (SPFx) library component extends the Aeqous Modern Data Visualizer search box with custom suggestion capabilities. It demonstrates how to create and implement custom search suggestions that integrate seamlessly with the Modern Data Visualizer solution.

## Features

- **Custom Suggestion Provider**: Provides intelligent search suggestions based on user input
- **Zero-Term Suggestions**: Display helpful suggestions even when the search box is empty
- **Configurable Options**: Admin-configurable settings through the property pane
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

- **Dynamic Suggestions**: Suggestions that update based on user input
- **Zero-Term Suggestions**: Pre-defined suggestions shown when search box is empty
- **Grouping**: Suggestions organized into logical groups
- **Custom Actions**: Handle suggestion selection events

**Location**: `src/libraries/AeqousCustomSuggestionProvider.ts`

## Configuration Options

The custom suggestion provider includes configurable properties accessible through the property pane:

- **Custom API Endpoint**: URL to a custom API for fetching suggestions
- **Number of Suggestions**: Maximum number of suggestions to display

## Customization

### Modifying Suggestions

To customize the suggestions, edit the `AeqousCustomSuggestionProvider.ts` file:

1. **Zero-Term Suggestions**: Update the `_zeroTermSuggestions` array in the `onInit()` method
2. **Query-Based Suggestions**: Modify the `_getCustomSuggestions()` method
3. **API Integration**: Uncomment and configure the API call example in `_getCustomSuggestions()`

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
