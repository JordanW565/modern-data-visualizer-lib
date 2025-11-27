# Building the Aeqous Search Extensibility Library

## Prerequisites

**IMPORTANT**: This project requires **Node.js v10.24.1** (or v12-14) as specified in the `.nvmrc` file. Building with Node.js v15+ will fail due to incompatibilities with SPFx 1.12.1 and node-sass.

### Required Software

- **Node.js**: v10.24.1 (recommended) or v12.13.0 - v14.x
- **npm**: v6.x or compatible with your Node version
- **gulp-cli**: Installed globally (optional, can use npx)

## Setting Up Your Environment

### Option 1: Using NVM (Recommended)

If you have nvm (Node Version Manager) installed:

```bash
# Install the correct Node version
nvm install 10.24.1

# Use the correct version
nvm use 10.24.1

# Verify the version
node --version
# Should output: v10.24.1
```

### Option 2: Using n (Node Version Manager)

If you have `n` installed:

```bash
# Install and use Node 10.24.1
n 10.24.1

# Verify
node --version
```

### Option 3: Manual Installation

Download and install Node.js v10.24.1 from:
- Official archive: https://nodejs.org/dist/v10.24.1/
- Or use Docker (see Docker section below)

## Build Instructions

Once you have the correct Node version:

### 1. Install Dependencies

```bash
cd aeqous-search-extensibility-library
npm install
```

This will install all required dependencies including:
- @aequos/extensibility v1.5.0
- SharePoint Framework v1.12.1 components
- Build tools and dependencies

### 2. Bundle the Solution

```bash
# Production bundle
gulp bundle --ship

# Or use npm script
npm run build
```

This command:
- Compiles TypeScript to JavaScript
- Bundles all assets
- Prepares the solution for packaging

### 3. Package the Solution

```bash
# Create the .sppkg package
gulp package-solution --ship
```

This command:
- Creates the SharePoint solution package (.sppkg)
- Outputs to: `sharepoint/solution/aeqous-search-extensibility-library.sppkg`

### 4. Combined Build

Use the npm script for a complete build:

```bash
npm run build
```

This runs both `gulp bundle --ship` and `gulp package-solution --ship`.

## Build Output

After a successful build, you'll find:

- **Compiled JavaScript**: `lib/` directory
- **SharePoint Package**: `sharepoint/solution/aeqous-search-extensibility-library.sppkg`
- **Temporary files**: `temp/` directory (can be ignored)

## Using Docker (Alternative)

If you can't install Node v10 locally, use Docker:

### Create a Dockerfile

```dockerfile
FROM node:10.24.1

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install

# Copy source code
COPY . .

# Build the solution
RUN npm run build

# The output will be in sharepoint/solution/
CMD ["bash"]
```

### Build with Docker

```bash
# Build the Docker image
docker build -t aeqous-search-lib-builder .

# Run the container and copy the output
docker run --rm -v $(pwd)/output:/app/sharepoint/solution aeqous-search-lib-builder
```

The `.sppkg` file will be in the `output/` directory.

## Troubleshooting

### node-sass Build Errors

**Symptom**: `Node Sass does not yet support your current environment`

**Solution**: You're using the wrong Node version. Switch to Node v10.24.1 or v12-14.

### Python Errors (gyp errors)

**Symptom**: `gyp ERR! configure error` or missing Python

**Solution**:
- Install Python 2.7 (required for node-gyp in older SPFx)
- Or use the Docker approach which includes Python

### Module Not Found Errors

**Symptom**: `Cannot find module '@aequos/extensibility'`

**Solution**:
```bash
# Clean and reinstall
rm -rf node_modules package-lock.json
npm install
```

### Build Hangs or Times Out

**Solution**:
```bash
# Clean build artifacts
gulp clean

# Or manually remove
rm -rf lib temp

# Rebuild
npm run build
```

## Verifying the Build

After building, verify the package:

```bash
# Check if .sppkg exists
ls -lh sharepoint/solution/*.sppkg

# Should output:
# aeqous-search-extensibility-library.sppkg
```

The package size should be approximately 300-500 KB.

## Deployment

Once built, deploy the `.sppkg` file:

1. Navigate to your SharePoint App Catalog
2. Upload `sharepoint/solution/aeqous-search-extensibility-library.sppkg`
3. Click **Deploy** when prompted
4. The library is now available for use in your Search Box web parts

## Version Management

To update the version number:

```bash
# Update to version 1.5.5 (for example)
gulp update-version --value 1.5.5

# Then rebuild
npm run build
```

This updates:
- `package-solution.json`
- `AeqousSearchLibrary.manifest.json`

## Clean Build

For a completely fresh build:

```bash
# Clean all build artifacts
gulp clean

# Remove dependencies
rm -rf node_modules

# Reinstall and build
npm install
npm run build
```

## CI/CD Integration

For automated builds in CI/CD pipelines:

```yaml
# Example GitHub Actions workflow
- uses: actions/setup-node@v2
  with:
    node-version: '10.24.1'

- run: npm ci
- run: npm run build

- uses: actions/upload-artifact@v2
  with:
    name: sppkg
    path: sharepoint/solution/*.sppkg
```

## Support

If you encounter build issues:

1. Verify Node version: `node --version` (must be v10.x or v12-14)
2. Check npm version: `npm --version` (should be v6.x)
3. Try clean build: `gulp clean && npm run build`
4. Check build logs in `temp/` directory for detailed errors

## Additional Resources

- [SharePoint Framework v1.12.1 Documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)
- [Node Version Manager (nvm)](https://github.com/nvm-sh/nvm)
- [SPFx Version Compatibility](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/compatibility)
