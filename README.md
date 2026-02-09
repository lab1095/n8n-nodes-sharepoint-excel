# n8n-nodes-sharepoint-excel

[![npm version](https://img.shields.io/npm/v/@lab1095/n8n-nodes-sharepoint-excel.svg)](https://www.npmjs.com/package/@lab1095/n8n-nodes-sharepoint-excel)

> ⚠️ **Warning:** Please review the [known limitations](#limitations) before using this node.

This is an n8n community node that provides Excel file operations for SharePoint via Microsoft Graph API. Unlike native n8n nodes that rely on WAC tokens, this node **downloads and modifies files directly**, bypassing common SharePoint Excel issues.

![n8n.io - Workflow Automation](https://raw.githubusercontent.com/n8n-io/n8n/master/assets/n8n-logo.png)

## Features

- 🔓 **Bypasses WAC Token Issues** - No more token expiration or permission errors with SharePoint files
- 📊 **Full Excel Support** - Read, write, append, update, and upsert rows in `.xlsx` files
- 📋 **Table Operations** - Read data from formal Excel Tables with column lookups
- 📁 **Workbook Management** - Create sheets, list workbooks, and manage files
- 🔍 **Smart Dropdowns** - Searchable site, drive, file, and sheet selectors
- ⚡ **Reliable** - Uses `exceljs` library for robust Excel parsing

## Why This Node?

The native n8n Microsoft Excel 365 node uses WAC (Web Application Companion) tokens which often fail with SharePoint:

- ❌ Token expiration issues
- ❌ Permission complexities with SharePoint sites
- ❌ WAC service availability problems

This node takes a different approach: **download → modify → upload**. Simple and reliable.

> 📖 For a deep technical analysis, see [Why Not WAC?](docs/research/wac-tokens-research.md)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

```bash
npm install @lab1095/n8n-nodes-sharepoint-excel
```

## Operations

### 📊 Sheet

| Operation       | Description                                      |
| --------------- | ------------------------------------------------ |
| **Get Sheets**  | List all worksheets in the workbook              |
| **Get Rows**    | Read rows with configurable header and start row |
| **Append Rows** | Add new rows to the end of a sheet               |
| **Update**      | Update a specific cell by reference (e.g., A1)   |
| **Upsert Rows** | Insert or update rows based on a key column      |
| **Clear**       | Clear all data from a sheet                      |
| **Delete**      | Delete a sheet from the workbook                 |

### 📋 Table

| Operation       | Description                           |
| --------------- | ------------------------------------- |
| **Get Rows**    | Retrieve all rows from an Excel table |
| **Get Columns** | Get column definitions from a table   |
| **Lookup**      | Find a row by column value            |

### 📁 Workbook

| Operation         | Description                        |
| ----------------- | ---------------------------------- |
| **Add Sheet**     | Create a new sheet in the workbook |
| **Delete**        | Delete the workbook file           |
| **Get Workbooks** | List all Excel files in the drive  |

## Credentials

This node requires **Microsoft Graph OAuth2** credentials.

### Setup

1. **Register an app in Azure AD:**
   - Go to Azure Portal → Azure Active Directory → App registrations
   - Click "New registration"
   - Set redirect URI to your n8n OAuth callback URL

2. **Configure API permissions:**
   - Add Microsoft Graph **delegated** permissions:
     - `Sites.Read.All` - Read sites
     - `Files.ReadWrite.All` - Read and write files
   - Grant admin consent

3. **Create a client secret:**
   - Go to "Certificates & secrets"
   - Create new client secret and copy the value

4. **Configure in n8n:**
   - Add credentials of type "Microsoft Graph OAuth2 API"
   - Enter Client ID and Client Secret
   - Complete OAuth2 flow

> 📖 See [Security and Permissions](docs/security-and-permissions.md) for enterprise considerations.

## Usage

### Finding Required IDs

The node provides searchable dropdowns to select:

- **Site** - Your SharePoint site
- **Drive** - The document library
- **File** - The Excel file
- **Sheet/Table** - The specific sheet or table

Or enter IDs manually:

- **Site ID**: `contoso.sharepoint.com,site-guid,web-guid`
- **Drive ID**: `b!xxxxx...`
- **File ID**: Item ID from SharePoint

### Reading Data

Configure:

- **Header Row** - Which row contains column headers (default: 1)
- **Start Row** - First data row to read (default: 2)
- **Max Rows** - Limit rows returned (0 = all)

### Writing Data

Provide data as JSON with column headers as keys:

```json
{ "Name": "John", "Email": "john@example.com" }
```

For multiple rows, use an array of objects.

## Limitations

| Limitation                        | Impact                                                                                                                                                 |
| --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Excel Tables may be corrupted** | Write operations on files with formal Excel Tables may corrupt them. Read operations are safe. [Details](docs/limitations/exceljs-table-limitation.md) |
| **Full file download/upload**     | Each operation downloads and re-uploads the entire file                                                                                                |
| **No table creation/deletion**    | Can only read table data, not create/delete tables                                                                                                     |

> **Note:** "Excel Tables" refers to Insert → Table, not regular cell data. Most users are unaffected.

## Documentation

| Document                                                               | Description                                |
| ---------------------------------------------------------------------- | ------------------------------------------ |
| [Security and Permissions](docs/security-and-permissions.md)           | OAuth scopes and enterprise considerations |
| [File Locking Behavior](docs/limitations/file-locking-behavior.md)     | Why files get locked and how to handle it  |
| [Excel Table Limitation](docs/limitations/exceljs-table-limitation.md) | Details on table corruption risk           |
| [Why Not WAC?](docs/research/wac-tokens-research.md)                   | Deep dive into WAC tokens and our approach |

## Compatibility

- n8n version 1.x and above
- Microsoft Graph API v1.0
- `.xlsx` files only

## Resources

- [n8n Community Nodes](https://docs.n8n.io/integrations/#community-nodes)
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview)
- [SharePoint REST API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)

## License

[MIT](LICENSE.md)
