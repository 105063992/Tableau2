# SharePoint Metadata Inheritance Flow

## Flow Overview
**Name:** SP - Inherit Parent Folder Metadata  
**Trigger:** When a file or folder is created in SharePoint  
**Purpose:** Automatically copy metadata from parent folder to newly created items/subfolders

---

## Flow Architecture

```
┌─────────────────────────────────────────┐
│  TRIGGER: When a file is created        │
└─────────────────┬───────────────────────┘
                  ▼
┌─────────────────────────────────────────┐
│  INITIALIZE: Set up variables           │
└─────────────────┬───────────────────────┘
                  ▼
┌─────────────────────────────────────────┐
│  EXTRACT: Get parent folder path        │
└─────────────────┬───────────────────────┘
                  ▼
┌─────────────────────────────────────────┐
│  GET: Retrieve parent folder metadata   │
└─────────────────┬───────────────────────┘
                  ▼
┌─────────────────────────────────────────┐
│  CONDITION: Check if metadata exists    │
└─────────────────┬───────────────────────┘
                  ▼
┌─────────────────────────────────────────┐
│  UPDATE: Apply metadata to new item     │
└─────────────────────────────────────────┘
```

---

## Detailed Step Configuration

### Step 1: TRIGGER - When a file is created (properties only)

**Action:** SharePoint - When a file is created (properties only)

| Setting | Value |
|---------|-------|
| Site Address | `[Your SharePoint Site URL]` |
| Library Name | `[Your Document Library]` |

**Outputs Used:**
- `{FullPath}` - Full path of the created item
- `ID` - Item ID
- `{IsFolder}` - Boolean indicating if item is folder

---

### Step 2: INIT - Initialize varParentFolderPath

**Action:** Initialize variable

| Setting | Value |
|---------|-------|
| Name | `varParentFolderPath` |
| Type | String |
| Value | *(empty)* |

---

### Step 3: INIT - Initialize varMetadataFound

**Action:** Initialize variable

| Setting | Value |
|---------|-------|
| Name | `varMetadataFound` |
| Type | Boolean |
| Value | `false` |

---

### Step 4: COMPOSE - Extract Item Name from Path

**Action:** Compose

**Expression:**
```
last(split(triggerOutputs()?['body/{FullPath}'], '/'))
```

**Rename to:** `COMPOSE - Extract Item Name`

---

### Step 5: COMPOSE - Calculate Parent Folder Path

**Action:** Compose

**Expression:**
```
substring(
    triggerOutputs()?['body/{FullPath}'],
    0,
    sub(
        length(triggerOutputs()?['body/{FullPath}']),
        add(length(outputs('COMPOSE_-_Extract_Item_Name')), 1)
    )
)
```

**Rename to:** `COMPOSE - Calculate Parent Folder Path`

---

### Step 6: SET - Assign Parent Folder Path

**Action:** Set variable

| Setting | Value |
|---------|-------|
| Name | `varParentFolderPath` |
| Value | `@{outputs('COMPOSE_-_Calculate_Parent_Folder_Path')}` |

---

### Step 7: SCOPE - Get Parent Folder Metadata

**Action:** Scope (to handle errors gracefully)

**Rename to:** `SCOPE - Get Parent Folder Metadata`

#### Inside Scope:

##### Step 7a: HTTP - Get Parent Folder Properties

**Action:** Send an HTTP request to SharePoint

| Setting | Value |
|---------|-------|
| Site Address | `[Your SharePoint Site URL]` |
| Method | `GET` |
| Uri | `_api/web/GetFolderByServerRelativeUrl('@{variables('varParentFolderPath')}')/ListItemAllFields` |
| Headers | `Accept: application/json;odata=verbose` |

**Rename to:** `HTTP - Get Parent Folder Properties`

##### Step 7b: PARSE - Parse Parent Folder Response

**Action:** Parse JSON

| Setting | Value |
|---------|-------|
| Content | `@{body('HTTP_-_Get_Parent_Folder_Properties')}` |
| Schema | *(see schema below)* |

**Schema:**
```json
{
    "type": "object",
    "properties": {
        "d": {
            "type": "object",
            "properties": {
                "Id": { "type": "integer" },
                "Title": { "type": ["string", "null"] },
                "YourMetadataField1": { "type": ["string", "null"] },
                "YourMetadataField2": { "type": ["string", "null"] },
                "YourMetadataField3": { "type": ["string", "null"] }
            }
        }
    }
}
```

> **Note:** Update schema with your actual metadata column internal names

**Rename to:** `PARSE - Parse Parent Folder Response`

##### Step 7c: SET - Mark Metadata Found

**Action:** Set variable

| Setting | Value |
|---------|-------|
| Name | `varMetadataFound` |
| Value | `true` |

---

### Step 8: CONDITION - Check If Metadata Exists

**Action:** Condition

**Condition:**
```
@equals(variables('varMetadataFound'), true)
```

**Rename to:** `CONDITION - Check If Parent Has Metadata`

---

### Step 9: (If Yes) UPDATE - Apply Metadata to New Item

**Action:** SharePoint - Update file properties

| Setting | Value |
|---------|-------|
| Site Address | `[Your SharePoint Site URL]` |
| Library Name | `[Your Document Library]` |
| Id | `@{triggerOutputs()?['body/ID']}` |
| **MetadataField1** | `@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField1']}` |
| **MetadataField2** | `@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField2']}` |
| **MetadataField3** | `@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField3']}` |

> **Note:** Add all metadata columns you want to inherit

**Rename to:** `UPDATE - Apply Parent Metadata to New Item`

---

### Step 10: (If No) COMPOSE - Log No Metadata

**Action:** Compose

| Setting | Value |
|---------|-------|
| Inputs | `No metadata found on parent folder or parent is library root` |

**Rename to:** `COMPOSE - Log No Metadata Found`

---

## Configure Run After Settings

For the **CONDITION** step, configure "Run After" to run even if the scope fails:

1. Click the three dots (...) on `CONDITION - Check If Parent Has Metadata`
2. Select "Configure run after"
3. Check:
   - ✅ is successful
   - ✅ has failed
   - ✅ is skipped

This ensures the flow continues gracefully if the parent folder has no metadata.

---

## Alternative: Handling Multiple Metadata Fields Dynamically

If you have many metadata fields, use this HTTP approach instead of Update file properties:

### Alternative Step 9: HTTP - Update Item Metadata

**Action:** Send an HTTP request to SharePoint

| Setting | Value |
|---------|-------|
| Site Address | `[Your SharePoint Site URL]` |
| Method | `POST` |
| Uri | `_api/web/lists/getbytitle('[Your Library Name]')/items(@{triggerOutputs()?['body/ID']})` |
| Headers | See below |
| Body | See below |

**Headers:**
```json
{
    "Accept": "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    "IF-MATCH": "*",
    "X-HTTP-Method": "MERGE"
}
```

**Body:**
```json
{
    "__metadata": {
        "type": "SP.Data.[YourLibraryName]Item"
    },
    "YourMetadataField1": "@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField1']}",
    "YourMetadataField2": "@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField2']}",
    "YourMetadataField3": "@{body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['YourMetadataField3']}"
}
```

---

## Full Flow JSON Export

Below is the flow definition you can import:

```json
{
    "definition": {
        "$schema": "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#",
        "contentVersion": "1.0.0.0",
        "triggers": {
            "TRIGGER_-_When_a_file_is_created_(properties_only)": {
                "type": "OpenApiConnectionWebhook",
                "inputs": {
                    "host": {
                        "connectionName": "shared_sharepointonline",
                        "operationId": "OnFileCreatedProperties",
                        "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                    },
                    "parameters": {
                        "dataset": "[YOUR_SITE_URL]",
                        "table": "[YOUR_LIBRARY_GUID]"
                    }
                }
            }
        },
        "actions": {
            "INIT_-_Initialize_varParentFolderPath": {
                "type": "InitializeVariable",
                "inputs": {
                    "variables": [
                        {
                            "name": "varParentFolderPath",
                            "type": "string",
                            "value": ""
                        }
                    ]
                },
                "runAfter": {}
            },
            "INIT_-_Initialize_varMetadataFound": {
                "type": "InitializeVariable",
                "inputs": {
                    "variables": [
                        {
                            "name": "varMetadataFound",
                            "type": "boolean",
                            "value": false
                        }
                    ]
                },
                "runAfter": {
                    "INIT_-_Initialize_varParentFolderPath": ["Succeeded"]
                }
            },
            "COMPOSE_-_Extract_Item_Name": {
                "type": "Compose",
                "inputs": "@last(split(triggerOutputs()?['body/{FullPath}'], '/'))",
                "runAfter": {
                    "INIT_-_Initialize_varMetadataFound": ["Succeeded"]
                }
            },
            "COMPOSE_-_Calculate_Parent_Folder_Path": {
                "type": "Compose",
                "inputs": "@substring(triggerOutputs()?['body/{FullPath}'], 0, sub(length(triggerOutputs()?['body/{FullPath}']), add(length(outputs('COMPOSE_-_Extract_Item_Name')), 1)))",
                "runAfter": {
                    "COMPOSE_-_Extract_Item_Name": ["Succeeded"]
                }
            },
            "SET_-_Assign_Parent_Folder_Path": {
                "type": "SetVariable",
                "inputs": {
                    "name": "varParentFolderPath",
                    "value": "@outputs('COMPOSE_-_Calculate_Parent_Folder_Path')"
                },
                "runAfter": {
                    "COMPOSE_-_Calculate_Parent_Folder_Path": ["Succeeded"]
                }
            },
            "SCOPE_-_Get_Parent_Folder_Metadata": {
                "type": "Scope",
                "actions": {
                    "HTTP_-_Get_Parent_Folder_Properties": {
                        "type": "OpenApiConnection",
                        "inputs": {
                            "host": {
                                "connectionName": "shared_sharepointonline",
                                "operationId": "HttpRequest",
                                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                            },
                            "parameters": {
                                "dataset": "[YOUR_SITE_URL]",
                                "parameters/method": "GET",
                                "parameters/uri": "_api/web/GetFolderByServerRelativeUrl('@{variables('varParentFolderPath')}')/ListItemAllFields",
                                "parameters/headers": {
                                    "Accept": "application/json;odata=verbose"
                                }
                            }
                        }
                    },
                    "PARSE_-_Parse_Parent_Folder_Response": {
                        "type": "ParseJson",
                        "inputs": {
                            "content": "@body('HTTP_-_Get_Parent_Folder_Properties')",
                            "schema": {
                                "type": "object",
                                "properties": {
                                    "d": {
                                        "type": "object",
                                        "properties": {
                                            "Id": {"type": "integer"},
                                            "Title": {"type": ["string", "null"]},
                                            "MetadataField1": {"type": ["string", "null"]},
                                            "MetadataField2": {"type": ["string", "null"]}
                                        }
                                    }
                                }
                            }
                        },
                        "runAfter": {
                            "HTTP_-_Get_Parent_Folder_Properties": ["Succeeded"]
                        }
                    },
                    "SET_-_Mark_Metadata_Found": {
                        "type": "SetVariable",
                        "inputs": {
                            "name": "varMetadataFound",
                            "value": true
                        },
                        "runAfter": {
                            "PARSE_-_Parse_Parent_Folder_Response": ["Succeeded"]
                        }
                    }
                },
                "runAfter": {
                    "SET_-_Assign_Parent_Folder_Path": ["Succeeded"]
                }
            },
            "CONDITION_-_Check_If_Parent_Has_Metadata": {
                "type": "If",
                "expression": {
                    "and": [
                        {
                            "equals": ["@variables('varMetadataFound')", true]
                        }
                    ]
                },
                "actions": {
                    "UPDATE_-_Apply_Parent_Metadata_to_New_Item": {
                        "type": "OpenApiConnection",
                        "inputs": {
                            "host": {
                                "connectionName": "shared_sharepointonline",
                                "operationId": "PatchFileItem",
                                "apiId": "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
                            },
                            "parameters": {
                                "dataset": "[YOUR_SITE_URL]",
                                "table": "[YOUR_LIBRARY_GUID]",
                                "id": "@triggerOutputs()?['body/ID']",
                                "item/MetadataField1": "@body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['MetadataField1']",
                                "item/MetadataField2": "@body('PARSE_-_Parse_Parent_Folder_Response')?['d']?['MetadataField2']"
                            }
                        }
                    }
                },
                "else": {
                    "actions": {
                        "COMPOSE_-_Log_No_Metadata_Found": {
                            "type": "Compose",
                            "inputs": "No metadata found on parent folder or parent is library root"
                        }
                    }
                },
                "runAfter": {
                    "SCOPE_-_Get_Parent_Folder_Metadata": ["Succeeded", "Failed", "Skipped"]
                }
            }
        }
    }
}
```

---

## Configuration Checklist

Before deploying, update these placeholders:

- [ ] `[Your SharePoint Site URL]` - e.g., `https://yourtenant.sharepoint.com/sites/yoursite`
- [ ] `[Your Document Library]` - e.g., `Documents` or `Shared Documents`
- [ ] `[YOUR_LIBRARY_GUID]` - The GUID of your document library
- [ ] `YourMetadataField1`, `YourMetadataField2`, etc. - Replace with your actual column internal names
- [ ] Update the Parse JSON schema with your actual metadata fields
- [ ] Update the SP.Data type in HTTP body (replace spaces with `_x0020_`)

---

## Testing Recommendations

1. **Test with a single file** in a subfolder first
2. **Check the flow run history** for any errors in the HTTP request
3. **Verify the parent folder path** is being calculated correctly
4. **Confirm metadata fields** are mapped to the correct internal names

---

## Troubleshooting Common Issues

| Issue | Solution |
|-------|----------|
| HTTP request fails with 404 | Verify the parent folder path is correct and URL-encoded |
| Metadata not updating | Check column internal names (use `_x0020_` for spaces) |
| Flow triggers multiple times | Add a trigger condition to exclude system files |
| Root folder items fail | The scope handles this - items at root won't have parent metadata |

---

## Optional Enhancements

1. **Add trigger condition** to exclude certain file types:
   ```
   @not(endsWith(triggerOutputs()?['body/{FilenameWithExtension}'], '.tmp'))
   ```

2. **Add concurrency control** to prevent race conditions with bulk uploads

3. **Add Teams/Email notification** on failure for monitoring
