import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as msal from "@azure/msal-node";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const schemaPath = path.join(rootDir, "dataverse", "schema.json");
const dryRun = process.argv.includes("--dry-run");
let token = process.env.DATAVERSE_ACCESS_TOKEN;

const lcid = 1033;

function label(text) {
  return {
    LocalizedLabels: [
      {
        Label: text,
        LanguageCode: lcid
      }
    ]
  };
}

function requiredLevel(value = "None") {
  return {
    Value: value,
    CanBeChanged: true,
    ManagedPropertyLogicalName: "canmodifyrequirementlevelsettings"
  };
}

function buildColumn(column) {
  const base = {
    SchemaName: column.schemaName,
    DisplayName: label(column.displayName),
    RequiredLevel: requiredLevel(),
    Description: label(column.displayName)
  };

  if (column.type === "text") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
      ...base,
      MaxLength: column.maxLength ?? 200,
      FormatName: { Value: "Text" }
    };
  }

  if (column.type === "date") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
      ...base,
      Format: "DateOnly",
      DateTimeBehavior: { Value: "DateOnly" }
    };
  }

  if (column.type === "currency") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.MoneyAttributeMetadata",
      ...base,
      MinValue: 0,
      MaxValue: 100000000000,
      Precision: 2,
      PrecisionSource: 1
    };
  }

  if (column.type === "decimal") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
      ...base,
      MinValue: -100000000000,
      MaxValue: 100000000000,
      Precision: 4
    };
  }

  if (column.type === "integer") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
      ...base,
      MinValue: 0,
      MaxValue: 2147483647,
      Format: "None"
    };
  }

  if (column.type === "choice") {
    return {
      "@odata.type": "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
      ...base,
      OptionSet: {
        "@odata.type": "Microsoft.Dynamics.CRM.OptionSetMetadata",
        IsGlobal: false,
        OptionSetType: "Picklist",
        Options: column.options.map((option, index) => ({
          Value: 100000000 + index,
          Label: label(option)
        }))
      }
    };
  }

  throw new Error(`Unsupported column type: ${column.type}`);
}

async function dataverseRequest(schema, method, relativeUrl, body) {
  const url = new URL(relativeUrl, `${schema.environmentUrl.replace(/\/$/, "")}/api/data/v9.2/`);
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "MSCRM.SolutionUniqueName": schema.solution.uniqueName
    },
    body: body ? JSON.stringify(body) : undefined
  });

  if (response.status === 204) return null;

  const text = await response.text();
  const data = text ? JSON.parse(text) : null;

  if (!response.ok) {
    const message = data?.error?.message ?? response.statusText;
    throw new Error(`${method} ${url} failed: ${message}`);
  }

  return data;
}

async function getAccessToken(schema) {
  if (token) return token;

  if (!schema.auth?.tenantId || !schema.auth?.clientId) {
    throw new Error("Missing auth.tenantId or auth.clientId in dataverse/schema.json.");
  }

  const app = new msal.PublicClientApplication({
    auth: {
      clientId: schema.auth.clientId,
      authority: `https://login.microsoftonline.com/${schema.auth.tenantId}`
    }
  });

  const result = await app.acquireTokenByDeviceCode({
    scopes: [`${schema.environmentUrl.replace(/\/$/, "")}/user_impersonation`],
    deviceCodeCallback: (response) => {
      console.log(response.message);
    }
  });

  token = result.accessToken;
  return token;
}

async function tableExists(schema, logicalName) {
  try {
    await dataverseRequest(schema, "GET", `EntityDefinitions(LogicalName='${logicalName}')?$select=LogicalName`);
    return true;
  } catch (error) {
    if (error.message.includes("Does Not Exist") || error.message.includes("not exist")) return false;
    throw error;
  }
}

async function columnExists(schema, tableLogicalName, columnLogicalName) {
  try {
    await dataverseRequest(
      schema,
      "GET",
      `EntityDefinitions(LogicalName='${tableLogicalName}')/Attributes(LogicalName='${columnLogicalName}')?$select=LogicalName`
    );
    return true;
  } catch (error) {
    if (error.message.includes("Does Not Exist") || error.message.includes("not exist")) return false;
    throw error;
  }
}

async function ensureTable(schema, table) {
  const tableLogicalName = table.schemaName.toLowerCase();
  const exists = await tableExists(schema, tableLogicalName);

  if (exists) {
    console.log(`OK table exists: ${tableLogicalName}`);
    return;
  }

  const primary = table.primaryColumn;
  const body = {
    "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
    SchemaName: table.schemaName,
    DisplayName: label(table.displayName),
    DisplayCollectionName: label(`${table.displayName}s`),
    Description: label(table.displayName),
    OwnershipType: "UserOwned",
    IsActivity: false,
    HasActivities: false,
    HasNotes: true,
    Attributes: [
      {
        "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
        SchemaName: primary.schemaName,
        DisplayName: label(primary.displayName),
        RequiredLevel: requiredLevel("ApplicationRequired"),
        MaxLength: 100,
        FormatName: { Value: "Text" },
        Description: label(primary.displayName),
        IsPrimaryName: true
      }
    ]
  };

  console.log(`CREATE table: ${tableLogicalName}`);
  if (!dryRun) await dataverseRequest(schema, "POST", "EntityDefinitions", body);
}

async function ensureColumns(schema, table) {
  const tableLogicalName = table.schemaName.toLowerCase();

  for (const column of table.columns) {
    const columnLogicalName = column.schemaName.toLowerCase();
    const exists = await columnExists(schema, tableLogicalName, columnLogicalName);

    if (exists) {
      console.log(`OK column exists: ${tableLogicalName}.${columnLogicalName}`);
      continue;
    }

    console.log(`CREATE column: ${tableLogicalName}.${columnLogicalName}`);
    if (!dryRun) {
      await dataverseRequest(
        schema,
        "POST",
        `EntityDefinitions(LogicalName='${tableLogicalName}')/Attributes`,
        buildColumn(column)
      );
    }
  }
}

async function publish(schema) {
  console.log("PUBLISH customizations");
  if (!dryRun) await dataverseRequest(schema, "POST", "PublishAllXml", {});
}

async function main() {
  const schema = JSON.parse(await fs.readFile(schemaPath, "utf8"));

  if (!dryRun && !token) {
    await getAccessToken(schema);
  }

  for (const table of schema.tables) {
    if (dryRun) {
      console.log(`PLAN table: ${table.schemaName.toLowerCase()}`);
      for (const column of table.columns) {
        console.log(`PLAN column: ${column.schemaName.toLowerCase()} (${column.type})`);
      }
      continue;
    }

    await ensureTable(schema, table);
    await ensureColumns(schema, table);
  }

  if (!dryRun) await publish(schema);
}

main().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
