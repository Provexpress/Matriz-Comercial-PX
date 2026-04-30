import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as msal from "@azure/msal-node";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const schemaPath = path.join(rootDir, "dataverse", "schema.json");

let schemaCache;
let tokenCache = process.env.DATAVERSE_ACCESS_TOKEN;
const entitySetCache = new Map();

export async function readSchema() {
  if (!schemaCache) {
    schemaCache = JSON.parse(await fs.readFile(schemaPath, "utf8"));
  }
  return schemaCache;
}

async function getAccessToken(schema) {
  if (tokenCache) return tokenCache;

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

  tokenCache = result.accessToken;
  return tokenCache;
}

export async function dataverseRequest(method, relativeUrl, body, extraHeaders = {}) {
  const schema = await readSchema();
  const token = await getAccessToken(schema);
  const url = new URL(relativeUrl, `${schema.environmentUrl.replace(/\/$/, "")}/api/data/v9.2/`);
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json; charset=utf-8",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      ...extraHeaders
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

export async function getTableEntitySet(schemaName) {
  const logicalName = schemaName.toLowerCase();
  if (entitySetCache.has(logicalName)) return entitySetCache.get(logicalName);

  const table = await dataverseRequest(
    "GET",
    `EntityDefinitions(LogicalName='${logicalName}')?$select=EntitySetName,LogicalName`
  );
  entitySetCache.set(logicalName, table.EntitySetName);
  return table.EntitySetName;
}

export function choiceValue(index) {
  return 100000000 + index;
}
