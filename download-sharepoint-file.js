#!/usr/bin/env node

const fs = require("node:fs/promises");
const path = require("node:path");

const envPath = path.join(__dirname, ".env");
const workbookPath = path.join(__dirname, "test.xlsx");

async function loadEnvFile(filePath) {
  try {
    const contents = await fs.readFile(filePath, "utf8");
    for (const rawLine of contents.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) {
        continue;
      }

      const separatorIndex = line.indexOf("=");
      if (separatorIndex === -1) {
        continue;
      }

      const key = line.slice(0, separatorIndex).trim();
      let value = line.slice(separatorIndex + 1).trim();

      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }

      if (key && !(key in process.env)) {
        process.env[key] = value;
      }
    }
  } catch (error) {
    if (error.code !== "ENOENT") {
      throw error;
    }
  }
}

function requireEnv(name) {
  const value = process.env[name];
  if (!value || !String(value).trim()) {
    throw new Error(`${name} is required.`);
  }

  return String(value).trim();
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let payload = null;

  if (text) {
    try {
      payload = JSON.parse(text);
    } catch {
      payload = text;
    }
  }

  if (!response.ok) {
    const detail = typeof payload === "string"
      ? payload
      : payload?.error_description || payload?.error?.message || `HTTP ${response.status}`;
    throw new Error(detail);
  }

  return payload;
}

async function fetchBuffer(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    const text = await response.text();
    let detail = text || `HTTP ${response.status}`;

    try {
      const payload = JSON.parse(text);
      detail = payload?.error?.message || payload?.error_description || detail;
    } catch {
      // Keep text fallback when the response body is not JSON.
    }

    throw new Error(detail);
  }

  return Buffer.from(await response.arrayBuffer());
}

async function getAccessToken(tenantId, clientId, clientSecret) {
  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default"
  });

  const payload = await fetchJson(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });

  if (!payload?.access_token) {
    throw new Error("Microsoft token response did not include an access token.");
  }

  return payload.access_token;
}

function normalizeFilePath(input) {
  let filePath = String(input || "").trim();
  if (!filePath) {
    throw new Error("SHAREPOINT_FILE_PATH is required.");
  }

  if (!filePath.startsWith("/")) {
    filePath = `/${filePath}`;
  }

  if (!/\.[A-Za-z0-9]+$/.test(filePath)) {
    filePath = `${filePath}.xlsx`;
  }

  return filePath;
}

function buildCandidatePaths(filePath) {
  const candidates = new Set([filePath]);

  if (filePath.startsWith("/Documents/")) {
    candidates.add(filePath.replace("/Documents/", "/Shared Documents/"));
  } else if (filePath.startsWith("/Shared Documents/")) {
    candidates.add(filePath.replace("/Shared Documents/", "/Documents/"));
  }

  return [...candidates];
}

async function resolveSite(siteUrl, accessToken) {
  const parsed = new URL(siteUrl);
  const relativePath = parsed.pathname.replace(/\/$/, "");
  if (!relativePath) {
    throw new Error("SHAREPOINT_SITE_URL must include a site path like /sites/Public.");
  }

  const url =
    `https://graph.microsoft.com/v1.0/sites/${parsed.hostname}:${encodeURI(relativePath)}`;

  return fetchJson(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
}

async function resolveFile(siteId, accessToken, candidatePath) {
  const encodedPath = encodeURI(candidatePath);
  const metadataUrl =
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}` +
    `/drive/root:${encodedPath}`;

  return fetchJson(metadataUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
}

async function downloadFile(siteId, accessToken, candidatePath) {
  const encodedPath = encodeURI(candidatePath);
  const downloadUrl =
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}` +
    `/drive/root:${encodedPath}:/content`;

  return fetchBuffer(downloadUrl, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
}

async function main() {
  await loadEnvFile(envPath);

  console.error(`Raw SHAREPOINT_SITE_URL: ${process.env.SHAREPOINT_SITE_URL || ""}`);
  console.error(`Raw SHAREPOINT_FILE_PATH: ${process.env.SHAREPOINT_FILE_PATH || ""}`);

  const tenantId = requireEnv("MS_TENANT_ID");
  const clientId = requireEnv("MS_CLIENT_ID");
  const clientSecret = requireEnv("MS_CLIENT_SECRET");
  const siteUrl = requireEnv("SHAREPOINT_SITE_URL");
  const filePath = normalizeFilePath(requireEnv("SHAREPOINT_FILE_PATH"));

  const accessToken = await getAccessToken(tenantId, clientId, clientSecret);
  const site = await resolveSite(siteUrl, accessToken);
  console.error(`Resolved site.id: ${site.id}`);
  console.error(`Resolved site.webUrl: ${site.webUrl || ""}`);
  const candidatePaths = buildCandidatePaths(filePath);

  let resolvedFile = null;
  let resolvedPath = "";
  let lastError = null;

  for (const candidatePath of candidatePaths) {
    try {
      const metadata = await resolveFile(site.id, accessToken, candidatePath);
      const buffer = await downloadFile(site.id, accessToken, candidatePath);

      await fs.writeFile(workbookPath, buffer);
      resolvedFile = metadata;
      resolvedPath = candidatePath;
      break;
    } catch (error) {
      console.error(`Failed candidate path: ${candidatePath}`);
      console.error(`Candidate error: ${error.message}`);
      lastError = error;
    }
  }

  if (!resolvedFile || !resolvedPath) {
    const reason = lastError ? ` ${lastError.message}` : "";
    throw new Error(
      `Could not locate the SharePoint workbook at ${filePath}.${reason}`.trim()
    );
  }

  console.error(`Resolved SharePoint site: ${site.webUrl || siteUrl}`);
  console.error(`Downloaded workbook: ${resolvedFile.name || path.basename(resolvedPath)}`);
  console.error(`Resolved file path: ${resolvedPath}`);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
