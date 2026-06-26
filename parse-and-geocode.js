#!/usr/bin/env node

const { execFileSync } = require("node:child_process");
const fs = require("node:fs/promises");
const path = require("node:path");

const workbookPath = path.join(__dirname, "test.xlsx");
const outputGeoJsonPath = path.join(__dirname, "locations.geojson");
const failedLogPath = path.join(__dirname, "failed.txt");
const envPath = path.join(__dirname, ".env");

async function loadEnvFile(filePath) {
  try {
    const contents = await fs.readFile(filePath, "utf8");
    for (const rawLine of contents.split(/\r?\n/)) {
      const line = rawLine.trim();
      if (!line || line.startsWith("#")) continue;

      const separatorIndex = line.indexOf("=");
      if (separatorIndex === -1) continue;

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
    if (error.code !== "ENOENT") throw error;
  }
}

function readZipEntry(zipPath, entryPath) {
  return execFileSync("unzip", ["-p", zipPath, entryPath], {
    encoding: "utf8"
  });
}

function decodeXml(text) {
  return String(text || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function parseSharedStrings(xml) {
  const sharedStrings = [];
  const matches = xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g);

  for (const match of matches) {
    const textParts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((part) => decodeXml(part[1]));
    sharedStrings.push(textParts.join(""));
  }

  return sharedStrings;
}

function columnLetters(cellRef) {
  const match = String(cellRef || "").match(/[A-Z]+/);
  return match ? match[0] : "";
}

function parseWorksheet(xml, sharedStrings) {
  const rows = [];
  const rowMatches = xml.matchAll(/<row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g);

  for (const rowMatch of rowMatches) {
    const rowNumber = Number(rowMatch[1]);
    const rowXml = rowMatch[2];
    const values = {};
    const cellMatches = rowXml.matchAll(/<c\b([^>]*?)\/>|<c\b([^>]*?)>([\s\S]*?)<\/c>/g);

    for (const cellMatch of cellMatches) {
      const attributes = cellMatch[1] || cellMatch[2] || "";
      const body = cellMatch[3] || "";
      const refMatch = attributes.match(/\br="([^"]+)"/);
      if (!refMatch) continue;

      const ref = refMatch[1];
      const col = columnLetters(ref);

      const typeMatch = attributes.match(/\bt="([^"]+)"/);
      const type = typeMatch ? typeMatch[1] : "";

      const valueMatch = body.match(/<v>([\s\S]*?)<\/v>/);
      const inlineMatch = body.match(/<t\b[^>]*>([\s\S]*?)<\/t>/);

      let value = "";
      if (type === "s" && valueMatch) {
        value = sharedStrings[Number(valueMatch[1])] || "";
      } else if (type === "inlineStr" && inlineMatch) {
        value = decodeXml(inlineMatch[1]);
      } else if (valueMatch) {
        value = decodeXml(valueMatch[1]);
      }

      values[col] = String(value).trim();
    }

    rows.push({ rowNumber, values });
  }

  return rows;
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s+/g, " ");
}

function findHeaderRow(rows, headerAliases) {
  for (const row of rows) {
    const normalizedByColumn = Object.fromEntries(
      Object.entries(row.values).map(([col, value]) => [
        col,
        normalizeHeader(value)
      ])
    );

    for (const [, aliases] of Object.entries(headerAliases)) {
      const foundColumn = Object.entries(normalizedByColumn).find(([, v]) =>
        aliases.includes(v)
      );
      if (!foundColumn) continue;

      const columnMap = {};
      let foundAll = true;

      for (const [, targetAliases] of Object.entries(headerAliases)) {
        const match = Object.entries(normalizedByColumn).find(([, v]) =>
          targetAliases.includes(v)
        );
        if (!match) {
          foundAll = false;
          break;
        }
        columnMap[targetField] = match[0];
      }

      if (foundAll) {
        return { rowNumber: row.rowNumber, columnMap };
      }
    }
  }

  return null;
}

function mapOptionalColumns(normalizedByColumn, optionalHeaderAliases) {
  const columnMap = {};

  for (const [targetField, targetAliases] of Object.entries(
    optionalHeaderAliases
  )) {
    const match = Object.entries(normalizedByColumn).find(([, v]) =>
      targetAliases.includes(v)
    );

    if (match) columnMap[targetField] = match[0];
  }

  return columnMap;
}

function parseAddress(rawAddress) {
  const parts = String(rawAddress || "")
    .split(/\s+-\s+/)
    .map((p) => p.trim())
    .filter(Boolean);

  if (parts.length < 2) return null;

  const [cityState, streetAddress] = parts;
  if (!cityState || !streetAddress) return null;

  return { cityState, streetAddress };
}

function validateParsedAddress(rawAddress, parsedAddress) {
  const cityState = String(parsedAddress.cityState || "").trim();
  const streetAddress = String(parsedAddress.streetAddress || "").trim();

  if (!cityState || !streetAddress) {
    return "missing city/state or street address";
  }

  if ((cityState.match(/,/g) || []).length > 1) {
    return "city/state looks like a region list";
  }

  if (/\bportfolio\b/i.test(rawAddress)) {
    return "portfolio entry is not a single address";
  }

  if (!/\d/.test(streetAddress)) {
    return "missing street number";
  }

  return "";
}

function buildProperties(rowValues, columnMap, parsedAddress) {
  const owner = rowValues[columnMap.owner] || ""; // broker name (KEEP)

  const dealName = rowValues[columnMap.addressDealNumber] || "";
  const recordType = rowValues[columnMap.recordType] || "";
  const askingPrice = rowValues[columnMap.askingPrice] || "";
  const leaseRate = rowValues[columnMap.leaseRate] || "";
  const listingEffectiveDate = rowValues[columnMap.listingEffectiveDate] || "";
  const listingExpirationDate = rowValues[columnMap.listingExpirationDate] || "";
  const squareFootage = rowValues[columnMap.squareFootage] || "";

  const leaseSquareFootage = columnMap.leaseSquareFootage
    ? rowValues[columnMap.leaseSquareFootage] || ""
    : "";

  const acreage = columnMap.acreage
    ? rowValues[columnMap.acreage] || ""
    : "";

  const address = `${parsedAddress.streetAddress}, ${parsedAddress.cityState}`;

  return {
    Owner: owner,
    broker_name: owner,

    Property: dealName,
    "Deal: Deal Name": dealName,
    "Deal: Record Type": recordType,
    "Asking Price": askingPrice,
    "Lease Rate": leaseRate,
    "Listing Effective Date": listingEffectiveDate,
    "Listing Expiration Date": listingExpirationDate,
    "Square Footage": squareFootage,
    Acreage: acreage,
    "Lease Square Footage": leaseSquareFootage,

    deal_name: dealName,
    record_type: recordType,
    asking_price: askingPrice,
    lease_rate: leaseRate,
    listing_effective_date: listingEffectiveDate,
    listing_expiration_date: listingExpirationDate,
    square_footage: squareFootage,
    acreage,

    address,
    city_state: parsedAddress.cityState,
    street_address: parsedAddress.streetAddress
  };
}

async function geocodeAddress(query, mapboxToken) {
  const url =
    `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(
      query
    )}.json` +
    `?access_token=${encodeURIComponent(mapboxToken)}` +
    "&limit=1&autocomplete=false";

  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const payload = await response.json();
  const feature = payload.features?.[0];

  if (!feature?.center) return null;
  return feature.center;
}

async function main() {
  await loadEnvFile(envPath);

  const mapboxToken = process.env.MAPBOX_TOKEN;
  if (!mapboxToken) {
    console.error("Missing MAPBOX_TOKEN");
    process.exit(1);
  }

  const sharedStringsXml = readZipEntry(workbookPath, "xl/sharedStrings.xml");
  const worksheetXml = readZipEntry(workbookPath, "xl/worksheets/sheet1.xml");

  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const rows = parseWorksheet(worksheetXml, sharedStrings);

  const requiredHeaderAliases = {
    owner: ["owner"], // broker retained
    addressDealNumber: ["deal: deal name", "property"],
    recordType: ["deal: record type"],
    askingPrice: ["asking price"],
    leaseRate: ["lease rate"],
    listingEffectiveDate: ["listing effective date"],
    listingExpirationDate: ["listing expiration date"],
    squareFootage: ["square footage"]
  };

  const optionalHeaderAliases = {
    leaseSquareFootage: ["lease square footage"],
    acreage: ["acreage"]
  };

  const header = findHeaderRow(rows, requiredHeaderAliases);
  if (!header) {
    console.error("Header row not found");
    process.exit(1);
  }

  const headerRow = rows.find((r) => r.rowNumber === header.rowNumber);

  const normalizedByColumn = Object.fromEntries(
    Object.entries(headerRow.values).map(([col, v]) => [
      col,
      normalizeHeader(v)
    ])
  );

  header.columnMap = {
    ...header.columnMap,
    ...mapOptionalColumns(normalizedByColumn, optionalHeaderAliases)
  };

  const features = [];
  const failedRows = [];

  for (const row of rows) {
    if (row.rowNumber <= header.rowNumber) continue;

    const rawAddress = row.values[header.columnMap.addressDealNumber] || "";
    if (!rawAddress) continue;

    const parsedAddress = parseAddress(rawAddress);
    if (!parsedAddress) {
      failedRows.push(`Row ${row.rowNumber}: bad address`);
      continue;
    }

    const invalidReason = validateParsedAddress(rawAddress, parsedAddress);
    if (invalidReason) {
      failedRows.push(`Row ${row.rowNumber}: ${invalidReason}`);
      continue;
    }

    const query = `${parsedAddress.streetAddress.split(/\s+(?:&|and)\s+\d/i)[0].trim()}, ${parsedAddress.cityState}`;
    const properties = buildProperties(row.values, header.columnMap, parsedAddress);

    try {
      const coordinates = await geocodeAddress(query, mapboxToken);
      if (!coordinates) {
        failedRows.push(`Row ${row.rowNumber}: no geocode`);
        continue;
      }

      features.push({
        type: "Feature",
        properties,
        geometry: {
          type: "Point",
          coordinates
        }
      });
    } catch (e) {
      failedRows.push(`Row ${row.rowNumber}: ${e.message}`);
    }
  }

  const geojson = { type: "FeatureCollection", features };

  await fs.writeFile(
    outputGeoJsonPath,
    JSON.stringify(geojson, null, 2) + "\n",
    "utf8"
  );

  await fs.writeFile(
    failedLogPath,
    failedRows.length ? failedRows.join("\n") + "\n" : "",
    "utf8"
  );

  console.error(
    `Wrote ${features.length} features. Failed ${failedRows.length} rows.`
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
