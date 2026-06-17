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
      if (!refMatch) {
        continue;
      }

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
      Object.entries(row.values).map(([col, value]) => [col, normalizeHeader(value)])
    );

    for (const [fieldName, aliases] of Object.entries(headerAliases)) {
      const foundColumn = Object.entries(normalizedByColumn).find(([, value]) =>
        aliases.includes(value)
      );
      if (!foundColumn) {
        continue;
      }

      const columnMap = {};
      let foundAll = true;

      for (const [targetField, targetAliases] of Object.entries(headerAliases)) {
        const match = Object.entries(normalizedByColumn).find(([, value]) =>
          targetAliases.includes(value)
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

function parseAddress(rawAddress) {
  const parts = String(rawAddress || "")
    .split(/\s+-\s+/)
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length < 2) {
    return null;
  }

  const [cityState, streetAddress] = parts;
  if (!cityState || !streetAddress) {
    return null;
  }

  return { cityState, streetAddress };
}

function validateParsedAddress(rawAddress, parsedAddress) {
  const cityState = String(parsedAddress.cityState || "").trim();
  const streetAddress = String(parsedAddress.streetAddress || "").trim();

  if (!cityState || !streetAddress) {
    return "missing city/state or street address";
  }

  // Regional rollups like "NC, VA, TN" are not geocodable street addresses.
  if ((cityState.match(/,/g) || []).length > 1) {
    return "city/state looks like a region list, not a single place";
  }

  // Portfolio and rollup entries should be skipped instead of guessed.
  if (/\bportfolio\b/i.test(rawAddress) || /\bportfolio\b/i.test(streetAddress)) {
    return "portfolio entry is not a single street address";
  }

  // For this dataset, valid listing addresses should include a street number.
  if (!/\d/.test(streetAddress)) {
    return "street address does not include a street number";
  }

  return "";
}

function buildProperties(rowValues, columnMap, parsedAddress) {
  const owner = rowValues[columnMap.owner] || "";
  const dealName = rowValues[columnMap.addressDealNumber] || "";
  const recordType = rowValues[columnMap.recordType] || "";
  const askingPrice = rowValues[columnMap.askingPrice] || "";
  const leaseRate = rowValues[columnMap.leaseRate] || "";
  const landlordCompany = rowValues[columnMap.landlordCompany] || "";
  const listingEffectiveDate = rowValues[columnMap.listingEffectiveDate] || "";
  const listingExpirationDate = rowValues[columnMap.listingExpirationDate] || "";
  const squareFootage = rowValues[columnMap.squareFootage] || "";
  const leaseSquareFootage = rowValues[columnMap.leaseSquareFootage] || "";
  const address = `${parsedAddress.streetAddress}, ${parsedAddress.cityState}`;

  return {
    Owner: owner,
    "Deal: Deal Name": dealName,
    "Deal: Record Type": recordType,
    "Asking Price": askingPrice,
    "Lease Rate": leaseRate,
    "Landlord Company/Seller Company": landlordCompany,
    "Listing Effective Date": listingEffectiveDate,
    "Listing Expiration Date": listingExpirationDate,
    "Square Footage": squareFootage,
    "Lease Square Footage": leaseSquareFootage,
    broker_name: owner,
    address,
    city_state: parsedAddress.cityState,
    street_address: parsedAddress.streetAddress,
    deal_name: dealName,
    record_type: recordType,
    asking_price: askingPrice,
    lease_rate: leaseRate,
    landlord_company: landlordCompany,
    listing_effective_date: listingEffectiveDate,
    listing_expiration_date: listingExpirationDate,
    square_footage: squareFootage,
    lease_square_footage: leaseSquareFootage
  };
}

async function geocodeAddress(query, mapboxToken) {
  const url =
    `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(query)}.json` +
    `?access_token=${encodeURIComponent(mapboxToken)}` +
    "&limit=1&autocomplete=false";

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Mapbox returned HTTP ${response.status}`);
  }

  const payload = await response.json();
  const feature = payload.features && payload.features[0];
  if (!feature || !Array.isArray(feature.center) || feature.center.length < 2) {
    return null;
  }

  return feature.center;
}

async function main() {
  await loadEnvFile(envPath);

  const mapboxToken = process.env.MAPBOX_TOKEN;
  if (!mapboxToken) {
    console.error("MAPBOX_TOKEN is required. Add it to .env or your shell environment.");
    process.exitCode = 1;
    return;
  }

  const sharedStringsXml = readZipEntry(workbookPath, "xl/sharedStrings.xml");
  const worksheetXml = readZipEntry(workbookPath, "xl/worksheets/sheet1.xml");

  const sharedStrings = parseSharedStrings(sharedStringsXml);
  const rows = parseWorksheet(worksheetXml, sharedStrings);

  const headerAliases = {
    owner: ["owner"],
    addressDealNumber: ["deal: deal name"],
    recordType: ["deal: record type"],
    askingPrice: ["asking price"],
    leaseRate: ["lease rate"],
    landlordCompany: ["landlord company/seller company"],
    listingEffectiveDate: ["listing effective date"],
    listingExpirationDate: ["listing expiration date"],
    squareFootage: ["square footage"],
    leaseSquareFootage: ["lease square footage"]
  };

  const header = findHeaderRow(rows, headerAliases);
  if (!header) {
    console.error("Could not find the expected header row in test.xlsx.");
    process.exitCode = 1;
    return;
  }

  const features = [];
  const failedRows = [];

  for (const row of rows) {
    if (row.rowNumber <= header.rowNumber) {
      continue;
    }

    const rawAddress = row.values[header.columnMap.addressDealNumber] || "";
    if (!rawAddress) {
      continue;
    }

    const parsedAddress = parseAddress(rawAddress);
    if (!parsedAddress) {
      failedRows.push(`Row ${row.rowNumber}: could not parse address "${rawAddress}"`);
      continue;
    }

    const invalidReason = validateParsedAddress(rawAddress, parsedAddress);
    if (invalidReason) {
      failedRows.push(`Row ${row.rowNumber}: skipped "${rawAddress}" (${invalidReason})`);
      continue;
    }

    const query = `${parsedAddress.streetAddress}, ${parsedAddress.cityState}`;
    const properties = buildProperties(row.values, header.columnMap, parsedAddress);

    try {
      const coordinates = await geocodeAddress(query, mapboxToken);
      if (!coordinates) {
        failedRows.push(`Row ${row.rowNumber}: no geocoding result for "${query}"`);
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
    } catch (error) {
      failedRows.push(`Row ${row.rowNumber}: geocoding failed for "${query}" (${error.message})`);
    }
  }

  const geojson = {
    type: "FeatureCollection",
    features
  };

  await fs.writeFile(outputGeoJsonPath, `${JSON.stringify(geojson, null, 2)}\n`, "utf8");
  await fs.writeFile(
    failedLogPath,
    failedRows.length ? `${failedRows.join("\n")}\n` : "",
    "utf8"
  );

  console.error(`Wrote ${features.length} geocoded features to ${path.basename(outputGeoJsonPath)}.`);
  console.error(`Logged ${failedRows.length} failed rows to ${path.basename(failedLogPath)}.`);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
