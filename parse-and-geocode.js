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

function mapOptionalColumns(normalizedByColumn, optionalHeaderAliases) {
  const columnMap = {};

  for (const [targetField, targetAliases] of Object.entries(optionalHeaderAliases)) {
    const match = Object.entries(normalizedByColumn).find(([, value]) =>
      targetAliases.includes(value)
    );

    if (match) {
      columnMap[targetField] = match[0];
    }
  }

  return columnMap;
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

const STREET_SUFFIX_PATTERN =
  /\b(st|street|rd|road|ave|avenue|blvd|boulevard|dr|drive|ln|lane|ct|court|cir|circle|pl|place|way|pkwy|parkway|hwy|highway|trl|trail|ter|terrace)\b/i;

function getPrimaryGeocodeStreetAddress(streetAddress) {
  const normalizedStreetAddress = String(streetAddress || "").trim();
  const multiAddressMatch = normalizedStreetAddress.match(/^(.*?)\s+(?:&|and)\s+(\d[\s\S]*)$/i);

  if (!multiAddressMatch) {
    return normalizedStreetAddress;
  }

  const primaryStreetAddress = multiAddressMatch[1].trim();
  if (!STREET_SUFFIX_PATTERN.test(primaryStreetAddress)) {
    return normalizedStreetAddress;
  }

  return primaryStreetAddress;
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

  // Parcel/block descriptions are too ambiguous to geocode reliably unless
  // they also include a conventional street suffix.
  if (/\bblock\b/i.test(streetAddress)) {
    const hasStreetSuffix = STREET_SUFFIX_PATTERN.test(streetAddress);
    if (!hasStreetSuffix) {
      return "parcel/block entry is not a normal street address";
    }
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
  const listingEffectiveDate = rowValues[columnMap.listingEffectiveDate] || "";
  const listingExpirationDate = rowValues[columnMap.listingExpirationDate] || "";
  const squareFootage = rowValues[columnMap.squareFootage] || "";
  const leaseSquareFootage = columnMap.leaseSquareFootage
    ? (rowValues[columnMap.leaseSquareFootage] || "")
    : "";
  const acreage = columnMap.acreage ? (rowValues[columnMap.acreage] || "") : "";
  const address = `${parsedAddress.streetAddress}, ${parsedAddress.cityState}`;

  return {
    Owner: owner,
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
    broker_name: owner,
    address,
    city_state: parsedAddress.cityState,
    street_address: parsedAddress.streetAddress,
    deal_name: dealName,
    record_type: recordType,
    asking_price: askingPrice,
    lease_rate: leaseRate,
    listing_effective_date: listingEffectiveDate,
    listing_expiration_date: listingExpirationDate,
    square_footage: squareFootage,
    acreage,
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

  const requiredHeaderAliases = {
    owner: ["owner"],
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
    console.error("Could not find the expected header row in test.xlsx.");
    process.exitCode = 1;
    return;
  }

  const headerRow = rows.find((row) => row.rowNumber === header.rowNumber);
  const normalizedByColumn = Object.fromEntries(
    Object.entries(headerRow?.values || {}).map(([col, value]) => [col, normalizeHeader(value)])
  );
  header.columnMap = {
    ...header.columnMap,
    ...mapOptionalColumns(normalizedByColumn, optionalHeaderAliases)
  };

  const features = [];
  const failedRows = [];
  const geocodeFailureCounts = new Map();
  let geocodeAttempts = 0;
  let geocodeFailures = 0;
  let skippedRows = 0;

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

    const geocodeStreetAddress = getPrimaryGeocodeStreetAddress(parsedAddress.streetAddress);
    const query = `${geocodeStreetAddress}, ${parsedAddress.cityState}`;
    const properties = buildProperties(row.values, header.columnMap, parsedAddress);
    geocodeAttempts += 1;

    try {
      const coordinates = await geocodeAddress(query, mapboxToken);
      if (!coordinates) {
        geocodeFailures += 1;
        geocodeFailureCounts.set(
          "no geocoding result",
          (geocodeFailureCounts.get("no geocoding result") || 0) + 1
        );
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
      geocodeFailures += 1;
      const failureReason = error && error.message ? error.message : "unknown error";
      geocodeFailureCounts.set(
        failureReason,
        (geocodeFailureCounts.get(failureReason) || 0) + 1
      );
      failedRows.push(`Row ${row.rowNumber}: geocoding failed for "${query}" (${error.message})`);
    }
  }

  skippedRows = failedRows.length - geocodeFailures;

  const systemicFailureReasons = [...geocodeFailureCounts.entries()]
    .sort((left, right) => right[1] - left[1])
    .map(([reason, count]) => `${reason}: ${count}`)
    .join("; ");
  const systemicGeocodeFailure =
    geocodeAttempts > 0 &&
    geocodeFailures === geocodeAttempts &&
    features.length === 0;

  if (systemicGeocodeFailure) {
    const summaryLines = [
      `Aborted update to avoid overwriting ${path.basename(outputGeoJsonPath)} with zero features.`,
      `Geocode attempts: ${geocodeAttempts}`,
      `Geocode failures: ${geocodeFailures}`,
      `Skipped rows before geocoding: ${Math.max(skippedRows, 0)}`,
      `Failure summary: ${systemicFailureReasons || "unknown"}`
    ];

    await fs.writeFile(failedLogPath, `${summaryLines.join("\n")}\n\n${failedRows.join("\n")}\n`, "utf8");
    console.error(summaryLines.join("\n"));
    process.exitCode = 1;
    return;
  }

  const geojson = {
    type: "FeatureCollection",
    features
  };

  await fs.writeFile(outputGeoJsonPath, `${JSON.stringify(geojson, null, 2)}\n`, "utf8");
  await fs.writeFile(
    failedLogPath,
    [
      `Geocode attempts: ${geocodeAttempts}`,
      `Features written: ${features.length}`,
      `Geocode failures: ${geocodeFailures}`,
      `Skipped rows before geocoding: ${Math.max(skippedRows, 0)}`,
      systemicFailureReasons ? `Failure summary: ${systemicFailureReasons}` : ""
    ].filter(Boolean).join("\n") + (failedRows.length ? `\n\n${failedRows.join("\n")}\n` : "\n"),
    "utf8"
  );

  console.error(`Wrote ${features.length} geocoded features to ${path.basename(outputGeoJsonPath)}.`);
  console.error(`Logged ${failedRows.length} failed rows to ${path.basename(failedLogPath)}.`);
}

main().catch((error) => {
  console.error(error.message);
  process.exitCode = 1;
});
