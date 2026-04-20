// ============================================================================
// handlers/properties_extended.ts — Property notes, listings, and directory.
//
// Exports:
//   handlePropertyNotes    — fetch property notes from DB API v0
//   handleListings         — fetch listings (for-rent/for-sale) from DB API v0
//   handlePropertiesDetail — fetch full property details with enrichment
//
// ============================================================================

import { cacheGet, cacheSet, rowsAsObjects, sqlite } from "../db.ts";
import { AF_DB, AF_REPORTS, dbHeaders } from "../config.ts";
import { fetchWithTimeout, snapDays } from "../lib/fetchUtils.ts";

type PropertyNote = {
  id: string;
  body: string;
  entity_id: string;
  last_updated_at: string;
};

type Listing = {
  id: string;
  property_id: string;
  unit_id: string;
  bedrooms: number;
  bathrooms: number;
  listed_rent: string;
  available_on: string;
  address1: string;
  city: string;
  state: string;
  zip: string;
  posted_to_website: boolean;
  [key: string]: any;
};

// Fetch property notes for a specific property
export async function handlePropertyNotes(
  params: Record<string, string>,
): Promise<any> {
  const propertyId = String(params.property_id || "").trim();
  if (!propertyId) {
    return { ok: false, error: "Missing property_id parameter" };
  }

  const cacheKey = `property_notes_${propertyId}`;
  const cached = await cacheGet(cacheKey, "property_notes");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  try {
    // Fetch from DB API v0
    const path = `/api/v0/properties/notes?filters%5BPropertyId%5D=${
      encodeURIComponent(propertyId)
    }&page%5Bsize%5D=100`;

    let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
      headers: dbHeaders(),
    });

    if ([401, 403, 404, 422].includes(resp.status)) {
      resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
        headers: dbHeaders(),
      });
    }

    if (!resp.ok) {
      return {
        ok: false,
        error: `Failed to fetch property notes: HTTP ${resp.status}`,
      };
    }

    const data = await resp.json();
    const notes: PropertyNote[] = (data.data || data.results || []).map(
      (n: any) => ({
        id: String(n.Id || n.id || ""),
        body: String(n.Body || n.body || ""),
        entity_id: String(n.EntityId || n.entity_id || propertyId),
        last_updated_at: String(n.LastUpdatedAt || n.last_updated_at || ""),
      }),
    );

    // Cache the results
    await cacheSet(cacheKey, "property_notes", notes, notes.length);

    // Store in database for cross-tab access
    try {
      const now = Date.now();
      for (const note of notes) {
        await sqlite.execute({
          sql: `INSERT OR REPLACE INTO property_notes
                (id, property_id, body, last_updated_at, cached_at)
                VALUES (?, ?, ?, ?, ?)`,
          args: [note.id, propertyId, note.body, note.last_updated_at, now],
        });
      }
    } catch {
      // Non-fatal database storage failure
    }

    return {
      ok: true,
      results: notes,
      count: notes.length,
      from_cache: false,
    };
  } catch (err: any) {
    return {
      ok: false,
      error: `Property notes fetch error: ${err.message}`,
    };
  }
}

// Fetch listings for properties
export async function handleListings(
  params: Record<string, string>,
): Promise<any> {
  const propertyId = String(params.property_id || "").trim();
  const days = snapDays(
    parseInt(params.days || "365", 10) || 365,
    "listings",
  );
  const cacheKey = propertyId ? `listings_${days}_${propertyId}` : `listings_${days}`;

  const cached = await cacheGet(cacheKey, "listings");
  if (cached) {
    return {
      ok: true,
      results: cached.data,
      count: cached.record_count,
      cached_at: cached.cached_at,
      from_cache: true,
    };
  }

  try {
    const fromDate = new Date();
    fromDate.setDate(fromDate.getDate() - days);

    const propertyFilter = propertyId
      ? `&filters%5BPropertyId%5D=${encodeURIComponent(propertyId)}`
      : "";
    const path = `/api/v0/listings?filters%5BLastUpdatedAtFrom%5D=${
      encodeURIComponent(fromDate.toISOString())
    }${propertyFilter}&page%5Bsize%5D=200`;

    let resp = await fetchWithTimeout(`${AF_DB}${path}`, {
      headers: dbHeaders(),
    });

    if ([401, 403, 404, 422].includes(resp.status)) {
      resp = await fetchWithTimeout(`${AF_REPORTS}${path}`, {
        headers: dbHeaders(),
      });
    }

    if (!resp.ok) {
      return {
        ok: false,
        error: `Failed to fetch listings: HTTP ${resp.status}`,
      };
    }

    const data = await resp.json();
    const allListings: Listing[] = [];

    // Process paginated results
    let listings = data.data || data.results || [];
    if (!Array.isArray(listings)) listings = [];

    for (const listing of listings) {
      allListings.push({
        id: String(listing.Id || listing.id || ""),
        property_id: String(listing.PropertyId || listing.property_id || ""),
        unit_id: String(listing.UnitId || listing.unit_id || ""),
        bedrooms: Number(listing.Bedrooms || listing.bedrooms || 0),
        bathrooms: Number(listing.Bathrooms || listing.bathrooms || 0),
        listed_rent: String(listing.ListedRent || listing.listed_rent || ""),
        available_on: String(
          listing.AvailableOn || listing.available_on || "",
        ),
        address1: String(listing.Address1 || listing.address1 || ""),
        city: String(listing.City || listing.city || ""),
        state: String(listing.State || listing.state || ""),
        zip: String(listing.Zip || listing.zip || ""),
        posted_to_website: Boolean(
          listing.PostedToWebsite || listing.posted_to_website || false,
        ),
        posted_to_internet: Boolean(
          listing.PostedToInternet || listing.posted_to_internet || false,
        ),
        unit_type: String(listing.UnitType || listing.unit_type || ""),
        square_feet: Number(listing.SquareFeet || listing.square_feet || 0),
        advertised_rent: Number(
          listing.AdvertisedRent || listing.advertised_rent || 0,
        ),
        deposit: Number(listing.Deposit || listing.deposit || 0),
        cats_allowed: listing.CatsAllowed || listing.cats_allowed || null,
        dogs_allowed: String(listing.DogPolicy || listing.dogs_allowed || ""),
        email: String(listing.Email || listing.email || ""),
        phone_number: String(
          listing.PhoneNumber || listing.phone_number || "",
        ),
        application_url: String(
          listing.ApplicationURL || listing.application_url || "",
        ),
        marketing_title: String(
          listing.MarketingTitle || listing.marketing_title || "",
        ),
        marketing_description: String(
          listing.MarketingDescription ||
            listing.marketing_description ||
            "",
        ),
        unit_amenities_json: Array.isArray(listing.UnitAmenities ||
          listing.unit_amenities)
          ? JSON.stringify(listing.UnitAmenities || listing.unit_amenities)
          : null,
        unit_photos_json: Array.isArray(listing.UnitPhotos ||
          listing.unit_photos)
          ? JSON.stringify(listing.UnitPhotos || listing.unit_photos)
          : null,
        utilities_included_json: Array.isArray(listing.UtilitiesIncluded ||
          listing.utilities_included)
          ? JSON.stringify(listing.UtilitiesIncluded ||
            listing.utilities_included)
          : null,
        youtube_url: String(listing.YouTubeURL || listing.youtube_url || ""),
        last_updated_at: String(
          listing.LastUpdatedAt || listing.last_updated_at || "",
        ),
      });
    }

    // Handle pagination if needed
    let nextPath = data.next_page_path || null;
    let pageCount = 1;
    while (nextPath && pageCount < 10) {
      const fullUrl = nextPath.startsWith("http")
        ? nextPath
        : `${AF_DB}${nextPath}`;
      const nextResp = await fetchWithTimeout(fullUrl, {
        headers: dbHeaders(),
      });
      if (!nextResp.ok) break;

      const nextData = await nextResp.json();
      const nextListings = nextData.data || nextData.results || [];
      if (!Array.isArray(nextListings) || nextListings.length === 0) break;

      for (const listing of nextListings) {
        allListings.push({
          id: String(listing.Id || listing.id || ""),
          property_id: String(listing.PropertyId || listing.property_id || ""),
          unit_id: String(listing.UnitId || listing.unit_id || ""),
          bedrooms: Number(listing.Bedrooms || listing.bedrooms || 0),
          bathrooms: Number(listing.Bathrooms || listing.bathrooms || 0),
          listed_rent: String(listing.ListedRent || listing.listed_rent || ""),
          available_on: String(
            listing.AvailableOn || listing.available_on || "",
          ),
          address1: String(listing.Address1 || listing.address1 || ""),
          city: String(listing.City || listing.city || ""),
          state: String(listing.State || listing.state || ""),
          zip: String(listing.Zip || listing.zip || ""),
          posted_to_website: Boolean(
            listing.PostedToWebsite || listing.posted_to_website || false,
          ),
          posted_to_internet: Boolean(
            listing.PostedToInternet || listing.posted_to_internet || false,
          ),
          unit_type: String(listing.UnitType || listing.unit_type || ""),
          square_feet: Number(listing.SquareFeet || listing.square_feet || 0),
          advertised_rent: Number(
            listing.AdvertisedRent || listing.advertised_rent || 0,
          ),
          deposit: Number(listing.Deposit || listing.deposit || 0),
          cats_allowed: listing.CatsAllowed || listing.cats_allowed || null,
          dogs_allowed: String(listing.DogPolicy || listing.dogs_allowed || ""),
          email: String(listing.Email || listing.email || ""),
          phone_number: String(
            listing.PhoneNumber || listing.phone_number || "",
          ),
          application_url: String(
            listing.ApplicationURL || listing.application_url || "",
          ),
          marketing_title: String(
            listing.MarketingTitle || listing.marketing_title || "",
          ),
          marketing_description: String(
            listing.MarketingDescription ||
              listing.marketing_description ||
              "",
          ),
          unit_amenities_json: Array.isArray(listing.UnitAmenities ||
            listing.unit_amenities)
            ? JSON.stringify(listing.UnitAmenities || listing.unit_amenities)
            : null,
          unit_photos_json: Array.isArray(listing.UnitPhotos ||
            listing.unit_photos)
            ? JSON.stringify(listing.UnitPhotos || listing.unit_photos)
            : null,
          utilities_included_json: Array.isArray(listing.UtilitiesIncluded ||
            listing.utilities_included)
            ? JSON.stringify(listing.UtilitiesIncluded ||
              listing.utilities_included)
            : null,
          youtube_url: String(listing.YouTubeURL || listing.youtube_url || ""),
          last_updated_at: String(
            listing.LastUpdatedAt || listing.last_updated_at || "",
          ),
        });
      }

      nextPath = nextData.next_page_path || null;
      pageCount++;
    }

    // Cache and store
    await cacheSet(cacheKey, "listings", allListings, allListings.length);

    // Store in database
    try {
      const now = Date.now();
      for (const listing of allListings) {
        await sqlite.execute({
          sql: `INSERT OR REPLACE INTO property_listings
                (id, property_id, unit_id, unit_type, bedrooms, bathrooms,
                 square_feet, advertised_rent, listed_rent, available_on,
                 address1, address2, city, state, zip, posted_to_website,
                 posted_to_internet, accepting_apps, application_url, email,
                 phone_number, cats_allowed, dogs_allowed, deposit,
                 application_fee, marketing_title, marketing_description,
                 unit_amenities_json, unit_photos_json, utilities_included_json,
                 youtube_url, last_updated_at, cached_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          args: [
            listing.id,
            listing.property_id,
            listing.unit_id,
            listing.unit_type,
            listing.bedrooms,
            listing.bathrooms,
            listing.square_feet,
            listing.advertised_rent,
            listing.listed_rent,
            listing.available_on,
            listing.address1,
            String(listing.Address2 || listing.address2 || ""),
            listing.city,
            listing.state,
            listing.zip,
            listing.posted_to_website ? 1 : 0,
            listing.posted_to_internet ? 1 : 0,
            0, // accepting_apps
            listing.application_url,
            listing.email,
            listing.phone_number,
            listing.cats_allowed ? 1 : 0,
            listing.dogs_allowed,
            listing.deposit,
            "", // application_fee
            listing.marketing_title,
            listing.marketing_description,
            listing.unit_amenities_json,
            listing.unit_photos_json,
            listing.utilities_included_json,
            listing.youtube_url,
            listing.last_updated_at,
            now,
          ],
        });
      }
    } catch {
      // Non-fatal database storage failure
    }

    const scopedResults = propertyId
      ? allListings.filter((row) => String(row.property_id || "") === propertyId)
      : allListings;

    return {
      ok: true,
      results: scopedResults,
      count: scopedResults.length,
      from_cache: false,
    };
  } catch (err: any) {
    return {
      ok: false,
      error: `Listings fetch error: ${err.message}`,
    };
  }
}

// Return per-property aggregate counts for directory badges.
export async function handlePropertyStats(
  _params: Record<string, string>,
): Promise<any> {
  try {
    const notesRs = await sqlite.execute({
      sql: "SELECT property_id, COUNT(*) AS c FROM property_notes GROUP BY property_id",
      args: [],
    });
    const listingsRs = await sqlite.execute({
      sql: "SELECT property_id, COUNT(*) AS c FROM property_listings GROUP BY property_id",
      args: [],
    });
    const billsRs = await sqlite.execute({
      sql: "SELECT property_id, COUNT(*) AS c FROM billing_map GROUP BY property_id",
      args: [],
    });

    const notesRows = rowsAsObjects(notesRs);
    const listingsRows = rowsAsObjects(listingsRs);
    const billsRows = rowsAsObjects(billsRs);

    const byProperty: Record<string, { notes: number; listings: number; bills: number }> = {};
    for (const row of notesRows) {
      const pid = String(row.property_id || "").trim();
      if (!pid) continue;
      if (!byProperty[pid]) byProperty[pid] = { notes: 0, listings: 0, bills: 0 };
      byProperty[pid].notes = Number(row.c || 0);
    }
    for (const row of listingsRows) {
      const pid = String(row.property_id || "").trim();
      if (!pid) continue;
      if (!byProperty[pid]) byProperty[pid] = { notes: 0, listings: 0, bills: 0 };
      byProperty[pid].listings = Number(row.c || 0);
    }
    for (const row of billsRows) {
      const pid = String(row.property_id || "").trim();
      if (!pid) continue;
      if (!byProperty[pid]) byProperty[pid] = { notes: 0, listings: 0, bills: 0 };
      byProperty[pid].bills = Number(row.c || 0);
    }

    return { ok: true, by_property: byProperty };
  } catch (err: any) {
    return { ok: false, error: `Property stats error: ${err.message || err}` };
  }
}

// Fetch property details from database (cross-tab access)
export async function getPropertiesFromDB(
  groupUuid?: string,
): Promise<any[]> {
  try {
    let sql =
      `SELECT * FROM property_reference WHERE property_id IS NOT NULL`;
    const args: any[] = [];

    if (groupUuid) {
      sql += ` AND property_group_uuid = ?`;
      args.push(groupUuid);
    }

    sql += ` ORDER BY property_name ASC`;

    const result = await sqlite.execute({ sql, args });
    return rowsAsObjects(result);
  } catch {
    return [];
  }
}

// Get property notes from database
export async function getPropertyNotesFromDB(
  propertyId: string,
): Promise<any[]> {
  try {
    const result = await sqlite.execute({
      sql: `SELECT * FROM property_notes WHERE property_id = ? ORDER BY last_updated_at DESC`,
      args: [propertyId],
    });
    return rowsAsObjects(result);
  } catch {
    return [];
  }
}

// Get listings for a property from database
export async function getPropertyListingsFromDB(
  propertyId: string,
): Promise<any[]> {
  try {
    const result = await sqlite.execute({
      sql: `SELECT * FROM property_listings WHERE property_id = ? ORDER BY available_on DESC`,
      args: [propertyId],
    });
    return rowsAsObjects(result);
  } catch {
    return [];
  }
}
