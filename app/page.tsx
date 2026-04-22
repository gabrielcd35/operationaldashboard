'use client';

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

// --- Types ---

type DashboardRow = Record<string, string | number | null | undefined>;
type PartsRow   = Record<string, string | number | null | undefined>;

type AlertCard = {
  id: string;
  title: string;
  count: number;
  rows: DashboardRow[];
  description: string;
  info: string;
  section: 'Operations' | 'Parts' | 'Conventional' | 'Inventory';
  detailType?: 'default' | 'sa-monthly-qc' | 'must-return' | 'missing-install' | 'we-have-parts';
};

type MustReturnGroup = {
  jobNumber: string;
  parts: string[];
  maxDays: number;
  statusPriority: string;
};

type MainCard = {
  id: string;
  title: string;
  count: number;
  rows: DashboardRow[];
  modalType: 'sa-chart' | 'job-list' | 'delivered-hail' | 'repair-approved-buckets';
  isDelayed?: boolean;
};

type SaMonthlyQcItem = {
  row: DashboardRow;
  sa: string;
  issues: string[];
};

type SaMonthlyQcBucket = {
  sa: string;
  items: SaMonthlyQcItem[];
};

// --- Helpers ---

function toText(value: unknown): string {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function formatDate(value: unknown): string {
  const s = toText(value);
  if (!s) return '';
  const d = new Date(s);
  if (Number.isNaN(d.getTime())) return s;
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const yyyy = d.getUTCFullYear();
  return `${mm}/${dd}/${yyyy}`;
}
function normalize(value: unknown): string {
  return toText(value).toLowerCase();
}

function toNumber(value: unknown): number {
  const num = Number(value);
  return Number.isNaN(num) ? 0 : num;
}

function includesAny(text: unknown, terms: string[]): boolean {
  const value = normalize(text);
  return terms.some((term) => value.includes(term.toLowerCase()));
}

function isBlank(value: unknown): boolean {
  return toText(value) === '';
}

function isPastDue(value: unknown): boolean {
  const raw = toText(value);
  if (!raw) return false;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const compare = new Date(date);
  compare.setHours(0, 0, 0, 0);
  return compare < today;
}

function sortByPriority(rows: DashboardRow[]): DashboardRow[] {
  return [...rows].sort((a, b) => {
    const pa = toNumber(a['Priority']);
    const pb = toNumber(b['Priority']);
    return pb - pa;
  });
}

function getColumnValue(row: DashboardRow, keys: string[]): unknown {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return row[key];
    }
  }

  const lowerMap = new Map<string, unknown>();
  for (const [key, value] of Object.entries(row)) {
    lowerMap.set(key.toLowerCase(), value);
  }

  for (const key of keys) {
    if (lowerMap.has(key.toLowerCase())) {
      return lowerMap.get(key.toLowerCase());
    }
  }

  return undefined;
}

function getSaName(row: DashboardRow): string {
  return toText(getColumnValue(row, ['SA'])) || 'Unassigned';
}

function getDateStartValue(row: DashboardRow): unknown {
  return getColumnValue(row, ['date_start', 'Date Start', 'date start']);
}

function getRepairApprovedDateValue(row: DashboardRow): unknown {
  return getColumnValue(row, ['Repair Approved', 'repair approved', 'Repair Approved Date']);
}

function getDateEndValue(row: DashboardRow): unknown {
  return getColumnValue(row, ['date_end', 'Date End', 'date end']);
}

function getQcNotCompletedValue(row: DashboardRow): unknown {
  return getColumnValue(row, ['QC Not Completed', 'Qc Not Completed', 'qc not completed']);
}

function parseExcelSerialDate(value: number): Date | null {
  if (!Number.isFinite(value) || value <= 59) return null;
  const parsed = new Date(Math.round((value - 25569) * 86400 * 1000));
  if (Number.isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function parseDateValue(value: unknown): Date | null {
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const copy = new Date(value);
    copy.setHours(0, 0, 0, 0);
    return copy;
  }

  if (typeof value === 'number') {
    return parseExcelSerialDate(value);
  }

  const raw = toText(value);
  if (!raw) return null;

  if (/^\d+(\.\d+)?$/.test(raw)) {
    const serialDate = parseExcelSerialDate(Number(raw));
    if (serialDate) return serialDate;
  }

  // Handle "YYYY-MM-DD HH:MM:SS" (space separator) explicitly before generic parse
  const spaceMatch = raw.match(/^(\d{4}-\d{2}-\d{2})\s\d{2}:\d{2}:\d{2}/);
  if (spaceMatch) {
    const parsed = new Date(`${spaceMatch[1]}T00:00:00`);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  const direct = new Date(raw);
  if (!Number.isNaN(direct.getTime())) {
    direct.setHours(0, 0, 0, 0);
    return direct;
  }

  const slashMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slashMatch) {
    const [, month, day, year] = slashMatch;
    const normalizedYear = year.length === 2 ? `20${year}` : year;
    const parsed = new Date(`${normalizedYear}-${month.padStart(2, '0')}-${day.padStart(2, '0')}T00:00:00`);
    if (!Number.isNaN(parsed.getTime())) {
      parsed.setHours(0, 0, 0, 0);
      return parsed;
    }
  }

  return null;
}

function getSaMonthlyQcIssues(row: DashboardRow): string[] {
  const issues: string[] = [];

  const dateStartRaw = getDateStartValue(row);
  const repairApprovedRaw = getRepairApprovedDateValue(row);
  const dateEndRaw = getDateEndValue(row);
  const qcFlagRaw = getQcNotCompletedValue(row);

  const dateStartText = toText(dateStartRaw);
  const repairApprovedText = toText(repairApprovedRaw);
  const dateEndText = toText(dateEndRaw);

  const dateStart = parseDateValue(dateStartRaw);
  const repairApprovedDate = parseDateValue(repairApprovedRaw);
  const dateEnd = parseDateValue(dateEndRaw);

  if (!dateStartText) {
    issues.push('Missing START DATE');
  }

  if (!repairApprovedText) {
    issues.push('Missing REPAIR APPROVED DATE');
  }

  if (!dateEndText) {
    issues.push('Missing END DATE');
  }

  if (dateStart && repairApprovedDate && repairApprovedDate.getTime() < dateStart.getTime()) {
    issues.push('Repair Approved inconsistent with Start Date');
  }

  if (dateStart && dateEnd && dateEnd.getTime() < dateStart.getTime()) {
    issues.push('End Date inconsistent with Start Date');
  }

  if (repairApprovedDate && dateEnd && dateEnd.getTime() < repairApprovedDate.getTime()) {
    issues.push('End Date inconsistent with Repair Approved');
  }

  const qcFlagNormalized = normalize(qcFlagRaw);

  if (qcFlagNormalized === 'flag') {
    issues.push('QC Flag is missing');
  }

  return issues;
}

function buildSaMonthlyQcBuckets(rows: DashboardRow[]): SaMonthlyQcBucket[] {
  const buckets = new Map<string, SaMonthlyQcItem[]>();

  for (const row of rows.filter(isVehicleDeliveredHail)) {
    const issues = getSaMonthlyQcIssues(row);
    if (issues.length === 0) continue;

    const sa = getSaName(row);
    const items = buckets.get(sa) ?? [];
    items.push({ row, sa, issues });
    buckets.set(sa, items);
  }

  return Array.from(buckets.entries())
    .map(([sa, items]) => ({
      sa,
      items: [...items].sort((a, b) => {
        const priorityDiff = toNumber(b.row['Priority']) - toNumber(a.row['Priority']);
        if (priorityDiff !== 0) return priorityDiff;
        return toText(a.row['Job Number']).localeCompare(toText(b.row['Job Number']), undefined, {
          numeric: true,
          sensitivity: 'base',
        });
      }),
    }))
    .sort((a, b) => a.sa.localeCompare(b.sa));
}

function buildSaMonthlyQcClipboardText(buckets: SaMonthlyQcBucket[]): string {
  return buckets
    .map((bucket) => {
      const lines = bucket.items.map(
        (item) => `${toText(item.row['Job Number'])} - *${item.issues.join(', ')}*`
      );
      return [bucket.sa, '', ...lines].join('\n');
    })
    .join('\n\n');
}

// --- Status Checks ---

function isPostRepair(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']) === 'post repair';
}

function isRepairApproved(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']).includes('repair approved');
}

function isInsuranceApproval(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']).includes('insurance approval');
}

function isPdrInProgress(row: DashboardRow): boolean {
  const status = normalize(row['Status + Priority']);
  return status === 'pdr in-progress' || status === 'e - ehi repair';
}

function isConventionalHail(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']) === 'conventional (hail)';
}

function isReadyToDeliver(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']).includes('ready to deliver');
}

function isVehicleOnSite(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']) === 'vehicle on-site';
}

function isVehicleDeliveredHail(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']) === 'vehicle delivered (hail)';
}

// --- Delay Logic ---

function isRowDelayed(row: DashboardRow): boolean {
  const days = toNumber(row['Status Days']);
  if (isVehicleOnSite(row)) return days >= 2;
  if (isInsuranceApproval(row)) return days >= 6;
  if (isPdrInProgress(row)) return days >= 3;
  if (isConventionalHail(row)) return days >= 5;
  if (isPostRepair(row)) return days >= 3;
  if (isReadyToDeliver(row)) return days > 2;
  return false;
}

// --- Parts Helpers ---

function getPartJobNumber(part: PartsRow): string {
  return toText(getColumnValue(part, ['Job', 'job', 'Job Number', 'job number']));
}

function getPartName(part: PartsRow): string {
  return toText(getColumnValue(part, ['Part', 'part', 'Part Name', 'part name']));
}

function getOrderedAt(part: PartsRow): Date | null {
  return parseDateValue(getColumnValue(part, ['Ordered At', 'ordered at', 'ordered_at']));
}

function getReceivedAt(part: PartsRow): Date | null {
  return parseDateValue(getColumnValue(part, ['Received At', 'received at', 'received_at']));
}

function getCheckedOutAt(part: PartsRow): Date | null {
  return parseDateValue(getColumnValue(part, ['Checked Out At', 'checked out at', 'checked_out_at']));
}

function getReturnedAt(part: PartsRow): Date | null {
  return parseDateValue(getColumnValue(part, ['Returned At', 'returned at', 'returned_at']));
}

function getPartModel(part: PartsRow): string {
  return toText(getColumnValue(part, ['Model', 'model']));
}

function getPartMake(part: PartsRow): string {
  return toText(getColumnValue(part, ['Make', 'make']));
}

function getPartYear(part: PartsRow): string {
  return toText(getColumnValue(part, ['Year', 'year']));
}

function getSetKey(part: PartsRow): string {
  return toText(getColumnValue(part, ['Set Key', 'set key', 'SetKey', 'set_key']));
}

function getPartEta(part: PartsRow): string {
  return toText(getColumnValue(part, ['ETA', 'eta', 'Eta']));
}

function daysSince(date: Date): number {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.floor((today.getTime() - date.getTime()) / (1000 * 60 * 60 * 24));
}

// Job numbers used for stock inventory rather than a real vehicle.
// Seen in the PARTS sheet as "000", "#000", or sometimes stored as plain 0.
function isStockJob(rawJobNumber: unknown): boolean {
  const jn = toText(rawJobNumber).replace(/^#/, '').toLowerCase();
  return jn === '000' || jn === '0';
}

function getPartsInfoForJob(
  jobNumber: string,
  partsData: PartsRow[]
): { arrived: string[]; missing: { name: string; eta: string }[] } {
  const jobParts = partsData.filter(
    (p) => normalize(getPartJobNumber(p)) === normalize(jobNumber)
  );
  const receivedAtKeys = ['Received At', 'received at', 'received_at'];
  const orderedAtKeys  = ['Ordered At', 'ordered at', 'ordered_at'];
  const arrived = jobParts
    .filter((p) => !isBlank(getColumnValue(p, receivedAtKeys)))
    .map((p) => getPartName(p))
    .filter(Boolean);
  const missing = jobParts
    .filter((p) =>
      !isBlank(getColumnValue(p, orderedAtKeys)) &&
      isBlank(getColumnValue(p, receivedAtKeys))
    )
    .map((p) => ({ name: getPartName(p), eta: getPartEta(p) }))
    .filter((m) => m.name);
  return { arrived, missing };
}

function getGlassPartsForJob(
  jobNumber: string,
  partsData: PartsRow[]
): { name: string; arrived: boolean }[] {
  const glassTerms = ['quarter glass', 'windshield'];
  const jobParts = partsData.filter(
    (p) => normalize(getPartJobNumber(p)) === normalize(jobNumber)
  );
  return jobParts
    .filter((p) => glassTerms.some((term) => normalize(getPartName(p)).includes(term)))
    .map((p) => ({
      name: getPartName(p),
      arrived: !isBlank(getColumnValue(p, ['Received At', 'received at', 'received_at'])),
    }));
}

type MissingInstallGroup = {
  jobNumber: string;
  statusPriority: string;
  items: { name: string; received: boolean }[];
};

function buildMissingInstallGroups(
  partsData: PartsRow[],
  rows: DashboardRow[]
): MissingInstallGroup[] {
  const qualifyingJobs = new Map<string, string>();
  for (const r of rows) {
    if (isPostRepair(r) || isReadyToDeliver(r) || isVehicleDeliveredHail(r)) {
      const jn = normalize(toText(r['Job Number']));
      if (jn && !isStockJob(jn)) qualifyingJobs.set(jn, toText(r['Status + Priority']));
    }
  }

  const orderedAtKeys   = ['Ordered At', 'ordered at', 'ordered_at'];
  const receivedAtKeys  = ['Received At', 'received at', 'received_at'];
  const checkedOutKeys  = ['Checked Out At', 'checked out at', 'checked_out_at'];

  const flaggedParts = partsData.filter((p) => {
    const jn = normalize(getPartJobNumber(p));
    if (!jn || isStockJob(getPartJobNumber(p))) return false;
    if (!qualifyingJobs.has(jn)) return false;
    // Part is "in the pipeline" if either Ordered At OR Received At is set.
    const hasOrdered  = !isBlank(getColumnValue(p, orderedAtKeys));
    const hasReceived = !isBlank(getColumnValue(p, receivedAtKeys));
    if (!hasOrdered && !hasReceived) return false;
    // Something is missing: Received At still blank OR Checked Out At still blank.
    return (
      !hasReceived ||
      isBlank(getColumnValue(p, checkedOutKeys))
    );
  });

  const groups = new Map<string, MissingInstallGroup>();
  for (const part of flaggedParts) {
    const jn = getPartJobNumber(part);
    const jnKey = normalize(jn);
    const received = !isBlank(getColumnValue(part, receivedAtKeys));
    const existing = groups.get(jnKey);
    const item = { name: getPartName(part), received };
    if (existing) {
      existing.items.push(item);
    } else {
      groups.set(jnKey, {
        jobNumber: jn,
        statusPriority: qualifyingJobs.get(jnKey) ?? '',
        items: [item],
      });
    }
  }

  return Array.from(groups.values()).sort((a, b) => b.items.length - a.items.length);
}

// Allow-list of Set Key UUIDs for parts currently sitting in stock.
// Update this list as inventory changes.
const STOCK_SET_KEYS = new Set<string>([
  '4d94cc5e-af46-44bb-b6c6-f80c17c65777',
  '00dfaae0-714f-4408-856c-375c1a782ea1',
  '8b751350-ede0-405c-b871-5994d8e3a7db',
  '75b7927a-27d1-4e02-972a-f33b98fef4a7',
  '693e8794-0323-431f-8926-37958aac0adf',
  '0d7b422a-6927-4a51-9c82-6f8963bc6969',
  'bb394ea7-9ac4-4f58-8904-e0ec9beeeb32',
  'f9675657-8530-430a-b2b0-3b541dfd8472',
  'f7f4ecf2-0a57-4b67-8858-e936c1a23e9d',
  '997943d5-b7db-499c-989b-a5463cb91a89',
  'c9101d06-1ef1-428f-9b4c-50546567fd60',
  '725d20ce-e247-4486-a85f-afd23e35fa69',
  '30e16e81-1c67-4966-a6bc-6214142b0cf8',
  'fd072635-1c03-416e-b512-a0a4ff48a1c6',
  '8424d2b5-5ef0-48be-9c15-801aa1336c28',
  '43490abc-1a91-4291-9eeb-a6a24c91790b',
  '997d3df5-6538-46f4-b930-768cd38956a2',
  '8b8e850a-cf15-48c4-8fc2-2931ea92124b',
  'ba3adc5b-79e2-4e1a-bd1e-454126f05d30',
  'b3379819-4b66-4050-8661-6905fb3297f3',
  'eaaf5b0d-f430-4fd6-ae17-68c688ff9dbd',
  'ee8418d0-414c-432c-a53d-8b06f8a81833',
  'd7bb5ee3-df60-4130-8914-20088b49aa61',
  'be5560e7-32bd-49aa-8600-13e79987ec4e',
  '2fefa46d-7d73-4d58-b2b0-df64317bdab1',
  '0bf6b628-9e8c-4bca-a9c0-db65f13ecde1',
  '4e00855f-5dcc-49d8-914c-aaeac1bd96e6',
  '4ce23c76-4bbb-4a4c-9660-c72d5611b02c',
  '309ed7c6-e1d7-4dee-946d-518b4f66116d',
  '65f0d676-908d-4c5a-befe-67a858f06847',
  '0ac56ed2-cb12-43f6-940c-9b1d43f2618e',
  'dd11958c-2492-4a27-a33e-ee28a8ca7738',
  'e86a604f-e6d3-4645-a889-c713e3930685',
  'ebdbc830-7a7a-4984-b680-edf49574e828',
  'b17815f6-0a75-447e-ae82-13c2135e8f55',
  '8fec3366-7e59-4e74-a8a2-a61a9cc4454b',
  '6279934a-3700-4f52-b511-3935091b16cb',
  '4a766a12-c433-4dbf-a316-833bac154852',
  'b66d9be5-4524-4bc3-aa72-9d83f0644b93',
  '3b88a109-6fe5-4947-8b3c-30f7e041d2e3',
  '621d89b8-6f87-46da-b464-e36c3ab44a7f',
  '1608b95d-c202-4d79-948a-2f8eb2d58ef0',
  'aee69aec-b98b-42ad-8268-0d367029cfe8',
  'd51ae9b1-6924-4744-b424-cbecebfb9b47',
  'bf573175-b95d-4d3c-9bed-77d16fb2936d',
  'fdf88b51-e6aa-48a7-bb38-fadfe2b84b53',
  'cc4427ce-0c95-40ed-b009-715e3e4504cc',
  'd34e9f49-83c5-4f2f-ac8f-2b2b74d8f161',
  'fb049603-01d7-4685-9dda-52ac00a9d631',
  '8c3b80b1-4ec9-44dc-8240-9c15b9c608ca',
  '8ba61365-7526-48da-b48d-2932ad5ba300',
  '96283fc5-f928-450f-8d14-70fc8fa78ec5',
  '671b2de1-96e6-4e7b-b7af-72873f000c6b',
  '466dda97-db3e-4972-aa6c-793706f61c54',
  '6ebc0afa-45ab-4926-a375-e07b9ac0ef5d',
  '8f346bac-12f5-4484-bb7e-47f373bac4fb',
  '9beff54a-6c05-43fc-83f9-b7ce7ee674be',
  'ef9e0ccf-b65f-4a8c-94e4-b755527cfd61',
  'ffba9a67-d66e-4716-9ac4-44018d20eeed',
  '59c38d01-1d0a-4cbe-be0b-ba7e9709a080',
  '4fcf9c01-4f7c-4b41-99f4-b5dfbd5bf705',
  '397b53b8-88a1-4b74-bf02-9c5f3d9e3325',
  '114a4a94-0609-4b76-b07f-94e8332b1037',
  'e9ac97ff-bc48-4cfa-a91b-2eee0ffd4d17',
  'bdce959c-be30-447d-b993-2c6346db67f2',
  '8ba764aa-88b9-42e5-b7c8-bac29eeb93a4',
  'b96e20c0-c3de-4cee-a120-e0659c67f83e',
  '8c807bd4-a90e-4ecf-85de-278a3221139b',
  '514cc20c-8bdd-4d15-af43-442354684dcd',
  '9ffa74e6-a098-4d90-ae44-b10df9d19cc2',
  'ea0ddc01-dce3-4295-aa33-8390981fee3b',
  'fbec0ce1-4124-463f-b740-eb8c7520defa',
  '356a8be0-6fd2-4e56-988b-8f2e13664239',
  '09b684fe-0315-46b5-b1fd-6b434a7460d5',
  '1fbcc1ce-16b6-4882-8cbe-8a209362a543',
  'a7140eed-4b26-4710-823e-b8c04343c59c',
  'a8108c4c-6e8e-46bf-aad5-d154d8177445',
  '82afdf0c-3e0c-447b-8a40-095ca2c246ec',
  'fe3e66cf-79b8-47f6-9bf1-a7ad82fcc824',
  'dde5ac3b-7233-4a8a-923f-20eb5a2c9896',
  '53381e40-5954-40d2-bdf3-49f4e53058f4',
  '5bcac775-b0fe-46f5-9b37-211b55485af1',
  '98c8b5cc-d972-413b-abd8-b764209458a4',
]);

type WeHavePartsMatch = {
  vehicleJobNumber: string;
  statusPriority: string;
  parts: { name: string; year: string; make: string; model: string }[];
};

function buildWeHavePartsMatches(
  partsData: PartsRow[],
  rows: DashboardRow[]
): WeHavePartsMatch[] {
  // Qualifying dashboard rows: Vehicle On-Site, Insurance Approval,
  // Repair Approved (incl. PDR Approved variants), PDR In-Progress, Post Repair
  const qualifying = rows.filter((r) =>
    isVehicleOnSite(r) ||
    isInsuranceApproval(r) ||
    isRepairApproved(r) ||
    isPdrInProgress(r) ||
    isPostRepair(r)
  );

  // Stock parts = rows whose Set Key UUID is in the hard-coded STOCK_SET_KEYS allow-list.
  // This list represents the physical inventory currently sitting in the container.
  const stockParts = partsData.filter((p) => {
    const key = getSetKey(p).toLowerCase();
    return STOCK_SET_KEYS.has(key);
  });

  // Forgiving match: equal, or either string contains the other.
  // Handles "Civic" vs "Civic LX" vs "Honda Civic" vs trailing/extra whitespace.
  const modelsMatch = (a: string, b: string) => {
    if (!a || !b) return false;
    return a === b || a.includes(b) || b.includes(a);
  };

  // Group by vehicle: one entry per job number, all matching stock parts collected together.
  const byJob = new Map<string, WeHavePartsMatch>();
  for (const r of qualifying) {
    const vehicleModel = normalize(r['Model']).replace(/\s+/g, ' ').trim();
    if (!vehicleModel) continue;
    for (const p of stockParts) {
      const partModel = normalize(getPartModel(p)).replace(/\s+/g, ' ').trim();
      const partMake  = normalize(getPartMake(p)).replace(/\s+/g, ' ').trim();
      if (modelsMatch(vehicleModel, partModel) || modelsMatch(vehicleModel, partMake)) {
        const jn = toText(r['Job Number']);
        const existing = byJob.get(jn);
        const partEntry = {
          name: getPartName(p),
          year: getPartYear(p),
          make: getPartMake(p),
          model: getPartModel(p),
        };
        if (existing) {
          existing.parts.push(partEntry);
        } else {
          byJob.set(jn, {
            vehicleJobNumber: jn,
            statusPriority: toText(r['Status + Priority']),
            parts: [partEntry],
          });
        }
      }
    }
  }

  return Array.from(byJob.values()).sort((a, b) =>
    a.vehicleJobNumber.localeCompare(b.vehicleJobNumber, undefined, { numeric: true })
  );
}

function buildMustReturnGroups(
  partsData: PartsRow[],
  rows: DashboardRow[]
): MustReturnGroup[] {
  const jobStatusMap = new Map<string, string>();
  for (const row of rows) {
    const jn = normalize(toText(row['Job Number']));
    if (jn) jobStatusMap.set(jn, toText(row['Status + Priority']));
  }

  const qualifying = partsData.filter((part) => {
    const received = getReceivedAt(part);
    if (!received) return false;
    if (getCheckedOutAt(part) !== null) return false;
    if (getReturnedAt(part) !== null) return false;
    return daysSince(received) >= 25;
  });

  const groups = new Map<string, { parts: string[]; maxDays: number }>();
  for (const part of qualifying) {
    const jn = getPartJobNumber(part);
    if (isStockJob(jn)) continue;
    if (!jn) continue;
    const received = getReceivedAt(part)!;
    const days = daysSince(received);
    const existing = groups.get(jn);
    if (existing) {
      existing.parts.push(getPartName(part));
      existing.maxDays = Math.max(existing.maxDays, days);
    } else {
      groups.set(jn, { parts: [getPartName(part)], maxDays: days });
    }
  }

  return Array.from(groups.entries())
    .map(([jobNumber, { parts, maxDays }]) => ({
      jobNumber,
      parts: parts.filter(Boolean),
      maxDays,
      statusPriority: jobStatusMap.get(normalize(jobNumber)) ?? '',
    }))
    .sort((a, b) => b.maxDays - a.maxDays);
}

// --- Logic Builders ---

function matchesPartsStage(row: DashboardRow): boolean {
  const status = normalize(row['Status + Priority']);
  const days = toNumber(row['Status Days']);
  return (
    status === 'post repair' ||
    status === 'pdr in-progress' ||
    status === 'e - ehi repair' ||
    status === 'conventional (hail)' ||
    (status.includes('repair approved') && days >= 3)
  );
}

function buildClipboardText(rows: DashboardRow[]): string {
  return rows
    .map((row) => {
      const jobNumber = toText(row['Job Number']);
      const model = toText(row['Model']);
      return model ? `${jobNumber} - ${model}` : jobNumber;
    })
    .join('\n');
}

function average(values: number[]): number {
  if (values.length === 0) return 0;
  const total = values.reduce((sum, value) => sum + value, 0);
  return total / values.length;
}

function formatDays(value: number): string {
  return value.toFixed(1);
}

function repairBonusRate(avg: number): number {
  if (avg <= 8) return 15;
  if (avg <= 9) return 12;
  if (avg <= 10) return 8;
  if (avg <= 11) return 3;
  return 0;
}

function approvalBonusRate(avg: number): number {
  if (avg <= 3) return 15;
  if (avg <= 4) return 12;
  if (avg <= 5) return 7;
  if (avg <= 6) return 4;
  if (avg <= 7) return 2;
  return 0;
}

function buildAlertCards(rows: DashboardRow[]): AlertCard[] {
  const needsSeverity = rows.filter((r) => {
    const filter = normalize(r['Filter']);
    // Logic updated to match "ana", "roy", or "roy / ana"
    return (filter === 'ana' || filter === 'roy' || filter === 'roy / ana') && isBlank(r['Severity']);
  });

  const escalationOnSite = rows.filter((r) => {
    return normalize(r['Status + Priority']) === 'escalation on-site';
  });

  const saMonthlyQc = rows.filter((r) => {
    return isVehicleDeliveredHail(r) && getSaMonthlyQcIssues(r).length > 0;
  });

  const generalParts = rows.filter((r) => {
    return (
      matchesPartsStage(r) &&
      includesAny(r['Task Titles'], ['order parts', 'parts received'])
    );
  });

  const glassParts = rows.filter((r) => {
    return (
      matchesPartsStage(r) &&
      includesAny(r['Task Titles'], [
        'glass order',
        'order windshield',
        'glass received',
      ])
    );
  });

  const scheduleGlassInstall = rows.filter((r) => {
    return (
      isPostRepair(r) &&
      includesAny(r['Task Titles'], [
        'glass install (both)',
        'glass install',
        'install quarter glass',
      ])
    );
  });

  const glassInstallAfterDelivery = rows.filter((r) => {
    return (
      isVehicleDeliveredHail(r) &&
      includesAny(r['Task Titles'], [
        'glass install (both)',
        'glass install',
        'order windshield',
        'glass received',
        'install quarter glass',
      ])
    );
  });

  const conventionalMissing = rows.filter(
    (r) => isConventionalHail(r) && isBlank(r['Body ECD'])
  );

  const conventionalPastDue = rows.filter(
    (r) => isConventionalHail(r) && !isBlank(r['Body ECD']) && isPastDue(r['Body ECD'])
  );

  return [
    {
      id: 'needs-severity',
      title: 'Needs Severity',
      count: needsSeverity.length,
      rows: sortByPriority(needsSeverity),
      description: 'Roy / Ana filters missing severity',
      info: 'This alert appears when the Filter column is Ana, Roy, or Roy / Ana and the Severity field is blank.',
      section: 'Operations',
    },
    {
      id: 'escalation-onsite',
      title: 'Escalation On-site',
      count: escalationOnSite.length,
      rows: sortByPriority(escalationOnSite),
      description: 'Jobs marked as Escalation On-site',
      info: 'This alert triggers whenever a vehicle has the "Status + Priority" set to "Escalation On-site".',
      section: 'Operations',
    },
    {
      id: 'sa-monthly-qc',
      title: 'SA’s Monthly QC',
      count: saMonthlyQc.length,
      rows: sortByPriority(saMonthlyQc),
      description: 'Delivered hail jobs with date or QC inconsistencies',
      info: 'This alert checks Vehicle Delivered (Hail) jobs and flags missing start, repair approved, or end dates, date order inconsistencies, and QC Not Completed entries marked FLAG. The detail view groups every flagged job by SA.',
      section: 'Operations',
      detailType: 'sa-monthly-qc',
    },
    {
      id: 'general-parts',
      title: 'General Parts Incomplete',
      count: generalParts.length,
      rows: sortByPriority(generalParts),
      description: 'General parts tasks still active',
      info: 'This alert appears when Order Parts or Parts Received still exists in Task Titles while the job is in Post Repair, PDR In-Progress, E - EHI Repair, Conventional (Hail), or Repair Approved with 3 or more status days.',
      section: 'Parts',
    },
    {
      id: 'glass-parts',
      title: 'Glass Parts Incomplete',
      count: glassParts.length,
      rows: sortByPriority(glassParts),
      description: 'Glass tasks still active',
      info: 'This alert appears when Glass Order, Order Windshield, or Glass Received still exists in Task Titles while the job is in Post Repair, PDR In-Progress, E - EHI Repair, Conventional (Hail), or Repair Approved with 3 or more status days.',
      section: 'Parts',
    },
    {
      id: 'schedule-glass-install',
      title: 'Schedule Glass Install',
      count: scheduleGlassInstall.length,
      rows: sortByPriority(scheduleGlassInstall),
      description: 'Glass install task exists while status is Post Repair',
      info: 'This alert appears when Task Titles contains Glass Install (Both), Glass Install, or Install Quarter Glass and Status + Priority is Post Repair.',
      section: 'Parts',
    },
    {
      id: 'glass-install-after-delivery',
      title: 'Glass Install AFTER-DELIVERY',
      count: glassInstallAfterDelivery.length,
      rows: sortByPriority(glassInstallAfterDelivery),
      description: 'Glass related tasks still active after delivery',
      info: 'This alert appears when Task Titles contains Glass Install (Both), Glass Install, Order Windshield, Glass Received, or Install Quarter Glass and Status + Priority is Vehicle Delivered (Hail).',
      section: 'Parts',
    },
    {
      id: 'conv-missing',
      title: 'Missing Body ECD',
      count: conventionalMissing.length,
      rows: sortByPriority(conventionalMissing),
      description: 'Conventional without Body ECD',
      info: 'This alert appears when Status + Priority is Conventional (Hail) and Body ECD is blank.',
      section: 'Conventional',
    },
    {
      id: 'conv-past-due',
      title: 'Past Due Body ECD',
      count: conventionalPastDue.length,
      rows: sortByPriority(conventionalPastDue),
      description: 'Conventional past due',
      info: 'This alert appears when Status + Priority is Conventional (Hail) and Body ECD is already past due.',
      section: 'Conventional',
    },
  ];
}

function buildMainCards(rows: DashboardRow[]): MainCard[] {
  const insuranceRows = sortByPriority(rows.filter(isInsuranceApproval));
  const repairApprovedRows = sortByPriority(rows.filter(isRepairApproved));
  const pdrRows = sortByPriority(rows.filter(isPdrInProgress));
  const conventionalHailRows = sortByPriority(rows.filter(isConventionalHail));
  const postRepairRows = sortByPriority(rows.filter(isPostRepair));
  const readyRows = sortByPriority(rows.filter(isReadyToDeliver));
  const onSiteRows = sortByPriority(rows.filter(isVehicleOnSite));
  const deliveredRows = sortByPriority(rows.filter(isVehicleDeliveredHail));

  return [
    {
      id: 'total-jobs',
      title: 'Total Jobs',
      count: rows.length,
      rows: sortByPriority(rows),
      modalType: 'sa-chart',
    },
    {
      id: 'vehicle-on-site',
      title: 'Vehicle On-Site',
      count: onSiteRows.length,
      rows: onSiteRows,
      modalType: 'job-list',
      isDelayed: onSiteRows.some(isRowDelayed),
    },
    {
      id: 'insurance-approval-main',
      title: 'Insurance Approval',
      count: insuranceRows.length,
      rows: insuranceRows,
      modalType: 'job-list',
      isDelayed: insuranceRows.some(isRowDelayed),
    },
    {
      id: 'repair-approved-main',
      title: 'Repair Approved',
      count: repairApprovedRows.length,
      rows: repairApprovedRows,
      modalType: 'repair-approved-buckets',
    },
    {
      id: 'pdr-in-progress-main',
      title: 'PDR In-Progress',
      count: pdrRows.length,
      rows: pdrRows,
      modalType: 'job-list',
      isDelayed: pdrRows.some(isRowDelayed),
    },
    {
      id: 'conventional-hail-main',
      title: 'Conventional (Hail)',
      count: conventionalHailRows.length,
      rows: conventionalHailRows,
      modalType: 'job-list',
      isDelayed: conventionalHailRows.some(isRowDelayed),
    },
    {
      id: 'post-repair-main',
      title: 'Post Repair',
      count: postRepairRows.length,
      rows: postRepairRows,
      modalType: 'job-list',
      isDelayed: postRepairRows.some(isRowDelayed),
    },
    {
      id: 'ready-to-deliver-main',
      title: 'Ready to Deliver',
      count: readyRows.length,
      rows: readyRows,
      modalType: 'job-list',
      isDelayed: readyRows.some(isRowDelayed),
    },
    {
      id: 'vehicle-delivered-hail-main',
      title: 'Vehicle Delivered (Hail)',
      count: deliveredRows.length,
      rows: deliveredRows,
      modalType: 'delivered-hail',
    },
  ];
}

function alertColorClasses(alertId: string, count: number): string {
  if (alertId === 'escalation-onsite') {
    return count > 0 ? 'bg-red-100 border-red-400' : 'bg-green-100 border-green-400';
  }
  if (count >= 5) return 'bg-red-100 border-red-400';
  if (count > 0) return 'bg-yellow-100 border-yellow-400';
  return 'bg-green-100 border-green-400';
}

function getSaCounts(rows: DashboardRow[]) {
  const counts = new Map<string, number>();
  for (const row of rows) {
    const sa = getSaName(row);
    counts.set(sa, (counts.get(sa) || 0) + 1);
  }
  return Array.from(counts.entries())
    .map(([sa, count]) => ({ sa, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);
}

// --- Flip Clock Components ---

const FLIP_W = 14, FLIP_H = 22, FLIP_FONT = 16;

function FlipDigit({ value }: { value: string }) {
  const [display, setDisplay] = useState(value);
  const [pending, setPending] = useState<string | null>(null);
  const [phase, setPhase] = useState<'idle' | 'top' | 'bottom'>('idle');
  const displayRef = useRef(value);

  useEffect(() => {
    if (value !== displayRef.current) {
      setPending(value);
      setPhase('top');
    }
  }, [value]);

  const handleTopEnd = useCallback(() => {
    if (pending !== null) {
      displayRef.current = pending;
      setDisplay(pending);
      setPhase('bottom');
    }
  }, [pending]);

  const handleBottomEnd = useCallback(() => {
    setPending(null);
    setPhase('idle');
  }, []);

  // Uses wrapper trick: centers digit in full FLIP_H space, clips to half
  function HalfPanel({ digit, half, extra }: { digit: string; half: 'top' | 'bottom'; extra?: React.CSSProperties }) {
    const isTop = half === 'top';
    return (
      <div style={{
        position: 'absolute',
        top: isTop ? 0 : FLIP_H / 2,
        left: 0, right: 0,
        height: FLIP_H / 2,
        overflow: 'hidden',
        background: isTop ? '#ffffff' : '#f2f2f2',
        borderRadius: isTop ? '2px 2px 0 0' : '0 0 2px 2px',
        border: '1px solid #222',
        borderTop: isTop ? '1px solid #222' : 'none',
        borderBottom: isTop ? 'none' : '1px solid #222',
        ...extra,
      }}>
        <div style={{
          height: FLIP_H,
          marginTop: isTop ? 0 : -(FLIP_H / 2),
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          <span style={{
            fontFamily: '"Courier New", Courier, monospace',
            fontSize: FLIP_FONT, fontWeight: 'bold', color: '#111',
            userSelect: 'none', lineHeight: 1,
          }}>{digit}</span>
        </div>
      </div>
    );
  }

  const shown  = phase === 'idle' ? display : (pending ?? display);
  const bottom = phase === 'bottom' ? (pending ?? display) : display;

  return (
    <div style={{ position: 'relative', width: FLIP_W, height: FLIP_H, display: 'inline-block' }}>
      <HalfPanel digit={shown}   half="top" />
      <HalfPanel digit={bottom}  half="bottom" />
      {/* Center divider */}
      <div style={{ position: 'absolute', top: FLIP_H / 2 - 0.5, left: 0, right: 0, height: 1, background: '#444', zIndex: 5 }} />
      {/* Flip top — old digit folding away */}
      {phase === 'top' && (
        <HalfPanel digit={display} half="top"
          extra={{ transformOrigin: 'bottom center', zIndex: 10 } as React.CSSProperties}
        />
      )}
      {/* We can't put className on HalfPanel directly, so wrap */}
      {phase === 'top' && (
        <div className="flip-top-out" onAnimationEnd={handleTopEnd}
          style={{ position: 'absolute', top: 0, left: 0, right: 0, height: FLIP_H / 2, transformOrigin: 'bottom center', zIndex: 20, overflow: 'hidden', background: '#ffffff', borderRadius: '2px 2px 0 0', border: '1px solid #222', borderBottom: 'none' }}>
          <div style={{ height: FLIP_H, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <span style={{ fontFamily: '"Courier New", Courier, monospace', fontSize: FLIP_FONT, fontWeight: 'bold', color: '#111', userSelect: 'none', lineHeight: 1 }}>{display}</span>
          </div>
        </div>
      )}
      {/* Flip bottom — new digit unfolding */}
      {phase === 'bottom' && (
        <div className="flip-bottom-in" onAnimationEnd={handleBottomEnd}
          style={{ position: 'absolute', top: FLIP_H / 2, left: 0, right: 0, height: FLIP_H / 2, transformOrigin: 'top center', zIndex: 20, overflow: 'hidden', background: '#f2f2f2', borderRadius: '0 0 2px 2px', border: '1px solid #222', borderTop: 'none' }}>
          <div style={{ height: FLIP_H, marginTop: -(FLIP_H / 2), display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <span style={{ fontFamily: '"Courier New", Courier, monospace', fontSize: FLIP_FONT, fontWeight: 'bold', color: '#111', userSelect: 'none', lineHeight: 1 }}>{pending ?? display}</span>
          </div>
        </div>
      )}
    </div>
  );
}

function FlipClock({ seconds }: { seconds: number }) {
  const mm = String(Math.floor(seconds / 60)).padStart(2, '0');
  const ss = String(seconds % 60).padStart(2, '0');
  return (
    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 2 }}>
      <FlipDigit value={mm[0]} />
      <FlipDigit value={mm[1]} />
      <div style={{ display: 'flex', flexDirection: 'column', gap: 3, margin: '0 1px' }}>
        <div style={{ width: 3, height: 3, borderRadius: '50%', background: '#333' }} />
        <div style={{ width: 3, height: 3, borderRadius: '50%', background: '#333' }} />
      </div>
      <FlipDigit value={ss[0]} />
      <FlipDigit value={ss[1]} />
    </div>
  );
}

// --- Holidays ---

const DMG_HOLIDAYS_2026 = [
  { name: 'New Year\'s Day',    date: new Date('2026-01-01T00:00:00'), note: '' },
  { name: 'Spring Holiday',     date: new Date('2026-03-06T00:00:00'), note: '*' },
  { name: 'Memorial Day',       date: new Date('2026-05-25T00:00:00'), note: '' },
  { name: 'Juneteenth',         date: new Date('2026-06-19T00:00:00'), note: '*' },
  { name: 'Independence Day',   date: new Date('2026-07-04T00:00:00'), note: '' },
  { name: 'Labor Day',          date: new Date('2026-09-07T00:00:00'), note: '' },
  { name: 'Thanksgiving',       date: new Date('2026-11-26T00:00:00'), note: '' },
  { name: 'Christmas Day',      date: new Date('2026-12-25T00:00:00'), note: '' },
  { name: 'New Year\'s Day',    date: new Date('2027-01-01T00:00:00'), note: '' },
];

function formatHolidayDate(d: Date): string {
  return new Intl.DateTimeFormat('en-US', {
    weekday: 'long', month: 'numeric', day: 'numeric', year: '2-digit',
  }).format(d);
}

function getUpcomingHoliday(): { name: string; note: string; weeksAway: number; daysAway: number } | null {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  for (const h of DMG_HOLIDAYS_2026) {
    const diff = Math.ceil((h.date.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    if (diff >= 0) {
      return { name: h.name, note: h.note, daysAway: diff, weeksAway: Math.floor(diff / 7) };
    }
  }
  return null;
}

// --- Main Page Component ---

export default function Page() {
  const [data, setData] = useState<{
    updatedAt?: string;
    rows?: DashboardRow[];
    partsHeaders?: string[];
    partsRows?: unknown[];
  }>({});
  const [selectedAlertId, setSelectedAlertId] = useState<string | null>(null);
  const [selectedMainId, setSelectedMainId] = useState<string | null>(null);
  const [selectedSa, setSelectedSa] = useState<string | null>(null);
  const [infoAlertId, setInfoAlertId] = useState<string | null>(null);
  const [copyMessage, setCopyMessage] = useState('');
  const [jobSearch, setJobSearch] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
  const searchRef = useRef<HTMLDivElement>(null);
  const alertDetailsRef = useRef<HTMLElement>(null);
  const [holidaysOpen, setHolidaysOpen] = useState(false);
  const [hideZeroAlerts, setHideZeroAlerts] = useState(false);
  const [showBonus, setShowBonus] = useState(false);
  const [bonusUnlocked, setBonusUnlocked] = useState(false);
  const [bonusInput, setBonusInput] = useState('');
  const [bonusError, setBonusError] = useState(false);

  const [countdown, setCountdown] = useState(300);
  const [easterEggActive, setEasterEggActive] = useState(false);
  const [easterEggInput, setEasterEggInput] = useState('');
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const fetchData = useCallback(() => {
    fetch('/api/dashboard')
      .then((r) => r.json())
      .then(setData);
  }, []);

  useEffect(() => {
    fetchData();
  }, [fetchData]);

  useEffect(() => {
    if (intervalRef.current) clearInterval(intervalRef.current);
    intervalRef.current = setInterval(() => {
      setCountdown((c) => {
        if (c <= 1) { fetchData(); return 300; }
        return c - 1;
      });
    }, 1000);
    return () => {
      if (intervalRef.current) clearInterval(intervalRef.current);
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (searchRef.current && !searchRef.current.contains(e.target as Node)) {
        setSearchOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const formattedLastPulled = useMemo(() => {
    if (!data.updatedAt) return 'N/A';
    try {
      return new Intl.DateTimeFormat('en-US', {
        timeZone: 'America/Chicago',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        hour12: true,
        timeZoneName: 'short'
      }).format(new Date(data.updatedAt));
    } catch {
      return data.updatedAt;
    }
  }, [data.updatedAt]);

  const rows = data.rows ?? [];

  const partsData = useMemo((): PartsRow[] => {
    const rawRows = data.partsRows;
    if (!rawRows) return [];
    return rawRows.map((row) => {
      if (Array.isArray(row)) {
        const headers = data.partsHeaders ?? [];
        const obj: PartsRow = {};
        headers.forEach((h, i) => { obj[h] = (row as unknown[])[i] as string | number | null | undefined; });
        return obj;
      }
      return row as PartsRow;
    });
  }, [data.partsHeaders, data.partsRows]);

  const jobSearchResults = useMemo(() => {
    const query = jobSearch.trim().toLowerCase();
    if (!query) return [];
    return rows
      .filter((r) => normalize(r['Job Number']).includes(query))
      .slice(0, 10);
  }, [rows, jobSearch]);

  useEffect(() => {
    if (selectedAlertId && alertDetailsRef.current) {
      alertDetailsRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [selectedAlertId]);

  const alertCards = useMemo(() => buildAlertCards(rows), [rows]);
  const mainCards = useMemo(() => buildMainCards(rows), [rows]);

  const mustReturnGroups = useMemo(() => buildMustReturnGroups(partsData, rows), [partsData, rows]);
  const missingInstallGroups = useMemo(() => buildMissingInstallGroups(partsData, rows), [partsData, rows]);
  const weHavePartsMatches = useMemo(() => buildWeHavePartsMatches(partsData, rows), [partsData, rows]);

  const wipCount = useMemo(() => rows.filter((r) =>
    isVehicleOnSite(r) || isInsuranceApproval(r) || isRepairApproved(r) ||
    isPdrInProgress(r) || isConventionalHail(r) || isPostRepair(r) || isReadyToDeliver(r)
  ).length, [rows]);

  const inventoryAlertCards = useMemo((): AlertCard[] => [
    {
      id: 'must-return',
      title: 'Must Return',
      count: mustReturnGroups.length,
      rows: [],
      description: 'Parts received ≥25 days — not checked out or returned',
      info: 'Flags parts that have been received for 25 or more days without being checked out or returned. Groups multiple parts under the same job number.',
      section: 'Inventory',
      detailType: 'must-return',
    },
    {
      id: 'missing-install',
      title: 'Missing Install / Not Checked Out',
      count: missingInstallGroups.length,
      rows: [],
      description: 'Post Repair / Ready to Deliver / Delivered jobs with parts not received or checked out',
      info: 'For jobs where Status + Priority is Post Repair, Ready to Deliver, or Vehicle Delivered (Hail), flags parts that have an Ordered At date but are missing either Received At or Checked Out At. Helps catch parts that need check-out before install, or install that was skipped before delivery. Ignores job 000.',
      section: 'Inventory',
      detailType: 'missing-install',
    },
    {
      id: 'we-have-parts',
      title: 'We have parts!',
      count: weHavePartsMatches.length,
      rows: [],
      description: 'Stock (Job 000) parts that match vehicles currently in the lot',
      info: 'Scans the PARTS sheet for the specific Set Key UUIDs that represent current stock inventory, then matches their Model and Make (sometimes the Make is written in the Model field on the cycle time sheet) to any vehicle currently in Vehicle On-Site, Insurance Approval, Repair Approved, PDR Approved, PDR In-Progress, or Post Repair. The goal is to use up old stock before ordering new parts. The stock Set Keys list is maintained in the code.',
      section: 'Inventory',
      detailType: 'we-have-parts',
    },
  ], [mustReturnGroups, missingInstallGroups, weHavePartsMatches]);

  const allAlertCards = useMemo(() => [...alertCards, ...inventoryAlertCards], [alertCards, inventoryAlertCards]);

  const grouped = useMemo(() => ({
    Operations: alertCards.filter((a) => a.section === 'Operations'),
    Parts: alertCards.filter((a) => a.section === 'Parts'),
    Conventional: alertCards.filter((a) => a.section === 'Conventional'),
    Inventory: inventoryAlertCards,
  }), [alertCards, inventoryAlertCards]);

  const selectedAlert = allAlertCards.find((a) => a.id === selectedAlertId) ?? null;
  const selectedMain = mainCards.find((m) => m.id === selectedMainId) ?? null;
  const saCounts = useMemo(() => getSaCounts(rows), [rows]);
  const maxSaCount = saCounts.length > 0 ? Math.max(...saCounts.map((x) => x.count)) : 1;

  const selectedSaRows = useMemo(() => {
    if (!selectedSa) return [];
    return sortByPriority(rows.filter((row) => getSaName(row) === selectedSa));
  }, [rows, selectedSa]);

  const selectedAlertMonthlyQcBuckets = useMemo(() => {
    if (!selectedAlert || selectedAlert.detailType !== 'sa-monthly-qc') return [];
    return buildSaMonthlyQcBuckets(selectedAlert.rows);
  }, [selectedAlert]);

  const selectedAlertMonthlyQcClipboardText = useMemo(() => {
    if (selectedAlertMonthlyQcBuckets.length === 0) return '';
    return buildSaMonthlyQcClipboardText(selectedAlertMonthlyQcBuckets);
  }, [selectedAlertMonthlyQcBuckets]);

  const repairBuckets = useMemo(() => {
    if (!selectedMain || selectedMain.modalType !== 'repair-approved-buckets') return null;
    
    const enterprise = sortByPriority(selectedMain.rows.filter(r => 
      normalize(r['Job Type']) === 'enterprise rental' && 
      normalize(r['Status + Priority']).includes('e - ehi repair')
    ));

    const progressive = sortByPriority(selectedMain.rows.filter(r => 
      normalize(r['Insurance']) === 'progressive'
    ));

    const progressiveIds = new Set(progressive.map(r => r['Job Number']));
    const enterpriseIds = new Set(enterprise.map(r => r['Job Number']));

    const general = sortByPriority(selectedMain.rows.filter(r => 
      !progressiveIds.has(r['Job Number']) && !enterpriseIds.has(r['Job Number'])
    ));

    return { enterprise, progressive, general };
  }, [selectedMain]);

  const deliveredHailStats = useMemo(() => {
    const deliveredRows = rows.filter(isVehicleDeliveredHail);
    
    const approvalTimeValues = deliveredRows.map((r) => toNumber(r['Approval Time'])).filter((v) => v > 0);
    const pendingPdrValues = deliveredRows.map((r) => toNumber(r['Approved Pending PDR'])).filter((v) => v > 0);
    const repairTimeValues = deliveredRows.map((r) => toNumber(r['Repair Time'])).filter((v) => v > 0);
    const deliveryTimeValues = deliveredRows.map((r) => toNumber(r['Delivery Time'])).filter((v) => v > 0);

    const totalTimeValues = deliveredRows.map((r) =>
      (toNumber(r['Approved Pending PDR']) / 2) + toNumber(r['Repair Time']) + toNumber(r['Delivery Time'])
    ).filter((v) => v > 0);

    const avgApproval = average(approvalTimeValues);
    const avgTotal = average(totalTimeValues);
    const count = deliveredRows.length;

    return {
      avgApprovalTime: avgApproval,
      avgApprovedPendingPdr: average(pendingPdrValues) / 2, 
      avgRepairTime: average(repairTimeValues),
      avgDeliveryTime: average(deliveryTimeValues),
      avgApprovedToDelivered: avgTotal,
      repairRate: repairBonusRate(avgTotal),
      approvalRate: approvalBonusRate(avgApproval),
      deliveredCount: count,
      repairBonusTotal: repairBonusRate(avgTotal) * count,
      approvalBonusTotal: approvalBonusRate(avgApproval) * count,
      combinedBonusTotal: (repairBonusRate(avgTotal) + approvalBonusRate(avgApproval)) * count,
      deliveredRows: sortByPriority(deliveredRows),
    };
  }, [rows]);

  async function handleCopyText(textToCopy: string) {
    if (!textToCopy) return;
    try {
      await navigator.clipboard.writeText(textToCopy);
      setCopyMessage('Copied');
      window.setTimeout(() => setCopyMessage(''), 2000);
    } catch {
      setCopyMessage('Copy failed');
      window.setTimeout(() => setCopyMessage(''), 2000);
    }
  }

  async function handleCopyRows(rowsToCopy: DashboardRow[]) {
    if (rowsToCopy.length === 0) return;
    await handleCopyText(buildClipboardText(rowsToCopy));
  }

  return (
    <main className="min-h-screen bg-slate-100 text-slate-900 p-6 md:p-8">
      <div className="mx-auto max-w-7xl space-y-8">
        
        {/* Header */}
        <section className="rounded-3xl bg-white border border-slate-300 p-6 shadow-sm">
          <div className="flex items-start justify-between gap-4">
            <h1 className="text-3xl font-bold">Operations Manager Dashboard</h1>
          </div>
          <div className="mt-2 flex items-center gap-2">
            <FlipClock seconds={countdown} />
            {easterEggActive ? (
              <input
                type="text"
                autoFocus
                value={easterEggInput}
                onChange={(e) => {
                  const v = e.target.value;
                  setEasterEggInput(v);
                  if (v === '4815162342') {
                    fetchData();
                    setCountdown(300);
                    setEasterEggActive(false);
                    setEasterEggInput('');
                  }
                }}
                onBlur={() => { setEasterEggActive(false); setEasterEggInput(''); }}
                onKeyDown={(e) => { if (e.key === 'Escape') { setEasterEggActive(false); setEasterEggInput(''); } }}
                className="w-28 rounded border border-slate-300 px-2 py-0.5 text-xs outline-none focus:ring-1 focus:ring-slate-400"
                placeholder="_ _ _ _ _ _ _ _ _ _"
              />
            ) : (
              <button
                onClick={() => setEasterEggActive(true)}
                className="text-xs text-slate-400 hover:text-slate-600 transition-colors"
              >
                next refresh
              </button>
            )}
          </div>
          <p className="mt-1 text-sm text-slate-600">Last pulled: {formattedLastPulled}</p>

          {/* Holidays panel — bottom right */}
          <div className="mt-3 flex justify-end">
            <div className="relative text-right">
              {/* Collapsed summary */}
              {!holidaysOpen && (() => {
                const upcoming = getUpcomingHoliday();
                return (
                  <button
                    onClick={() => setHolidaysOpen(true)}
                    className="text-xs text-slate-400 hover:text-slate-600 transition-colors"
                  >
                    {upcoming
                      ? <>🗓 <span className="font-medium">{upcoming.name}{upcoming.note}</span> in {upcoming.daysAway === 0 ? 'today!' : upcoming.weeksAway < 1 ? `${upcoming.daysAway}d` : `${upcoming.weeksAway}w`}</>
                      : '🗓 No upcoming holidays'}
                  </button>
                );
              })()}

              {/* Expanded holiday list */}
              {holidaysOpen && (
                <div className="absolute right-0 top-6 z-50 rounded-xl border border-slate-200 bg-white shadow-lg p-3 text-left min-w-[260px]">
                  <div className="flex items-center justify-between mb-2">
                    <p className="text-xs font-bold text-slate-600 uppercase tracking-wide">DMG Observed Holidays 2026</p>
                    <button onClick={() => setHolidaysOpen(false)} className="text-slate-400 hover:text-slate-600 text-xs ml-3">✕</button>
                  </div>
                  <ul className="space-y-1">
                    {DMG_HOLIDAYS_2026.map((h, i) => {
                      const today = new Date(); today.setHours(0,0,0,0);
                      const isPast = h.date < today;
                      const isToday = h.date.getTime() === today.getTime();
                      return (
                        <li key={i} className={`flex items-baseline justify-between gap-4 text-xs ${isPast ? 'text-slate-300 line-through' : isToday ? 'text-green-600 font-bold' : 'text-slate-700'}`}>
                          <span>{h.name}{h.note && <span className="text-slate-400">{h.note}</span>}</span>
                          <span className="tabular-nums text-right whitespace-nowrap">{formatHolidayDate(h.date)}</span>
                        </li>
                      );
                    })}
                  </ul>
                  <p className="mt-2 text-[10px] text-slate-400">*Subject to change with Ice Closures / Business Need</p>
                </div>
              )}
            </div>
          </div>
        </section>

        {/* Main Stat Cards */}
        <section className="space-y-4">
          <div className="flex items-center justify-between gap-4">
            <h2 className="text-2xl font-semibold">Main Information</h2>
            <div ref={searchRef} className="relative w-64">
              <input
                type="text"
                inputMode="numeric"
                placeholder="Search job number…"
                value={jobSearch}
                onChange={(e) => { setJobSearch(e.target.value); setSearchOpen(true); }}
                onFocus={() => setSearchOpen(true)}
                className="w-full rounded-xl border border-slate-300 bg-white px-4 py-2 text-sm shadow-sm outline-none focus:ring-2 focus:ring-slate-400"
              />
              {searchOpen && jobSearchResults.length > 0 && (
                <ul className="absolute right-0 z-50 mt-1 w-full rounded-xl border border-slate-200 bg-white shadow-lg overflow-hidden">
                  {jobSearchResults.map((r, i) => (
                    <li
                      key={i}
                      className="flex flex-col gap-0.5 px-4 py-2.5 text-sm hover:bg-slate-50 cursor-default border-b border-slate-100 last:border-0"
                    >
                      <span className="font-semibold text-slate-900">{toText(r['Job Number'])}</span>
                      <span className="text-xs text-slate-500">{toText(r['Status + Priority'])}</span>
                    </li>
                  ))}
                </ul>
              )}
              {searchOpen && jobSearch.trim() !== '' && jobSearchResults.length === 0 && (
                <div className="absolute right-0 z-50 mt-1 w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-500 shadow-lg">
                  No matching jobs found.
                </div>
              )}
            </div>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-9 gap-4">
            {mainCards.map((card) => (
              <button
                key={card.id}
                type="button"
                onClick={() => { setSelectedMainId(card.id); setSelectedSa(null); }}
                className={`rounded-2xl border p-4 text-left shadow-sm transition hover:shadow-md ${
                  card.isDelayed ? 'bg-red-100 border-red-300' : 'bg-white border-slate-300'
                } ${selectedMainId === card.id ? 'ring-2 ring-slate-900' : ''}`}
              >
                <p className="text-[10px] font-semibold uppercase tracking-wide text-slate-700 leading-tight">{card.title}</p>
                <p className="mt-2 text-3xl font-bold text-slate-900">{card.count}</p>
                {card.id === 'total-jobs' && (
                  <p className="mt-1 text-[10px] text-slate-400">WIP: {wipCount}</p>
                )}
              </button>
            ))}
          </div>
        </section>

        {/* Global alert filter toggle */}
        <div className="flex justify-end -mb-4">
          <button
            onClick={() => setHideZeroAlerts((v) => !v)}
            className={`rounded-full border px-3 py-1 text-xs font-medium transition-colors ${
              hideZeroAlerts
                ? 'bg-slate-800 text-white border-slate-800 hover:bg-slate-700'
                : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title={hideZeroAlerts ? 'Show all alerts' : 'Hide alerts with 0 count'}
          >
            {hideZeroAlerts ? '👁 Show all alerts' : '🙈 Hide zero alerts'}
          </button>
        </div>

        {/* Alert Sections */}
        {(['Operations', 'Parts', 'Conventional'] as const).map((section) => {
          const visibleAlerts = hideZeroAlerts
            ? grouped[section].filter((a) => a.count > 0)
            : grouped[section];
          const visibleInventory = hideZeroAlerts
            ? grouped.Inventory.filter((a) => a.count > 0)
            : grouped.Inventory;
          // If hide is on and there's nothing left in this section (including Inventory for Parts), skip rendering
          const sectionHasContent = visibleAlerts.length > 0 ||
            (section === 'Parts' && visibleInventory.length > 0);
          if (hideZeroAlerts && !sectionHasContent) return null;
          return (
          <section key={section} className="space-y-4">
            <h2 className="text-2xl font-semibold">{section}</h2>
            <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3">
              {visibleAlerts.map((alert) => (
                <div key={alert.id} className={`rounded-xl border p-3 shadow-sm min-h-[108px] ${alertColorClasses(alert.id, alert.count)} ${selectedAlertId === alert.id ? 'ring-2 ring-slate-900' : ''}`}>
                  <div className="flex items-start justify-between gap-2">
                    <button type="button" onClick={() => setSelectedAlertId(alert.id)} className="min-w-0 flex-1 text-left">
                      <p className="text-xs font-semibold text-slate-700 leading-tight">{alert.title}</p>
                      <p className="mt-2 text-2xl font-bold text-slate-900">{alert.count}</p>
                    </button>
                    <button onClick={() => setInfoAlertId(infoAlertId === alert.id ? null : alert.id)} className="h-5 w-5 rounded-full border border-slate-400 bg-white text-xs font-semibold flex items-center justify-center">i</button>
                  </div>
                  <p className="mt-2 text-xs text-slate-700 leading-tight">{alert.description}</p>
                  {infoAlertId === alert.id && (
                    <div className="mt-3 rounded-lg border border-slate-300 bg-white p-2 text-[11px] leading-snug text-slate-700">{alert.info}</div>
                  )}
                </div>
              ))}
            </div>

            {/* Inventory Control subsection — shown inside Parts */}
            {section === 'Parts' && visibleInventory.length > 0 && (
              <div className="space-y-3 pt-2">
                <h3 className="text-lg font-semibold text-slate-600">Inventory Control</h3>
                <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3">
                  {visibleInventory.map((alert) => (
                    <div key={alert.id} className={`rounded-xl border p-3 shadow-sm min-h-[108px] ${alertColorClasses(alert.id, alert.count)} ${selectedAlertId === alert.id ? 'ring-2 ring-slate-900' : ''}`}>
                      <div className="flex items-start justify-between gap-2">
                        <button type="button" onClick={() => setSelectedAlertId(alert.id)} className="min-w-0 flex-1 text-left">
                          <p className="text-xs font-semibold text-slate-700 leading-tight">{alert.title}</p>
                          <p className="mt-2 text-2xl font-bold text-slate-900">{alert.count}</p>
                        </button>
                        <button onClick={() => setInfoAlertId(infoAlertId === alert.id ? null : alert.id)} className="h-5 w-5 rounded-full border border-slate-400 bg-white text-xs font-semibold flex items-center justify-center">i</button>
                      </div>
                      <p className="mt-2 text-xs text-slate-700 leading-tight">{alert.description}</p>
                      {infoAlertId === alert.id && (
                        <div className="mt-3 rounded-lg border border-slate-300 bg-white p-2 text-[11px] leading-snug text-slate-700">{alert.info}</div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}
          </section>
        );
        })}

        {/* Selected Alert Details Table */}
        <section ref={alertDetailsRef} className="rounded-3xl bg-white border border-slate-300 p-6 shadow-sm">
          {!selectedAlert ? (
            <>
              <h3 className="text-xl font-semibold">Alert Details</h3>
              <p className="mt-3 text-slate-600">Select an alert card to see the matching jobs.</p>
            </>
          ) : selectedAlert.detailType === 'sa-monthly-qc' ? (
            <>
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h3 className="text-xl font-semibold">{selectedAlert.title}</h3>
                  <p className="mt-1 text-slate-600">{selectedAlert.count} matching job(s)</p>
                </div>
                <div className="flex items-center gap-3">
                  <button onClick={() => handleCopyText(selectedAlertMonthlyQcClipboardText)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium flex items-center gap-2">
                    <span>📋</span><span>Copy Report</span>
                  </button>
                  <button onClick={() => setSelectedAlertId(null)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Clear</button>
                </div>
              </div>
              {copyMessage && <p className="mt-3 text-sm text-slate-600">{copyMessage}</p>}
              {selectedAlertMonthlyQcBuckets.length === 0 ? (
                <p className="mt-6 text-slate-600">No delivered hail inconsistencies were found for this report.</p>
              ) : (
                <div className="mt-6 space-y-6">
                  {selectedAlertMonthlyQcBuckets.map((bucket) => (
                    <div key={bucket.sa} className="rounded-2xl border border-slate-300 overflow-hidden">
                      <div className="border-b border-slate-300 bg-slate-50 px-4 py-3">
                        <h4 className="text-lg font-semibold">{bucket.sa}</h4>
                        <p className="mt-1 text-sm text-slate-600">{bucket.items.length} flagged job(s)</p>
                      </div>
                      {/* Mobile card view */}
                      <div className="md:hidden divide-y divide-slate-200">
                        {bucket.items.map((item, index) => (
                          <div key={`${bucket.sa}-monthly-qc-m-${index}`} className="p-3 bg-white space-y-1 text-sm">
                            <p className="font-semibold text-blue-700">{toText(item.row['Job Number'])}</p>
                            {toText(getDateStartValue(item.row)) && <p><span className="font-semibold text-slate-500">Start:</span> {toText(getDateStartValue(item.row))}</p>}
                            {toText(getRepairApprovedDateValue(item.row)) && <p><span className="font-semibold text-slate-500">Repair Approved:</span> {toText(getRepairApprovedDateValue(item.row))}</p>}
                            {formatDate(getDateEndValue(item.row)) && <p><span className="font-semibold text-slate-500">End:</span> {formatDate(getDateEndValue(item.row))}</p>}
                            {toText(getQcNotCompletedValue(item.row)) && <p><span className="font-semibold text-slate-500">QC:</span> {toText(getQcNotCompletedValue(item.row))}</p>}
                            {item.issues.length > 0 && <p className="italic text-slate-700">{item.issues.join(', ')}</p>}
                          </div>
                        ))}
                      </div>
                      {/* Desktop table view */}
                      <div className="hidden sm:block overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead className="bg-white">
                            <tr className="border-b border-slate-300">
                              <th className="p-3 text-left font-semibold">Job Number</th>
                              <th className="p-3 text-left font-semibold">date_start</th>
                              <th className="p-3 text-left font-semibold">Repair Approved</th>
                              <th className="p-3 text-left font-semibold">date_end</th>
                              <th className="p-3 text-left font-semibold">QC Not Completed</th>
                              <th className="p-3 text-left font-semibold">Notes</th>
                            </tr>
                          </thead>
                          <tbody>
                            {bucket.items.map((item, index) => (
                              <tr key={`${bucket.sa}-monthly-qc-${index}`} className="border-b border-slate-200 align-top bg-white">
                                <td className="p-3 font-medium">{toText(item.row['Job Number'])}</td>
                                <td className="p-3">{toText(getDateStartValue(item.row))}</td>
                                <td className="p-3">{toText(getRepairApprovedDateValue(item.row))}</td>
                                <td className="p-3">{formatDate(getDateEndValue(item.row))}</td>
                                <td className="p-3">{toText(getQcNotCompletedValue(item.row))}</td>
                                <td className="p-3 italic text-slate-700">{item.issues.join(', ')}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </>
          ) : selectedAlert.detailType === 'must-return' ? (
            <>
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h3 className="text-xl font-semibold">{selectedAlert.title}</h3>
                  <p className="mt-1 text-slate-600">{mustReturnGroups.length} job(s) with parts to return</p>
                </div>
                <button onClick={() => setSelectedAlertId(null)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Clear</button>
              </div>
              {mustReturnGroups.length === 0 ? (
                <p className="mt-6 text-slate-600">No parts pending return at this time.</p>
              ) : (
                <div className="mt-6 rounded-2xl border border-slate-300 overflow-hidden">
                  {/* Mobile card view */}
                  <div className="md:hidden divide-y divide-slate-200">
                    {mustReturnGroups.map((group, i) => (
                      <div key={`must-return-m-${i}`} className="p-3 bg-white space-y-1 text-sm">
                        <p className="font-semibold text-blue-700">{group.jobNumber}</p>
                        <p><span className="font-semibold text-slate-500">Days:</span> <span className="font-bold text-red-600">{group.maxDays}</span></p>
                        <p><span className="font-semibold text-slate-500">Part(s):</span> {group.parts.join(', ')}</p>
                        <p><span className="font-semibold text-slate-500">Status:</span> {group.statusPriority}</p>
                      </div>
                    ))}
                  </div>
                  {/* Desktop table view */}
                  <table className="hidden md:table w-full text-sm">
                    <thead className="bg-slate-50">
                      <tr className="border-b border-slate-300">
                        <th className="p-3 text-left font-semibold">Job Number</th>
                        <th className="p-3 text-left font-semibold">Days Past Received</th>
                        <th className="p-3 text-left font-semibold">Part(s)</th>
                        <th className="p-3 text-left font-semibold">Status + Priority</th>
                      </tr>
                    </thead>
                    <tbody>
                      {mustReturnGroups.map((group, i) => (
                        <tr key={`must-return-${i}`} className="border-b border-slate-200 align-top bg-white">
                          <td className="p-3 font-semibold text-blue-700">{group.jobNumber}</td>
                          <td className="p-3 font-bold text-red-600">{group.maxDays}</td>
                          <td className="p-3">{group.parts.join(', ')}</td>
                          <td className="p-3">{group.statusPriority}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          ) : selectedAlert.detailType === 'missing-install' ? (
            <>
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h3 className="text-xl font-semibold">{selectedAlert.title}</h3>
                  <p className="mt-1 text-slate-600">{missingInstallGroups.length} job(s) with parts outstanding</p>
                </div>
                <button onClick={() => setSelectedAlertId(null)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Clear</button>
              </div>
              {missingInstallGroups.length === 0 ? (
                <p className="mt-6 text-slate-600">All parts accounted for. Nothing outstanding at this time.</p>
              ) : (
                <div className="mt-6 rounded-2xl border border-slate-300 overflow-hidden">
                  {/* Mobile card view */}
                  <div className="md:hidden divide-y divide-slate-200">
                    {missingInstallGroups.map((group, i) => (
                      <div key={`mi-m-${i}`} className="p-3 bg-white space-y-1 text-sm">
                        <p className="font-semibold text-blue-700">{group.jobNumber}</p>
                        <p><span className="font-semibold text-slate-500">Status:</span> {group.statusPriority}</p>
                        <ul className="space-y-0.5 mt-1">
                          {group.items.map((it, idx) => (
                            <li key={idx} className="flex items-center gap-1.5 text-xs">
                              <span className={it.received ? 'text-orange-500 font-bold' : 'text-red-600 font-bold'}>
                                {it.received ? '📦' : '⏳'}
                              </span>
                              <span>{it.name}</span>
                              <span className={`text-[10px] font-semibold ${it.received ? 'text-orange-500' : 'text-red-600'}`}>
                                {it.received ? 'Received — Not Checked Out' : 'Not Received'}
                              </span>
                            </li>
                          ))}
                        </ul>
                      </div>
                    ))}
                  </div>
                  {/* Desktop table view */}
                  <table className="hidden md:table w-full text-sm">
                    <thead className="bg-slate-50">
                      <tr className="border-b border-slate-300">
                        <th className="p-3 text-left font-semibold">Job Number</th>
                        <th className="p-3 text-left font-semibold">Status + Priority</th>
                        <th className="p-3 text-left font-semibold">Part(s) Outstanding</th>
                      </tr>
                    </thead>
                    <tbody>
                      {missingInstallGroups.map((group, i) => (
                        <tr key={`mi-${i}`} className="border-b border-slate-200 align-top bg-white">
                          <td className="p-3 font-semibold text-blue-700">{group.jobNumber}</td>
                          <td className="p-3">{group.statusPriority}</td>
                          <td className="p-3">
                            <ul className="space-y-0.5">
                              {group.items.map((it, idx) => (
                                <li key={idx} className="flex items-center gap-1.5 text-xs">
                                  <span className={it.received ? 'text-orange-500 font-bold' : 'text-red-600 font-bold'}>
                                    {it.received ? '📦' : '⏳'}
                                  </span>
                                  <span>{it.name}</span>
                                  <span className={`text-[10px] font-semibold ${it.received ? 'text-orange-500' : 'text-red-600'}`}>
                                    {it.received ? 'Received — Not Checked Out' : 'Not Received'}
                                  </span>
                                </li>
                              ))}
                            </ul>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          ) : selectedAlert.detailType === 'we-have-parts' ? (
            <>
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h3 className="text-xl font-semibold">{selectedAlert.title}</h3>
                  <p className="mt-1 text-slate-600">{weHavePartsMatches.length} vehicle(s) with matching stock</p>
                </div>
                <div className="flex items-center gap-3">
                  <button
                    onClick={() => handleCopyText(
                      weHavePartsMatches.map((m) => {
                        const partsStr = m.parts.map((p) => p.name).join(', ');
                        const fitsStr = Array.from(new Set(
                          m.parts.map((p) => [p.year, p.make, p.model].filter(Boolean).join(' ')).filter(Boolean)
                        )).join('; ');
                        return `${m.vehicleJobNumber} | ${m.statusPriority} | ${partsStr} | ${fitsStr}`;
                      }).join('\n')
                    )}
                    className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium flex items-center gap-2"
                  >
                    <span>📋</span><span>Copy Matches</span>
                  </button>
                  <button onClick={() => setSelectedAlertId(null)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Clear</button>
                </div>
              </div>
              {copyMessage && <p className="mt-3 text-sm text-slate-600">{copyMessage}</p>}
              {weHavePartsMatches.length === 0 ? (
                <p className="mt-6 text-slate-600">No stock parts currently match any vehicle in the lot.</p>
              ) : (
                <div className="mt-6 rounded-2xl border border-slate-300 overflow-hidden">
                  {/* Mobile card view */}
                  <div className="md:hidden divide-y divide-slate-200">
                    {weHavePartsMatches.map((m, i) => {
                      const fits = Array.from(new Set(
                        m.parts.map((p) => [p.year, p.make, p.model].filter(Boolean).join(' ')).filter(Boolean)
                      ));
                      return (
                        <div key={`wh-m-${i}`} className="p-3 bg-white space-y-1 text-sm">
                          <p className="font-semibold text-blue-700">{m.vehicleJobNumber}</p>
                          <p><span className="font-semibold text-slate-500">Status:</span> {m.statusPriority}</p>
                          <p><span className="font-semibold text-slate-500">Parts ({m.parts.length}):</span> {m.parts.map((p) => p.name).join(', ')}</p>
                          {fits.length > 0 && (
                            <p><span className="font-semibold text-slate-500">Fits:</span> {fits.join('; ')}</p>
                          )}
                        </div>
                      );
                    })}
                  </div>
                  {/* Desktop table view */}
                  <table className="hidden md:table w-full text-sm">
                    <thead className="bg-slate-50">
                      <tr className="border-b border-slate-300">
                        <th className="p-3 text-left font-semibold">Job Number</th>
                        <th className="p-3 text-left font-semibold">Status + Priority</th>
                        <th className="p-3 text-left font-semibold">Parts Available</th>
                        <th className="p-3 text-left font-semibold">Year / Make / Model</th>
                      </tr>
                    </thead>
                    <tbody>
                      {weHavePartsMatches.map((m, i) => {
                        const fits = Array.from(new Set(
                          m.parts.map((p) => [p.year, p.make, p.model].filter(Boolean).join(' ')).filter(Boolean)
                        ));
                        return (
                          <tr key={`wh-${i}`} className="border-b border-slate-200 align-top bg-white">
                            <td className="p-3 font-semibold text-blue-700">{m.vehicleJobNumber}</td>
                            <td className="p-3">{m.statusPriority}</td>
                            <td className="p-3">
                              <span className="text-xs font-semibold text-slate-500">({m.parts.length}) </span>
                              {m.parts.map((p) => p.name).join(', ')}
                            </td>
                            <td className="p-3 text-xs text-slate-600">{fits.join('; ')}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          ) : (
            <>
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h3 className="text-xl font-semibold">{selectedAlert.title}</h3>
                  <p className="mt-1 text-slate-600">{selectedAlert.count} matching job(s)</p>
                </div>
                <div className="flex items-center gap-3">
                  <button onClick={() => handleCopyRows(selectedAlert.rows)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium flex items-center gap-2">
                    <span>📋</span><span>Copy Jobs</span>
                  </button>
                  <button onClick={() => setSelectedAlertId(null)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Clear</button>
                </div>
              </div>
              {copyMessage && <p className="mt-3 text-sm text-slate-600">{copyMessage}</p>}
              <div className="mt-6 rounded-2xl border border-slate-300 overflow-hidden">
                {/* Desktop table view */}
                <table className="hidden md:table w-full text-sm">
                  <thead className="bg-slate-50">
                    <tr className="border-b border-slate-300">
                      <th className="p-3 text-left font-semibold">Job Number</th>
                      <th className="p-3 text-left font-semibold">Priority</th>
                      <th className="p-3 text-left font-semibold">Model</th>
                      <th className="p-3 text-left font-semibold">Status + Priority</th>
                      <th className="p-3 text-left font-semibold">Status Days</th>
                      {selectedAlert.id === 'general-parts' ? (
                        <th className="p-3 text-left font-semibold">Parts</th>
                      ) : selectedAlert.id === 'glass-install-after-delivery' ? (
                        <>
                          <th className="p-3 text-left font-semibold">Task Titles</th>
                          <th className="p-3 text-left font-semibold">Glass Parts</th>
                        </>
                      ) : (
                        <>
                          <th className="p-3 text-left font-semibold">Task Titles</th>
                          <th className="p-3 text-left font-semibold">Body ECD</th>
                        </>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {selectedAlert.rows.map((r, i) => {
                      const delayed = isRowDelayed(r);
                      const jobNum = toText(r['Job Number']);
                      const partsInfo = selectedAlert.id === 'general-parts'
                        ? getPartsInfoForJob(jobNum, partsData)
                        : null;
                      const glassParts = selectedAlert.id === 'glass-install-after-delivery'
                        ? getGlassPartsForJob(jobNum, partsData)
                        : null;
                      return (
                        <tr
                          key={`${selectedAlert.id}-row-${i}`}
                          className={`border-b border-slate-200 align-top ${delayed ? 'bg-red-50' : 'bg-white'}`}
                        >
                          <td className="p-3 font-medium">{jobNum}</td>
                          <td className="p-3 font-semibold">{toText(r['Priority'])}</td>
                          <td className="p-3">{toText(r['Model'])}</td>
                          <td className="p-3">{toText(r['Status + Priority'])}</td>
                          <td className={`p-3 ${delayed ? 'font-bold text-red-600' : ''}`}>
                            {toText(r['Status Days'])}
                          </td>
                          {partsInfo ? (
                            <td className="p-3 max-w-md">
                              {partsInfo.arrived.length > 0 && (
                                <div className="mb-1"><span className="font-bold">Arrived:</span> {partsInfo.arrived.join(', ')}</div>
                              )}
                              {partsInfo.missing.length > 0 && (
                                <div>
                                  <span className="font-bold">Missing:</span>
                                  <ul className="mt-0.5 space-y-0.5">
                                    {partsInfo.missing.map((m, mi) => (
                                      <li key={mi} className="text-xs">
                                        {m.name}
                                        {m.eta ? <span className="text-slate-500"> · ETA {m.eta}</span> : <span className="italic text-slate-400"> · no ETA</span>}
                                      </li>
                                    ))}
                                  </ul>
                                </div>
                              )}
                              {partsInfo.arrived.length === 0 && partsInfo.missing.length === 0 && (
                                <span className="text-slate-400 italic">No parts data</span>
                              )}
                            </td>
                          ) : glassParts ? (
                            <>
                              <td className="p-3 max-w-md">{toText(r['Task Titles'])}</td>
                              <td className="p-3 max-w-xs">
                                {glassParts.length === 0 ? (
                                  <span className="text-slate-400 italic">None found</span>
                                ) : (
                                  <ul className="space-y-0.5">
                                    {glassParts.map((gp, gi) => (
                                      <li key={gi} className="flex items-center gap-1.5 text-xs">
                                        <span className={gp.arrived ? 'text-green-600 font-bold' : 'text-orange-500 font-bold'}>
                                          {gp.arrived ? '✓' : '⏳'}
                                        </span>
                                        <span>{gp.name}</span>
                                        <span className={`text-[10px] font-semibold ${gp.arrived ? 'text-green-600' : 'text-orange-500'}`}>
                                          {gp.arrived ? 'Arrived' : 'Pending'}
                                        </span>
                                      </li>
                                    ))}
                                  </ul>
                                )}
                              </td>
                            </>
                          ) : (
                            <>
                              <td className="p-3 max-w-md">{toText(r['Task Titles'])}</td>
                              <td className="p-3">{formatDate(r['Body ECD'])}</td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Mobile card list */}
              <div className="mt-6 space-y-3 md:hidden">
                {selectedAlert.rows.map((r, i) => {
                  const delayed = isRowDelayed(r);
                  const jobNum = toText(r['Job Number']);
                  const partsInfo = selectedAlert.id === 'general-parts'
                    ? getPartsInfoForJob(jobNum, partsData)
                    : null;
                  const glassParts = selectedAlert.id === 'glass-install-after-delivery'
                    ? getGlassPartsForJob(jobNum, partsData)
                    : null;
                  return (
                    <div key={`m-${selectedAlert.id}-${i}`} className={`rounded-2xl border p-4 text-sm ${delayed ? 'bg-red-50 border-red-300' : 'bg-white border-slate-300'}`}>
                      <div className="flex items-center justify-between gap-3">
                        <span className="text-lg font-bold text-blue-700">{jobNum}</span>
                        <span className="text-xs font-bold uppercase tracking-wide">P: {toText(r['Priority'])}</span>
                      </div>
                      <p className="mt-1 text-slate-700">{toText(r['Model'])}</p>
                      <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-xs">
                        <span className="rounded-full bg-slate-100 px-2 py-0.5">{toText(r['Status + Priority'])}</span>
                        <span className={delayed ? 'font-bold text-red-600' : 'font-medium text-slate-700'}>Days: {toText(r['Status Days'])}</span>
                      </div>
                      {partsInfo ? (
                        <div className="mt-2 text-xs text-slate-700">
                          {partsInfo.arrived.length > 0 && (
                            <p><span className="font-semibold">Arrived:</span> {partsInfo.arrived.join(', ')}</p>
                          )}
                          {partsInfo.missing.length > 0 && (
                            <div className="mt-1">
                              <span className="font-semibold">Missing:</span>
                              <ul className="mt-0.5 ml-1 space-y-0.5">
                                {partsInfo.missing.map((m, mi) => (
                                  <li key={mi}>
                                    · {m.name}
                                    {m.eta ? <span className="text-slate-500"> — ETA {m.eta}</span> : <span className="italic text-slate-400"> — no ETA</span>}
                                  </li>
                                ))}
                              </ul>
                            </div>
                          )}
                          {partsInfo.arrived.length === 0 && partsInfo.missing.length === 0 && (
                            <p className="italic text-slate-400">No parts data</p>
                          )}
                        </div>
                      ) : glassParts ? (
                        <div className="mt-2 text-xs text-slate-700 space-y-1">
                          <p><span className="font-semibold">Task Titles:</span> {toText(r['Task Titles']) || <span className="italic text-slate-400">—</span>}</p>
                          <div>
                            <p className="font-semibold">Glass Parts:</p>
                            {glassParts.length === 0 ? (
                              <p className="italic text-slate-400">None found</p>
                            ) : (
                              <ul className="space-y-0.5 mt-1">
                                {glassParts.map((gp, gi) => (
                                  <li key={gi} className="flex items-center gap-1.5">
                                    <span className={gp.arrived ? 'text-green-600 font-bold' : 'text-orange-500 font-bold'}>
                                      {gp.arrived ? '✓' : '⏳'}
                                    </span>
                                    <span>{gp.name}</span>
                                    <span className={`text-[10px] font-semibold ${gp.arrived ? 'text-green-600' : 'text-orange-500'}`}>
                                      {gp.arrived ? 'Arrived' : 'Pending'}
                                    </span>
                                  </li>
                                ))}
                              </ul>
                            )}
                          </div>
                        </div>
                      ) : (
                        <div className="mt-2 text-xs text-slate-700 space-y-1">
                          <p><span className="font-semibold">Task Titles:</span> {toText(r['Task Titles']) || <span className="italic text-slate-400">—</span>}</p>
                          <p><span className="font-semibold">Body ECD:</span> {formatDate(r['Body ECD']) || <span className="italic text-slate-400">—</span>}</p>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </>
          )}
        </section>
      </div>

      {/* Main Stat Modal */}
      {selectedMain && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="w-full max-w-5xl rounded-3xl bg-white border border-slate-300 shadow-xl p-6 max-h-[85vh] overflow-auto">
            <div className="flex items-start justify-between gap-4">
              <div>
                <h3 className="text-2xl font-bold">{selectedMain.title}</h3>
                <p className="mt-1 text-slate-600">{selectedMain.count} total matching job(s)</p>
              </div>
              <div className="flex items-center gap-2">
                {selectedMain.modalType === 'delivered-hail' && (
                  <button
                    onClick={() => { setShowBonus((v) => !v); setBonusError(false); setBonusInput(''); }}
                    className="rounded-xl border border-slate-300 bg-slate-50 px-3 py-2 text-sm font-medium flex items-center gap-1.5"
                  >
                    <span>{bonusUnlocked ? '💰' : '🔒'}</span>
                    <span>Bonus</span>
                  </button>
                )}
                <button onClick={() => { setSelectedMainId(null); setSelectedSa(null); setBonusUnlocked(false); setShowBonus(false); setBonusInput(''); setBonusError(false); }} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Close</button>
              </div>
            </div>

            {/* Modal Content - Repair Approved Buckets */}
            {selectedMain.modalType === 'repair-approved-buckets' && repairBuckets ? (
              <div className="mt-8 space-y-12">
                {[
                  { name: 'Enterprise', rows: repairBuckets.enterprise },
                  { name: 'Progressive', rows: repairBuckets.progressive },
                  { name: 'General Retail', rows: repairBuckets.general }
                ].map((bucket) => (
                  <div key={bucket.name} className="space-y-4">
                    <div className="flex items-center justify-between">
                      <h4 className="text-xl font-bold text-slate-800">{bucket.name} ({bucket.rows.length})</h4>
                      <button onClick={() => handleCopyRows(bucket.rows)} className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs font-semibold flex items-center gap-2"><span>📋</span><span>Copy Bucket</span></button>
                    </div>
                    <div className="rounded-2xl border border-slate-300 overflow-hidden">
                      {/* Mobile card view */}
                      <div className="md:hidden divide-y divide-slate-200">
                        {bucket.rows.map((r, i) => {
                          const delayed = isRowDelayed(r);
                          return (
                            <div key={i} className={`p-3 space-y-1 text-sm ${delayed ? 'bg-red-50' : 'bg-white'}`}>
                              <p className="font-bold text-blue-700">{toText(r['Job Number'])}</p>
                              <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                                <p><span className="font-semibold text-slate-500">Priority:</span> <span className="font-bold">{toText(r['Priority'])}</span></p>
                                <p><span className="font-semibold text-slate-500">Model:</span> {toText(r['Model'])}</p>
                              </div>
                              <p><span className="font-semibold text-slate-500">Status:</span> {toText(r['Status + Priority'])}</p>
                              <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                                {toText(r['Severity']) && <p><span className="font-semibold text-slate-500">Severity:</span> {toText(r['Severity'])}</p>}
                                {toText(r['Insurance']) && <p><span className="font-semibold text-slate-500">Insurance:</span> {toText(r['Insurance'])}</p>}
                                <p><span className="font-semibold text-slate-500">Days:</span> <span className={delayed ? 'font-bold text-red-600' : 'font-medium'}>{toText(r['Status Days'])}</span></p>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                      {/* Desktop table view */}
                      <table className="hidden md:table w-full text-sm text-left">
                        <thead className="bg-slate-50">
                          <tr className="border-b border-slate-300">
                            <th className="p-3 font-semibold">Job Number</th>
                            <th className="p-3 font-semibold">Priority</th>
                            <th className="p-3 font-semibold">Model</th>
                            <th className="p-3 font-semibold">Status + Priority</th>
                            <th className="p-3 font-semibold">Severity</th>
                            <th className="p-3 font-semibold">Insurance</th>
                            <th className="p-3 font-semibold">Status Days</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bucket.rows.map((r, i) => {
                             const delayed = isRowDelayed(r);
                             return (
                              <tr key={i} className={`border-b border-slate-200 ${delayed ? 'bg-red-50' : 'bg-white'}`}>
                                <td className="p-3 font-bold text-blue-700">{toText(r['Job Number'])}</td>
                                <td className="p-3 font-bold">{toText(r['Priority'])}</td>
                                <td className="p-3">{toText(r['Model'])}</td>
                                <td className="p-3">{toText(r['Status + Priority'])}</td>
                                <td className="p-3">{toText(r['Severity'])}</td>
                                <td className="p-3">{toText(r['Insurance'])}</td>
                                <td className={`p-3 ${delayed ? 'font-bold text-red-600' : 'font-medium'}`}>{toText(r['Status Days'])}</td>
                              </tr>
                             );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            ) : selectedMain.modalType === 'sa-chart' ? (
              <div className="mt-8 space-y-6">
                <h4 className="text-lg font-semibold">Jobs by SA</h4>
                <div className="space-y-4">
                  {saCounts.map((item) => (
                    <button key={item.sa} onClick={() => setSelectedSa(item.sa)} className={`w-full rounded-2xl border p-4 text-left transition ${selectedSa === item.sa ? 'border-slate-900 bg-slate-50' : 'border-slate-200 bg-white'}`}>
                      <div className="flex items-center justify-between text-sm font-medium">
                        <span>{item.sa}</span><span>{item.count}</span>
                      </div>
                      <div className="mt-3 h-5 rounded-full bg-slate-200 overflow-hidden">
                        <div className="h-full bg-blue-500" style={{ width: `${(item.count / maxSaCount) * 100}%` }} />
                      </div>
                    </button>
                  ))}
                </div>
                {selectedSa && (
                  <div className="space-y-4 mt-6">
                    <div className="flex items-center justify-between">
                      <h4 className="text-lg font-semibold">{selectedSa}</h4>
                      <button onClick={() => handleCopyRows(selectedSaRows)} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium flex items-center gap-2"><span>📋</span><span>Copy Jobs</span></button>
                    </div>
                    <div className="rounded-2xl border border-slate-300 overflow-hidden">
                      {/* Mobile card view */}
                      <div className="md:hidden divide-y divide-slate-200">
                        {selectedSaRows.map((r, i) => (
                          <div key={`sa-card-${i}`} className="p-3 space-y-1 text-sm bg-white">
                            <p className="font-medium text-blue-700">{toText(r['Job Number'])}</p>
                            <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                              <p><span className="font-semibold text-slate-500">Priority:</span> <span className="font-semibold">{toText(r['Priority'])}</span></p>
                              <p><span className="font-semibold text-slate-500">Model:</span> {toText(r['Model'])}</p>
                            </div>
                            <p><span className="font-semibold text-slate-500">Status:</span> {toText(r['Status + Priority'])}</p>
                          </div>
                        ))}
                      </div>
                      {/* Desktop table view */}
                      <table className="hidden md:table w-full text-sm">
                        <thead className="bg-slate-50">
                          <tr className="border-b border-slate-300">
                            <th className="p-3 text-left font-semibold">Job Number</th>
                            <th className="p-3 text-left font-semibold">Priority</th>
                            <th className="p-3 text-left font-semibold">Model</th>
                            <th className="p-3 text-left font-semibold">Status + Priority</th>
                          </tr>
                        </thead>
                        <tbody>
                          {selectedSaRows.map((r, i) => (
                            <tr key={`sa-row-${i}`} className="border-b border-slate-200">
                              <td className="p-3 font-medium">{toText(r['Job Number'])}</td>
                              <td className="p-3 font-semibold">{toText(r['Priority'])}</td>
                              <td className="p-3">{toText(r['Model'])}</td>
                              <td className="p-3">{toText(r['Status + Priority'])}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            ) : selectedMain.modalType === 'delivered-hail' ? (
              <div className="mt-8 space-y-6">

                {/* Bonus Panel */}
                {showBonus && (
                  <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                    {!bonusUnlocked ? (
                      <div className="flex flex-col gap-3">
                        <p className="text-sm font-semibold text-slate-700">Enter password to view bonus</p>
                        <div className="flex gap-2">
                          <input
                            type="password"
                            value={bonusInput}
                            onChange={(e) => { setBonusInput(e.target.value); setBonusError(false); }}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                if (bonusInput === 'Gf080417') { setBonusUnlocked(true); setBonusError(false); }
                                else { setBonusError(true); setBonusInput(''); }
                              }
                            }}
                            placeholder="Password"
                            className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-400 w-44"
                            autoFocus
                          />
                          <button
                            onClick={() => {
                              if (bonusInput === 'Gf080417') { setBonusUnlocked(true); setBonusError(false); }
                              else { setBonusError(true); setBonusInput(''); }
                            }}
                            className="rounded-xl bg-slate-800 text-white px-4 py-2 text-sm font-medium"
                          >
                            Unlock
                          </button>
                        </div>
                        {bonusError && <p className="text-xs text-red-600 font-medium">Incorrect password.</p>}
                      </div>
                    ) : (
                      <div className="flex flex-wrap items-center gap-4">
                        <p className="text-sm font-bold text-slate-700 mr-2">💰 Bonus Breakdown</p>
                        <div className="flex gap-3 flex-wrap">
                          <div className="rounded-xl border border-green-200 bg-green-50 px-4 py-2 text-center">
                            <p className="text-[10px] uppercase font-bold text-green-700 tracking-wider">Approval Bonus</p>
                            <p className="text-lg font-bold text-green-800">${deliveredHailStats.approvalBonusTotal}</p>
                          </div>
                          <div className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-2 text-center">
                            <p className="text-[10px] uppercase font-bold text-blue-700 tracking-wider">Repair Bonus</p>
                            <p className="text-lg font-bold text-blue-800">${deliveredHailStats.repairBonusTotal}</p>
                          </div>
                          <div className="rounded-xl border border-slate-300 bg-white px-4 py-2 text-center">
                            <p className="text-[10px] uppercase font-bold text-slate-600 tracking-wider">Total</p>
                            <p className="text-lg font-bold text-slate-900">${deliveredHailStats.combinedBonusTotal}</p>
                          </div>
                        </div>
                        <button onClick={() => { setBonusUnlocked(false); setBonusInput(''); setShowBonus(false); }} className="ml-auto text-xs text-slate-400 hover:text-slate-600">🔒 Lock</button>
                      </div>
                    )}
                  </div>
                )}

                <div className="grid grid-cols-1 md:grid-cols-4 lg:grid-cols-7 gap-4">
                  {[
                    { label: 'Units', val: deliveredHailStats.deliveredCount },
                    { label: 'Avg Appr', val: formatDays(deliveredHailStats.avgApprovalTime) },
                    { label: 'Avg PDR/2', val: formatDays(deliveredHailStats.avgApprovedPendingPdr) },
                    { label: 'Avg Repair', val: formatDays(deliveredHailStats.avgRepairTime) },
                    { label: 'Avg Deliv', val: formatDays(deliveredHailStats.avgDeliveryTime) },
                  ].map((stat, i) => (
                    <div key={i} className="rounded-2xl border border-slate-200 bg-slate-50 p-4">
                      <p className="text-[10px] uppercase font-bold text-slate-500 tracking-wider">{stat.label}</p>
                      <p className="mt-1 text-xl font-bold">{stat.val}</p>
                    </div>
                  ))}
                </div>
                <div className="rounded-2xl border border-slate-300 overflow-hidden">
                  {/* Mobile card view */}
                  <div className="md:hidden divide-y divide-slate-200">
                    {deliveredHailStats.deliveredRows.map((r, i) => (
                      <div key={i} className="p-3 space-y-1 text-sm bg-white">
                        <p className="font-medium">{toText(r['Job Number'])}</p>
                        <p><span className="font-semibold text-slate-500">Model:</span> {toText(r['Model'])}</p>
                        <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                          {toText(r['Approval Time']) && <p><span className="font-semibold text-slate-500">Appr:</span> {toText(r['Approval Time'])}</p>}
                          <p><span className="font-semibold text-slate-500">PDR/2:</span> {toNumber(r['Approved Pending PDR']) / 2}</p>
                          {toText(r['Repair Time']) && <p><span className="font-semibold text-slate-500">Repair:</span> {toText(r['Repair Time'])}</p>}
                          {toText(r['Delivery Time']) && <p><span className="font-semibold text-slate-500">Deliv:</span> {toText(r['Delivery Time'])}</p>}
                        </div>
                        <p><span className="font-semibold text-blue-700">Appr to Deliv:</span> <span className="font-bold text-blue-700">{(toNumber(r['Approved Pending PDR']) / 2) + toNumber(r['Repair Time']) + toNumber(r['Delivery Time'])}</span></p>
                      </div>
                    ))}
                  </div>
                  {/* Desktop table view */}
                  <table className="hidden md:table w-full text-sm text-left">
                    <thead className="bg-slate-50">
                      <tr className="border-b border-slate-300">
                        <th className="p-3 font-semibold">Job Number</th>
                        <th className="p-3 font-semibold">Model</th>
                        <th className="p-3 font-semibold">Appr Time</th>
                        <th className="p-3 font-semibold">PDR Wait/2</th>
                        <th className="p-3 font-semibold">Repair</th>
                        <th className="p-3 font-semibold">Deliv</th>
                        <th className="p-3 font-semibold text-blue-700">Appr to Deliv</th>
                      </tr>
                    </thead>
                    <tbody>
                      {deliveredHailStats.deliveredRows.map((r, i) => (
                        <tr key={i} className="border-b border-slate-200">
                          <td className="p-3 font-medium">{toText(r['Job Number'])}</td>
                          <td className="p-3">{toText(r['Model'])}</td>
                          <td className="p-3">{toText(r['Approval Time'])}</td>
                          <td className="p-3">{toNumber(r['Approved Pending PDR']) / 2}</td>
                          <td className="p-3">{toText(r['Repair Time'])}</td>
                          <td className="p-3">{toText(r['Delivery Time'])}</td>
                          <td className="p-3 font-bold text-blue-700">
                            {(toNumber(r['Approved Pending PDR']) / 2) + toNumber(r['Repair Time']) + toNumber(r['Delivery Time'])}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div className="mt-8 rounded-2xl border border-slate-300 overflow-hidden">
                {/* Mobile card view */}
                <div className="md:hidden divide-y divide-slate-200">
                  {selectedMain.rows.map((r, i) => {
                    const delayed = isRowDelayed(r);
                    return (
                      <div key={i} className={`p-3 space-y-1 text-sm ${delayed ? 'bg-red-50' : 'bg-white'}`}>
                        <p className="font-bold text-blue-700">{toText(r['Job Number'])}</p>
                        <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                          <p><span className="font-semibold text-slate-500">Priority:</span> <span className="font-bold">{toText(r['Priority'])}</span></p>
                          <p><span className="font-semibold text-slate-500">Model:</span> {toText(r['Model'])}</p>
                        </div>
                        <p><span className="font-semibold text-slate-500">Status:</span> {toText(r['Status + Priority'])}</p>
                        <div className="flex flex-wrap gap-x-4 gap-y-0.5">
                          <p><span className="font-semibold text-slate-500">Days:</span> <span className={delayed ? 'font-bold text-red-600' : 'font-medium'}>{toText(r['Status Days'])}</span></p>
                          {toText(r['SA']) && <p><span className="font-semibold text-slate-500">SA:</span> {toText(r['SA'])}</p>}
                        </div>
                        {selectedMain.id === 'post-repair-main' && (
                          <p><span className="font-semibold text-slate-500">Task Titles:</span> {toText(r['Task Titles']) || <span className="italic text-slate-400">—</span>}</p>
                        )}
                        {selectedMain.id === 'conventional-hail-main' && (
                          <p><span className="font-semibold text-slate-500">Body ECD:</span> {toText(r['Body ECD']) || <span className="italic text-slate-400">—</span>}</p>
                        )}
                        {selectedMain.id === 'ready-to-deliver-main' && (
                          <p><span className="font-semibold text-slate-500">date_end:</span> {toText(getDateEndValue(r)) || <span className="italic text-slate-400">—</span>}</p>
                        )}
                      </div>
                    );
                  })}
                </div>
                {/* Desktop table view */}
                <table className="hidden md:table w-full text-sm text-left">
                  <thead className="bg-slate-50">
                    <tr className="border-b border-slate-300">
                      <th className="p-3 font-semibold">Job Number</th>
                      <th className="p-3 font-semibold">Priority</th>
                      <th className="p-3 font-semibold">Model</th>
                      <th className="p-3 font-semibold">Status + Priority</th>
                      <th className="p-3 font-semibold">Status Days</th>
                      <th className="p-3 font-semibold">SA</th>
                      {selectedMain.id === 'post-repair-main' && <th className="p-3 font-semibold">Task Titles</th>}
                      {selectedMain.id === 'conventional-hail-main' && <th className="p-3 font-semibold">Body ECD</th>}
                      {selectedMain.id === 'ready-to-deliver-main' && <th className="p-3 font-semibold">date_end</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {selectedMain.rows.map((r, i) => {
                      const delayed = isRowDelayed(r);
                      return (
                        <tr key={i} className={`border-b border-slate-200 align-top ${delayed ? 'bg-red-50' : 'bg-white'}`}>
                          <td className="p-3 font-bold text-blue-700">{toText(r['Job Number'])}</td>
                          <td className="p-3 font-bold">{toText(r['Priority'])}</td>
                          <td className="p-3">{toText(r['Model'])}</td>
                          <td className="p-3">{toText(r['Status + Priority'])}</td>
                          <td className={`p-3 ${delayed ? 'font-bold text-red-600' : 'font-medium'}`}>{toText(r['Status Days'])}</td>
                          <td className="p-3">{toText(r['SA'])}</td>
                          {selectedMain.id === 'post-repair-main' && <td className="p-3 max-w-md">{toText(r['Task Titles'])}</td>}
                          {selectedMain.id === 'conventional-hail-main' && <td className="p-3">{formatDate(r['Body ECD'])}</td>}
                          {selectedMain.id === 'ready-to-deliver-main' && <td className="p-3">{formatDate(getDateEndValue(r))}</td>}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>
      )}
    </main>
  );
}
