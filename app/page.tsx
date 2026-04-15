'use client';

import { useEffect, useMemo, useRef, useState } from 'react';

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
  detailType?: 'default' | 'sa-monthly-qc' | 'must-return';
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

function daysSince(date: Date): number {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.floor((today.getTime() - date.getTime()) / (1000 * 60 * 60 * 24));
}

function getPartsInfoForJob(
  jobNumber: string,
  partsData: PartsRow[]
): { arrived: string[]; missing: string[] } {
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
    .map((p) => getPartName(p))
    .filter(Boolean);
  return { arrived, missing };
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
    if (normalize(jn) === '000') continue;
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
      info: 'This alert appears when Order Parts or Parts Received still exists in Task Titles while the job is in Post Repair, PDR In-Progress, E - EHI Repair, or Repair Approved with 3 or more status days.',
      section: 'Parts',
    },
    {
      id: 'glass-parts',
      title: 'Glass Parts Incomplete',
      count: glassParts.length,
      rows: sortByPriority(glassParts),
      description: 'Glass tasks still active',
      info: 'This alert appears when Glass Order, Order Windshield, or Glass Received still exists in Task Titles while the job is in Post Repair, PDR In-Progress, E - EHI Repair, or Repair Approved with 3 or more status days.',
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
  const [showBonus, setShowBonus] = useState(false);
  const [bonusUnlocked, setBonusUnlocked] = useState(false);
  const [bonusInput, setBonusInput] = useState('');
  const [bonusError, setBonusError] = useState(false);

  useEffect(() => {
    fetch('/api/dashboard')
      .then((r) => r.json())
      .then(setData);
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

  const alertCards = useMemo(() => buildAlertCards(rows), [rows]);
  const mainCards = useMemo(() => buildMainCards(rows), [rows]);

  const mustReturnGroups = useMemo(() => buildMustReturnGroups(partsData, rows), [partsData, rows]);

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
  ], [mustReturnGroups]);

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
          <h1 className="text-3xl font-bold">Operations Manager Dashboard</h1>
          <p className="mt-2 text-sm text-slate-600">Last pulled: {formattedLastPulled}</p>
          {/* TEMP DEBUG */}
          <details className="mt-3">
            <summary className="cursor-pointer text-xs text-slate-400 hover:text-slate-600 select-none">Parts debug info</summary>
            <div className="mt-2 rounded-lg border border-slate-200 bg-slate-50 p-3 text-[11px] font-mono text-slate-700 space-y-1">
              <p>API keys: <strong>{Object.keys(data).join(', ') || '(loading...)'}</strong></p>
              <p>partsRows: <strong>{Array.isArray(data.partsRows) ? `array(${data.partsRows.length})` : String(typeof data.partsRows)}</strong></p>
              <p>partsHeaders: <strong>{Array.isArray(data.partsHeaders) ? `[${(data.partsHeaders as string[]).join(', ')}]` : String(typeof data.partsHeaders)}</strong></p>
              <p>partsData length: <strong>{partsData.length}</strong></p>
              {Array.isArray(data.partsRows) && data.partsRows.length > 0 && (
                <p>partsRows[0]: <strong>{JSON.stringify(data.partsRows[0]).slice(0, 300)}</strong></p>
              )}
              {partsData.length > 0 && rows.length > 0 && (
                <p>Job match test — dashboard[0]: <strong>&quot;{String(rows[0]['Job Number'])}&quot;</strong> vs parts[0] job: <strong>&quot;{String((partsData[0] as Record<string,unknown>)['Job'] ?? (partsData[0] as Record<string,unknown>)['Job Number'] ?? '(none)')}&quot;</strong></p>
              )}
            </div>
          </details>
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
              </button>
            ))}
          </div>
        </section>

        {/* Alert Sections */}
        {(['Operations', 'Parts', 'Conventional'] as const).map((section) => (
          <section key={section} className="space-y-4">
            <h2 className="text-2xl font-semibold">{section}</h2>
            <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3">
              {grouped[section].map((alert) => (
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
            {section === 'Parts' && (
              <div className="space-y-3 pt-2">
                <h3 className="text-lg font-semibold text-slate-600">Inventory Control</h3>
                <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3">
                  {grouped.Inventory.map((alert) => (
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
        ))}

        {/* Selected Alert Details Table */}
        <section className="rounded-3xl bg-white border border-slate-300 p-6 shadow-sm">
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
                      <div className="overflow-x-auto">
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
                                <td className="p-3">{toText(getDateEndValue(item.row))}</td>
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
                <div className="mt-6 overflow-x-auto rounded-2xl border border-slate-300">
                  <table className="w-full text-sm">
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
              <div className="mt-6 overflow-x-auto rounded-2xl border border-slate-300">
                <table className="w-full text-sm">
                  <thead className="bg-slate-50">
                    <tr className="border-b border-slate-300">
                      <th className="p-3 text-left font-semibold">Job Number</th>
                      <th className="p-3 text-left font-semibold">Priority</th>
                      <th className="p-3 text-left font-semibold">Model</th>
                      <th className="p-3 text-left font-semibold">Status + Priority</th>
                      <th className="p-3 text-left font-semibold">Status Days</th>
                      {selectedAlert.id === 'general-parts' ? (
                        <th className="p-3 text-left font-semibold">Parts</th>
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
                                <span><span className="font-bold">Arrived:</span> {partsInfo.arrived.join(', ')}</span>
                              )}
                              {partsInfo.arrived.length > 0 && partsInfo.missing.length > 0 && (
                                <span className="mx-1">·</span>
                              )}
                              {partsInfo.missing.length > 0 && (
                                <span><span className="font-bold">Missing:</span> {partsInfo.missing.join(', ')}</span>
                              )}
                              {partsInfo.arrived.length === 0 && partsInfo.missing.length === 0 && (
                                <span className="text-slate-400 italic">No parts data</span>
                              )}
                            </td>
                          ) : (
                            <>
                              <td className="p-3 max-w-md">{toText(r['Task Titles'])}</td>
                              <td className="p-3">{toText(r['Body ECD'])}</td>
                            </>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
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
                    <div className="overflow-x-auto rounded-2xl border border-slate-300">
                      <table className="w-full text-sm text-left">
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
                    <div className="overflow-x-auto rounded-2xl border border-slate-300">
                      <table className="w-full text-sm">
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
                <div className="overflow-x-auto rounded-2xl border border-slate-300">
                  <table className="w-full text-sm text-left">
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
              <div className="mt-8 overflow-x-auto rounded-2xl border border-slate-300">
                <table className="w-full text-sm text-left">
                  <thead className="bg-slate-50">
                    <tr className="border-b border-slate-300">
                      <th className="p-3 font-semibold">Job Number</th>
                      <th className="p-3 font-semibold">Priority</th>
                      <th className="p-3 font-semibold">Model</th>
                      <th className="p-3 font-semibold">Status + Priority</th>
                      <th className="p-3 font-semibold">Status Days</th>
                      <th className="p-3 font-semibold">SA</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedMain.rows.map((r, i) => {
                      const delayed = isRowDelayed(r);
                      return (
                        <tr key={i} className={`border-b border-slate-200 ${delayed ? 'bg-red-50' : 'bg-white'}`}>
                          <td className="p-3 font-bold text-blue-700">{toText(r['Job Number'])}</td>
                          <td className="p-3 font-bold">{toText(r['Priority'])}</td>
                          <td className="p-3">{toText(r['Model'])}</td>
                          <td className="p-3">{toText(r['Status + Priority'])}</td>
                          <td className={`p-3 ${delayed ? 'font-bold text-red-600' : 'font-medium'}`}>{toText(r['Status Days'])}</td>
                          <td className="p-3">{toText(r['SA'])}</td>
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
