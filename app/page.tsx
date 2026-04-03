'use client';

import { useEffect, useMemo, useState } from 'react';

type DashboardRow = Record<string, string | number | null | undefined>;

type AlertCard = {
  id: string;
  title: string;
  count: number;
  rows: DashboardRow[];
  description: string;
  info: string;
  section: 'Operations' | 'Parts' | 'Conventional';
};

type MainCard = {
  id: string;
  title: string;
  count: number;
  rows: DashboardRow[];
  modalType: 'sa-chart' | 'job-list' | 'delivered-hail' | 'repair-approved-buckets';
  isDelayed?: boolean;
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
  if (isPostRepair(row)) return days >= 3;
  if (isReadyToDeliver(row)) return days > 2;
  return false;
}

// --- Logic ---

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

function isConventional(row: DashboardRow): boolean {
  return normalize(row['Status + Priority']) === 'conventional';
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
    return (filter === 'ana' || filter === 'roy' || filter === 'ana/roy') && isBlank(r['Severity']);
  });

  const escalationOnSite = rows.filter((r) => {
    return normalize(r['Status + Priority']) === 'escalation on-site';
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

  const conventionalMissing = rows.filter(
    (r) => isConventional(r) && isBlank(r['Body ECD'])
  );

  const conventionalPastDue = rows.filter(
    (r) => isConventional(r) && !isBlank(r['Body ECD']) && isPastDue(r['Body ECD'])
  );

  return [
    {
      id: 'needs-severity',
      title: 'Needs Severity',
      count: needsSeverity.length,
      rows: sortByPriority(needsSeverity),
      description: 'Ana / Roy filters missing severity',
      info: 'This alert appears when the Filter column is Ana, Roy, or Ana/Roy and the Severity field is blank.',
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
      id: 'conv-missing',
      title: 'Missing Body ECD',
      count: conventionalMissing.length,
      rows: sortByPriority(conventionalMissing),
      description: 'Conventional without Body ECD',
      info: 'This alert appears when Status + Priority is Conventional and Body ECD is blank.',
      section: 'Conventional',
    },
    {
      id: 'conv-past-due',
      title: 'Past Due Body ECD',
      count: conventionalPastDue.length,
      rows: sortByPriority(conventionalPastDue),
      description: 'Conventional past due',
      info: 'This alert appears when Status + Priority is Conventional and Body ECD is already past due.',
      section: 'Conventional',
    },
  ];
}

function buildMainCards(rows: DashboardRow[]): MainCard[] {
  const insuranceRows = sortByPriority(rows.filter(isInsuranceApproval));
  const repairApprovedRows = sortByPriority(rows.filter(isRepairApproved));
  const pdrRows = sortByPriority(rows.filter(isPdrInProgress));
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
  // Logic for Escalation On-site: green if 0, red if > 0
  if (alertId === 'escalation-onsite') {
    return count > 0 ? 'bg-red-100 border-red-400' : 'bg-green-100 border-green-400';
  }

  // Standard Logic for other cards
  if (count >= 5) return 'bg-red-100 border-red-400';
  if (count > 0) return 'bg-yellow-100 border-yellow-400';
  return 'bg-green-100 border-green-400';
}

function getSaCounts(rows: DashboardRow[]) {
  const counts = new Map<string, number>();
  for (const row of rows) {
    const sa = toText(row['SA']) || 'Unassigned';
    counts.set(sa, (counts.get(sa) || 0) + 1);
  }
  return [...counts.entries()]
    .map(([sa, count]) => ({ sa, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 5);
}

export default function Page() {
  const [data, setData] = useState<{ updatedAt?: string; rows?: DashboardRow[] }>({});
  const [selectedAlertId, setSelectedAlertId] = useState<string | null>(null);
  const [selectedMainId, setSelectedMainId] = useState<string | null>(null);
  const [selectedSa, setSelectedSa] = useState<string | null>(null);
  const [infoAlertId, setInfoAlertId] = useState<string | null>(null);
  const [copyMessage, setCopyMessage] = useState('');

  useEffect(() => {
    fetch('/api/dashboard')
      .then((r) => r.json())
      .then(setData);
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
    } catch (e) {
      return data.updatedAt;
    }
  }, [data.updatedAt]);

  const rows = data.rows ?? [];
  const alertCards = useMemo(() => buildAlertCards(rows), [rows]);
  const mainCards = useMemo(() => buildMainCards(rows), [rows]);

  const grouped = useMemo(() => ({
    Operations: alertCards.filter((a) => a.section === 'Operations'),
    Parts: alertCards.filter((a) => a.section === 'Parts'),
    Conventional: alertCards.filter((a) => a.section === 'Conventional'),
  }), [alertCards]);

  const selectedAlert = alertCards.find((a) => a.id === selectedAlertId) ?? null;
  const selectedMain = mainCards.find((m) => m.id === selectedMainId) ?? null;
  const saCounts = useMemo(() => getSaCounts(rows), [rows]);
  const maxSaCount = saCounts.length > 0 ? Math.max(...saCounts.map((x) => x.count)) : 1;

  const selectedSaRows = useMemo(() => {
    if (!selectedSa) return [];
    return sortByPriority(rows.filter((row) => (toText(row['SA']) || 'Unassigned') === selectedSa));
  }, [rows, selectedSa]);

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

  async function handleCopyRows(rowsToCopy: DashboardRow[]) {
    if (rowsToCopy.length === 0) return;
    try {
      await navigator.clipboard.writeText(buildClipboardText(rowsToCopy));
      setCopyMessage('Copied');
      window.setTimeout(() => setCopyMessage(''), 2000);
    } catch {
      setCopyMessage('Copy failed');
      window.setTimeout(() => setCopyMessage(''), 2000);
    }
  }

  return (
    <main className="min-h-screen bg-slate-100 text-slate-900 p-6 md:p-8">
      <div className="mx-auto max-w-7xl space-y-8">
        <section className="rounded-3xl bg-white border border-slate-300 p-6 shadow-sm">
          <h1 className="text-3xl font-bold">Operations Manager Dashboard</h1>
          <p className="mt-2 text-sm text-slate-600">Last pulled: {formattedLastPulled}</p>
        </section>

        <section className="space-y-4">
          <h2 className="text-2xl font-semibold">Main Information</h2>
          <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-8 gap-4">
            {mainCards.map((card) => (
              <button
                key={card.id}
                type="button"
                onClick={() => { setSelectedMainId(card.id); setSelectedSa(null); }}
                className={`rounded-2xl border p-4 text-left shadow-sm transition hover:shadow-md ${
                  card.isDelayed ? 'bg-red-100 border-red-300' : 'bg-white border-slate-300'
                } ${selectedMainId === card.id ? 'ring-2 ring-slate-900' : ''}`}
              >
                <p className="text-xs font-semibold uppercase tracking-wide text-slate-700">{card.title}</p>
                <p className="mt-2 text-3xl font-bold text-slate-900">{card.count}</p>
              </button>
            ))}
          </div>
        </section>

        {Object.entries(grouped).map(([section, items]) => (
          <section key={section} className="space-y-4">
            <h2 className="text-2xl font-semibold">{section}</h2>
            <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-3">
              {items.map((alert) => (
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
          </section>
        ))}

        <section className="rounded-3xl bg-white border border-slate-300 p-6 shadow-sm">
          {!selectedAlert ? (
            <>
              <h3 className="text-xl font-semibold">Alert Details</h3>
              <p className="mt-3 text-slate-600">Select an alert card to see the matching jobs.</p>
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
                      <th className="p-3 text-left font-semibold">Task Titles</th>
                      <th className="p-3 text-left font-semibold">Body ECD</th>
                    </tr>
                  </thead>
                  <tbody>
                    {selectedAlert.rows.map((r, i) => (
                      <tr key={`${selectedAlert.id}-row-${i}`} className="border-b border-slate-200 align-top">
                        <td className="p-3 font-medium">{toText(r['Job Number'])}</td>
                        <td className="p-3 font-semibold">{toText(r['Priority'])}</td>
                        <td className="p-3">{toText(r['Model'])}</td>
                        <td className="p-3">{toText(r['Status + Priority'])}</td>
                        <td className="p-3">{toText(r['Status Days'])}</td>
                        <td className="p-3 max-w-md">{toText(r['Task Titles'])}</td>
                        <td className="p-3">{toText(r['Body ECD'])}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}
        </section>
      </div>

      {selectedMain && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
          <div className="w-full max-w-5xl rounded-3xl bg-white border border-slate-300 shadow-xl p-6 max-h-[85vh] overflow-auto">
            <div className="flex items-start justify-between gap-4">
              <div>
                <h3 className="text-2xl font-bold">{selectedMain.title}</h3>
                <p className="mt-1 text-slate-600">{selectedMain.count} total matching job(s)</p>
              </div>
              <button onClick={() => { setSelectedMainId(null); setSelectedSa(null); }} className="rounded-xl border border-slate-300 bg-slate-50 px-4 py-2 text-sm font-medium">Close</button>
            </div>

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
                            <th className="p-3 font-semibold">Insurance</th>
                            <th className="p-3 font-semibold">Status Days</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bucket.rows.map((r, i) => (
                            <tr key={i} className="border-b border-slate-200">
                              <td className="p-3 font-bold text-blue-700">{toText(r['Job Number'])}</td>
                              <td className="p-3 font-bold">{toText(r['Priority'])}</td>
                              <td className="p-3">{toText(r['Model'])}</td>
                              <td className="p-3">{toText(r['Status + Priority'])}</td>
                              <td className="p-3">{toText(r['Insurance'])}</td>
                              <td className="p-3 font-medium">{toText(r['Status Days'])}</td>
                            </tr>
                          ))}
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
                <div className="grid grid-cols-1 md:grid-cols-4 lg:grid-cols-7 gap-4">
                  {[
                    { label: 'Units', val: deliveredHailStats.deliveredCount },
                    { label: 'Avg Appr', val: formatDays(deliveredHailStats.avgApprovalTime) },
                    { label: 'Appr Bonus', val: `$${deliveredHailStats.approvalBonusTotal}` },
                    { label: 'Avg PDR/2', val: formatDays(deliveredHailStats.avgApprovedPendingPdr) },
                    { label: 'Avg Repair', val: formatDays(deliveredHailStats.avgRepairTime) },
                    { label: 'Avg Deliv', val: formatDays(deliveredHailStats.avgDeliveryTime) },
                    { label: 'Repair Bonus', val: `$${deliveredHailStats.repairBonusTotal}` },
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
                        <th className="p-3 font-semibold">Approval</th>
                        <th className="p-3 font-semibold">Repair</th>
                        <th className="p-3 font-semibold">Delivery</th>
                      </tr>
                    </thead>
                    <tbody>
                      {deliveredHailStats.deliveredRows.map((r, i) => (
                        <tr key={i} className="border-b border-slate-200">
                          <td className="p-3 font-medium">{toText(r['Job Number'])}</td>
                          <td className="p-3">{toText(r['Approval Time'])}</td>
                          <td className="p-3">{toText(r['Repair Time'])}</td>
                          <td className="p-3">{toText(r['Delivery Time'])}</td>
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
                    </tr>
                  </thead>
                  <tbody>
                    {selectedMain.rows.map((r, i) => (
                      <tr key={i} className={`border-b border-slate-200 ${isRowDelayed(r) ? 'bg-red-50' : ''}`}>
                        <td className="p-3 font-medium">{toText(r['Job Number'])}</td>
                        <td className="p-3 font-semibold">{toText(r['Priority'])}</td>
                        <td className="p-3">{toText(r['Model'])}</td>
                        <td className="p-3">{toText(r['Status + Priority'])}</td>
                        <td className="p-3 font-medium">{toText(r['Status Days'])}</td>
                      </tr>
                    ))}
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