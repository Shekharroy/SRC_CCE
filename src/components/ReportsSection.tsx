import { useMemo, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  Download,
  Save,
  Search,
  Check,
  FileSpreadsheet,
} from 'lucide-react';
import { CustomReportBuilder } from './CustomReportBuilder';
import { StatusText } from './StatusText';
import type {
  AllReportsBundle,
  CanonicalCounterName,
  CustomReportConfig,
  EnquiryRow,
} from '../shared/types';

export type ActiveReportTab = 'report1' | 'report2' | 'report3' | 'report4' | 'custom';

interface ReportsSectionProps {
  bundle: AllReportsBundle;
  rawRows: EnquiryRow[];
  targetDate: string;
  onChangeTargetDate: (date: string) => void;
  strictDateMatch: boolean;
  onToggleStrictDate: (strict: boolean) => void;
  crmTarget: string;
  onChangeCrmTarget: (crm: string) => void;
  onExportExcel: (scope: 'all' | 'report1' | 'report2' | 'report3' | 'report4') => Promise<void>;
  onExportCustomExcel: (config: CustomReportConfig) => Promise<void>;
  onSaveSnapshot: (title: string) => Promise<void>;
  isSavingSnapshot: boolean;
  snapshotSavedFeedback: boolean;
  onOpenUpload: () => void;
}

export function ReportsSection({
  bundle,
  rawRows,
  targetDate,
  onChangeTargetDate,
  strictDateMatch,
  onToggleStrictDate,
  crmTarget,
  onChangeCrmTarget,
  onExportExcel,
  onExportCustomExcel,
  onSaveSnapshot,
  isSavingSnapshot,
  snapshotSavedFeedback,
  onOpenUpload,
}: ReportsSectionProps) {
  const [activeTab, setActiveTab] = useState<ActiveReportTab>('report1');
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDseFilter, setSelectedDseFilter] = useState<string>('ALL');
  const [selectedModelFilter, setSelectedModelFilter] = useState<string | null>(null);
  const [expandedCounters, setExpandedCounters] = useState<Record<string, boolean>>({
    Chandrahiya: true,
  });
  const [snapshotTitle, setSnapshotTitle] = useState('');

  const q = searchQuery.trim().toLowerCase();

  const toggleCounterExpand = (name: CanonicalCounterName) => {
    setExpandedCounters((prev) => ({ ...prev, [name]: !prev[name] }));
  };

  // Filtered rows for Report 1
  const filteredReport1Rows = useMemo(() => {
    return bundle.report1DseFollowup.flatRows.filter((row) => {
      if (selectedDseFilter !== 'ALL' && row.dseName !== selectedDseFilter) return false;
      if (!q) return true;
      return (
        row.dseName.toLowerCase().includes(q) ||
        row.customerName.toLowerCase().includes(q) ||
        row.principleModel.toLowerCase().includes(q) ||
        row.counterName.toLowerCase().includes(q) ||
        row.remarks.toLowerCase().includes(q) ||
        row.mobile.toLowerCase().includes(q)
      );
    });
  }, [bundle.report1DseFollowup.flatRows, selectedDseFilter, q]);

  // Filtered summaries for Report 2
  const filteredModelSummaries = useMemo(() => {
    return bundle.report2PrincipleModel.summaries.filter((m) =>
      !q ? true : m.principleModel.toLowerCase().includes(q) || m.topCounter.toLowerCase().includes(q)
    );
  }, [bundle.report2PrincipleModel.summaries, q]);

  const modelDetailRows = useMemo(() => {
    if (!selectedModelFilter) return [];
    return bundle.report2PrincipleModel.rows.filter(
      (r) =>
        (r.principleModel || 'Unspecified Model') === selectedModelFilter &&
        (!q ||
          r.customerName.toLowerCase().includes(q) ||
          r.dseName.toLowerCase().includes(q) ||
          r.counterName.toLowerCase().includes(q) ||
          r.remarks.toLowerCase().includes(q))
    );
  }, [bundle.report2PrincipleModel.rows, selectedModelFilter, q]);

  // Filtered rows for Report 3
  const filteredCallRows = useMemo(() => {
    return bundle.report3CrmDailyCall.callRows.filter((row) => {
      if (!q) return true;
      return (
        row.customerName.toLowerCase().includes(q) ||
        row.dseName.toLowerCase().includes(q) ||
        row.counterName.toLowerCase().includes(q) ||
        row.principleModel.toLowerCase().includes(q) ||
        row.remarks.toLowerCase().includes(q)
      );
    });
  }, [bundle.report3CrmDailyCall.callRows, q]);

  const filteredConversions = useMemo(() => {
    return bundle.report3CrmDailyCall.conversionList.filter((row) => {
      if (!q) return true;
      return (
        row.customerName.toLowerCase().includes(q) ||
        row.dseName.toLowerCase().includes(q) ||
        row.counterName.toLowerCase().includes(q) ||
        row.principleModel.toLowerCase().includes(q) ||
        row.remarks.toLowerCase().includes(q)
      );
    });
  }, [bundle.report3CrmDailyCall.conversionList, q]);

  const maxModelCount = useMemo(() => {
    return Math.max(1, ...bundle.report2PrincipleModel.summaries.map((s) => s.totalEnquiries));
  }, [bundle.report2PrincipleModel.summaries]);

  if (bundle.totalRecords === 0) {
    return (
      <section
        id="reports"
        className="bg-white border border-slate-200 rounded-lg p-12 text-center"
      >
        <div className="mx-auto w-12 h-12 rounded-md bg-slate-100 border border-slate-200 flex items-center justify-center text-slate-500">
          <FileSpreadsheet className="w-6 h-6" />
        </div>
        <h2 className="mt-4 text-base font-semibold text-slate-900">
          No Enquiry Records Loaded Yet
        </h2>
        <p className="mt-1.5 text-sm text-slate-600 max-w-md mx-auto">
          Upload an Excel sheet, CSV file, text log, or register image above—or download the sample Excel sheet to inspect all 4 reports immediately.
        </p>
        <div className="mt-6 flex flex-wrap items-center justify-center gap-3">
          <button
            type="button"
            onClick={onOpenUpload}
            className="px-4 py-2 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md transition-colors whitespace-nowrap cursor-pointer"
          >
            Upload Document
          </button>
          <a
            href="/api/sample-template"
            download
            className="px-4 py-2 text-xs font-medium text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-md transition-colors whitespace-nowrap"
          >
            Download Sample Excel
          </a>
        </div>
      </section>
    );
  }

  return (
    <section id="reports" className="bg-white border border-slate-200 rounded-lg">
      {/* Toolbar: Segmented Tabs + Controls */}
      <div className="p-5 border-b border-slate-200 space-y-4">
        <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4">
          <div className="flex flex-wrap items-center gap-1 p-1 bg-slate-100 rounded-md">
            <button
              type="button"
              onClick={() => setActiveTab('report1')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report1'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              1. DSE Today Follow-up ({bundle.report1DseFollowup.totalFollowups})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report2')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report2'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              2. Principle Model Wise ({bundle.report2PrincipleModel.totalModels})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report3')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report3'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              3. CRM Chandrashekhar ({bundle.report3CrmDailyCall.totalCalls})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('report4')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'report4'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              4. Counter Wise Summary (7)
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('custom')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                activeTab === 'custom'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              5. Custom Report Builder
            </button>
          </div>

          <div className="flex flex-wrap items-center gap-2.5 shrink-0">
            <div className="flex items-center gap-1.5">
              <input
                type="text"
                value={snapshotTitle}
                onChange={(e) => setSnapshotTitle(e.target.value)}
                placeholder="Report title (optional)"
                className="w-40 px-2.5 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
              <button
                type="button"
                disabled={isSavingSnapshot}
                onClick={() => onSaveSnapshot(snapshotTitle)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 hover:bg-slate-50 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                {snapshotSavedFeedback ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-700" />
                    <span className="text-emerald-700">Saved to SRC_CCE</span>
                  </>
                ) : (
                  <>
                    <Save className="w-3.5 h-3.5" />
                    <span>{isSavingSnapshot ? 'Saving...' : 'Save Report to DB'}</span>
                  </>
                )}
              </button>
            </div>

            {activeTab !== 'custom' && (
              <button
                type="button"
                onClick={() => onExportExcel(activeTab)}
                className="inline-flex items-center gap-1.5 px-3.5 py-1.5 text-xs font-medium text-white bg-blue-700 hover:bg-blue-800 rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Download This Report (.xlsx)</span>
              </button>
            )}
          </div>
        </div>

        {/* Filter Bar */}
        {activeTab !== 'custom' && (
          <div className="flex flex-wrap items-center justify-between gap-3 pt-1">
            <div className="relative flex-1 min-w-[220px] max-w-md">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search customer, DSE, model, counter, or remarks..."
                className="w-full pl-9 pr-3 py-1.5 text-xs border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
              />
            </div>

          <div className="flex flex-wrap items-center gap-3 text-xs">
            {activeTab === 'report1' && (
              <>
                <label className="inline-flex items-center gap-2 text-slate-600">
                  <span>Target Date:</span>
                  <input
                    type="date"
                    value={targetDate}
                    onChange={(e) => onChangeTargetDate(e.target.value)}
                    className="px-2.5 py-1 text-xs font-mono border border-slate-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-600"
                  />
                </label>
                <label className="inline-flex items-center gap-1.5 text-slate-600 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={strictDateMatch}
                    onChange={(e) => onToggleStrictDate(e.target.checked)}
                    className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                  />
                  <span>Strict date match only</span>
                </label>
                <select
                  value={selectedDseFilter}
                  onChange={(e) => setSelectedDseFilter(e.target.value)}
                  className="px-2.5 py-1 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  <option value="ALL">All DSEs ({bundle.report1DseFollowup.dseCount})</option>
                  {bundle.report1DseFollowup.groups.map((g) => (
                    <option key={`${g.dseName}-${g.counterName}`} value={g.dseName}>
                      {g.dseName} ({g.totalFollowups})
                    </option>
                  ))}
                </select>
              </>
            )}

            {activeTab === 'report3' && (
              <label className="inline-flex items-center gap-2 text-slate-600">
                <span>CRM Caller:</span>
                <select
                  value={crmTarget}
                  onChange={(e) => onChangeCrmTarget(e.target.value)}
                  className="px-2.5 py-1 text-xs border border-slate-300 rounded-md bg-white focus:outline-none focus:ring-2 focus:ring-blue-600"
                >
                  {bundle.report3CrmDailyCall.availableCrmNames.map((name) => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                  <option value="ALL">All CRM Callers</option>
                </select>
              </label>
            )}
          </div>
          </div>
        )}
      </div>

      {/* REPORT 1: DSE Name Wise Today Followup Data List with Remarks */}
      {activeTab === 'report1' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              First Report: DSE Name Wise Today Followup Data List with Remarks
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Target Date: {bundle.report1DseFollowup.targetDate} · {bundle.report1DseFollowup.dseCount} DSEs · {bundle.report1DseFollowup.totalFollowups} follow-up entries
            </p>
          </div>

          {/* DSE Summary Matrix */}
          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3">DSE Name</th>
                  <th className="py-2.5 px-3">Counter</th>
                  <th className="py-2.5 px-3 text-right">Follow-ups</th>
                  <th className="py-2.5 px-3 text-right">Hot Leads</th>
                  <th className="py-2.5 px-3 text-right">Converted / Booked</th>
                  <th className="py-2.5 px-3 text-right">Filter</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {bundle.report1DseFollowup.groups.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="py-6 text-center text-slate-500">
                      No follow-up records found for {targetDate}. Uncheck &quot;Strict date match only&quot; to inspect all follow-ups across uploaded dates.
                    </td>
                  </tr>
                ) : (
                  bundle.report1DseFollowup.groups.map((g) => (
                    <tr
                      key={`${g.dseName}-${g.counterName}`}
                      className="hover:bg-slate-50/80 transition-colors"
                    >
                      <td className="py-2.5 px-3 font-medium text-slate-900">{g.dseName}</td>
                      <td className="py-2.5 px-3 text-slate-600">{g.counterName}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-slate-900">
                        {g.totalFollowups}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700">
                        {g.hotCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700">
                        {g.convertedCount}
                      </td>
                      <td className="py-2.5 px-3 text-right">
                        <button
                          type="button"
                          onClick={() =>
                            setSelectedDseFilter((prev) =>
                              prev === g.dseName ? 'ALL' : g.dseName
                            )
                          }
                          className="text-xs font-medium text-blue-700 hover:underline cursor-pointer"
                        >
                          {selectedDseFilter === g.dseName ? 'Show All' : 'Filter Rows'}
                        </button>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          {/* Detailed Follow-up Rows with Remarks */}
          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-10">Sl</th>
                  <th className="py-2.5 px-3">DSE Name</th>
                  <th className="py-2.5 px-3">Counter</th>
                  <th className="py-2.5 px-3">Followup Date</th>
                  <th className="py-2.5 px-3">Customer</th>
                  <th className="py-2.5 px-3">Mobile</th>
                  <th className="py-2.5 px-3">Principle Model</th>
                  <th className="py-2.5 px-3">Status</th>
                  <th className="py-2.5 px-3 min-w-[240px]">Remarks</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredReport1Rows.map((row, idx) => (
                  <tr key={row.id} className="hover:bg-slate-50/80 transition-colors">
                    <td className="py-2.5 px-3 font-mono tabular-nums text-slate-500">
                      {idx + 1}
                    </td>
                    <td className="py-2.5 px-3 font-medium text-slate-900 whitespace-nowrap">
                      {row.dseName}
                    </td>
                    <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">
                      {row.counterName}
                    </td>
                    <td className="py-2.5 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                      {row.followupDate || targetDate}
                    </td>
                    <td className="py-2.5 px-3 font-medium text-slate-900 whitespace-nowrap">
                      {row.customerName}
                    </td>
                    <td className="py-2.5 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                      {row.mobile || '-'}
                    </td>
                    <td className="py-2.5 px-3 text-slate-800 whitespace-nowrap">
                      {row.principleModel}
                    </td>
                    <td className="py-2.5 px-3">
                      <StatusText status={row.status} />
                    </td>
                    <td className="py-2.5 px-3 text-slate-700 leading-relaxed">
                      {row.remarks || '-'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* REPORT 2: Principle Model Wise Enquiry Report */}
      {activeTab === 'report2' && (
        <div className="p-5 space-y-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-base font-semibold text-slate-900">
                Second Report: Principle Model Wise Enquiry Report
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                {bundle.report2PrincipleModel.totalModels} models · {bundle.report2PrincipleModel.totalEnquiries} total enquiries · {bundle.report2PrincipleModel.totalConverted} converted/booked
              </p>
            </div>
            {selectedModelFilter && (
              <button
                type="button"
                onClick={() => setSelectedModelFilter(null)}
                className="text-xs font-medium text-blue-700 hover:underline cursor-pointer"
              >
                Clear model drill-down ({selectedModelFilter})
              </button>
            )}
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-10">Sl</th>
                  <th className="py-2.5 px-3 min-w-[180px]">Principle Model</th>
                  <th className="py-2.5 px-3 min-w-[120px]">Share</th>
                  <th className="py-2.5 px-3 text-right">Total</th>
                  <th className="py-2.5 px-3 text-right">Open</th>
                  <th className="py-2.5 px-3 text-right">Follow-up</th>
                  <th className="py-2.5 px-3 text-right">Hot</th>
                  <th className="py-2.5 px-3 text-right">Warm</th>
                  <th className="py-2.5 px-3 text-right">Cold</th>
                  <th className="py-2.5 px-3 text-right">Booked</th>
                  <th className="py-2.5 px-3 text-right">Converted</th>
                  <th className="py-2.5 px-3 text-right">Lost</th>
                  <th className="py-2.5 px-3 text-right">Conv %</th>
                  <th className="py-2.5 px-3">Top Counter</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredModelSummaries.map((m, idx) => {
                  const pctWidth = Math.max(4, Math.round((m.totalEnquiries / maxModelCount) * 100));
                  const isSelected = selectedModelFilter === m.principleModel;
                  return (
                    <tr
                      key={m.principleModel}
                      onClick={() =>
                        setSelectedModelFilter((prev) =>
                          prev === m.principleModel ? null : m.principleModel
                        )
                      }
                      className={`cursor-pointer transition-colors ${
                        isSelected ? 'bg-blue-50/70' : 'hover:bg-slate-50/80'
                      }`}
                    >
                      <td className="py-2.5 px-3 font-mono tabular-nums text-slate-500">
                        {idx + 1}
                      </td>
                      <td className="py-2.5 px-3 font-medium text-slate-900">
                        {m.principleModel}
                      </td>
                      <td className="py-2.5 px-3">
                        <div className="w-24 h-2 bg-slate-100 rounded-xs overflow-hidden">
                          <div
                            className="h-full bg-blue-700"
                            style={{ width: `${pctWidth}%` }}
                          />
                        </div>
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                        {m.totalEnquiries}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600">
                        {m.openCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700">
                        {m.followupCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700">
                        {m.hotCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-600">
                        {m.warmCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-500">
                        {m.coldCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700">
                        {m.bookedCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700">
                        {m.convertedCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700">
                        {m.lostCount}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-slate-900">
                        {m.conversionRate}%
                      </td>
                      <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">
                        {m.topCounter}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {selectedModelFilter && (
            <div className="space-y-2 pt-2">
              <h3 className="text-sm font-semibold text-slate-900">
                Enquiry Drill-Down: {selectedModelFilter} ({modelDetailRows.length} rows)
              </h3>
              <div className="overflow-x-auto border border-slate-200 rounded-md">
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                      <th className="py-2 px-3">Enquiry No</th>
                      <th className="py-2 px-3">Date</th>
                      <th className="py-2 px-3">Customer</th>
                      <th className="py-2 px-3">Counter</th>
                      <th className="py-2 px-3">DSE Name</th>
                      <th className="py-2 px-3">Status</th>
                      <th className="py-2 px-3">Remarks</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {modelDetailRows.map((r) => (
                      <tr key={r.id} className="hover:bg-slate-50/80">
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600">
                          {r.enquiryNo}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600">
                          {r.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-medium text-slate-900">
                          {r.customerName}
                        </td>
                        <td className="py-2 px-3 text-slate-600">{r.counterName}</td>
                        <td className="py-2 px-3 text-slate-800">{r.dseName}</td>
                        <td className="py-2 px-3">
                          <StatusText status={r.status} />
                        </td>
                        <td className="py-2 px-3 text-slate-600">{r.remarks || '-'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      )}

      {/* REPORT 3: Daily Called Data Count Report of CRM Chandrashekhar Kumar with Remarks & Daily Lead Conversion List */}
      {activeTab === 'report3' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Third Report: Daily Called Data Count Report of CRM {bundle.report3CrmDailyCall.crmName} with Remarks &amp; Daily Lead Conversion List
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Total Called Records: {bundle.report3CrmDailyCall.totalCalls} · Converted Leads: {bundle.report3CrmDailyCall.totalConversions} · Conversion Rate: {bundle.report3CrmDailyCall.conversionRate}%
            </p>
          </div>

          {/* 3A. Daily Called Count Summary */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3A. Daily Called Data Count &amp; Consolidated Remarks
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <th className="py-2.5 px-3">Call Date</th>
                    <th className="py-2.5 px-3 text-right">Called Count</th>
                    <th className="py-2.5 px-3 text-right">Follow-up</th>
                    <th className="py-2.5 px-3 text-right">Hot</th>
                    <th className="py-2.5 px-3 text-right">Converted</th>
                    <th className="py-2.5 px-3 text-right">Lost</th>
                    <th className="py-2.5 px-3 min-w-[280px]">Remarks Summary</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {bundle.report3CrmDailyCall.dailySummaries.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-6 text-center text-slate-500">
                        No call log rows matched CRM &quot;{bundle.report3CrmDailyCall.crmName}&quot;. Ensure your sheet includes a CRM / Caller column or switch the CRM filter above.
                      </td>
                    </tr>
                  ) : (
                    bundle.report3CrmDailyCall.dailySummaries.map((d) => (
                      <tr key={d.date} className="hover:bg-slate-50/80">
                        <td className="py-2.5 px-3 font-mono tabular-nums font-medium text-slate-900 whitespace-nowrap">
                          {d.date}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {d.totalCalled}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700">
                          {d.followupCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700">
                          {d.hotCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700">
                          {d.convertedCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700">
                          {d.lostCount}
                        </td>
                        <td className="py-2.5 px-3 text-slate-600 leading-relaxed">
                          {d.remarksList.join(' | ') || '-'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3B. Detailed Call Remarks Log */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3B. Individual Call Log with Remarks ({filteredCallRows.length})
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <th className="py-2.5 px-3 w-10">Sl</th>
                    <th className="py-2.5 px-3">Call Date</th>
                    <th className="py-2.5 px-3">Customer</th>
                    <th className="py-2.5 px-3">Mobile</th>
                    <th className="py-2.5 px-3">Counter</th>
                    <th className="py-2.5 px-3">DSE Name</th>
                    <th className="py-2.5 px-3">Principle Model</th>
                    <th className="py-2.5 px-3">Status</th>
                    <th className="py-2.5 px-3 min-w-[220px]">Remarks</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredCallRows.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-5 text-center text-slate-500">
                        No matching call log entries.
                      </td>
                    </tr>
                  ) : (
                    filteredCallRows.map((row, idx) => (
                      <tr key={row.id} className="hover:bg-slate-50/80">
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-500">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                          {row.calledDate || row.followupDate || row.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-medium text-slate-900 whitespace-nowrap">
                          {row.customerName}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.mobile || '-'}
                        </td>
                        <td className="py-2 px-3 text-slate-600 whitespace-nowrap">
                          {row.counterName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.dseName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.principleModel}
                        </td>
                        <td className="py-2 px-3">
                          <StatusText status={row.status} />
                        </td>
                        <td className="py-2 px-3 text-slate-700">{row.remarks || '-'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3C. Daily Lead Conversion List */}
          <div className="space-y-2">
            <h3 className="text-xs font-semibold text-slate-700">
              3C. Daily Lead Conversion List ({filteredConversions.length})
            </h3>
            <div className="overflow-x-auto border border-slate-200 rounded-md">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                    <th className="py-2.5 px-3 w-10">Sl</th>
                    <th className="py-2.5 px-3">Conversion Date</th>
                    <th className="py-2.5 px-3">Enquiry No</th>
                    <th className="py-2.5 px-3">Customer</th>
                    <th className="py-2.5 px-3">Mobile</th>
                    <th className="py-2.5 px-3">Counter</th>
                    <th className="py-2.5 px-3">DSE Name</th>
                    <th className="py-2.5 px-3">Principle Model</th>
                    <th className="py-2.5 px-3">Stage</th>
                    <th className="py-2.5 px-3 min-w-[200px]">Remarks</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {filteredConversions.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="py-5 text-center text-slate-500">
                        No converted or booked leads found in current dataset.
                      </td>
                    </tr>
                  ) : (
                    filteredConversions.map((row, idx) => (
                      <tr key={row.id} className="hover:bg-slate-50/80">
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-500">
                          {idx + 1}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-700 whitespace-nowrap">
                          {row.conversionDate || row.calledDate || row.enquiryDate || '-'}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.enquiryNo}
                        </td>
                        <td className="py-2 px-3 font-medium text-slate-900 whitespace-nowrap">
                          {row.customerName}
                        </td>
                        <td className="py-2 px-3 font-mono tabular-nums text-slate-600 whitespace-nowrap">
                          {row.mobile || '-'}
                        </td>
                        <td className="py-2 px-3 text-slate-600 whitespace-nowrap">
                          {row.counterName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.dseName}
                        </td>
                        <td className="py-2 px-3 text-slate-800 whitespace-nowrap">
                          {row.principleModel}
                        </td>
                        <td className="py-2 px-3">
                          <StatusText status={row.status} />
                        </td>
                        <td className="py-2 px-3 text-slate-700">{row.remarks || '-'}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* REPORT 4: Counter Wise Enquiry Report with Status Summaries */}
      {activeTab === 'report4' && (
        <div className="p-5 space-y-6">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Fourth Report: Counter Wise Enquiry Report with Status Summaries
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              All 7 dealership counters · {bundle.report4CounterWise.totalEnquiries} total enquiries · {bundle.report4CounterWise.totalConverted} converted/booked ({bundle.report4CounterWise.overallConversionRate}% conversion rate). Click any counter row to expand DSE breakdown.
            </p>
          </div>

          <div className="overflow-x-auto border border-slate-200 rounded-md">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                  <th className="py-2.5 px-3 w-8"></th>
                  <th className="py-2.5 px-3">Counter Name</th>
                  <th className="py-2.5 px-3 text-right">Total</th>
                  <th className="py-2.5 px-3 text-right">Open</th>
                  <th className="py-2.5 px-3 text-right">Follow-up</th>
                  <th className="py-2.5 px-3 text-right">Hot</th>
                  <th className="py-2.5 px-3 text-right">Warm</th>
                  <th className="py-2.5 px-3 text-right">Cold</th>
                  <th className="py-2.5 px-3 text-right">Booked</th>
                  <th className="py-2.5 px-3 text-right">Converted</th>
                  <th className="py-2.5 px-3 text-right">Lost</th>
                  <th className="py-2.5 px-3 text-right">Conv %</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {bundle.report4CounterWise.summaries.map((counter) => {
                  const isExpanded = Boolean(expandedCounters[counter.counterName]);
                  return (
                    <>
                      <tr
                        key={counter.counterName}
                        onClick={() => toggleCounterExpand(counter.counterName)}
                        className="cursor-pointer bg-white hover:bg-slate-50 transition-colors"
                      >
                        <td className="py-2.5 px-3 text-slate-500">
                          {isExpanded ? (
                            <ChevronDown className="w-4 h-4" />
                          ) : (
                            <ChevronRight className="w-4 h-4" />
                          )}
                        </td>
                        <td className="py-2.5 px-3 font-semibold text-slate-900 whitespace-nowrap">
                          {counter.counterNumber <= 7
                            ? `${counter.counterNumber}. ${counter.counterName}`
                            : counter.counterName}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {counter.totalEnquiries}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600">
                          {counter.openCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-blue-700">
                          {counter.followupCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-700">
                          {counter.hotCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-amber-600">
                          {counter.warmCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-500">
                          {counter.coldCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-emerald-700">
                          {counter.bookedCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-medium text-emerald-700">
                          {counter.convertedCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums text-red-700">
                          {counter.lostCount}
                        </td>
                        <td className="py-2.5 px-3 text-right font-mono tabular-nums font-semibold text-slate-900">
                          {counter.conversionRate}%
                        </td>
                      </tr>
                      {isExpanded &&
                        counter.dseBreakdowns.map((dse) => (
                          <tr
                            key={`${counter.counterName}-${dse.dseName}`}
                            className="bg-slate-50/70 text-slate-700"
                          >
                            <td className="py-2 px-3"></td>
                            <td className="py-2 px-3 pl-8 text-slate-700 whitespace-nowrap">
                              {dse.dseName}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.totalEnquiries}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.openCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.followupCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.hotCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.warmCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.coldCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.bookedCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.convertedCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.lostCount}
                            </td>
                            <td className="py-2 px-3 text-right font-mono tabular-nums">
                              {dse.conversionRate}%
                            </td>
                          </tr>
                        ))}
                    </>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {activeTab === 'custom' && (
        <CustomReportBuilder
          rows={rawRows}
          onExportCustomExcel={onExportCustomExcel}
        />
      )}
    </section>
  );
}
