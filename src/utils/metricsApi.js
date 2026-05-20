/**
 * metricsApi.js — Fetch helpers for the observability metrics endpoints.
 *
 * All functions use the same VITE_AI_EXTRACT base URL that the rest of the
 * application uses, ensuring consistent environment configuration.
 *
 * Time range
 * ----------
 * Pass either { hours } (relative, default 24) or { startTime, endTime }
 * (absolute ISO strings).  startTime / endTime take precedence when both
 * are provided.
 */

import { buildAiExtractUrl } from "../config/apiBase";

/**
 * Build a URLSearchParams string from a time-range descriptor object.
 *
 * @param {{ template: string, hours?: number, startTime?: string, endTime?: string }} params
 * @returns {URLSearchParams}
 */
function buildParams({ template, hours, startTime, endTime, granularity }) {
  const p = new URLSearchParams({ template });
  if (startTime && endTime) {
    p.append("start_time", startTime);
    p.append("end_time", endTime);
  } else if (hours) {
    p.append("hours", String(hours));
  }
  if (granularity) {
    p.append("granularity", granularity);
  }
  return p;
}

async function _fetchJson(path) {
  const res = await fetch(buildAiExtractUrl(path));
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`${path} → HTTP ${res.status}: ${body}`);
  }
  return res.json();
}

/**
 * Fetch aggregated summary stat-card values for a template + time window.
 *
 * @param {{ template: string, hours?: number, startTime?: string, endTime?: string }} params
 * @returns {Promise<{
 *   business: {
 *     total_cost_usd: number,
 *     cost_per_page_usd: number|null,
 *     avg_confidence_after: number|null,
 *     avg_confidence_before: number|null,
 *     fields_extracted: number
 *   },
 *   operational: {
 *     total_runs: number,
 *     pages_processed: number,
 *     avg_duration_ms: number|null,
 *     total_tokens: number
 *   }
 * }>}
 */
export async function fetchMetricsSummary({ template, hours, startTime, endTime }) {
  const p = buildParams({ template, hours, startTime, endTime });
  return _fetchJson(`/api/metrics/summary?${p}`);
}

/**
 * Fetch the list of pipeline runs for a template, most-recent first.
 *
 * @param {{ template: string, hours?: number, startTime?: string, endTime?: string }} params
 * @returns {Promise<Array<{ run_id: string, document_name: string, model: string,
 *   LastSeen: string, AvgDurationMs: number, TotalRuns: number }>>}
 */
export async function fetchMetricsRuns({ template, hours, startTime, endTime }) {
  const p = buildParams({ template, hours, startTime, endTime });
  return _fetchJson(`/api/metrics/runs?${p}`);
}

/**
 * Fetch all metric rows for a single run (drill-down detail).
 *
 * @param {string} runId
 * @returns {Promise<Array<{ TimeGenerated: string, Name: string, Sum: number,
 *   run_id: string, template: string, document_name: string, model: string,
 *   cost_component: string, token_type: string, page_tier: string }>>}
 */
export async function fetchRunDetail(runId) {
  return _fetchJson(`/api/metrics/run/${encodeURIComponent(runId)}`);
}

/**
 * Fetch all chart datasets for a template in a single request.
 *
 * @param {{ template: string, hours?: number, startTime?: string, endTime?: string,
 *   granularity?: string }} params
 * @returns {Promise<{
 *   business: {
 *     charts: {
 *       cost_per_day: Array,
 *       cost_per_page: Array,
 *       fields_extracted: Array,
 *       confidence_after: Array
 *     }
 *   },
 *   operational: {
 *     charts: {
 *       runs_per_day: Array,
 *       pages_per_day: Array,
 *       tokens_per_day: Array,
 *       avg_duration: Array
 *     }
 *   },
 *   granularity: string
 * }>}
 */
export async function fetchMetricsCharts({
  template,
  hours,
  startTime,
  endTime,
  granularity,
}) {
  const p = buildParams({ template, hours, startTime, endTime, granularity });
  return _fetchJson(`/api/metrics/charts?${p}`);
}
