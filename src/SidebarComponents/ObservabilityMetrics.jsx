/**
 * ObservabilityMetrics.jsx
 */

import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Button,
  Card,
  Col,
  DatePicker,
  Divider,
  Empty,
  Row,
  Spin,
  Table,
  Tag,
  Typography,
} from "antd";
import {
  ClockCircleOutlined,
  DollarOutlined,
  FileTextOutlined,
  NumberOutlined,
  ThunderboltOutlined,
} from "@ant-design/icons";
import dayjs from "dayjs";

import useMetaData from "../context/metaData";
import AreaChart from "../components/Graphs/AreaChart";
import LineChart from "../components/Graphs/LineChart";
import ColumnChart from "../components/Graphs/ColumnChart";
import {
  fetchMetricsCharts,
  fetchMetricsSummary,
  fetchMetricsRuns,
  fetchRunDetail,
} from "../utils/metricsApi";

import "./ObservabilityMetrics.css";

const { RangePicker } = DatePicker;
const { Text, Title } = Typography;

function fmtDuration(ms) {
  if (ms == null) return "-";
  if (ms < 1000) return `${Math.round(ms)} ms`;
  return `${(ms / 1000).toFixed(1)} s`;
}

function fmtCost(usd) {
  if (usd == null) return "-";
  if (usd === 0) return "$0.00";
  if (usd < 0.0001) return `$${usd.toExponential(2)}`;
  return `$${usd.toFixed(4)}`;
}

function fmtCount(value) {
  if (value == null) return "-";
  return Number(value).toLocaleString();
}

function fmtConfidence(value) {
  if (value == null) return "-";
  return `${(Number(value) * 100).toFixed(1)}%`;
}

function toEpochMs(value) {
  if (value == null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    const abs = Math.abs(value);
    if (abs < 1e11) return Math.round(value * 1000); // seconds
    if (abs < 1e14) return Math.round(value); // milliseconds
    if (abs < 1e17) return Math.round(value / 1000); // microseconds
    return Math.round(value / 1000000); // nanoseconds
  }

  const asNumber = Number(value);
  if (Number.isFinite(asNumber)) {
    return toEpochMs(asNumber);
  }

  const parsed = new Date(value).getTime();
  return Number.isNaN(parsed) ? null : parsed;
}

function fmtBinTs(isoStr) {
  if (!isoStr) return "";
  const ts = toEpochMs(isoStr);
  if (ts == null) return String(isoStr);
  const d = new Date(ts);
  const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dy = String(d.getUTCDate()).padStart(2, "0");
  const hr = String(d.getUTCHours()).padStart(2, "0");
  const mn = String(d.getUTCMinutes()).padStart(2, "0");
  return `${mo}/${dy} ${hr}:${mn}`;
}

function shortLabel(value, max = 18) {
  if (!value) return "";
  return value.length > max ? `${value.slice(0, max - 1)}…` : value;
}

function fmtRunLabel(row) {
  const stamp = fmtBinTs(row?.ts);
  const doc = shortLabel(row?.document_name);
  if (doc && stamp) return `${doc} • ${stamp}`;
  if (doc) return doc;
  if (stamp) return stamp;
  return shortLabel(row?.run_id || "run");
}

function autoGranularity(start, end) {
  const hours = end.diff(start, "hour");
  if (hours <= 2) return "15m";
  if (hours <= 12) return "30m";
  if (hours <= 72) return "1h";
  if (hours <= 336) return "6h";
  return "1d";
}

function granularityToMs(granularity) {
  switch (granularity) {
    case "15m": return 15 * 60 * 1000;
    case "30m": return 30 * 60 * 1000;
    case "6h": return 6 * 60 * 60 * 1000;
    case "1d": return 24 * 60 * 60 * 1000;
    case "1h":
    default:
      return 60 * 60 * 1000;
  }
}

function buildBins(min, max, stepMs) {
  if (!Number.isFinite(min) || !Number.isFinite(max) || !Number.isFinite(stepMs) || stepMs <= 0 || max <= min) {
    return [];
  }
  const bins = [];
  for (let ts = min; ts <= max; ts += stepMs) bins.push(ts);
  if (!bins.length || bins[bins.length - 1] < max) bins.push(max);
  return bins;
}

function withTimeout(promise, timeoutMs = 25000) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Request timed out after ${timeoutMs}ms`)), timeoutMs)
    ),
  ]);
}

const MetricsCards = ({ title, cards, loading }) => (
  <Card className="obs-runs-card" title={title} bordered={false} style={{ marginTop: 16 }}>
    <Spin spinning={loading}>
      <Row gutter={[12, 12]} className="obs-summary-row">
        {cards.map((c) => (
          <Col xs={24} sm={12} md={8} lg={6} key={c.key}>
            <Card className="obs-stat-card" bordered={false}>
              <div className="obs-stat-icon" style={{ color: c.color }}>
                {c.icon}
              </div>
              <div className="obs-stat-title">{c.title}</div>
              <div className="obs-stat-value" style={{ color: c.color }}>
                {c.value}
              </div>
              {c.subValue && <div className="obs-stat-subvalue">{c.subValue}</div>}
            </Card>
          </Col>
        ))}
      </Row>
    </Spin>
  </Card>
);

const RunDetailPanel = ({ runId, runMeta, rows, loading, onClose }) => {
  if (!runId) return null;

  const toNumber = (v) => {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  };

  const latestValue = (metricName) => {
    const metricRows = rows.filter((r) => r.Name === metricName);
    if (!metricRows.length) return null;
    const ordered = [...metricRows].sort(
      (a, b) => new Date(a.TimeGenerated || 0) - new Date(b.TimeGenerated || 0)
    );
    return toNumber(ordered[ordered.length - 1].Sum);
  };

  const fieldsExtracted = latestValue("slide.fields.extracted.count") || 0;
  const confBefore = latestValue("slide.confidence.avg.before_validation");
  const confAfter = latestValue("slide.confidence.avg.after_validation");
  const confidenceLift =
    confBefore != null && confAfter != null
      ? `${((confAfter - confBefore) * 100).toFixed(1)}pp lift from validation`
      : "-";

  const latestTime = rows
    .map((r) => r.TimeGenerated)
    .filter(Boolean)
    .sort((a, b) => new Date(a) - new Date(b))
    .pop();

  const runInfo = {
    document: runMeta?.document_name || rows.find((r) => r.document_name)?.document_name || "-",
    template: runMeta?.template || rows.find((r) => r.template)?.template || "-",
    runId,
    time: latestTime ? new Date(latestTime).toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }) : "-",
    duration: fmtDuration(latestValue("slide.run.duration.ms")),
  };

  const tokenRows = rows.filter((r) => r.Name === "slide.llm.tokens.by_model");
  // Collect contextualization total (stored under the analyzer model)
  let contextualizationTokens = 0;
  tokenRows.forEach((r) => {
    if ((r.token_type || "").toLowerCase() === "contextualization") {
      contextualizationTokens += toNumber(r.Sum);
    }
  });
  const tokensByModelMap = {};
  tokenRows.forEach((r) => {
    const model = r.model || "unknown";
    // Skip the analyzer model row (e.g. mortgage_analyzer_0604, wind_mit_analyzer_*)
    // — it is an aggregate label, not an actual LLM model
    if (/_analyzer_|_0\d{3}$/.test(model)) return;
    if (!tokensByModelMap[model]) {
      tokensByModelMap[model] = { model, input: 0, output: 0 };
    }
    const tokenType = (r.token_type || "").toLowerCase();
    const value = toNumber(r.Sum);
    if (tokenType === "input") tokensByModelMap[model].input += value;
    if (tokenType === "output") tokensByModelMap[model].output += value;
  });
  const tokensByModel = Object.values(tokensByModelMap).map((r) => ({
    ...r,
    total: r.input + r.output,
  }));
  tokensByModel.push({
    model: "Contextualization",
    input: null,
    output: contextualizationTokens,
    total: contextualizationTokens,
  });

  const pagesByTier = { minimal: 0, basic: 0, standard: 0 };
  rows
    .filter((r) => r.Name === "slide.document.pages.processed")
    .forEach((r) => {
      const tier = (r.page_tier || "").toLowerCase();
      if (Object.prototype.hasOwnProperty.call(pagesByTier, tier)) {
        pagesByTier[tier] += Math.round(toNumber(r.Sum));
      }
    });

  const costLabels = {
    input: "LLM input",
    output: "LLM output",
    contextualization: "Contextualization",
    "page.minimal": "Page fees",
    "page.basic": "Page fees",
    "page.standard": "Page fees",
  };

  const costBreakdownMap = {
    "LLM input": 0,
    "LLM output": 0,
    Contextualization: 0,
    "Page fees": 0,
  };

  rows
    .filter((r) => r.Name === "slide.cost.component.usd")
    .forEach((r) => {
      const label = costLabels[r.cost_component] || r.cost_component || "Other";
      if (!Object.prototype.hasOwnProperty.call(costBreakdownMap, label)) {
        costBreakdownMap[label] = 0;
      }
      costBreakdownMap[label] += toNumber(r.Sum);
    });

  const totalCostFromMetric = latestValue("slide.cost.estimated.usd");
  const totalCost = totalCostFromMetric && totalCostFromMetric > 0
    ? totalCostFromMetric
    : Object.values(costBreakdownMap).reduce((acc, n) => acc + n, 0);

  const costBreakdown = Object.entries(costBreakdownMap)
    .filter(([, value]) => value > 0)
    .map(([label, value]) => ({
      label,
      value,
      width: totalCost > 0 ? Math.max((value / totalCost) * 100, 4) : 0,
    }));

  return (
    <Card
      className="obs-detail-card"
      title={runInfo.document}
      extra={<Button size="small" onClick={onClose}>Close</Button>}
      bordered={false}
    >
      <Spin spinning={loading}>
        {!rows.length && !loading && <Empty description="No metric data found for this run" />}

        {!!rows.length && (
          <div className="obs-run-detail-layout">
            <Row gutter={[12, 12]}>
              <Col xs={24} lg={12}>
                <div className="obs-run-detail-block">
                  <div className="obs-run-detail-block-title">Document</div>
                  <div className="obs-run-detail-row"><span>File</span><span>{runInfo.document}</span></div>
                  <div className="obs-run-detail-row"><span>Template</span><span>{runInfo.template}</span></div>
                  <div className="obs-run-detail-row"><span>Run ID</span><span>{runInfo.runId}</span></div>
                  <div className="obs-run-detail-row"><span>Time</span><span>{runInfo.time}</span></div>
                  <div className="obs-run-detail-row"><span>Duration</span><span>{runInfo.duration}</span></div>
                </div>
              </Col>

              <Col xs={24} lg={12}>
                <div className="obs-run-detail-block">
                  <div className="obs-run-detail-block-title">Extraction Quality</div>
                  <div className="obs-run-detail-progress-row">
                    <span>Total fields extracted</span>
                    <span>{fmtCount(fieldsExtracted)}</span>
                  </div>
                  <div className="obs-run-detail-bar"><div style={{ width: "100%" }} /></div>

                  <div className="obs-run-detail-progress-row">
                    <span>Before validation</span>
                    <span>{fmtConfidence(confBefore)}</span>
                  </div>
                  <div className="obs-run-detail-bar obs-before"><div style={{ width: `${Math.max((confBefore || 0) * 100, 2)}%` }} /></div>

                  <div className="obs-run-detail-progress-row">
                    <span>After validation</span>
                    <span>{fmtConfidence(confAfter)}</span>
                  </div>
                  <div className="obs-run-detail-bar obs-after"><div style={{ width: `${Math.max((confAfter || 0) * 100, 2)}%` }} /></div>

                  <div className="obs-run-detail-footnote">{confidenceLift}</div>
                </div>
              </Col>

              <Col xs={24} lg={12}>
                <div className="obs-run-detail-block">
                  <div className="obs-run-detail-block-title">Tokens by model</div>
                  {tokensByModel.length ? (
                    <div className="obs-run-detail-token-table">
                      <div className="obs-run-detail-token-header">
                        <span>Model</span><span>In</span><span>Out</span><span>Total</span>
                      </div>
                      {tokensByModel.map((r) => (
                        <div className="obs-run-detail-token-row" key={r.model}>
                          <span>{r.model}</span>
                          <span>{r.input == null ? "-" : fmtCount(r.input)}</span>
                          <span>{r.output == null ? "-" : fmtCount(r.output)}</span>
                          <span>{r.total == null ? "-" : fmtCount(r.total)}</span>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <Text type="secondary">No token breakdown</Text>
                  )}

                  <div className="obs-pages-tier-row">
                    <span>Pages by tier</span>
                    <Tag>minimal {pagesByTier.minimal}</Tag>
                    <Tag>basic {pagesByTier.basic}</Tag>
                    <Tag>standard {pagesByTier.standard}</Tag>
                  </div>
                </div>
              </Col>

              <Col xs={24} lg={12}>
                <div className="obs-run-detail-block">
                  <div className="obs-run-detail-block-title">Cost breakdown</div>
                  {costBreakdown.length ? (
                    costBreakdown.map((r) => (
                      <div className="obs-cost-break-row" key={r.label}>
                        <div className="obs-cost-break-head">
                          <span>{r.label}</span>
                          <span>{fmtCost(r.value)}</span>
                        </div>
                        <div className="obs-run-detail-bar"><div style={{ width: `${r.width}%` }} /></div>
                      </div>
                    ))
                  ) : (
                    <Text type="secondary">No cost breakdown</Text>
                  )}
                  <div className="obs-run-detail-row obs-total-row"><span>Total</span><span>{fmtCost(totalCost)}</span></div>
                </div>
              </Col>
            </Row>
          </div>
        )}
      </Spin>
    </Card>
  );
};

const ObservabilityMetrics = ({ template }) => {
  const { theme } = useMetaData();

  const [summary, setSummary] = useState(null);
  const [summaryLoading, setSummaryLoading] = useState(false);
  const [runs, setRuns] = useState([]);
  const [runsLoading, setRunsLoading] = useState(false);
  const [selectedRunId, setSelectedRunId] = useState(null);
  const [runDetail, setRunDetail] = useState([]);
  const [runDetailLoading, setRunDetailLoading] = useState(false);
  const [chartData, setChartData] = useState(null);
  const [chartsLoading, setChartsLoading] = useState(false);

  const [dateRange, setDateRange] = useState([dayjs().subtract(24, "hour"), dayjs()]);

  const mainFetchId = useRef(0);
  const detailFetchId = useRef(0);

  const timeParams = useMemo(() => {
    if (!dateRange || !dateRange[0] || !dateRange[1]) return { hours: 24 };
    return {
      startTime: dateRange[0].toISOString(),
      endTime: dateRange[1].toISOString(),
    };
  }, [dateRange]);

  const granularity = useMemo(() => {
    if (!dateRange || !dateRange[0] || !dateRange[1]) return "1h";
    return autoGranularity(dateRange[0], dateRange[1]);
  }, [dateRange]);

  const selectedRangeMs = useMemo(() => {
    if (!dateRange || !dateRange[0] || !dateRange[1]) return { min: null, max: null };
    return {
      min: dateRange[0].valueOf(),
      max: dateRange[1].valueOf(),
    };
  }, [dateRange]);
  const chartStepMs = useMemo(() => granularityToMs(granularity), [granularity]);

  const fetchAll = useCallback(async () => {
    const fetchId = ++mainFetchId.current;
    setSummaryLoading(true);
    setRunsLoading(true);
    setChartsLoading(true);
    fetchMetricsSummary({ template, ...timeParams })
      .then((data) => {
        if (fetchId !== mainFetchId.current) return;
        setSummary(data);
      })
      .catch(() => {})
      .finally(() => {
        if (fetchId !== mainFetchId.current) return;
        setSummaryLoading(false);
      });

    fetchMetricsRuns({ template, ...timeParams })
      .then((runsData) => {
        if (fetchId !== mainFetchId.current) return;
        setRuns(runsData);
        if (runsData.length === 0) {
          setSelectedRunId(null);
          setRunDetail([]);
        } else if (selectedRunId && !runsData.some((r) => r.run_id === selectedRunId)) {
          setSelectedRunId(null);
          setRunDetail([]);
        }
      })
      .catch(() => {})
      .finally(() => {
        if (fetchId !== mainFetchId.current) return;
        setRunsLoading(false);
      });

    withTimeout(fetchMetricsCharts({ template, ...timeParams, granularity }), 90000)
      .then((data) => {
        if (fetchId !== mainFetchId.current) return;
        setChartData(data);
      })
      .catch(() => {})
      .finally(() => {
        if (fetchId !== mainFetchId.current) return;
        setChartsLoading(false);
      });
  }, [template, timeParams, granularity, selectedRunId]);

  useEffect(() => {
    fetchAll();
  }, [fetchAll]);

  useEffect(() => {
    if (!selectedRunId) {
      setRunDetail([]);
      setRunDetailLoading(false);
      return;
    }

    const fetchId = ++detailFetchId.current;
    setRunDetailLoading(true);

    fetchRunDetail(selectedRunId)
      .then((data) => {
        if (fetchId !== detailFetchId.current) return;
        setRunDetail(data);
        setRunDetailLoading(false);
      })
      .catch(() => {
        if (fetchId !== detailFetchId.current) return;
        setRunDetailLoading(false);
      });
  }, [selectedRunId]);

  const runsColumns = useMemo(
    () => [
      {
        title: "Document",
        dataIndex: "document_name",
        width: 200,
        ellipsis: true,
        render: (v) => v || "-",
      },
      {
        title: "Model",
        dataIndex: "model",
        width: 120,
        render: (v) => (v ? <Tag>{v}</Tag> : "-"),
      },
      {
        title: "Last Seen",
        dataIndex: "LastSeen",
        width: 160,
        render: (v) => (v ? new Date(v).toLocaleString(undefined, { hour12: false }) : "-"),
        sorter: (a, b) => new Date(a.LastSeen) - new Date(b.LastSeen),
        defaultSortOrder: "descend",
      },
      {
        title: "Avg Duration",
        dataIndex: "AvgDurationMs",
        width: 120,
        render: (v) => fmtDuration(v),
        sorter: (a, b) => (a.AvgDurationMs || 0) - (b.AvgDurationMs || 0),
      },
    ],
    []
  );

  const businessSummary = summary?.business || {};
  const operationalSummary = summary?.operational || {};
  const businessCharts = chartData?.business?.charts || {};
  const operationalCharts = chartData?.operational?.charts || {};

  const businessCards = [
    {
      key: "b_cost",
      title: "Total Cost",
      value: fmtCost(businessSummary.total_cost_usd),
      subValue: `avg cost/run ${fmtCost(businessSummary.avg_cost_per_run)}`,
      icon: <DollarOutlined />,
      color: "#1677ff",
    },
    {
      key: "b_cpp",
      title: "Avg. Cost Per Page",
      value: fmtCost(businessSummary.cost_per_page_usd),
      icon: <DollarOutlined />,
      color: "#52c41a",
    },
    {
      key: "b_conf_after",
      title: "Avg Confidence (After)",
      value: fmtConfidence(businessSummary.avg_confidence_after),
      subValue: `Before ${fmtConfidence(businessSummary.avg_confidence_before)}`,
      icon: <ThunderboltOutlined />,
      color: "#fa8c16",
    },
    {
      key: "b_fields",
      title: "Total Fields Extracted",
      value: fmtCount(businessSummary.fields_extracted || 0),
      icon: <FileTextOutlined />,
      color: "#13c2c2",
    },
  ];

  const operationalCards = [
    {
      key: "o_runs",
      title: "Total Runs",
      value: fmtCount(operationalSummary.total_runs || 0),
      icon: <FileTextOutlined />,
      color: "#1677ff",
    },
    {
      key: "o_pages",
      title: "Total Pages Processed",
      value: fmtCount(operationalSummary.pages_processed || 0),
      icon: <NumberOutlined />,
      color: "#52c41a",
    },
    {
      key: "o_duration",
      title: "Avg Duration",
      value: fmtDuration(operationalSummary.avg_duration_ms),
      icon: <ClockCircleOutlined />,
      color: "#fa8c16",
    },
    {
      key: "o_tokens",
      title: "Total Tokens",
      value: fmtCount(operationalSummary.total_tokens || 0),
      icon: <ThunderboltOutlined />,
      color: "#722ed1",
    },
  ];

  const costPerDaySeries = useMemo(() => {
    const rows = businessCharts.cost_per_day || [];
    const min = selectedRangeMs.min;
    const max = selectedRangeMs.max;
    const bins = buildBins(min, max, chartStepMs);
    if (!bins.length) return { categories: [], xValues: [], series: [] };

    const seriesDefs = [
      { name: "Total Cost", key: "cost_usd" },
      { name: "LLM Input", key: "input_cost_usd" },
      { name: "LLM Output", key: "output_cost_usd" },
      { name: "Contextualization", key: "contextualization_cost_usd" },
      { name: "Page Minimal", key: "page_cost_minimal_usd" },
      { name: "Page Basic", key: "page_cost_basic_usd" },
      { name: "Page Standard", key: "page_cost_standard_usd" },
    ];
    const bySeries = Object.fromEntries(seriesDefs.map((d) => [d.key, new Array(bins.length).fill(0)]));

    rows.forEach((r) => {
      const ts = toEpochMs(r.ts);
      if (!Number.isFinite(ts) || ts < min || ts > max) return;
      const idx = Math.max(0, Math.min(bins.length - 1, Math.floor((ts - min) / chartStepMs)));
      seriesDefs.forEach((d) => {
        bySeries[d.key][idx] += Number(r[d.key] || 0);
      });
    });

    const series = seriesDefs
      .map((d) => ({ name: d.name, data: bySeries[d.key] }))
      .filter((s, idx) => idx === 0 || s.data.some((v) => v > 0));

    return {
      categories: bins.map((ts) => fmtBinTs(ts)),
      xValues: bins,
      series,
    };
  }, [businessCharts, selectedRangeMs.min, selectedRangeMs.max, chartStepMs]);

  const costPerPageSeries = useMemo(() => {
    const rows = businessCharts.cost_per_page || [];
    const min = selectedRangeMs.min;
    const max = selectedRangeMs.max;
    const bins = buildBins(min, max, chartStepMs);
    if (!bins.length) return { categories: [], xValues: [], series: [] };

    const sums = new Array(bins.length).fill(0);
    const counts = new Array(bins.length).fill(0);
    rows.forEach((r) => {
      const ts = toEpochMs(r.ts);
      if (!Number.isFinite(ts) || ts < min || ts > max) return;
      const idx = Math.max(0, Math.min(bins.length - 1, Math.floor((ts - min) / chartStepMs)));
      const val = Number(r.cost_per_page_usd);
      if (Number.isFinite(val)) {
        sums[idx] += val;
        counts[idx] += 1;
      }
    });
    const data = sums.map((s, i) => (counts[i] > 0 ? s / counts[i] : 0));

    return {
      categories: bins.map((ts) => fmtBinTs(ts)),
      xValues: bins,
      series: [
        {
          name: "Cost / Page (USD)",
          type: "line",
          data,
        },
      ],
    };
  }, [businessCharts, selectedRangeMs.min, selectedRangeMs.max, chartStepMs]);

  const fieldsSeries = useMemo(() => {
    const rows = businessCharts.fields_extracted || [];
    const min = selectedRangeMs.min;
    const max = selectedRangeMs.max;
    const bins = buildBins(min, max, chartStepMs);
    if (!bins.length) return { categories: [], xValues: [], series: [] };

    const fields = new Array(bins.length).fill(0);
    rows.forEach((r) => {
      const ts = toEpochMs(r.ts);
      if (!Number.isFinite(ts) || ts < min || ts > max) return;
      const idx = Math.max(0, Math.min(bins.length - 1, Math.floor((ts - min) / chartStepMs)));
      fields[idx] += Number(r.fields || 0);
    });

    return {
      categories: bins.map((ts) => fmtBinTs(ts)),
      xValues: bins,
      series: [
        {
          name: "Fields Extracted",
          type: "line",
          data: fields,
        },
      ],
    };
  }, [businessCharts, selectedRangeMs.min, selectedRangeMs.max, chartStepMs]);

  const confidenceSeries = useMemo(() => {
    const rows = businessCharts.confidence_after || [];
    if (!rows.length) return { categories: [], xValues: [], series: [] };
    return {
      categories: rows.map((r) => fmtRunLabel(r)),
      xValues: rows.map((r) => toEpochMs(r.ts)),
      series: [
        { name: "Before Validation", data: rows.map((r) => Number(r.confidence_before || 0)) },
        { name: "After Validation", data: rows.map((r) => Number(r.confidence || 0)) },
      ],
    };
  }, [businessCharts]);

  const runsSeries = useMemo(() => {
    const rows = operationalCharts.runs_per_day || [];
    const min = selectedRangeMs.min;
    const max = selectedRangeMs.max;
    const bins = buildBins(min, max, chartStepMs);
    if (!bins.length) return { categories: [], xValues: [], series: [] };

    const runCounts = new Array(bins.length).fill(0);
    rows.forEach((r) => {
      const ts = toEpochMs(r.ts);
      if (!Number.isFinite(ts) || ts < min || ts > max) return;
      const idx = Math.max(0, Math.min(bins.length - 1, Math.floor((ts - min) / chartStepMs)));
      runCounts[idx] += Number(r.run_count || 0);
    });

    return {
      categories: bins.map((ts) => fmtBinTs(ts)),
      xValues: bins,
      series: [
        { name: "Run Count", type: "line", data: runCounts },
      ],
    };
  }, [operationalCharts, selectedRangeMs.min, selectedRangeMs.max, chartStepMs]);

  const pagesSeries = useMemo(() => {
    const rows = operationalCharts.pages_per_day || [];
    if (!rows.length) return { categories: [], xValues: [], series: [] };
    return {
      categories: rows.map((r) => fmtRunLabel(r)),
      xValues: rows.map((r) => toEpochMs(r.ts)),
      series: [
        { name: "Pages Processed", data: rows.map((r) => Number(r.pages || 0)) },
      ],
    };
  }, [operationalCharts]);

  const tokensSeries = useMemo(() => {
    const rows = operationalCharts.tokens_per_day || [];
    if (!rows.length) return { categories: [], xValues: [], series: [] };
    return {
      categories: rows.map((r) => fmtRunLabel(r)),
      xValues: rows.map((r) => toEpochMs(r.ts)),
      series: [
        { name: "Input", data: rows.map((r) => Number(r.input || 0)) },
        { name: "Output", data: rows.map((r) => Number(r.output || 0)) },
        { name: "Contextualization", data: rows.map((r) => Number(r.contextualization || 0)) },
      ],
    };
  }, [operationalCharts]);

  const durationSeries = useMemo(() => {
    const rows = operationalCharts.avg_duration || [];
    if (!rows.length) return { categories: [], xValues: [], series: [] };
    return {
      categories: rows.map((r) => fmtRunLabel(r)),
      xValues: rows.map((r) => toEpochMs(r.ts)),
      series: [
        {
          name: "Duration (s)",
          data: rows.map((r) => Number(r.avg_duration_ms || 0) / 1000),
        },
      ],
    };
  }, [operationalCharts]);

  const rangePresets = [
    { label: "Last 1 h", value: [dayjs().subtract(1, "hour"), dayjs()] },
    { label: "Last 6 h", value: [dayjs().subtract(6, "hour"), dayjs()] },
    { label: "Last 24 h", value: [dayjs().subtract(24, "hour"), dayjs()] },
    { label: "Last 7 d", value: [dayjs().subtract(7, "day"), dayjs()] },
    { label: "Last 30 d", value: [dayjs().subtract(30, "day"), dayjs()] },
  ];

  return (
    <div className="obs-metrics-root">
      <Divider orientation="left" className="obs-section-divider">
        <Title level={5} style={{ margin: 0 }}>Azure Monitor Observability</Title>
      </Divider>

      <div className="obs-range-row">
        <div className="obs-range-label">Date range for metrics, runs, and charts</div>
        <RangePicker
          showTime={{ format: "HH:mm" }}
          format="YYYY-MM-DD HH:mm"
          value={dateRange}
          presets={rangePresets}
          onChange={(range) => {
            if (range && range[0] && range[1]) setDateRange(range);
          }}
          allowClear={false}
          size="small"
        />
      </div>

      <MetricsCards title="Business Metrics" cards={businessCards} loading={summaryLoading} />
      <MetricsCards title="Operational Metrics" cards={operationalCards} loading={summaryLoading} />

      <Card className="obs-runs-card" title="Pipeline Runs" bordered={false} style={{ marginTop: 16 }}>
        <Spin spinning={runsLoading}>
          {runs.length === 0 && !runsLoading ? (
            <Empty description="No runs found in this time range" />
          ) : (
            <Table
              rowKey="run_id"
              columns={runsColumns}
              dataSource={runs}
              size="small"
              pagination={{ pageSize: 5, size: "small" }}
              scroll={{ x: 600 }}
              rowClassName={(row) => (row.run_id === selectedRunId ? "obs-run-row-selected" : "")}
              onRow={(row) => ({
                onClick: () => setSelectedRunId((prev) => (prev === row.run_id ? null : row.run_id)),
                style: { cursor: "pointer" },
              })}
            />
          )}
        </Spin>
      </Card>

      <RunDetailPanel
        runId={selectedRunId}
        runMeta={runs.find((r) => r.run_id === selectedRunId) || null}
        rows={runDetail}
        loading={runDetailLoading}
        onClose={() => setSelectedRunId(null)}
      />

      <Card className="obs-charts-card" title="Business Charts" bordered={false} style={{ marginTop: 16 }}>
        <Spin spinning={chartsLoading}>
          <Row gutter={[16, 16]}>
            <Col xs={24} lg={12}>
              <Card title="Total Cost" className="obs-inner-chart-card" bordered={false}>
                {costPerDaySeries.series.length ? (
                  <AreaChart
                    data={costPerDaySeries.categories.map((month, i) => ({
                      month,
                      submitted: costPerDaySeries.series[0]?.data?.[i] || 0,
                      converted: costPerDaySeries.series[1]?.data?.[i] || 0,
                    }))}
                    seriesOverride={costPerDaySeries.series}
                    categoriesOverride={costPerDaySeries.categories}
                    xValuesOverride={costPerDaySeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    showRunMarkersOnly
                    theme={theme}
                    height={280}
                    yLabelFormatter={(v) => `$${Number(v || 0).toFixed(4)}`}
                    tooltipYFormatter={(v) => `$${Number(v || 0).toFixed(6)}`}
                  />
                ) : (
                  <Empty description="No total cost data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Avg. Cost Per Page (USD)" className="obs-inner-chart-card" bordered={false}>
                {costPerPageSeries.series.length ? (
                  <AreaChart
                    data={costPerPageSeries.categories.map((month, i) => ({
                      month,
                      submitted: costPerPageSeries.series[0]?.data?.[i] || 0,
                      converted: 0,
                    }))}
                    seriesOverride={costPerPageSeries.series}
                    categoriesOverride={costPerPageSeries.categories}
                    xValuesOverride={costPerPageSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    showRunMarkersOnly
                    theme={theme}
                    height={280}
                    yLabelFormatter={(v) => `$${Number(v || 0).toFixed(4)}`}
                    tooltipYFormatter={(v) => `$${Number(v || 0).toFixed(6)}`}
                  />
                ) : (
                  <Empty description="No cost per page data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Total Fields Extracted" className="obs-inner-chart-card" bordered={false}>
                {fieldsSeries.series.length ? (
                  <AreaChart
                    data={fieldsSeries.categories.map((month, i) => ({
                      month,
                      submitted: fieldsSeries.series[0]?.data?.[i] || 0,
                      converted: 0,
                    }))}
                    seriesOverride={fieldsSeries.series}
                    categoriesOverride={fieldsSeries.categories}
                    xValuesOverride={fieldsSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    showRunMarkersOnly
                    theme={theme}
                    height={280}
                  />
                ) : (
                  <Empty description="No fields extracted data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Confidence Before vs After" className="obs-inner-chart-card" bordered={false}>
                {confidenceSeries.series.length ? (
                  <ColumnChart
                    series={confidenceSeries.series}
                    categories={confidenceSeries.categories}
                    xValues={confidenceSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    yAxisTitle="Confidence (0-1)"
                    theme={theme}
                    height={300}
                  />
                ) : (
                  <Empty description="No confidence data" />
                )}
              </Card>
            </Col>
          </Row>
        </Spin>
      </Card>

      <Card className="obs-charts-card" title="Operational Charts" bordered={false} style={{ marginTop: 16 }}>
        <Spin spinning={chartsLoading}>
          <Row gutter={[16, 16]}>
            <Col xs={24} lg={12}>
              <Card title="Runs" className="obs-inner-chart-card" bordered={false}>
                {runsSeries.series.length ? (
                  <AreaChart
                    data={runsSeries.categories.map((month, i) => ({
                      month,
                      submitted: runsSeries.series[0]?.data?.[i] || 0,
                      converted: 0,
                    }))}
                    seriesOverride={runsSeries.series}
                    categoriesOverride={runsSeries.categories}
                    xValuesOverride={runsSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    showRunMarkersOnly
                    theme={theme}
                    height={280}
                  />
                ) : (
                  <Empty description="No runs data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Total Pages Processed" className="obs-inner-chart-card" bordered={false}>
                {pagesSeries.series.length ? (
                  <AreaChart
                    data={pagesSeries.categories.map((month, i) => ({
                      month,
                      submitted: pagesSeries.series[0].data[i],
                      converted: 0,
                    }))}
                    seriesOverride={pagesSeries.series}
                    categoriesOverride={pagesSeries.categories}
                    xValuesOverride={pagesSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    padWithZeroRange
                    theme={theme}
                    height={280}
                    tooltipYFormatter={(v) => `${Math.round(Number(v || 0)).toLocaleString()} pages`}
                  />
                ) : (
                  <Empty description="No pages processed data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Token Usage" className="obs-inner-chart-card" bordered={false}>
                {tokensSeries.series.length ? (
                  <AreaChart
                    data={tokensSeries.categories.map((month, i) => ({
                      month,
                      submitted: tokensSeries.series[0].data[i],
                      converted: tokensSeries.series[1].data[i],
                    }))}
                    seriesOverride={tokensSeries.series}
                    categoriesOverride={tokensSeries.categories}
                    xValuesOverride={tokensSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    padWithZeroRange
                    theme={theme}
                    height={280}
                    tooltipYFormatter={(v) => `${Math.round(Number(v || 0)).toLocaleString()} tokens`}
                  />
                ) : (
                  <Empty description="No token usage data" />
                )}
              </Card>
            </Col>

            <Col xs={24} lg={12}>
              <Card title="Duration of Runs" className="obs-inner-chart-card" bordered={false}>
                {durationSeries.series.length ? (
                  <AreaChart
                    data={durationSeries.categories.map((month, i) => ({
                      month,
                      submitted: durationSeries.series[0].data[i],
                      converted: 0,
                    }))}
                    seriesOverride={durationSeries.series}
                    categoriesOverride={durationSeries.categories}
                    xValuesOverride={durationSeries.xValues}
                    xMin={selectedRangeMs.min}
                    xMax={selectedRangeMs.max}
                    padWithZeroRange
                    theme={theme}
                    height={280}
                    yLabelFormatter={(v) => `${Number(v || 0).toFixed(1)} s`}
                    tooltipYFormatter={(v) => `${Number(v || 0).toFixed(2)} s`}
                  />
                ) : (
                  <Empty description="No duration data" />
                )}
              </Card>
            </Col>
          </Row>
        </Spin>
      </Card>
    </div>
  );
};

export default ObservabilityMetrics;
