import React from "react";
import ReactApexChart from "react-apexcharts";

const toEpochMs = (value) => {
  if (value == null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    const abs = Math.abs(value);
    if (abs < 1e11) return Math.round(value * 1000);
    if (abs < 1e14) return Math.round(value);
    if (abs < 1e17) return Math.round(value / 1000);
    return Math.round(value / 1000000);
  }
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return toEpochMs(numeric);
  const parsed = new Date(value).getTime();
  return Number.isNaN(parsed) ? null : parsed;
};

/**
 * LineChart — mixed area + line chart with optional dual Y-axis.
 *
 * Designed for the run-trend chart: the first series is rendered as a filled
 * area (run count) and the second as a plain line on a secondary right-hand
 * Y-axis (avg duration).
 *
 * Props
 * -----
 * series       [{name, type?, data[]}]  ApexCharts series array.
 *               type per-series defaults to "line"; pass "area" for fill.
 * categories   string[]               X-axis labels.
 * yaxis        object[]               ApexCharts yaxis config array.
 *               Omit to get a single auto-scaled axis.
 * height       number                 Chart height in px.  Default 280.
 * theme        "light" | "dark"
 */
const LineChart = ({
  series = [],
  categories = [],
  xValues = [],
  xMin,
  xMax,
  padWithZeroRange = false,
  yaxis,
  height = 280,
  theme = "light",
  yLabelFormatter,
  tooltipYFormatter,
}) => {
  const isDark = theme === "dark";
  const normalizedX = (xValues || []).map((x) => toEpochMs(x));
  const minX = toEpochMs(xMin);
  const maxX = toEpochMs(xMax);
  const hasDatetime =
    Array.isArray(normalizedX) &&
    normalizedX.length === categories.length &&
    normalizedX.every((value) => Number.isFinite(value));

  const baseSeries = hasDatetime
    ? series.map((item) => ({
        ...item,
        data: (item.data || []).map((y, index) => ({
          x: normalizedX[index],
          y,
        })),
      }))
    : series;

  const chartSeries = hasDatetime
    ? baseSeries.map((item) => {
        const sorted = [...(item.data || [])].sort((a, b) => Number(a.x) - Number(b.x));
        if (!padWithZeroRange || !Number.isFinite(minX) || !Number.isFinite(maxX)) {
          return { ...item, data: sorted };
        }

        const span = Math.max(maxX - minX, 1);
        const edgeMs = Math.max(Math.floor(span * 0.01), 60 * 1000);
        const impulse = [];

        impulse.push({ x: minX, y: 0 });
        sorted.forEach((point) => {
          const px = Number(point.x);
          const py = Number(point.y || 0);
          const before = Math.max(minX, px - edgeMs);
          const after = Math.min(maxX, px + edgeMs);
          impulse.push({ x: before, y: 0 });
          impulse.push({ x: px, y: py });
          impulse.push({ x: after, y: 0 });
        });
        impulse.push({ x: maxX, y: 0 });

        const deduped = [];
        impulse
          .sort((a, b) => Number(a.x) - Number(b.x))
          .forEach((p) => {
            const prev = deduped[deduped.length - 1];
            if (!prev || prev.x !== p.x || prev.y !== p.y) deduped.push(p);
          });

        return { ...item, data: deduped };
      })
    : baseSeries;

  const yValues = chartSeries
    .flatMap((s) => (s.data || []).map((p) => (typeof p === "object" ? Number(p.y) : Number(p))))
    .filter((v) => Number.isFinite(v));
  const yMin = yValues.length ? Math.min(...yValues) : 0;
  const yMax = yValues.length ? Math.max(...yValues) : 0;
  const autoYMin = yValues.length && yMin >= 0
    ? -Math.max(yMax * 0.05, yMax > 0 ? 1e-9 : 0.1)
    : undefined;

  const options = {
    chart: {
      type: "line",
      height,
      toolbar: { show: false },
      zoom: { enabled: false },
      animations: { enabled: false },
    },
    stroke: {
      curve: padWithZeroRange ? "straight" : "smooth",
      width: chartSeries.map((s) => (s.type === "area" ? 2 : 2)),
    },
    fill: {
      type: chartSeries.map((s) => (s.type === "area" ? "gradient" : "solid")),
      gradient: {
        opacityFrom: 0.4,
        opacityTo: 0.05,
      },
      opacity: chartSeries.map((s) => (s.type === "area" ? 1 : 0)),
    },
    colors: isDark
      ? ["#36AFFA", "#ff9f43", "#a29bfe", "#55efc4", "#ffe066", "#ff6b6b", "#74c0fc"]
      : ["#1677ff", "#fa8c16", "#722ed1", "#52c41a", "#13c2c2", "#eb2f96", "#2f54eb"],
    markers: {
      size: 3,
      strokeWidth: 1,
      hover: { size: 5 },
    },
    dataLabels: {
      enabled: categories.length > 0 && categories.length <= 6,
      formatter: (val, { seriesIndex }) => {
        if (yLabelFormatter) {
          return yLabelFormatter(val, { seriesIndex, series });
        }
        const name = series?.[seriesIndex]?.name || "";
        if (name.toLowerCase().includes("duration")) {
          return `${Math.round(Number(val || 0)).toLocaleString()} ms`;
        }
        return `${Math.round(Number(val || 0)).toLocaleString()}`;
      },
      style: {
        fontSize: "10px",
        fontWeight: 500,
      },
      background: {
        enabled: true,
        opacity: 0.8,
        borderRadius: 4,
      },
    },
    legend: {
      position: "bottom",
      horizontalAlign: "center",
    },
    grid: {
      borderColor: isDark ? "#4f4f4f" : "#e5e7eb",
      strokeDashArray: 3,
    },
    xaxis: hasDatetime
      ? {
          type: "datetime",
          labels: {
            datetimeUTC: false,
            formatter: (value) => {
              const ts = toEpochMs(value);
              if (!Number.isFinite(ts)) return String(value ?? "");
              return new Date(ts).toLocaleString(undefined, {
                month: "short",
                day: "2-digit",
                hour: "2-digit",
                minute: "2-digit",
                hour12: false,
              });
            },
            style: {
              colors: isDark ? "#ccc" : "#666",
              fontSize: "11px",
            },
          },
          axisBorder: { show: false },
          axisTicks: { show: false },
          ...(Number.isFinite(minX) && Number.isFinite(maxX)
            ? { min: minX, max: maxX }
            : {}),
        }
      : {
          categories,
          labels: {
            rotate: -35,
            style: {
              colors: isDark ? "#ccc" : "#666",
              fontSize: "11px",
            },
          },
          axisBorder: { show: false },
          axisTicks: { show: false },
        },
    tooltip: {
      shared: true,
      intersect: false,
      x: {
        formatter: (value) => {
          if (!hasDatetime) return String(value ?? "");
          const ts = toEpochMs(value);
          if (!Number.isFinite(ts)) {
            return String(value ?? "");
          }
          const d = new Date(ts);
          return Number.isNaN(d.getTime())
            ? String(value ?? "")
            : d.toLocaleString(undefined, {
                year: "numeric",
                month: "short",
                day: "numeric",
                hour: "2-digit",
                minute: "2-digit",
                second: "2-digit",
                hour12: false,
              });
        },
      },
      y: {
        formatter: (value, { seriesIndex }) => {
          if (tooltipYFormatter) {
            return tooltipYFormatter(value, { seriesIndex, series });
          }
          const name = series?.[seriesIndex]?.name || "";
          const numeric = Number(value || 0);
          if (name.toLowerCase().includes("duration")) {
            return `${Math.round(numeric).toLocaleString()} ms`;
          }
          return `${Math.round(numeric).toLocaleString()}`;
        },
      },
    },
    ...(yaxis
      ? { yaxis }
      : {
          yaxis: {
            ...(Number.isFinite(autoYMin) ? { min: autoYMin } : {}),
            ...(yLabelFormatter
              ? {
                  labels: {
                    formatter: yLabelFormatter,
                  },
                }
              : {}),
          },
        }),
  };

  return (
    <ReactApexChart
      type="line"
      series={chartSeries}
      options={options}
      height={height}
    />
  );
};

export default LineChart;
