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
 * ColumnChart — grouped vertical bar chart.
 *
 * Used for the confidence before-vs-after chart: each time bucket has two
 * side-by-side bars so the validation improvement is immediately visible.
 *
 * Props
 * -----
 * series      [{name, data[]}]   ApexCharts series array.
 * categories  string[]           X-axis bucket labels.
 * yAxisTitle  string             Optional left Y-axis title.
 * height      number             Chart height in px.  Default 300.
 * theme       "light" | "dark"
 */
const ColumnChart = ({
  series = [],
  categories = [],
  xValues = [],
  xMin,
  xMax,
  yAxisTitle = "",
  height = 300,
  theme = "light",
}) => {
  const isDark = theme === "dark";
  const normalizedX = (xValues || []).map((x) => toEpochMs(x));
  const minX = toEpochMs(xMin);
  const maxX = toEpochMs(xMax);
  const hasDatetime =
    Array.isArray(normalizedX) &&
    normalizedX.length === categories.length &&
    normalizedX.every((value) => Number.isFinite(value));

  const chartSeries = hasDatetime
    ? series.map((item) => ({
        ...item,
        data: (item.data || []).map((y, index) => ({
          x: normalizedX[index],
          y,
        })),
      }))
    : series;

  const options = {
    chart: {
      type: "bar",
      height,
      toolbar: { show: false },
      animations: { enabled: false },
    },
    plotOptions: {
      bar: {
        horizontal: false,
        columnWidth: "55%",
        borderRadius: 3,
      },
    },
    colors: isDark
      ? ["#6c7ae0", "#36AFFA"]
      : ["#aab8ff", "#1677ff"],
    dataLabels: { enabled: false },
    stroke: {
      show: true,
      width: 1,
      colors: ["transparent"],
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
    yaxis: {
      title: yAxisTitle ? { text: yAxisTitle } : undefined,
      labels: {
        formatter: (val) => (val !== null && val !== undefined ? val.toFixed(1) : ""),
      },
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
        formatter: (val) =>
          val !== null && val !== undefined ? `${(val * 100).toFixed(1)}%` : "—",
      },
    },
  };

  return (
    <ReactApexChart
      type="bar"
      series={chartSeries}
      options={options}
      height={height}
    />
  );
};

export default ColumnChart;
