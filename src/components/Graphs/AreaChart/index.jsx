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

const AreaChart = ({
  data = [],
  theme = "light",
  height = 260,
  seriesOverride,
  categoriesOverride,
  xValuesOverride,
  xMin,
  xMax,
  padWithZeroRange = false,
  showRunMarkersOnly = false,
  yLabelFormatter,
  tooltipYFormatter,
}) => {
  const months = data.map((item) => item.month);
  const submitted = data.map((item) => Number(item.submitted || 0));
  const converted = data.map((item) => Number(item.converted || 0));

  const series =
    seriesOverride ||
    [
      {
        name: "Submitted",
        data: submitted,
      },
      {
        name: "High confidence docs",
        data: converted,
      },
    ];

  const categories = categoriesOverride || months;
  const xValues = xValuesOverride || [];
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

        const anchored = [...sorted];
        if (!anchored.length) {
          anchored.push({ x: minX, y: 0 });
          anchored.push({ x: maxX, y: 0 });
          return { ...item, data: anchored };
        }
        if (anchored[0].x > minX) anchored.unshift({ x: minX, y: 0 });
        if (anchored[anchored.length - 1].x < maxX) anchored.push({ x: maxX, y: 0 });
        return { ...item, data: anchored };
      })
    : baseSeries;

  const markerIndexes = showRunMarkersOnly
    ? ((chartSeries[0]?.data || [])
        .map((p, idx) => ({ idx, y: Number((typeof p === "object" ? p.y : p) || 0) }))
        .filter((p) => p.y > 0)
        .map((p) => p.idx))
    : [];

  const options = {
    chart: {
      type: "area",
      height,
      stacked: false,
      toolbar: {
        show: false,
      },
    },
    colors:
      theme === "dark"
        ? ["#36AFFA", "#63D471", "#FF8F6B", "#C3A6FF", "#6EE7D7"]
        : ["#1677ff", "#52c41a", "#fa8c16", "#722ed1", "#13c2c2"],
    dataLabels: {
      enabled: false,
    },
    stroke: {
      curve: "smooth",
      width: 2,
    },
    markers: showRunMarkersOnly
      ? {
          size: 0,
          strokeWidth: 2,
          discrete: markerIndexes.map((dataPointIndex) => ({
            seriesIndex: 0,
            dataPointIndex,
            size: 5,
            fillColor: theme === "dark" ? "#36AFFA" : "#1677ff",
            strokeColor: "#ffffff",
          })),
          hover: {
            size: 6,
          },
        }
      : {
          size: 4,
          strokeWidth: 2,
          hover: {
            size: 6,
          },
        },
    fill: {
      type: "gradient",
      gradient: {
        opacityFrom: 0.45,
        opacityTo: 0.12,
      },
    },
    legend: {
      position: "bottom",
      horizontalAlign: "center",
    },
    grid: {
      borderColor: theme === "dark" ? "#4f4f4f" : "#e5e7eb",
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
          },
          ...(Number.isFinite(minX) && Number.isFinite(maxX)
            ? { min: minX, max: maxX }
            : {}),
        }
      : {
          categories,
          labels: {
            rotate: -35,
          },
        },
    yaxis: {
      min: 0,
      forceNiceScale: true,
      labels: {
        formatter: yLabelFormatter || ((value) => Math.round(value)),
      },
    },
    tooltip: {
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
        formatter:
          tooltipYFormatter ||
          ((value) => Math.round(Number(value || 0)).toLocaleString()),
      },
    },
  };

  return <ReactApexChart options={options} series={chartSeries} type="area" height={height} />;
};

export default AreaChart;
