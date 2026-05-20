import React from "react";
import ReactApexChart from "react-apexcharts";

/**
 * DonutChart — proportional donut chart for cost breakdown.
 *
 * Props
 * -----
 * series   number[]   Slice values (e.g. costs in USD).
 * labels   string[]   Slice labels matching series order.
 * height   number     Chart height in px.  Default 300.
 * theme    "light" | "dark"
 */
const DonutChart = ({
  series = [],
  labels = [],
  height = 300,
  theme = "light",
}) => {
  const isDark = theme === "dark";

  const options = {
    chart: {
      type: "donut",
      height,
      toolbar: { show: false },
      animations: { enabled: false },
    },
    labels,
    colors: isDark
      ? ["#36AFFA", "#ff9f43", "#a29bfe", "#55efc4", "#fd79a8", "#fdcb6e"]
      : ["#1677ff", "#fa8c16", "#722ed1", "#52c41a", "#eb2f96", "#faad14"],
    dataLabels: {
      enabled: true,
      formatter: (val) => `${val.toFixed(1)}%`,
      style: {
        fontSize: "11px",
      },
    },
    plotOptions: {
      pie: {
        donut: {
          size: "60%",
          labels: {
            show: true,
            total: {
              show: true,
              label: "Total",
              formatter: (w) => {
                const total = w.globals.seriesTotals.reduce((a, b) => a + b, 0);
                return `$${total.toFixed(4)}`;
              },
            },
          },
        },
      },
    },
    legend: {
      position: "bottom",
      horizontalAlign: "center",
      fontSize: "12px",
    },
    tooltip: {
      y: {
        formatter: (val) => `$${val.toFixed(6)}`,
      },
    },
    theme: {
      mode: isDark ? "dark" : "light",
    },
  };

  return (
    <ReactApexChart
      type="donut"
      series={series}
      options={options}
      height={height}
    />
  );
};

export default DonutChart;
