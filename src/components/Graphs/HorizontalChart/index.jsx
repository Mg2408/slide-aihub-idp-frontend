import React from "react";
import ReactApexChart from "react-apexcharts";

const HorizontalChart = ({
  series = [],
  categories = [],
  theme = "light",
  height = 320,
}) => {
  const options = {
    chart: {
      type: "bar",
      height,
      toolbar: {
        show: false,
      },
    },
    colors: theme === "dark" ? ["#36AFFA", "#ff9f43"] : ["#1677ff", "#fa8c16"],
    plotOptions: {
      bar: {
        horizontal: true,
        borderRadius: 4,
        barHeight: "56%",
      },
    },
    dataLabels: {
      enabled: false,
    },
    stroke: {
      show: true,
      width: 1,
      colors: ["transparent"],
    },
    xaxis: {
      categories,
      labels: {
        formatter: (value) => Math.round(value),
      },
    },
    legend: {
      position: "top",
      horizontalAlign: "left",
    },
    grid: {
      borderColor: theme === "dark" ? "#4f4f4f" : "#e5e7eb",
    },
    tooltip: {
      shared: true,
      intersect: false,
    },
  };

  return <ReactApexChart options={options} series={series} type="bar" height={height} />;
};

export default HorizontalChart;
