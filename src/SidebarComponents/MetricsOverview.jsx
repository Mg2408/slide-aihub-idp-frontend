import React, { useMemo } from "react";
import { Card, Col, Empty, Row } from "antd";

import AreaChart from "../components/Graphs/AreaChart";
import HorizontalChart from "../components/Graphs/HorizontalChart";
import useMetaData from "../context/metaData";
import { buildDashboardMetrics } from "../utils/dashboardMetrics";
import "./MetricsOverview.css";

const statCards = (summary) => [
  {
    key: "documents",
    title: "Documents Processed",
    value: summary.totalDocuments,
  },
  {
    key: "fields",
    title: "Fields Extracted",
    value: summary.totalFields,
  },
  {
    key: "confidence",
    title: "Avg Confidence",
    value: `${summary.averageConfidence.toFixed(1)}%`,
    helper: `${summary.highConfidenceRate.toFixed(1)}% high-confidence fields`,
  },
];

const MetricsOverview = ({ submissions = [], className = "" }) => {
  const { theme } = useMetaData();

  const metrics = useMemo(() => buildDashboardMetrics(submissions), [submissions]);

  return (
    <div className={`metrics-overview ${className}`.trim()}>
      <Row gutter={[12, 12]}>
        {statCards(metrics.summary).map((stat) => (
          <Col xs={24} sm={12} lg={8} key={stat.key}>
            <Card className="metrics-stat-card" bordered={false}>
              <div className="metrics-stat-title">{stat.title}</div>
              <div className="metrics-stat-value">{stat.value}</div>
              {stat.helper ? <div className="metrics-stat-helper">{stat.helper}</div> : null}
            </Card>
          </Col>
        ))}
      </Row>

      <Row gutter={[12, 12]} style={{ marginTop: 8 }}>
        <Col xs={24} lg={12}>
          <Card title="Monthly Extraction Trend" className="metrics-chart-card" bordered={false}>
            {metrics.monthlyData.length ? (
              <AreaChart data={metrics.monthlyData} theme={theme} height={280} />
            ) : (
              <Empty description="No monthly trend data" />
            )}
          </Card>
        </Col>
        <Col xs={24} lg={12}>
          <Card
            title="Top Fields By Confidence"
            className="metrics-chart-card"
            bordered={false}
          >
            {metrics.fieldChart.categories.length ? (
              <HorizontalChart
                categories={metrics.fieldChart.categories}
                series={metrics.fieldChart.series}
                theme={theme}
                height={320}
              />
            ) : (
              <Empty description="No field confidence data" />
            )}
          </Card>
        </Col>
      </Row>
    </div>
  );
};

export default MetricsOverview;
