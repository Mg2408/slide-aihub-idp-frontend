import React, { useState } from "react";
import { Card, Select, Space, Typography } from "antd";

import ObservabilityMetrics from "./ObservabilityMetrics";

const { Title, Text } = Typography;

const TEMPLATE_OPTIONS = [
  { label: "Wind Mitigation", value: "wind_mit" },
  { label: "Mortgage", value: "mortgage" },
];

const ObservabilityDashboard = () => {
  const [template, setTemplate] = useState("wind_mit");

  return (
    <div>
      <Card bordered={false} style={{ marginBottom: 12 }}>
        <Space direction="vertical" size={8} style={{ width: "100%" }}>
          <Title level={4} style={{ margin: 0 }}>
            Dashboard
          </Title>
          <Text type="secondary">
            Select a template to view observability metrics and run drilldowns.
          </Text>
          <Select
            style={{ maxWidth: 320 }}
            options={TEMPLATE_OPTIONS}
            value={template}
            onChange={setTemplate}
            placeholder="Select template"
          />
        </Space>
      </Card>

      <ObservabilityMetrics template={template} />
    </div>
  );
};

export default ObservabilityDashboard;
