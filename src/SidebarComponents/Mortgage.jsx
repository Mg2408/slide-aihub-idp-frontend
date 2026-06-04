import React, { useEffect, useRef, useState } from "react";
import {
  Table,
  Button,
  Row,
  Col,
  Modal,
  Upload,
  message,
  Card,
  List,
  Tag,
  Input,
  Switch,
} from "antd";
import {
  UploadOutlined,
  InfoCircleOutlined,
  DownloadOutlined,
  FileExcelOutlined,
  FilePdfOutlined,
} from "@ant-design/icons";
import XLSX from "xlsx-js-style";

import "./Dashboard.css";
import "./Table.css";

import { buildAiExtractUrl } from "../config/apiBase";
import { TableContainer } from "../styles/components/TableComponent";
import { Container } from "../styles/components/Layout";
import useMetaData from "../context/metaData";
import {
  getDocumentName,
  getExtractionViewModel,
  getOwnerName,
} from "../utils/extractionResponse";

/* ─────────────────────────────────────────
   Constants
   ─── Aligned with backend JSON response ───
   Address_of_Mortgagee_Company.valueObject keys: PO_Box, City, State, ZIP
   Policies[].valueObject keys: Policy_Number, Loan_Number, Borrower_Name, Borrowers_Address
───────────────────────────────────────── */

// Top-level non-policy fields — merged across all policy rows in Excel preview
const TOP_LEVEL_MERGE_HEADERS = new Set([
  "Document Name",
  "Effective Date",
  "Current Mortgagee Company",
  "Mortgage Clause",
  "Mortgagee PO Box",   // ← was "Mortgagee Street"; JSON uses PO_Box
  "Mortgagee City",
  "Mortgagee State",
  "Mortgagee ZIP",
]);

// Policy sub-fields in display order — keys match JSON response
const POLICY_FIELDS = [
  { label: "Policy Number", key: "Policy_Number" },
  { label: "Loan Number", key: "Loan_Number" },
  { label: "Borrower Name", key: "Borrower_Name" }, // ← was key:"Borrowers"
  { label: "Payee Position / Rank", key: "Payee_Position_or_Rank" },
];

// Address object sub-fields — keys match Address_of_Mortgagee_Company.valueObject
const ADDRESS_SUB_FIELDS = [
  { label: "Mortgage Clause", key: "Mortgage_Clause" },
  { label: "PO Box", key: "PO_Box" }, // ← was { label:"Street", key:"Street" }
  { label: "City", key: "City" },
  { label: "State", key: "State" },
  { label: "ZIP", key: "ZIP" },
];

/* ─────────────────────────────────────────
   Helpers
───────────────────────────────────────── */

// Returns the display value from any field type (string, date, number)
const getFieldValue = (field) =>
  field?.valueString ?? field?.valueDate ?? field?.valueNumber ?? "";

// Normalise a field group name for matching regardless of casing/spaces/underscores
const normKey = (name) => name?.toLowerCase().replace(/[\s_]/g, "") ?? "";

// ── Shared confidence + page helpers ────────────────────────────────────────

/**
 * Reads confidence directly from a field object.
 * Returns 0–100 integer, or null if not present.
 */
const getConfValue = (field) =>
  field?.confidence != null ? Math.round(field.confidence * 100) : null;

/**
 * Maps a 0–100 confidence integer to a colour.
 * ≥90 → green  |  ≥70 → orange  |  <70 → red
 */
const getConfColor = (conf) =>
  conf == null ? null : conf >= 90 ? "#52bb2c" : conf >= 70 ? "#d37f30" : "#cf3e4a";

/**
 * Parses the page number from a source string like
 * "D(1,6.04,15.28,...)"  →  "1"
 * Works for multi-source strings separated by ";"
 */
const getPageFromSource = (source) => {
  if (!source) return null;
  const match = String(source).match(/D\((\d+)/);
  return match ? match[1] : null;
};

/* ─────────────────────────────────────────
   Build expanded rows for Excel
   Columns (aligned to JSON response):
     Document Name | Effective Date | Current Mortgagee Company |
     Mortgagee PO Box | Mortgagee City | Mortgagee State | Mortgagee ZIP |
     Policy Number | Loan Number | Borrower Name | Borrowers Address | Reference
───────────────────────────────────────── */
const buildExpandedExcelData = (json, documentName) => {
  if (!json) return { headers: [], rows: [], rowCount: 0 };

  const policies = json?.Policies?.valueArray || [];
  const addrObj = json?.Address_of_Mortgagee_Company?.valueObject || {};

  const headers = [
    "Document Name",
    "Effective Date",
    "Current Mortgagee Company",
    "Mortgage Clause",
    "Mortgagee PO Box",
    "Mortgagee City",
    "Mortgagee State",
    "Mortgagee ZIP",
    "Policy Number",
    "Loan Number",
    "Borrower Name",
    "Payee Position / Rank",
    "Reference",
  ];

  const baseRow = {
    "Document Name": documentName,
    "Effective Date": getFieldValue(json?.Effective_Date),
    "Current Mortgagee Company": getFieldValue(json?.Current_Mortgagee_Company),
    "Mortgage Clause": getFieldValue(addrObj.Mortgage_Clause),
    "Mortgagee PO Box": getFieldValue(addrObj.PO_Box),  // ← was addrObj.Street
    "Mortgagee City": getFieldValue(addrObj.City),
    "Mortgagee State": getFieldValue(addrObj.State),
    "Mortgagee ZIP": getFieldValue(addrObj.ZIP),
  };

  const rows =
    policies.length > 0
      ? policies.map((p) => {
        const obj = p.valueObject || {};
        const page = getPageFromSource(obj.Policy_Number?.source ?? "");
        return {
          ...baseRow,
          "Policy Number": getFieldValue(obj.Policy_Number),
          "Loan Number": getFieldValue(obj.Loan_Number),
          "Borrower Name": getFieldValue(obj.Borrower_Name),
          "Payee Position / Rank": getFieldValue(obj.Payee_Position_or_Rank),
          "Reference": page ? `Page: ${page}` : "",
        };
      })
      : [
        {
          ...baseRow,
          "Policy Number": "",
          "Loan Number": "",
          "Borrower Name": "",
          "Payee Position / Rank": "",
          "Reference": "",
        },
      ];

  return { headers, rows, rowCount: rows.length };
};

/* ─────────────────────────────────────────
   Excel styles
───────────────────────────────────────── */
const headerStyle = () => ({
  font: { bold: true, color: { rgb: "FFFFFF" } },
  fill: { fgColor: { rgb: "217346" } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
  border: {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" },
  },
});

const cellStyle = () => ({
  alignment: { vertical: "top", wrapText: true },
  border: {
    top: { style: "thin" },
    bottom: { style: "thin" },
    left: { style: "thin" },
    right: { style: "thin" },
  },
});

/* ─────────────────────────────────────────
   Download Excel
───────────────────────────────────────── */
const downloadMatrixExcel = (json, documentName) => {
  if (!json) return;

  const { headers, rows, rowCount } = buildExpandedExcelData(json, documentName);
  const workbook = XLSX.utils.book_new();
  const worksheet = {};

  // Header row (row 0)
  headers.forEach((header, colIndex) => {
    worksheet[XLSX.utils.encode_cell({ r: 0, c: colIndex })] = {
      v: header, s: headerStyle(),
    };
  });

  // Data rows (rows 1…rowCount)
  rows.forEach((row, rowIndex) => {
    const sheetRow = rowIndex + 1;
    headers.forEach((header, colIndex) => {
      worksheet[XLSX.utils.encode_cell({ r: sheetRow, c: colIndex })] = {
        v: row[header] ?? "", s: cellStyle(),
      };
    });
  });

  // Merge top-level fields across all policy rows
  const merges = [];
  if (rowCount > 1) {
    headers.forEach((header, colIndex) => {
      if (TOP_LEVEL_MERGE_HEADERS.has(header)) {
        merges.push({ s: { r: 1, c: colIndex }, e: { r: rowCount, c: colIndex } });
      }
    });
  }

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowCount, 1), c: headers.length - 1 },
  });
  worksheet["!merges"] = merges;
  worksheet["!cols"] = headers.map(() => ({ wch: 28 }));

  XLSX.utils.book_append_sheet(workbook, worksheet, "Mortgage Extracted Data");
  XLSX.writeFile(workbook, `${documentName}_extracted.xlsx`);
};

/* ─────────────────────────────────────────
   FieldRow — reusable field renderer
   Works for valueString, valueDate, valueNumber
───────────────────────────────────────── */
const FieldRow = ({ label, field, style = {} }) => {
  if (!field) return null;
  const value = getFieldValue(field) || "-";
  const conf = getConfValue(field);
  const confColor = getConfColor(conf);
  const page = getPageFromSource(field.source);

  return (
    <Row gutter={[16, 8]} style={{ marginBottom: 8, ...style }}>
      <Col span={14}>
        {value.length > 120 ? (
          <Input.TextArea
            value={value}
            readOnly
            autoSize={{ minRows: 2, maxRows: 6 }}
          />
        ) : (
          <Input
            value={value}
            readOnly
            addonBefore={
              <span style={{ fontWeight: 600, width: 145, display: "inline-block" }}>
                {label}
              </span>
            }
          />
        )}
      </Col>
      <Col span={10} style={{ textAlign: "right" }}>
        {conf != null && <Tag color={confColor}>Confidence: {conf}%</Tag>}
        {page && <Tag>Page: {page}</Tag>}
      </Col>
    </Row>
  );
};

/* ─────────────────────────────────────────
   AddressObjectRenderer
   Renders PO Box / City / State / ZIP
   each with its own confidence + page tag
   Keys match Address_of_Mortgagee_Company.valueObject
───────────────────────────────────────── */
const AddressObjectRenderer = ({ valueObject }) => {
  if (!valueObject) return null;
  return (
    <div
      style={{
        marginLeft: 16,
        paddingLeft: 12,
        borderLeft: "3px solid #e6f4ff",
        marginTop: 4,
      }}
    >
      {ADDRESS_SUB_FIELDS.map(({ label, key }) => {
        const field = valueObject[key];
        if (!field) return null;
        return <FieldRow key={key} label={label} field={field} style={{ marginBottom: 6 }} />;
      })}
    </div>
  );
};

/* ─────────────────────────────────────────
   Table Wrapper
───────────────────────────────────────── */
const MyTableComponent = ({ columns, dataSource, loading, selectedSubmissionId }) => {
  const { theme } = useMetaData();
  return (
    <TableContainer theme={theme}>
      <Table
        rowKey="key"
        className="custom-table-header"
        columns={columns}
        dataSource={dataSource}
        loading={loading}
        pagination={{ pageSize: 5 }}
        tableLayout="fixed"
        scroll={{ x: 1200 }}
        onRow={(record) => ({
          style:
            record.key === selectedSubmissionId
              ? { backgroundColor: "#e6f4ff", transition: "background-color 0.3s ease" }
              : {},
        })}
        components={{
          header: {
            cell: (props) => (
              <th {...props} style={{ color: "#fff", fontFamily: "inherit" }} />
            ),
          },
        }}
      />
    </TableContainer>
  );
};

const scrollCellStyle = {
  maxHeight: 55,
  overflowX: "auto",
  whiteSpace: "normal",
  wordBreak: "break-word",
};

/* ─────────────────────────────────────────
   Main Component
───────────────────────────────────────── */
const Mortgage = () => {
  const [apiData, setApiData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [excelModalOpen, setExcelModalOpen] = useState(false);
  const [selectedExcelData, setSelectedExcelData] = useState(null);
  const [selectedSubmissionId, setSelectedSubmissionId] = useState(null);
  const [fileList, setFileList] = useState([]);
  const [preprocess, setPreprocess] = useState(false);
  const hasFetchedRef = useRef(false);
  const detailsRef = useRef(null);

  useEffect(() => {
    if (hasFetchedRef.current) return;
    const fetchDocuments = async () => {
      try {
        setLoading(true);
        const response = await fetch(
          buildAiExtractUrl("/api/get_extracted_documents")
        );
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        setApiData(result.submission_list || []);
        hasFetchedRef.current = true;
      } catch (error) {
        console.error(error);
        // message.error("Failed to fetch data");
      } finally {
        setLoading(false);
      }
    };
    fetchDocuments();
  }, []);

  useEffect(() => {
    if (selectedSubmissionId && detailsRef.current) {
      setTimeout(() => {
        detailsRef.current.scrollIntoView({ behavior: "smooth", block: "start" });
      }, 100);
    }
  }, [selectedSubmissionId]);

  const getDocumentNameFromUri = (uri) => {
    if (!uri) return null;
    try {
      const rawFileName = uri.split("?")[0].split("/").pop();
      return decodeURIComponent(rawFileName).replace(/\s+/g, "");
    } catch {
      return null;
    }
  };

  const tableData = apiData.map((item) => {
    const extractionResponse = item.llm_response || {};
    const documentName =
      getDocumentNameFromUri(item.document_uri) ||
      getDocumentName(extractionResponse, [
        item.file_name,
        item.document_uri?.split("/").pop(),
        item.submission_id,
      ]);
    return {
      key: item.submission_id,
      submission: item.submission_id?.slice(0, 8),
      submittedBy: getOwnerName(extractionResponse, ["-"]),
      document: documentName,
      documentName,
      date: item.last_modified
        ? new Date(item.last_modified).toLocaleDateString()
        : "-",
      source: item.document_uri,
      json: extractionResponse,
      output: item.submission_id,
    };
  });

  const getColumnFilters = (dataIndex) => {
    const uniqueValues = Array.from(
      new Set(tableData.map((row) => row[dataIndex]).filter(Boolean))
    );
    return uniqueValues.map((value) => ({
      text: String(value).length > 40 ? `${String(value).slice(0, 40)}...` : value,
      value,
    }));
  };

  const columns = [
    { title: "SubmissionID", dataIndex: "submission", width: 100 },
    {
      title: "Document", dataIndex: "document", width: 150,
      filters: getColumnFilters("document"),
      onFilter: (value, record) => record.document === value,
      render: (text) => <div style={scrollCellStyle}>{text}</div>,
    },
    {
      title: "Date", dataIndex: "date", width: 100,
      filters: getColumnFilters("date"),
      onFilter: (value, record) => record.date === value,
    },
    {
      title: "Source", dataIndex: "source", width: 150, align: "left",
      render: (url, record) =>
        url ? (
          <a href={url} target="_blank" rel="noopener noreferrer">
            <FilePdfOutlined style={{ color: "#f84434" }} />
            {record.documentName}
          </a>
        ) : "-",
    },
    {
      title: "Excel", dataIndex: "json", width: 80,
      render: (json, record) =>
        json ? (
          <Button
            type="link"
            icon={<FileExcelOutlined style={{ color: "#217346" }} />}
            onClick={() => {
              setSelectedExcelData({ json, documentName: record.documentName });
              setExcelModalOpen(true);
            }}
          >
            View
          </Button>
        ) : "-",
    },
    {
      title: "Output", dataIndex: "output", width: 50, align: "center",
      render: (submissionId) => (
        <InfoCircleOutlined
          style={{ fontSize: 18, color: "#1677ff", cursor: "pointer" }}
          onClick={() => setSelectedSubmissionId(submissionId)}
        />
      ),
    },
  ];

  const submission = apiData.find((item) => item.submission_id === selectedSubmissionId);
  const llm = submission?.llm_response ?? {};
  const { metadataRows, fieldGroups } = getExtractionViewModel(llm);

  /* ── Excel preview ── */
  const { headers: previewHeaders, rows: previewRows, rowCount: previewRowCount } =
    selectedExcelData
      ? buildExpandedExcelData(
        selectedExcelData.json,
        selectedExcelData.documentName || "extracted"
      )
      : { headers: [], rows: [], rowCount: 0 };

  /* ── Preview table columns with merge logic ── */
  const previewTableColumns = previewHeaders.map((header) => ({
    title: header,
    dataIndex: header,
    onHeaderCell: () => ({ style: { backgroundColor: "#217346", color: "#fff" } }),
    render: (value, _row, index) => {
      if (TOP_LEVEL_MERGE_HEADERS.has(header)) {
        return {
          children: value ?? "-",
          props: { rowSpan: index === 0 ? previewRowCount : 0 },
        };
      }
      return {
        children: (
          <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
            {value || "-"}
          </span>
        ),
        props: { rowSpan: 1 },
      };
    },
  }));

  /* ─────────────────────────────────────────
     Extracted Fields — per-group renderer
  ───────────────────────────────────────── */
  const renderFieldGroup = (group) => {
    const nk = normKey(group.fieldName);

    /* ── Policies ── */
    if (nk === "policies") {
      const valueArray = llm?.Policies?.valueArray || [];
      return (
        <List.Item key={group.id}>
          <Row gutter={[16, 8]}>
            <Col span={24}><strong>Policies</strong></Col>
            <Col span={24}>
              {valueArray.map((policy, pIdx) => {
                const obj = policy.valueObject || {};
                return (
                  <div
                    key={pIdx}
                    style={{
                      marginBottom: pIdx < valueArray.length - 1 ? 20 : 0,
                      paddingBottom: pIdx < valueArray.length - 1 ? 20 : 0,
                      borderBottom: pIdx < valueArray.length - 1
                        ? "1px solid #f0f0f0"
                        : "none",
                    }}
                  >
                    <div
                      style={{
                        fontSize: 12, fontWeight: 600, color: "#5d9de2",
                        marginBottom: 10, textTransform: "uppercase", letterSpacing: "0.5px",
                      }}
                    >
                      Policy {pIdx + 1}
                    </div>
                    {POLICY_FIELDS.map(({ label, key }) => (
                      <FieldRow
                        key={key}
                        label={label}
                        field={obj[key]}   /* obj.Borrower_Name now correctly maps */
                        style={{ marginBottom: 6 }}
                      />
                    ))}
                  </div>
                );
              })}
            </Col>
          </Row>
        </List.Item>
      );
    }

    /* ── Address of Mortgagee Company (PO Box / City / State / ZIP) ── */
    if (nk === "addressofmortgageecompany") {
      const valueObject = llm?.Address_of_Mortgagee_Company?.valueObject ?? null;
      return (
        <List.Item key={group.id}>
          <Row gutter={[16, 8]}>
            <Col span={24}><strong>Address of Mortgagee Company</strong></Col>
            <Col span={24}>
              <AddressObjectRenderer valueObject={valueObject} />
            </Col>
          </Row>
        </List.Item>
      );
    }

    /* ── Effective Date (valueDate) ── */
    if (nk === "effectivedate") {
      const field = llm?.Effective_Date ?? null;
      return (
        <List.Item key={group.id}>
          <Row gutter={[16, 8]}>
            <Col span={24}><strong>Effective Date</strong></Col>
            <Col span={24}>
              <FieldRow label="Effective Date" field={field} />
            </Col>
          </Row>
        </List.Item>
      );
    }

    /* ── Current Mortgagee Company ── */
    if (nk === "currentmortgageecompany") {
      const field = llm?.Current_Mortgagee_Company ?? null;
      return (
        <List.Item key={group.id}>
          <Row gutter={[16, 8]}>
            <Col span={24}><strong>Current Mortgagee Company</strong></Col>
            <Col span={24}>
              <FieldRow label="Company" field={field} />
            </Col>
          </Row>
        </List.Item>
      );
    }

    /* ── Default: all other flat fields ── */
    return (
      <List.Item key={group.id}>
        <Row gutter={[16, 8]}>
          <Col span={24}><strong>{group.fieldName}</strong></Col>
          <Col span={24}>
            {group.entries.map((entry, index) => {
              const conf = getConfValue(entry);
              const confColor = getConfColor(conf);
              const page = getPageFromSource(entry.source);
              return (
                <Row
                  key={`${group.id}-${index}`}
                  gutter={[16, 8]}
                  style={{ marginBottom: index === group.entries.length - 1 ? 0 : 12 }}
                >
                  <Col span={14}>
                    {String(entry.value || "").length > 120 ? (
                      <Input.TextArea
                        value={entry.value}
                        readOnly
                        autoSize={{ minRows: 2, maxRows: 6 }}
                      />
                    ) : (
                      <Input value={entry.value} readOnly />
                    )}
                  </Col>
                  <Col span={10} style={{ textAlign: "right" }}>
                    {conf != null && (
                      <Tag color={confColor}>Confidence: {conf}%</Tag>
                    )}
                    {page && <Tag color="#1d3461">Page: {page}</Tag>}
                  </Col>
                </Row>
              );
            })}
          </Col>
        </Row>
      </List.Item>
    );
  };

  return (
    <Container>
      <MyTableComponent
        columns={columns}
        dataSource={tableData}
        loading={loading}
        selectedSubmissionId={selectedSubmissionId}
      />

      <Row>
        <Col span={24} style={{ textAlign: "right", marginTop: 16 }}>
          <Button type="primary" icon={<UploadOutlined />} onClick={() => setIsModalOpen(true)}>
            Upload
          </Button>
        </Col>
      </Row>

      {submission && (
        <div ref={detailsRef} style={{ marginTop: 24 }}>
          {/* ── Document Metadata ── */}
          {metadataRows.length > 0 && (
            <Card
              title="Document Metadata"
              headStyle={{ backgroundColor: "#5d9de2", color: "#fff" }}
            >
              <Table
                columns={[
                  { title: "", dataIndex: "keyName", width: "30%" },
                  { title: "", dataIndex: "value" },
                ]}
                dataSource={metadataRows}
                pagination={false}
                bordered
                size="small"
              />
            </Card>
          )}

          {/* ── Extracted Fields ── */}
          <Card
            style={{ marginTop: metadataRows.length > 0 ? 16 : 0 }}
            title={`Extracted Fields (${fieldGroups.length})`}
            headStyle={{ backgroundColor: "#5d9de2", color: "#fff" }}
          >
            <List
              itemLayout="vertical"
              dataSource={fieldGroups}
              renderItem={renderFieldGroup}
            />
          </Card>
        </div>
      )}

      {/* ── Upload Modal ── */}
      <Modal
        title="Upload File"
        open={isModalOpen}
        destroyOnClose
        onCancel={() => { setIsModalOpen(false); setFileList([]); }}
        afterClose={() => setFileList([])}
        footer={null}
      >
        {/* <div style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
          <span>Enable Preprocessing</span>
          <Switch checked={preprocess} onChange={(checked) => setPreprocess(checked)} />
        </div> */}
        <Upload.Dragger
          accept=".pdf,.tif,.tiff"
          multiple={false}
          maxCount={1}
          fileList={fileList}
          beforeUpload={(file) => {
            if (fileList.length >= 1) {
              message.error("Multiple files can't be uploaded");
              return Upload.LIST_IGNORE;
            }
            const fileName = file.name.toLowerCase();
            const isValid =
              file.type === "application/pdf" ||
              file.type === "image/tiff" ||
              file.type === "image/tif" ||
              fileName.endsWith(".pdf") ||
              fileName.endsWith(".tif") ||
              fileName.endsWith(".tiff");
            if (!isValid) {
              message.error("Only PDF, TIF, TIFF files are allowed");
              return Upload.LIST_IGNORE;
            }
            return true;
          }}
          onChange={({ fileList: updatedFileList }) => {
            if (updatedFileList.length <= 1) setFileList(updatedFileList);
          }}
          onRemove={() => setFileList([])}
          customRequest={async ({ file, onSuccess, onError }) => {
            try {
              const formData = new FormData();
              formData.append("file", file);
              const response = await fetch(
                buildAiExtractUrl("/api/extract_mortgage_document"),
                { method: "POST", body: formData }
              );
              const result = await response.json();
              setApiData((prev) => [...prev, result]);
              setSelectedSubmissionId(result.submission_id);
              setFileList([]);
              setIsModalOpen(false);
              message.success("PDF processed successfully");
              onSuccess();
            } catch (error) {
              message.error("Upload failed");
              onError(error);
            }
          }}
        >
          <p className="ant-upload-drag-icon"><UploadOutlined /></p>
          <p>Click or drag PDF / TIF / TIFF file to upload</p>
        </Upload.Dragger>
      </Modal>

      {/* ── Excel Preview Modal ── */}
      <Modal
        title={
          <Row justify="space-between" align="middle">
            <Col>
              <FileExcelOutlined style={{ color: "#217346", marginRight: 8 }} />
              Extracted Data (Excel Preview)
            </Col>
            <Col>
              <Button
                type="primary"
                icon={<DownloadOutlined />}
                onClick={() =>
                  downloadMatrixExcel(
                    selectedExcelData?.json,
                    selectedExcelData?.documentName || "extracted"
                  )
                }
                style={{ marginRight: 24 }}
              >
                Download Excel
              </Button>
            </Col>
          </Row>
        }
        open={excelModalOpen}
        onCancel={() => { setExcelModalOpen(false); setSelectedExcelData(null); }}
        footer={null}
        width={1300}
      >
        <Table
          rowKey={(_, index) => index}
          columns={previewTableColumns}
          dataSource={previewRows}
          pagination={{ pageSize: 10 }}
          bordered
          size="small"
          scroll={{ x: true }}
        />
      </Modal>
    </Container>
  );
};

export default Mortgage;