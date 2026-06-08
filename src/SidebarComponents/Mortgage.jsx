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
   New JSON structure: ALL fields are FLAT inside each Policies[].valueObject
   There is NO top-level Effective_Date / Current_Mortgagee_Company /
   Address_of_Mortgagee_Company — those fields now live per-policy.
───────────────────────────────────────── */

/**
 * Fields that are SHARED across all policies (same value on every row).
 * In the Excel preview these columns are merged vertically.
 */
const TOP_LEVEL_MERGE_HEADERS = new Set([
  "Document Name",
  "Current Mortgagee Company",
  "Mortgage Clause",
  "Mortgagee PO Box",
  "Mortgagee City",
  "Mortgagee State",
  "Mortgagee ZIP",
]);

/**
 * Per-policy display fields — rendered once per policy block in the UI
 * and as individual rows in Excel.
 */
const POLICY_FIELDS = [
  { label: "Policy Number",        key: "Policy_Number" },
  { label: "Loan Number",          key: "Loan_Number" },
  { label: "Borrower Name",        key: "Borrower_Name" },
  { label: "Payee Position / Rank",key: "Payee_Position_or_Rank" },
  { label: "Effective Date",       key: "Transaction_Effective_Date" },
];

/**
 * Address / shared fields that are read from the FIRST policy's valueObject
 * (they repeat identically across all policies in the sample data).
 */
const SHARED_POLICY_FIELDS = [
  { label: "Current Mortgagee Company", key: "Current_Mortgagee_Company" },
  { label: "Mortgage Clause",           key: "Mortgage_Clause" },
  { label: "PO Box",                    key: "PO_Box" },
  { label: "City",                      key: "City" },
  { label: "State",                     key: "State" },
  { label: "ZIP",                       key: "ZIP" },
];

/* ─────────────────────────────────────────
   Helpers
───────────────────────────────────────── */

const getFieldValue = (field) =>
  field?.valueString ?? field?.valueDate ?? field?.valueNumber ?? "";

const normKey = (name) => name?.toLowerCase().replace(/[\s_]/g, "") ?? "";

const getConfValue = (field) =>
  field?.confidence != null ? Math.round(field.confidence * 100) : null;

const getConfColor = (conf) =>
  conf == null ? null : conf >= 90 ? "#52bb2c" : conf >= 70 ? "#d37f30" : "#cf3e4a";

const getPageFromSource = (source) => {
  if (!source) return null;
  const match = String(source).match(/D\((\d+)/);
  return match ? match[1] : null;
};

/* ─────────────────────────────────────────
   Build Excel rows
   Columns:
     Document Name | Current Mortgagee Company | Mortgage Clause |
     Mortgagee PO Box | Mortgagee City | Mortgagee State | Mortgagee ZIP |
     Policy Number | Loan Number | Borrower Name | Payee Position / Rank |
     Effective Date | Reference
───────────────────────────────────────── */
const buildExpandedExcelData = (json, documentName) => {
  if (!json) return { headers: [], rows: [], rowCount: 0 };

  const policies = json?.Policies?.valueArray || [];

  const headers = [
    "Document Name",
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
    "Effective Date",
    "Reference",
  ];

  // Read shared fields from first policy (they repeat on every policy)
  const firstObj = policies[0]?.valueObject || {};

  const sharedBase = {
    "Document Name": documentName,
    "Current Mortgagee Company": getFieldValue(firstObj.Current_Mortgagee_Company),
    "Mortgage Clause":           getFieldValue(firstObj.Mortgage_Clause),
    "Mortgagee PO Box":          getFieldValue(firstObj.PO_Box),
    "Mortgagee City":            getFieldValue(firstObj.City),
    "Mortgagee State":           getFieldValue(firstObj.State),
    "Mortgagee ZIP":             getFieldValue(firstObj.ZIP),
  };

  const rows =
    policies.length > 0
      ? policies.map((p) => {
          const obj = p.valueObject || {};
          const page = getPageFromSource(obj.Policy_Number?.source ?? "");
          return {
            ...sharedBase,
            "Policy Number":         getFieldValue(obj.Policy_Number),
            "Loan Number":           getFieldValue(obj.Loan_Number),
            "Borrower Name":         getFieldValue(obj.Borrower_Name),
            "Payee Position / Rank": getFieldValue(obj.Payee_Position_or_Rank),
            "Effective Date":        getFieldValue(obj.Transaction_Effective_Date),
            "Reference":             page ? `Page: ${page}` : "",
          };
        })
      : [
          {
            ...sharedBase,
            "Policy Number": "",
            "Loan Number": "",
            "Borrower Name": "",
            "Payee Position / Rank": "",
            "Effective Date": "",
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
    top: { style: "thin" }, bottom: { style: "thin" },
    left: { style: "thin" }, right: { style: "thin" },
  },
});

const cellStyle = () => ({
  alignment: { vertical: "top", wrapText: true },
  border: {
    top: { style: "thin" }, bottom: { style: "thin" },
    left: { style: "thin" }, right: { style: "thin" },
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

  headers.forEach((header, colIndex) => {
    worksheet[XLSX.utils.encode_cell({ r: 0, c: colIndex })] = {
      v: header, s: headerStyle(),
    };
  });

  rows.forEach((row, rowIndex) => {
    const sheetRow = rowIndex + 1;
    headers.forEach((header, colIndex) => {
      worksheet[XLSX.utils.encode_cell({ r: sheetRow, c: colIndex })] = {
        v: row[header] ?? "", s: cellStyle(),
      };
    });
  });

  // Merge shared/top-level columns vertically across all policy rows
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
   FieldRow — single field with conf + page
───────────────────────────────────────── */
const FieldRow = ({ label, field, style = {} }) => {
  if (!field) return null;
  const value = getFieldValue(field) || "-";
  const conf  = getConfValue(field);
  const confColor = getConfColor(conf);
  const page  = getPageFromSource(field.source);

  return (
    <Row gutter={[16, 8]} style={{ marginBottom: 8, ...style }}>
      <Col span={14}>
        {value.length > 120 ? (
          <Input.TextArea value={value} readOnly autoSize={{ minRows: 2, maxRows: 6 }} />
        ) : (
          <Input
            value={value}
            readOnly
            addonBefore={
              <span style={{ fontWeight: 600, width: 200, display: "inline-block" }}>
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
  const [apiData, setApiData]                     = useState([]);
  const [loading, setLoading]                     = useState(false);
  const [isModalOpen, setIsModalOpen]             = useState(false);
  const [excelModalOpen, setExcelModalOpen]       = useState(false);
  const [selectedExcelData, setSelectedExcelData] = useState(null);
  const [selectedSubmissionId, setSelectedSubmissionId] = useState(null);
  const [fileList, setFileList]                   = useState([]);
  const hasFetchedRef = useRef(false);
  const detailsRef    = useRef(null);

  useEffect(() => {
    if (hasFetchedRef.current) return;
    const fetchDocuments = async () => {
      try {
        setLoading(true);
        const response = await fetch(buildAiExtractUrl("/api/get_extracted_documents"));
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const result = await response.json();
        setApiData(result.submission_list || []);
        hasFetchedRef.current = true;
      } catch (error) {
        console.error(error);
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

  /* ── Selected submission data ── */
  const submission = apiData.find((item) => item.submission_id === selectedSubmissionId);
  const llm        = submission?.llm_response ?? {};
  const policies   = llm?.Policies?.valueArray || [];

  // Shared fields sourced from first policy's valueObject
  const firstPolicyObj = policies[0]?.valueObject || {};

  /* ── Excel preview ── */
  const { headers: previewHeaders, rows: previewRows, rowCount: previewRowCount } =
    selectedExcelData
      ? buildExpandedExcelData(
          selectedExcelData.json,
          selectedExcelData.documentName || "extracted"
        )
      : { headers: [], rows: [], rowCount: 0 };

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
     UI Renderer
  ───────────────────────────────────────── */
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
          <Button
            type="primary"
            icon={<UploadOutlined />}
            onClick={() => setIsModalOpen(true)}
          >
            Upload
          </Button>
        </Col>
      </Row>

      {submission && (
        <div ref={detailsRef} style={{ marginTop: 24 }}>

          {/* ── Shared / Document-level Fields (from first policy) ── */}
          {policies.length > 0 && (
            <Card
              title="Document Level Fields"
              headStyle={{ backgroundColor: "#5d9de2", color: "#fff" }}
              style={{ marginBottom: 16 }}
            >
              {SHARED_POLICY_FIELDS.map(({ label, key }) => (
                <FieldRow
                  key={key}
                  label={label}
                  field={firstPolicyObj[key]}
                  style={{ marginBottom: 6 }}
                />
              ))}
            </Card>
          )}

          {/* ── Per-Policy Extracted Fields ── */}
          <Card
            title={`Policies (${policies.length})`}
            headStyle={{ backgroundColor: "#5d9de2", color: "#fff" }}
          >
            {policies.length === 0 ? (
              <p style={{ color: "#999" }}>No policies found.</p>
            ) : (
              <List
                itemLayout="vertical"
                dataSource={policies}
                renderItem={(policy, pIdx) => {
                  const obj = policy.valueObject || {};
                  return (
                    <List.Item key={pIdx}>
                      {/* Policy header badge */}
                      <div
                        style={{
                          fontSize: 12,
                          fontWeight: 600,
                          color: "#5d9de2",
                          marginBottom: 12,
                          textTransform: "uppercase",
                          letterSpacing: "0.5px",
                          borderBottom: "1px solid #e6f0ff",
                          paddingBottom: 6,
                        }}
                      >
                        Policy {pIdx + 1}
                      </div>

                      {/* Per-policy unique fields */}
                      <div style={{ marginBottom: 12 }}>
                        <div
                          style={{
                            fontSize: 11,
                            color: "#888",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: "0.4px",
                            marginBottom: 8,
                          }}
                        >
                          Policy Details
                        </div>
                        {POLICY_FIELDS.map(({ label, key }) => (
                          <FieldRow
                            key={key}
                            label={label}
                            field={obj[key]}
                            style={{ marginBottom: 6 }}
                          />
                        ))}
                      </div>

                      {/* Shared fields repeated per policy (shown for transparency / per-policy confidence) */}
                      <div
                        style={{
                          marginLeft: 16,
                          paddingLeft: 12,
                          borderLeft: "3px solid #e6f4ff",
                        }}
                      >
                        <div
                          style={{
                            fontSize: 11,
                            color: "#888",
                            fontWeight: 600,
                            textTransform: "uppercase",
                            letterSpacing: "0.4px",
                            marginBottom: 8,
                          }}
                        >
                          Mortgagee / Address Fields (as extracted for this policy)
                        </div>
                        {SHARED_POLICY_FIELDS.map(({ label, key }) => (
                          <FieldRow
                            key={key}
                            label={label}
                            field={obj[key]}
                            style={{ marginBottom: 6 }}
                          />
                        ))}
                      </div>
                    </List.Item>
                  );
                }}
              />
            )}
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