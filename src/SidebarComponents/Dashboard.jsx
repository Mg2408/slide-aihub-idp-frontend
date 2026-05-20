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
  Switch
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
  buildFlatFieldRows,
  getConfidenceDisplay,
  getDocumentName,
  getExtractionViewModel,
  getOwnerName,
  getReferenceDisplay,
} from "../utils/extractionResponse";

const getPageNumber = (referenceValue) => {
  if (!referenceValue) return null;
  const match = String(referenceValue).match(/D\((\d+)/);
  return match ? match[1] : null;
};

const MyTableComponent = ({
  columns,
  dataSource,
  loading,
  selectedSubmissionId,
}) => {
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
              ? {
                backgroundColor: "#e6f4ff",
                transition: "background-color 0.3s ease",
              }
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

const downloadFlatRowsExcel = (json, documentName) => {
  if (!json) return;

  const rows = buildFlatFieldRows(json);
  const workbook = XLSX.utils.book_new();
  const worksheet = {};
  const headers = ["Document Name", "Field", "Value", "Reference"];
  const merges = [];

  headers.forEach((header, colIndex) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
    worksheet[cellRef] = {
    };
  });

  let rowIndex = 1;
  let currentFieldStartRow = 1;

  rows.forEach((row, index) => {
    if (index > 0 && row.Field) {
      const previousRow = rowIndex - 1;
      if (previousRow > currentFieldStartRow) {
        merges.push({
          s: { r: currentFieldStartRow, c: 1 },
          e: { r: previousRow, c: 1 },
        });
      }
      currentFieldStartRow = rowIndex;
    }

    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = {
      v: index === 0 ? documentName : "",
      s: cellStyle(),
    };
    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 1 })] = {
      v: row.Field || "",
      s: cellStyle(),
    };
    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 2 })] = {
      v: row.Value ?? "",
      s: cellStyle(),
    };
    worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 3 })] = {
      v: row.Reference ?? "",
      s: cellStyle(),
    };

    rowIndex += 1;
  });

  const finalFieldEndRow = rowIndex - 1;
  if (finalFieldEndRow > currentFieldStartRow) {
    merges.push({
      s: { r: currentFieldStartRow, c: 1 },
      e: { r: finalFieldEndRow, c: 1 },
    });
  }

  if (rows.length > 1) {
    merges.push({
      s: { r: 1, c: 0 },
      e: { r: rowIndex - 1, c: 0 },
    });
  }

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rowIndex - 1, 1), c: 3 },
  });
  worksheet["!merges"] = merges;
  worksheet["!cols"] = [
    { wch: 28 },
    { wch: 30 },
    { wch: 70 },
    { wch: 36 },
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, "Extracted Data");
  XLSX.writeFile(workbook, `${documentName}_extracted.xlsx`);
};

const Dashboard = () => {
  const [apiData, setApiData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [excelModalOpen, setExcelModalOpen] = useState(false);
  const [selectedExcelData, setSelectedExcelData] = useState(null);
  const [selectedSubmissionId, setSelectedSubmissionId] = useState(null);
  const [fileList, setFileList] = useState([]);
  const hasFetchedRef = useRef(false);
  const detailsRef = useRef(null);
  const [preprocess, setPreprocess] = useState(false);

  useEffect(() => {
    if (hasFetchedRef.current) return;

    const fetchDocuments = async () => {
      try {
        setLoading(true);
        const response = await fetch(
          buildAiExtractUrl("/api/get_extracted_documents?template=wind_mit")
        );

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const result = await response.json();
        setApiData(result.submission_list || []);
        hasFetchedRef.current = true;
      } catch (error) {
        console.error(error);
        message.error("Failed to fetch data");
      } finally {
        setLoading(false);
      }
    };

    fetchDocuments();
  }, []);

  useEffect(() => {
    if (selectedSubmissionId && detailsRef.current) {
      setTimeout(() => {
        detailsRef.current.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }, 100);
    }
  }, [selectedSubmissionId]);

  const getDocumentNameFromUri = (uri) => {
    if (!uri) return null;
    try {
      // Strip query string, grab last path segment, then decode %20 etc.
      const rawFileName = uri.split("?")[0].split("/").pop();
      const decoded = decodeURIComponent(rawFileName);
      // Remove file extension (.pdf / .tif / .tiff), then collapse all spaces
      // return decoded.replace(/\.(pdf|tiff?)$/i, "").replace(/\s+/g, "");
      return decoded.replace(/\s+/g, "");
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
    // {
    //   title: "Submitted by",
    //   dataIndex: "submittedBy",
    //   width: 120,
    //   filters: getColumnFilters("submittedBy"),
    //   onFilter: (value, record) => record.submittedBy === value,
    //   render: (text) => <div style={scrollCellStyle}>{text}</div>,
    // },
    {
      title: "Document",
      dataIndex: "document",
      width: 150,
      filters: getColumnFilters("document"),
      onFilter: (value, record) => record.document === value,
      render: (text) => <div style={scrollCellStyle}>{text}</div>,
    },
    {
      title: "Date",
      dataIndex: "date",
      width: 100,
      filters: getColumnFilters("date"),
      onFilter: (value, record) => record.date === value,
    },
    {
      title: "Source",
      dataIndex: "source",
      width: 150,
      align: "left",
      render: (url, record) =>
        url ? (
          <a href={url} target="_blank" rel="noopener noreferrer">
            <FilePdfOutlined style={{ color: "#f84434" }} />
            {record.documentName}
          </a>
        ) : (
          "-"
        ),
    },
    {
      title: "Excel",
      dataIndex: "json",
      width: 80,
      render: (json, record) =>
        json ? (
          <Button
            type="link"
            icon={<FileExcelOutlined style={{ color: "#217346" }} />}
            onClick={() => {
              setSelectedExcelData({
                json,
                documentName: record.documentName,
              });
              setExcelModalOpen(true);
            }}
          >
            View
          </Button>
        ) : (
          "-"
        ),
    },
    {
      title: "Output",
      dataIndex: "output",
      width: 50,
      align: "center",
      render: (submissionId) => (
        <InfoCircleOutlined
          style={{ fontSize: 18, color: "#1677ff", cursor: "pointer" }}
          onClick={() => setSelectedSubmissionId(submissionId)}
        />
      ),
    },
  ];

  const submission = apiData.find(
    (item) => item.submission_id === selectedSubmissionId
  );
  const { metadataRows, fieldGroups } = getExtractionViewModel(
    submission?.llm_response
  );

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

          <Card
            style={{ marginTop: metadataRows.length > 0 ? 16 : 0 }}
            title={`Extracted Fields (${fieldGroups.length})`}
            headStyle={{ backgroundColor: "#5d9de2", color: "#fff" }}
          >
            <List
              itemLayout="vertical"
              dataSource={fieldGroups}
              renderItem={(group) => (
                <List.Item key={group.id}>
                  <Row gutter={[16, 8]}>
                    <Col span={24}>
                      <strong>{group.fieldName}</strong>
                    </Col>
                    <Col span={24}>
                      {group.entries.map((entry, index) => {
                        const confidence = getConfidenceDisplay(entry);
                        const reference = getReferenceDisplay(entry);

                        return (
                          <Row
                            key={`${group.id}-${index}`}
                            gutter={[16, 8]}
                            style={{
                              marginBottom:
                                index === group.entries.length - 1 ? 0 : 12,
                            }}
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
                              {confidence && (
                                <Tag color={confidence.color}>
                                  Confidence: {confidence.value}
                                </Tag>
                              )}
                              {reference && (
                                <Tag>
                                  {(() => {
                                    const page = getPageNumber(reference.value);
                                    return page ? `Page: ${page}` : `${reference.label}: ${reference.value}`;
                                  })()}
                                </Tag>
                              )}
                            </Col>
                          </Row>
                        );
                      })}
                    </Col>
                  </Row>
                </List.Item>
              )}
            />
          </Card>
        </div>
      )}

      <Modal
        title="Upload File"
        open={isModalOpen}
        destroyOnClose
        onCancel={() => {
          setIsModalOpen(false);
          setFileList([]);
        }}
        afterClose={() => setFileList([])}
        footer={null}
      >
        <div style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
          <span>Enable Preprocessing</span>
          <Switch
            checked={preprocess}
            onChange={(checked) => {
              console.log("Preprocessing toggle:", checked);
              setPreprocess(checked)
            }}
          />
        </div>
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

            const allowedTypes = [
              "application/pdf",
              "image/tiff",
              "image/tif",
            ];

            const fileName = file.name.toLowerCase();
            const isValidType =
              allowedTypes.includes(file.type) ||
              fileName.endsWith(".pdf") ||
              fileName.endsWith(".tif") ||
              fileName.endsWith(".tiff");

            if (!isValidType) {
              message.error("Only PDF, TIF, TIFF files are allowed");
              return Upload.LIST_IGNORE;
            }

            return true;
          }}
          onChange={({ fileList: updatedFileList }) => {
            if (updatedFileList.length <= 1) {
              setFileList(updatedFileList);
            }
          }}
          onRemove={() => setFileList([])}
          customRequest={async ({ file, onSuccess, onError }) => {
            try {
              const formData = new FormData();
              formData.append("file", file);

              const response = await fetch(
                buildAiExtractUrl(`/api/extract_document?template=wind_mit&preprocess=${preprocess}`),
                {
                  method: "POST",
                  body: formData,
                }
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
          <p className="ant-upload-drag-icon">
            <UploadOutlined />
          </p>
          <p>Click or drag PDF / TIF / TIFF file to upload</p>
        </Upload.Dragger>
      </Modal>

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
                  downloadFlatRowsExcel(
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
        onCancel={() => {
          setExcelModalOpen(false);
          setSelectedExcelData(null);
        }}
        footer={null}
        width={860}
      >
        {selectedExcelData && (
          <Table
            columns={[
              {
                title: "Field",
                dataIndex: "Field",
                key: "Field",
                width: 220,
                onHeaderCell: () => ({
                  style: { backgroundColor: "#217346", color: "#fff" },
                }),
              },
              {
                title: "Value",
                dataIndex: "Value",
                key: "Value",
                render: (value) => (
                  <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                    {value}
                  </span>
                ),
                onHeaderCell: () => ({
                  style: { backgroundColor: "#217346", color: "#fff" },
                }),
              },
              {
                title: "Reference",
                dataIndex: "Reference",
                key: "Reference",
                width: 220,
                render: (value) => (value ? value : "-"),
                onHeaderCell: () => ({
                  style: { backgroundColor: "#217346", color: "#fff" },
                }),
              },
            ]}
            dataSource={buildFlatFieldRows(selectedExcelData.json).map((row, index) => ({
              ...row,
              key: index,
            }))}
            pagination={{ pageSize: 15 }}
            size="small"
            bordered
            scroll={{ y: 420 }}
          />
        )}
      </Modal>
    </Container>
  );
};

export default Dashboard;
