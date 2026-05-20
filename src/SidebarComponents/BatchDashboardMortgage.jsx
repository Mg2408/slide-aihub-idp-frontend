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
  FilePdfOutlined,
  FileExcelOutlined,
  DownloadOutlined,
} from "@ant-design/icons";
import XLSX from "xlsx-js-style";

import { buildAiExtractUrl } from "../config/apiBase";
import { TableContainer } from "../styles/components/TableComponent";
import { Container } from "../styles/components/Layout";
import useMetaData from "../context/metaData";
import {
  buildMatrixPreviewRows,
  getConfidenceDisplay,
  getDocumentName,
  getExtractionViewModel,
  getOwnerName,
  getReferenceDisplay,
} from "../utils/extractionResponse";

const MAX_BATCH_FILES = 10;
const MAX_PDF_FILE_SIZE_MB = 3;
const MAX_TIFF_FILE_SIZE_KB = 300;
const MAX_PDF_FILE_SIZE_BYTES = MAX_PDF_FILE_SIZE_MB * 1024 * 1024;
const MAX_TIFF_FILE_SIZE_BYTES = MAX_TIFF_FILE_SIZE_KB * 1024;
const SPECIAL_MERGE_FIELDS = new Set([
  "Current Mortgagee Company",
  "Address of Mortgagee Company",
]);

const scrollCellStyle = {
  maxHeight: 55,
  overflowX: "auto",
  whiteSpace: "normal",
  wordBreak: "break-word",
};

const MyTableComponent = ({ columns, dataSource, loading, selectedKey }) => {
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
            record.key === selectedKey
              ? {
                backgroundColor: "#e6f4ff",
                transition: "background-color 0.3s ease",
              }
              : {},
        })}
        components={{
          header: {
            cell: (props) => <th {...props} style={{ color: "#fff" }} />,
          },
        }}
      />
    </TableContainer>
  );
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

const downloadMatrixExcel = (json, documentName) => {
  if (!json) return;

  const workbook = XLSX.utils.book_new();
  const worksheet = {};
  const { fieldGroups } = getExtractionViewModel(json);
  const fieldNames = fieldGroups.map((group) => group.fieldName);
  const rows = buildMatrixPreviewRows(json, documentName);
  const headers = ["Document Name", ...fieldNames, "Reference"];

  headers.forEach((header, colIndex) => {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
    worksheet[cellRef] = { v: header, s: headerStyle() };
  });

  rows.forEach((row, rowIndex) => {
    const sheetRowIndex = rowIndex + 1;

    worksheet[XLSX.utils.encode_cell({ r: sheetRowIndex, c: 0 })] = {
      v: row["Document Name"] ?? "",
      s: cellStyle(),
    };

    fieldNames.forEach((fieldName, columnIndex) => {
      worksheet[XLSX.utils.encode_cell({ r: sheetRowIndex, c: columnIndex + 1 })] = {
        v: row[fieldName] ?? "",
        s: cellStyle(),
      };
    });

    worksheet[XLSX.utils.encode_cell({ r: sheetRowIndex, c: fieldNames.length + 1 })] = {
      v: row.Reference ?? "",
      s: cellStyle(),
    };
  });

  const merges = [];

  if (rows.length > 1) {
    merges.push({
      s: { r: 1, c: 0 },
      e: { r: rows.length, c: 0 },
    });
  }

  fieldGroups.forEach((group, index) => {
    if (SPECIAL_MERGE_FIELDS.has(group.fieldName) && group.entries.length <= 1 && rows.length > 1) {
      merges.push({
        s: { r: 1, c: index + 1 },
        e: { r: rows.length, c: index + 1 },
      });
    }
  });

  worksheet["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(rows.length, 1), c: headers.length - 1 },
  });
  worksheet["!merges"] = merges;
  worksheet["!cols"] = headers.map(() => ({ wch: 28 }));

  XLSX.utils.book_append_sheet(workbook, worksheet, "Mortgage Extracted Data");
  XLSX.writeFile(workbook, `${documentName}_extracted.xlsx`);
};

const BatchDashboardMortgage = () => {
  const [batchData, setBatchData] = useState([]);
  const [selectedKey, setSelectedKey] = useState(null);
  const [fileList, setFileList] = useState([]);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [excelModalOpen, setExcelModalOpen] = useState(false);
  const [selectedExcelData, setSelectedExcelData] = useState(null);
  const detailsRef = useRef(null);
  const [preprocess, setPreprocess] = useState(false);

  useEffect(() => {
    if (selectedKey !== null && detailsRef.current) {
      setTimeout(() => {
        detailsRef.current.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
      }, 100);
    }
  }, [selectedKey]);

  const handleBatchUpload = async () => {
    if (!fileList.length) {
      message.warning("Please select at least one file");
      return;
    }

    if (fileList.length > MAX_BATCH_FILES) {
      message.error(`You can upload up to ${MAX_BATCH_FILES} files at a time`);
      return;
    }

    const oversizedFile = fileList.find((file) => {
      const fileSize = file?.originFileObj?.size ?? file?.size ?? 0;
      const fileName = (file?.name || "").toLowerCase();
      const isTiffFile =
        file?.type === "image/tiff" ||
        fileName.endsWith(".tif") ||
        fileName.endsWith(".tiff");
      const maxAllowedSize = isTiffFile
        ? MAX_TIFF_FILE_SIZE_BYTES
        : MAX_PDF_FILE_SIZE_BYTES;
      return fileSize > maxAllowedSize;
    });

    if (oversizedFile) {
      const fileName = (oversizedFile?.name || "").toLowerCase();
      const isTiffFile =
        oversizedFile?.type === "image/tiff" ||
        fileName.endsWith(".tif") ||
        fileName.endsWith(".tiff");
      const typeLabel = isTiffFile ? "TIF/TIFF" : "PDF";
      const maxSizeLabel = isTiffFile
        ? MAX_TIFF_FILE_SIZE_KB
        : MAX_PDF_FILE_SIZE_MB;

      message.error(
        `${typeLabel} files must be ${maxSizeLabel} ${isTiffFile ? "KB" : "MB"
        } or smaller. Please upload it in individual process.`
      );
      return;
    }

    try {
      setLoading(true);
      const formData = new FormData();

      fileList.forEach((file) => {
        formData.append("files", file.originFileObj);
      });

      const response = await fetch(
        buildAiExtractUrl(`/api/extract_document_batch?template=mortgage`),
        {
          method: "POST",
          body: formData,
        }
      );

      const result = await response.json();
      setBatchData((prev) => [...prev, ...(result.results || [])]);
      setIsModalOpen(false);
      setFileList([]);
      message.success("Batch processed successfully");
    } catch (error) {
      message.error("Batch upload failed");
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadConsolidatedExcel = () => {
    if (!batchData.length) {
      message.warning("No batch data available to download");
      return;
    }

    const workbook = XLSX.utils.book_new();
    const worksheet = {};
    const allFieldNames = Array.from(
      new Set(
        batchData.flatMap((item) =>
          getExtractionViewModel(item.data || {}).fieldGroups.map(
            (group) => group.fieldName
          )
        )
      )
    );
    const headers = ["Document Name", ...allFieldNames, "Reference"];
    const merges = [];
    let rowIndex = 1;

    headers.forEach((header, colIndex) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      worksheet[cellRef] = { v: header, s: headerStyle() };
    });

    batchData.forEach((item) => {
      const extractionResponse = item.data || {};
      const documentName = getDocumentName(extractionResponse, [item.file_name]);
      const { fieldGroups } = getExtractionViewModel(extractionResponse);
      const rows = buildMatrixPreviewRows(extractionResponse, documentName);
      const documentStartRow = rowIndex;

      rows.forEach((row) => {
        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: 0 })] = {
          v: row["Document Name"] ?? "",
          s: cellStyle(),
        };

        allFieldNames.forEach((fieldName, columnIndex) => {
          worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex + 1 })] = {
            v: row[fieldName] ?? "",
            s: cellStyle(),
          };
        });

        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: allFieldNames.length + 1 })] = {
          v: row.Reference ?? "",
          s: cellStyle(),
        };

        rowIndex += 1;
      });

      if (rows.length > 1) {
        merges.push({
          s: { r: documentStartRow, c: 0 },
          e: { r: rowIndex - 1, c: 0 },
        });
      }

      fieldGroups.forEach((group) => {
        if (
          SPECIAL_MERGE_FIELDS.has(group.fieldName) &&
          group.entries.length <= 1 &&
          rows.length > 1
        ) {
          const columnIndex = allFieldNames.indexOf(group.fieldName);
          if (columnIndex !== -1) {
            merges.push({
              s: { r: documentStartRow, c: columnIndex + 1 },
              e: { r: rowIndex - 1, c: columnIndex + 1 },
            });
          }
        }
      });
    });

    worksheet["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(rowIndex - 1, 1), c: headers.length - 1 },
    });
    worksheet["!merges"] = merges;
    worksheet["!cols"] = headers.map(() => ({ wch: 28 }));

    XLSX.utils.book_append_sheet(workbook, worksheet, "Mortgage Consolidated Data");
    XLSX.writeFile(workbook, "mortgage_consolidated_extracted.xlsx");
  };

  const tableData = batchData.map((item, index) => {
    const extractionResponse = item.data || {};
    const documentName = getDocumentName(extractionResponse, [item.file_name]);

    return {
      key: index,
      file: item.file_name,
      document: documentName,
      owner: getOwnerName(extractionResponse, ["-"]),
      pdf: item.pdf_url || null,
      json: extractionResponse,
      documentName,
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
    {
      title: "File",
      dataIndex: "file",
      width: 160,
      render: (text) => <div style={scrollCellStyle}>{text}</div>,
    },
    {
      title: "Document",
      dataIndex: "document",
      width: 200,
      filters: getColumnFilters("document"),
      onFilter: (value, record) => record.document === value,
      render: (text) => <div style={scrollCellStyle}>{text}</div>,
    },
    {
      title: "Owner",
      dataIndex: "owner",
      width: 160,
      filters: getColumnFilters("owner"),
      onFilter: (value, record) => record.owner === value,
      render: (text) => <div style={scrollCellStyle}>{text}</div>,
    },
    {
      title: "PDF",
      dataIndex: "pdf",
      width: 80,
      align: "center",
      render: (url) =>
        url ? (
          <a href={url} target="_blank" rel="noopener noreferrer">
            <FilePdfOutlined style={{ color: "#f84434" }} />
            View
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
      dataIndex: "key",
      width: 60,
      align: "center",
      render: (key) => (
        <InfoCircleOutlined
          style={{ fontSize: 18, color: "#1677ff", cursor: "pointer" }}
          onClick={() => setSelectedKey(key)}
        />
      ),
    },
  ];

  const selectedItem = selectedKey !== null ? batchData[selectedKey] : null;
  const { metadataRows, fieldGroups } = getExtractionViewModel(selectedItem?.data);

  const previewRows = selectedExcelData
    ? buildMatrixPreviewRows(
      selectedExcelData.json,
      selectedExcelData.documentName || "extracted"
    )
    : [];
  const previewFieldGroups = selectedExcelData
    ? getExtractionViewModel(selectedExcelData.json).fieldGroups
    : [];

  return (
    <Container>
      <MyTableComponent
        columns={columns}
        dataSource={tableData}
        loading={loading}
        selectedKey={selectedKey}
      />

      <Row>
        <Col span={24} style={{ textAlign: "right", marginTop: 16 }}>
          <Button
            type="primary"
            icon={<DownloadOutlined />}
            onClick={handleDownloadConsolidatedExcel}
            disabled={!batchData.length}
            style={{ marginRight: 8 }}
          >
            Download Excel
          </Button>
          <Button
            type="primary"
            icon={<UploadOutlined />}
            onClick={() => setIsModalOpen(true)}
          >
            Upload Batch PDF / TIF / TIFF
          </Button>
        </Col>
      </Row>

      {selectedItem && (
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
                                  {reference.label}: {reference.value}
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
        title="Upload Multiple PDFs"
        open={isModalOpen}
        onCancel={() => {
          setIsModalOpen(false);
          setFileList([]);
        }}
        onOk={handleBatchUpload}
        okText="Process Batch"
        confirmLoading={loading}
      >
        <div style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 8 }}>
          <span>Enable Preprocessing</span>
          <Switch
            checked={preprocess}
            onChange={(checked) => setPreprocess(checked)}
          />
        </div>
        <Upload.Dragger
          accept=".pdf,.tif,.tiff"
          multiple
          maxCount={MAX_BATCH_FILES}
          fileList={fileList}
          beforeUpload={(file) => {
            const fileName = file.name.toLowerCase();
            const isValid =
              file.type === "application/pdf" ||
              file.type === "image/tiff" ||
              fileName.endsWith(".pdf") ||
              fileName.endsWith(".tif") ||
              fileName.endsWith(".tiff");

            if (!isValid) {
              message.error("Only PDF, TIF, TIFF files are allowed");
              return Upload.LIST_IGNORE;
            }

            const isTiffFile =
              file.type === "image/tiff" ||
              fileName.endsWith(".tif") ||
              fileName.endsWith(".tiff");
            const maxAllowedSize = isTiffFile
              ? MAX_TIFF_FILE_SIZE_BYTES
              : MAX_PDF_FILE_SIZE_BYTES;
            const maxAllowedSizeLabel = isTiffFile
              ? MAX_TIFF_FILE_SIZE_KB
              : MAX_PDF_FILE_SIZE_MB;
            const typeLabel = isTiffFile ? "TIF/TIFF" : "PDF";

            if (file.size > maxAllowedSize) {
              message.error(
                `${file.name} exceeds ${maxAllowedSizeLabel} ${isTiffFile ? "KB" : "MB"
                } limit for ${typeLabel} files. Please upload it in individual process.`
              );
              return Upload.LIST_IGNORE;
            }

            return false;
          }}
          onChange={({ fileList: updatedFileList }) => {
            if (updatedFileList.length > MAX_BATCH_FILES) {
              message.error(
                `You can upload up to ${MAX_BATCH_FILES} files at a time`
              );
            }

            const limitedFileList = updatedFileList.slice(0, MAX_BATCH_FILES);
            const validatedFileList = limitedFileList.filter((file) => {
              const fileSize = file?.originFileObj?.size ?? file?.size ?? 0;
              const fileName = (file?.name || "").toLowerCase();
              const isTiffFile =
                file?.type === "image/tiff" ||
                fileName.endsWith(".tif") ||
                fileName.endsWith(".tiff");
              const maxAllowedSize = isTiffFile
                ? MAX_TIFF_FILE_SIZE_BYTES
                : MAX_PDF_FILE_SIZE_BYTES;
              return fileSize <= maxAllowedSize;
            });

            setFileList(validatedFileList);
          }}
        >
          <p className="ant-upload-drag-icon">
            <UploadOutlined />
          </p>
          <p>
            Click or drag PDF / TIF / TIFF files to upload (up to {MAX_BATCH_FILES}
            , PDF {MAX_PDF_FILE_SIZE_MB} MB, TIF/TIFF {MAX_TIFF_FILE_SIZE_KB} KB)
          </p>
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
        onCancel={() => {
          setExcelModalOpen(false);
          setSelectedExcelData(null);
        }}
        footer={null}
        width={860}
      >
        <Table
          rowKey={(_, index) => index}
          columns={[
            {
              title: "Document Name",
              dataIndex: "Document Name",
              onHeaderCell: () => ({
                style: { backgroundColor: "#217346", color: "#fff" },
              }),
              render: (value, _row, index) => ({
                children: value,
                props: {
                  rowSpan: index === 0 ? previewRows.length || 1 : 0,
                },
              }),
            },
            ...previewFieldGroups.map((group) => ({
              title: group.fieldName,
              dataIndex: group.fieldName,
              onHeaderCell: () => ({
                style: { backgroundColor: "#217346", color: "#fff" },
              }),
              render: (value, _row, index) => {
                if (
                  SPECIAL_MERGE_FIELDS.has(group.fieldName) &&
                  group.entries.length <= 1 &&
                  previewRows.length > 1
                ) {
                  return {
                    children: value,
                    props: {
                      rowSpan: index === 0 ? previewRows.length : 0,
                    },
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
            })),
            {
              title: "Reference",
              dataIndex: "Reference",
              width: 180,
              onHeaderCell: () => ({
                style: { backgroundColor: "#217346", color: "#fff" },
              }),
              render: (value) => (value ? value : "-"),
            },
          ]}
          dataSource={previewRows}
          pagination={{ pageSize: 100 }}
          bordered
          size="small"
          scroll={{ x: true }}
        />
      </Modal>
    </Container>
  );
};

export default BatchDashboardMortgage;
