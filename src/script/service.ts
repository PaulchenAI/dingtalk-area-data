/*global DingdocsScript*/
import type {} from 'dingtalk-docs-cool-app';

type RegionLevel = "province" | "city" | "district";

type RegionPayload = {
  level: RegionLevel;
  levelName: string;
  provinceCode: string;
  provinceName: string;
  cityCode: string;
  cityName: string;
  districtCode: string;
  districtName: string;
  regionCode: string;
  regionName: string;
};

const REQUIRED_FIELDS = [
  "区域层级",
  "省份",
  "城市",
  "区县",
];

function getActiveSheet() {
  const sheet = DingdocsScript.base.getActiveSheet();

  if (!sheet) {
    throw new Error("未获取到当前数据表，请先打开一个 AI 表格数据表");
  }

  return sheet;
}

function buildRegionFields(region: RegionPayload) {
  return {
    "区域层级": region.levelName,
    "省份": region.provinceName,
    "城市": region.cityName || "",
    "区县": region.districtName || "",
  };
}

async function ensureRegionFields() {
  const sheet = getActiveSheet();

  for (const fieldName of REQUIRED_FIELDS) {
    let field = null;

    try {
      field = sheet.getField(fieldName);
    } catch (error) {
      field = null;
    }

    if (!field) {
      await sheet.insertField({ name: fieldName, type: "text" });
    }
  }

  return {
    success: true,
    fields: REQUIRED_FIELDS,
  };
}

function getRecordId(record: any) {
  if (!record) return "";

  if (typeof record.getId === "function") {
    return record.getId();
  }

  return record.id || record.recordId || "";
}

function getRecordValues(record: any) {
  if (!record) return {};

  if (typeof record.getCellValues === "function") {
    return record.getCellValues();
  }

  if (typeof record.getFields === "function") {
    return record.getFields();
  }

  return record.fields || record.values || {};
}

function stringifyCellValue(value: any): string {
  if (value === null || value === undefined) return "";

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (Array.isArray(value)) {
    return value.map(stringifyCellValue).filter(Boolean).join("、");
  }

  if (typeof value === "object") {
    return value.text || value.name || value.title || value.value || JSON.stringify(value);
  }

  return String(value);
}

function getRecordTitle(record: any, index: number) {
  const values = getRecordValues(record);
  const valueList = Object.values(values || {});
  const firstValue = valueList.find((item) => stringifyCellValue(item));

  return firstValue ? stringifyCellValue(firstValue) : `第 ${index + 1} 行`;
}

async function getActiveSheetInfo() {
  const sheet = getActiveSheet();

  return {
    id: typeof sheet.getId === "function" ? sheet.getId() : "",
    name: typeof sheet.getName === "function" ? sheet.getName() : "当前数据表",
  };
}

async function getRecordOptions(limit = 100) {
  const sheet = getActiveSheet();
  const result = await sheet.getRecordsAsync({ pageSize: limit });

  return (result.records || []).slice(0, limit).map((record: any, index: number) => {
    return {
      id: getRecordId(record),
      title: getRecordTitle(record, index),
    };
  });
}

async function getCurrentSelectedRecord() {
  const sheet = getActiveSheet();
  const selection = DingdocsScript.base.getSelection?.();
  const recordId = selection?.recordId;

  if (!recordId) {
    return null;
  }

  const record = await sheet.getRecordAsync(recordId);
  if (!record) {
    return {
      id: recordId,
      title: "当前选中行",
      rowNumber: null,
    };
  }

  let rowNumber: number | null = null;
  let cursor: string | undefined = undefined;
  let scannedCount = 0;

  while (rowNumber === null) {
    const result = await sheet.getRecordsAsync({ pageSize: 100, cursor });
    const recordList = result?.records || [];

    const index = recordList.findIndex((item: any) => getRecordId(item) === recordId);
    if (index >= 0) {
      rowNumber = scannedCount + index + 1;
      break;
    }

    scannedCount += recordList.length;

    if (!result?.hasMore || !result?.cursor) {
      break;
    }

    cursor = result.cursor;
  }

  return {
    id: recordId,
    title: getRecordTitle(record, 0),
    rowNumber,
  };
}

async function insertRegionRecord(region: RegionPayload) {
  const sheet = getActiveSheet();

  await ensureRegionFields();

  await sheet.insertRecordsAsync([
    {
      fields: buildRegionFields(region),
    },
  ]);

  return {
    success: true,
    message: `已新增记录：${region.regionName}`,
  };
}

async function updateRegionRecord(recordId: string, region: RegionPayload) {
  const sheet = getActiveSheet();

  if (!recordId) {
    throw new Error("请先选择要更新的记录");
  }

  await ensureRegionFields();

  await sheet.updateRecordsAsync([
    {
      id: recordId,
      fields: buildRegionFields(region),
    },
  ]);

  return {
    success: true,
    message: `已更新记录：${region.regionName}`,
  };
}

DingdocsScript.registerScript("getActiveSheetInfo", getActiveSheetInfo);
DingdocsScript.registerScript("ensureRegionFields", ensureRegionFields);
DingdocsScript.registerScript("getRecordOptions", getRecordOptions);
DingdocsScript.registerScript("getCurrentSelectedRecord", getCurrentSelectedRecord);
DingdocsScript.registerScript("insertRegionRecord", insertRegionRecord);
DingdocsScript.registerScript("updateRegionRecord", updateRegionRecord);
