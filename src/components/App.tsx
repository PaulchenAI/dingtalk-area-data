/* global Dingdocs */

import { useEffect, useMemo, useRef, useState } from "react";
import { initView } from "dingtalk-docs-cool-app";
import { areaList } from "@vant/area-data";
import "./style.css";

declare const Dingdocs: any;

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

const provinceMap = areaList.province_list as Record<string, string>;
const cityMap = areaList.city_list as Record<string, string>;
const countyMap = areaList.county_list as Record<string, string>;

function App() {
  const [ready, setReady] = useState(false);
  const [sheetName, setSheetName] = useState("");
  const [selectedRecordId, setSelectedRecordId] = useState("");
  const [selectedRowNumber, setSelectedRowNumber] = useState<number | null>(null);

  const [level, setLevel] = useState<RegionLevel>("city");
  const [provinceCode, setProvinceCode] = useState("");
  const [cityCode, setCityCode] = useState("");
  const [districtCode, setDistrictCode] = useState("");

  const [autoWrite, setAutoWrite] = useState(false);
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] = useState<"success" | "error" | "info">("info");

  const lastAutoWriteKeyRef = useRef("");
  const isWritingRef = useRef(false);

  useEffect(() => {
    initView({
      onReady: async () => {
        setReady(true);
        await initData();
      },
      onError: (error) => {
        console.error(error);
        setMessage(`插件初始化失败：${error?.message || String(error)}`);
        setMessageType("error");
      },
    });
  }, []);

  const provinceEntries = useMemo(() => {
    return Object.entries(provinceMap);
  }, []);

  const cityEntries = useMemo(() => {
    if (!provinceCode) return [];

    return Object.entries(cityMap).filter(([code]) => {
      return code.slice(0, 2) === provinceCode.slice(0, 2);
    });
  }, [provinceCode]);

  const districtEntries = useMemo(() => {
    if (!cityCode) return [];

    return Object.entries(countyMap).filter(([code]) => {
      return code.slice(0, 4) === cityCode.slice(0, 4);
    });
  }, [cityCode]);

  const region = useMemo<RegionPayload>(() => {
    const provinceName = provinceCode ? provinceMap[provinceCode] : "";
    const cityName = cityCode ? cityMap[cityCode] : "";
    const districtName = districtCode ? countyMap[districtCode] : "";

    let regionName = provinceName;
    let regionCode = provinceCode;

    if (level === "city") {
      regionName = [provinceName, cityName].filter(Boolean).join("-");
      regionCode = cityCode || provinceCode;
    }

    if (level === "district") {
      regionName = [provinceName, cityName, districtName].filter(Boolean).join("-");
      regionCode = districtCode || cityCode || provinceCode;
    }

    return {
      level,
      levelName:
        level === "province" ? "省" : level === "city" ? "省市" : "省市区县",
      provinceCode,
      provinceName,
      cityCode: level === "province" ? "" : cityCode,
      cityName: level === "province" ? "" : cityName,
      districtCode: level === "district" ? districtCode : "",
      districtName: level === "district" ? districtName : "",
      regionCode,
      regionName,
    };
  }, [level, provinceCode, cityCode, districtCode]);

  const isRegionComplete = useMemo(() => {
    if (!provinceCode) return false;
    if (level === "city" && !cityCode) return false;
    if (level === "district" && !districtCode) return false;
    return true;
  }, [level, provinceCode, cityCode, districtCode]);

  useEffect(() => {
    if (!ready) return;

    const offSelectionChanged = Dingdocs?.base?.event?.onSelectionChanged?.(
      async (selection: any) => {
        const recordId = selection?.recordId;
        if (!recordId || recordId === selectedRecordId) return;

        try {
          await syncSelectedRecordFromSheet(false);
        } catch (error) {
          console.warn("sync selected record failed", error);
        }
      }
    );

    return () => {
      if (typeof offSelectionChanged === "function") {
        offSelectionChanged();
      }
    };
  }, [ready, selectedRecordId]);

  useEffect(() => {
    if (!ready) return;
    syncSelectedRecordFromSheet(false);
  }, [ready]);

  useEffect(() => {
    if (!autoWrite || !isRegionComplete || loading) return;

    if (!selectedRecordId) return;

    const key = JSON.stringify({
      selectedRecordId,
      level: region.level,
      regionCode: region.regionCode,
    });

    if (lastAutoWriteKeyRef.current === key) return;
    if (isWritingRef.current) return;

    lastAutoWriteKeyRef.current = key;
    isWritingRef.current = true;
    handleWrite().finally(() => {
      isWritingRef.current = false;
    });
  }, [autoWrite, isRegionComplete, selectedRecordId, region.level, region.regionCode, loading]);

  async function initData() {
    try {
      await Dingdocs.script.run("ensureRegionFields");

      const info = await Dingdocs.script.run("getActiveSheetInfo");
      setSheetName(info?.name || "当前数据表");
    } catch (error: any) {
      console.error(error);
      setMessage(`初始化数据失败：${error?.message || String(error)}`);
      setMessageType("error");
    }
  }

  function validate() {
    if (!provinceCode) return "请选择省份";
    if (level === "city" && !cityCode) return "请选择城市";
    if (level === "district" && !districtCode) return "请选择区县";
    if (!selectedRecordId) return "请先在表格中点击要更新的行";
    return "";
  }

  async function handleWrite() {
    const error = validate();

    if (error) {
      setMessage(error);
      setMessageType("error");
      return;
    }

    setLoading(true);
    setMessage("");

    try {
      await Dingdocs.script.run("updateRegionRecord", selectedRecordId, region);

      const rowText = selectedRowNumber ? `第 ${selectedRowNumber} 行` : "当前行";
      const successMessage = `已更新${rowText}：${region.regionName}`;

      setMessage(successMessage);
      setMessageType("success");
    } catch (error: any) {
      console.error(error);
      const errorMessage = `写入失败：${error?.message || String(error)}`;
      setMessage(errorMessage);
      setMessageType("error");
    } finally {
      setLoading(false);
    }
  }

  async function syncSelectedRecordFromSheet(showError = true) {
    try {
      const selected = await Dingdocs.script.run("getCurrentSelectedRecord");

      if (!selected?.id) {
        setSelectedRecordId("");
        setSelectedRowNumber(null);
        if (showError) {
          setMessage("请先在表格中点击要更新的行");
          setMessageType("error");
        }
        return;
      }

      setSelectedRecordId(selected.id);
      setSelectedRowNumber(
        typeof selected.rowNumber === "number" && selected.rowNumber > 0
          ? selected.rowNumber
          : null
      );

      if (showError) {
        const rowText =
          typeof selected.rowNumber === "number" && selected.rowNumber > 0
            ? `第 ${selected.rowNumber} 行`
            : "当前行";
        setMessage(`已同步${rowText}`);
        setMessageType("info");
      }
    } catch (error: any) {
      console.error(error);
      if (showError) {
        setMessage(`同步选中行失败：${error?.message || String(error)}`);
        setMessageType("error");
      }
    }
  }

  function handleLevelChange(nextLevel: RegionLevel) {
    setLevel(nextLevel);

    if (nextLevel === "province") {
      setCityCode("");
      setDistrictCode("");
    }

    if (nextLevel === "city") {
      setDistrictCode("");
    }
  }

  function handleProvinceChange(nextProvinceCode: string) {
    setProvinceCode(nextProvinceCode);
    setCityCode("");
    setDistrictCode("");
  }

  function handleCityChange(nextCityCode: string) {
    setCityCode(nextCityCode);
    setDistrictCode("");
  }

  return (
    <div className="region-plugin">
      <header className="header">
        <div>
          <h1>
            <svg
              className="title-icon"
              viewBox="0 0 1024 1024"
              xmlns="http://www.w3.org/2000/svg"
              aria-hidden="true"
            >
              <path d="M920 286.628c0-20.228-16.4-36.628-36.628-36.628H197.092C145.68 250 104 291.68 104 343.092v470.66a96.26 96.26 0 0 0 28.188 68.06 96.26 96.26 0 0 0 68.06 28.188h683.124c20.228 0 36.628-16.4 36.628-36.628V286.628z" fill="#D5F0FE" />
              <path d="M868 320.708a22.708 22.708 0 0 0-22.708-22.708h-394.584A22.708 22.708 0 0 0 428 320.708v250.584a22.708 22.708 0 0 0 22.708 22.708h394.584a22.708 22.708 0 0 0 22.708-22.708V320.708z" fill="#63B6F0" />
              <path d="M868 664.708a22.708 22.708 0 0 0-22.708-22.708h-181.92a22.7 22.7 0 0 0-21.08 14.276l-61.836 154.584a22.704 22.704 0 0 0 21.084 31.14h243.752a22.708 22.708 0 0 0 22.708-22.708v-154.584z" fill="#93EB3B" />
              <path d="M368 320.708a22.708 22.708 0 0 0-22.708-22.708H234.708A22.708 22.708 0 0 0 212 320.708v250.584a22.708 22.708 0 0 0 22.708 22.708h110.584A22.708 22.708 0 0 0 368 571.292V320.708z" fill="#00DF8E" />
              <path d="M590.504 655.744a10.032 10.032 0 0 0-9.308-13.744H130.708c-6.024 0-11.8 2.392-16.056 6.652A22.692 22.692 0 0 0 108 664.708v109.672c0 17.932 7.124 35.132 19.804 47.816a67.636 67.636 0 0 0 47.816 19.804h333.596a10.024 10.024 0 0 0 9.304-6.3l71.984-179.956z" fill="#EC9982" />
              <path d="M280 730H192.416A88.424 88.424 0 0 0 104 818V207.684A93.692 93.692 0 0 1 197.684 114h52.824A29.492 29.492 0 0 1 280 143.492V730z" fill="#C0E1F7" />
              <path d="M654 118c84.996 0 154 69.004 154 154 0 112.996-141.004 226-154 226-12.996 0-154-109.004-154-226 0-84.996 69.004-154 154-154z" fill="#FF2865" />
              <path d="M654 276m-66 0a66 66 0 1 0 132 0 66 66 0 1 0-132 0Z" fill="#E70049" />
            </svg>
            <span>区域选择助手</span>
          </h1>
          <p>{sheetName ? `当前数据表：${sheetName}` : "用于写入省 / 省市 / 省市区县字段"}</p>
        </div>
        <button className="icon-btn" onClick={initData} disabled={!ready || loading}>
          刷新
        </button>
      </header>

      {!ready && <div className="card muted">插件初始化中...</div>}

      <section className="card">
        <div className="section-title">1. 当前选中行</div>
        <div className="update-tip">在表格中点击要更新的行，插件会自动跟随。</div>
        <div className="record-hint">
          当前：{selectedRowNumber ? `第 ${selectedRowNumber} 行` : "未选中任何行"}
        </div>
        <div className="record-hint">
          点击刷新会自动添加“省份 / 城市 / 区县”表格列；如果已有列，请检查列名是否一致。
        </div>
      </section>

      <section className="card">
        <div className="section-title">2. 选择区域粒度</div>

        <div className="segmented three">
          <button
            className={level === "province" ? "active" : ""}
            onClick={() => handleLevelChange("province")}
            type="button"
          >
            省
          </button>
          <button
            className={level === "city" ? "active" : ""}
            onClick={() => handleLevelChange("city")}
            type="button"
          >
            省市
          </button>
          <button
            className={level === "district" ? "active" : ""}
            onClick={() => handleLevelChange("district")}
            type="button"
          >
            省市区县
          </button>
        </div>

        <label>省份</label>
        <select value={provinceCode} onChange={(event) => handleProvinceChange(event.target.value)}>
          <option value="">请选择省份</option>
          {provinceEntries.map(([code, name]) => (
            <option key={code} value={code}>
              {name}
            </option>
          ))}
        </select>

        {level !== "province" && (
          <>
            <label>城市</label>
            <select
              value={cityCode}
              onChange={(event) => handleCityChange(event.target.value)}
              disabled={!provinceCode}
            >
              <option value="">请选择城市</option>
              {cityEntries.map(([code, name]) => (
                <option key={code} value={code}>
                  {name}
                </option>
              ))}
            </select>
          </>
        )}

        {level === "district" && (
          <>
            <label>区县</label>
            <select
              value={districtCode}
              onChange={(event) => setDistrictCode(event.target.value)}
              disabled={!cityCode}
            >
              <option value="">请选择区县</option>
              {districtEntries.map(([code, name]) => (
                <option key={code} value={code}>
                  {name}
                </option>
              ))}
            </select>
          </>
        )}
      </section>

      <section className="card">
        <label className="check-row">
          <input
            type="checkbox"
            checked={autoWrite}
            onChange={(event) => setAutoWrite(event.target.checked)}
          />
          <span>选择完成后自动写入</span>
        </label>

        <button
          className="primary-btn"
          onClick={handleWrite}
          disabled={loading || !isRegionComplete || !selectedRecordId}
          type="button"
        >
          {loading ? "写入中..." : "更新当前选中行"}
        </button>

        {message && (
          <div className={`message ${messageType}`}>
            {messageType === "success" && <span className="message-check" aria-hidden="true" />}
            <span>{message}</span>
          </div>
        )}
      </section>

    </div>
  );
}

export default App;
