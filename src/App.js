import React, { useState } from "react";
import Plot from "react-plotly.js";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";

function App() {
  const [dept, setDept] = useState("");
  const [yearCols, setYearCols] = useState([]);
  const [rawData, setRawData] = useState([]);
  const [processedData, setProcessedData] = useState([]);
  const [chartData, setChartData] = useState(null);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheetRaw = workbook.Sheets[sheetName];

      // 取得欄位名稱，確保空欄也保留
      const headers = [];
      const range = XLSX.utils.decode_range(sheetRaw["!ref"]);
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell = sheetRaw[XLSX.utils.encode_cell({ r: 0, c: C })];
        let hdr = cell && cell.t ? XLSX.utils.format_cell(cell) : `UNKNOWN ${C}`;
        headers.push(hdr);
      }

      // 年份欄位
      const years = headers.filter((h) => h.includes("年"));
      setYearCols(years);

      // 原始資料（空值補 0）
      const raw = XLSX.utils.sheet_to_json(sheetRaw, { defval: 0 });
      setRawData(raw);

      // 計算結果（轉換成萬元）
      const processed = raw.map((row) => {
        const projName = `${row["計畫編號"]}：${row["計畫名稱"] || "未填"}`;
        const contractAmt = row["合約金額"] || 0;
        const yearlyValues = years.map(
          (y) => ((contractAmt * (row[y] || 0)) / 10000).toFixed(2)
        );
        const percents = years.map(
          (y) => ((row[y] || 0) * 100).toFixed(0) + "%"
        );
        return { projName, contractAmt, yearlyValues, percents };
      });
      setProcessedData(processed);

      // 堆疊圖資料
      const traces = processed.map((proj) => ({
        x: years,
        y: proj.yearlyValues.map((v) => parseFloat(v)),
        name: proj.projName,
        type: "bar",
        text: years.map(
          (y, i) => `${proj.projName.split("：")[0]}<br>${proj.percents[i]}`
        ),
        textposition: "inside",
        insidetextanchor: "middle"
      }));
      setChartData(traces);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDownloadImage = () => {
    const chartElement = document.getElementById("chart-container");
    html2canvas(chartElement, { scale: 2 }).then((canvas) => {
      const link = document.createElement("a");
      link.download = "計畫金額累計圖.png";
      link.href = canvas.toDataURL("image/png");
      link.click();
    });
  };

  return (
    <div style={{ padding: "20px", fontFamily: "Arial" }}>
      <h1>📊 計畫金額累計圖工具</h1>

      {/* 說明框 */}
      <div
        style={{
          border: "2px solid #ccc",
          borderRadius: "8px",
          padding: "10px",
          background: "#f9f9f9",
          marginBottom: "20px",
          maxWidth: "600px"
        }}
      >
        <ol>
          <li>點擊「📥 下載 Excel 範例檔案」取得正確格式。</li>
          <li>在 Excel 檔案中填寫計畫編號、名稱、合約金額與年度比例。</li>
          <li>選擇部門。</li>
          <li>上傳填寫完成的 Excel 檔後自動繪圖。</li>
          <li>可下載 PNG 圖檔。</li>
        </ol>
      </div>

      {/* 下載按鈕 */}
      <a
        href="計畫列表範例.xlsx"
        download="計畫列表範例.xlsx"
        style={{
          display: "inline-block",
          padding: "8px 12px",
          background: "#4CAF50",
          color: "white",
          borderRadius: "4px",
          textDecoration: "none",
          marginBottom: "10px"
        }}
      >
        📥 下載 Excel 範例檔案
      </a>

      {/* 部門選擇 */}
      <div style={{ marginTop: "10px" }}>
        <label>部門：</label>
        <select
          value={dept}
          onChange={(e) => setDept(e.target.value)}
          style={{ padding: "5px", fontSize: "14px" }}
        >
          <option value="">請選擇</option>
          <option value="運輸二部">運輸二部</option>
          <option value="運輸三部">運輸三部</option>
          <option value="研究室">研究室</option>
          <option value="AITS">AITS</option>
          <option value="行政室">行政室</option>
        </select>
      </div>

      <br />
      <input type="file" accept=".xlsx" onChange={handleFileUpload} />

      {/* 原始檔案內容 */}
      {rawData.length > 0 && (
        <>
          <h2>📄 原始檔案內容</h2>
          <table border="1" cellPadding="5">
            <thead>
              <tr>
                {Object.keys(rawData[0]).map((col) => (
                  <th key={col}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rawData.map((row, i) => (
                <tr key={i}>
                  {Object.values(row).map((val, j) => (
                    <td key={j}>{val}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      {/* 計算結果 */}
      {processedData.length > 0 && (
        <>
          <h2>📊 計算結果（單位：萬元）</h2>
          <table border="1" cellPadding="5">
            <thead>
              <tr>
                <th>計畫</th>
                <th>合約金額</th>
                {yearCols.map((y) => (
                  <th key={y}>{y}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {processedData.map((proj, i) => (
                <tr key={i}>
                  <td>{proj.projName}</td>
                  <td>{proj.contractAmt}</td>
                  {proj.yearlyValues.map((v, j) => (
                    <td key={j}>{v}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      {/* 圖表 */}
      {chartData && (
        <div>
          <div id="chart-container">
            <Plot
              data={chartData}
              layout={{
                barmode: "stack",
                title: {
                  text: `${dept ? dept + " - " : ""}計畫金額累計圖`,
                  font: { size: 24 },
                  x: 0.5,
                },
                xaxis: { title: { text: "年度", font: { size: 18 } } },
                yaxis: { title: { text: "金額（萬元）", font: { size: 18 } } },
                height: 600,
                legend: { orientation: "h", y: -0.2 },
              }}
              useResizeHandler={true}
              style={{ width: "100%" }}
            />
          </div>
          <button
            onClick={handleDownloadImage}
            style={{
              padding: "8px 12px",
              background: "#2196F3",
              color: "white",
              border: "none",
              borderRadius: "4px",
              cursor: "pointer",
              marginTop: "10px",
            }}
          >
            📥 下載圖表 PNG
          </button>
        </div>
      )}
    </div>
  );
}

export default App;
