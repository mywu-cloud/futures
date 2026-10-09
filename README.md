# futures — 台指期外資持倉成本每日追蹤

每個交易日由 GitHub Actions 自動執行，以**臺灣期貨交易所官方資料**自算「台指期外資持倉成本」，兩種算法並列對照，寫回本 repo 的 `data/` 目錄。

| 算法 | 說明 |
|---|---|
| 結算重置法 | 與玩股網「外資持倉成本」相同算法（2026/10/08 驗證：成本 46,319.72、未實現 −5,040,356.31 萬、已實現 −504,718.16 萬，完全一致）。每個結算日重置成本 |
| 連續移動平均法 | 以每日成交均價做移動平均，不因結算重置，反映較長期的部位成本 |

> 玩股網的成本是在瀏覽器端以 JavaScript 計算，伺服器回傳的頁面沒有數值，無法以一般 HTTP 抓取；本專案直接用期交所資料重算，因此可回補完整歷史，也不涉及轉載第三方加工資料。

## 執行排程

- **自動**：週一至週五，台灣時間 18:00（UTC 10:00）。GitHub 排程常會延遲數分鐘到數十分鐘。
- **手動**：Actions → 外資持倉成本每日更新 → Run workflow（可設定首次回補天數，預設 730 天；可勾選強制重建）。
- 假日或資料未變動時不改寫任何檔案，也不會 commit。

## 資料檔（`data/`）

| 檔案 | 內容 |
|---|---|
| `taifex_raw.csv` | 期交所每日原始資料（外資淨部位、TX 近月高低收／結算價、次月收盤） |
| `taifex_calc.csv` | 原始資料＋兩種自算持倉成本與損益 |
| `futures_foreign_cost.xlsx` | Excel 報表：對照總表、期交所原始、走勢圖、說明（由 CSV 重建） |
| `latest.json` | 最新一日摘要＋近 120 個交易日走勢（網站前端讀取） |

網站前端讀取（repo 為 public 時）：

```
https://raw.githubusercontent.com/mywu-cloud/futures/main/data/latest.json
```

`latest.json` 結構：

```json
{
  "updated_at": "2026-10-09 18:05",
  "status": {"taifex": "ok"},
  "latest": {
    "date": "2026/10/08", "close": 49349, "contract": "202610",
    "net_oi": -83194, "net_oi_change": -4093, "settlement_day": false,
    "reset":      {"cost": 46319.72, "unrealized": -5040356.31, "realized": -504718.16, "total": -5545074.47},
    "continuous": {"cost": 46796.6,  "unrealized": -4246894.17, "realized": 1704771.54, "total": -2542122.64}
  },
  "history": {"fields": ["date","close","reset_cost","continuous_cost","net_oi"], "rows": [["2026/10/08",49349,46319.72,46796.6,-83194]]}
}
```

損益單位為萬元。（`continuous` 的數值取決於回補起始日，此處僅示意格式。）

## 網站卡片

`web/foreign-cost-card.html` 是可直接嵌入網站的單檔卡片（深色、紅漲綠跌），讀取上面的 `latest.json` 顯示成本、收盤距成本、淨口數、損益與走勢圖。可用 iframe 嵌入，或以 `?src=網址` 指定其他 JSON 位置；讀取失敗時會顯示標示為「範例資料」的內建資料。

## 算法

**結算重置法**（與玩股網相同）

1. 結算日（近月到期日）：成本 =（近月最高 + 最低）÷ 2；未實現以次月收盤計
2. 非結算日：
   - 淨口數與前日同號且同向加碼 → 成本 =（前日成本 × 前日口數 + 變動口數 × 近月收盤）÷ 今日口數
   - 同號減碼 → 成本不變，並計入已實現
   - 翻多／翻空 → 成本 =（最高 + 最低）÷ 2
3. 已實現 = 減碼口數 ×（收盤 − 前日成本）× 200，每個合約月重新累計

**連續移動平均法**

1. 淨部位變動 q ＝ 今日未平倉淨口數 − 昨日未平倉淨口數
2. 當日成交均價 p ＝ 多空交易契約金額淨額（千元）× 1000 ÷（多空交易口數淨額 × 200）；淨交易口數為 0 或偏離收盤超過 15% 時改用結算價
3. 同向加碼 → 加權平均；反向 → 先沖銷計已實現，超出部分以 p 為新成本
4. 已實現自回補起始日起累計，絕對值受起始日影響

## 本機執行

```bash
pip install -r requirements.txt
python foreign_cost_tracker.py --backfill-days 730   # 首次回補
python foreign_cost_tracker.py                       # 每日例行
```

| 參數 | 說明 |
|---|---|
| `--data-dir` | 資料輸出目錄，預設 `data/` |
| `--backfill-days` | 首次執行（無 `taifex_raw.csv`）時回補天數，預設 365 |
| `--rebuild` | 資料未變動也強制重建 xlsx / json |

## Repo 結構

```
futures/
├── .github/workflows/daily.yml   # 每日排程
├── data/                         # 自動產生的資料檔
├── web/foreign-cost-card.html    # 網站嵌入卡片
├── foreign_cost_tracker.py       # 主程式
├── requirements.txt
├── .gitattributes
├── README.md
└── SETUP.md                      # 安裝、部署與疑難排解
```

## 免責聲明

本專案數據與計算結果僅供參考，不構成任何投資建議，投資人應自行判斷並承擔風險。資料以臺灣期貨交易所公告為準，使用者須遵守期交所資料使用規範。
