# 安裝、部署與疑難排解

本文件說明如何把專案部署到 GitHub、首次執行、每日排程，以及抓取失敗時的處理方式。專案概要請見 [README.md](README.md)。

## 1. 推上 GitHub

把專案檔案放在同一個資料夾（須包含 `.github/workflows/daily.yml`），於該資料夾執行：

```bash
git init -b main
git add .
git commit -m "init: 外資持倉成本每日追蹤"
git remote add origin https://github.com/mywu-cloud/futures.git
git push -u origin main
```

> 若用網頁上傳（Add file → Upload files），注意 `.github` 是隱藏資料夾，部分作業系統預設不顯示，請確認 `daily.yml` 有一起上傳，否則排程不會啟動。

## 2. 確認 Actions 寫入權限

workflow 已宣告 `permissions: contents: write`。如果 commit 步驟出現 403，請到：

Settings → Actions → General → Workflow permissions → 勾選 **Read and write permissions** → Save

## 3. 首次執行（回補歷史）

1. 到 repo 的 **Actions** 頁，選「外資持倉成本每日更新」
2. 按 **Run workflow**，`backfill_days` 預設 730（回補期交所兩年）
3. 約 1～2 分鐘後完成，`data/` 目錄會出現 4 個資料檔

回補天數只在 `data/taifex_raw.csv` 不存在時生效；之後每次執行只抓最後一筆日期前 3 天起的新資料。回補越長，自算成本受起始日的影響越小。

## 4. 每日排程

| 項目 | 設定 |
|---|---|
| 時間 | 週一至週五 UTC 10:00（台灣 18:00） |
| 修改方式 | 編輯 `daily.yml` 的 `cron: "0 10 * * 1-5"` |
| 資料公布時間 | 期交所三大法人資料約 15:00 後公布，建議不要早於 16:00 |

注意事項：

- GitHub 排程常延遲數分鐘到數十分鐘，屬正常現象
- **repo 連續 60 天沒有任何活動，GitHub 會自動停用排程**。本專案每個交易日都有資料 commit，正常情況不會被停用；若曾停用，到 Actions 頁按 Enable workflow 即可
- 假日執行不會產生新資料，也不會 commit

## 5. 執行結果與通知

| 狀況 | 結果 |
|---|---|
| 抓取成功 | 綠勾，資料 commit |
| 抓取失敗 | `latest.json` 記錄錯誤並 commit，該次執行標示紅叉，GitHub 寄信通知 |
| 無新資料（假日／尚未公布） | 綠勾，不改寫檔案、不 commit |

`data/latest.json` 的 `status.taifex` 欄位記錄當次是 `ok` 或 `error: 原因`。

## 6. 疑難排解

### 期交所一直抓取失敗（HTTP 403 / 連線逾時 / 查無任何資料）

GitHub Actions 主機位於美國機房，期交所可能限制海外連線。若持續失敗，改用 **self-hosted runner**，在自己的電腦上以台灣家用 IP 執行：

1. repo → Settings → Actions → Runners → **New self-hosted runner**
2. 選作業系統（Windows / macOS / Linux），依頁面顯示的指令下載並設定 runner
3. 建議安裝成服務，開機自動啟動（Windows 設定時選擇以服務執行）
4. 電腦需安裝 Python 3.10 以上
5. 修改 `daily.yml`：

```yaml
jobs:
  update:
    runs-on: self-hosted   # 原為 ubuntu-latest
```

6. self-hosted runner 若 `actions/setup-python` 步驟出錯可移除，改用電腦本身的 Python

> 排程時間到時電腦須開機並連網；資料一樣會自動 commit 回 repo。

### 期交所出現「找不到欄位」

期交所 CSV 欄位名稱改版時會發生。程式以關鍵字比對欄位，對小幅改版有容錯；若仍失敗，log 會印出實際表頭，依表頭調整 `parse_taifex_inst()` 或 `parse_taifex_price()` 中 `find_col()` 的關鍵字。

### 數值與玩股網不同

重置法與玩股網算法相同，但玩股網的資料起點約為兩年前；結算日會重置成本，因此只要回補涵蓋最近一個結算日，最新數值應一致。若不一致，先確認期交所當日資料是否已更新（約 15:00 後公布），或以 `--rebuild` 重跑。

### commit 時出現衝突

workflow 在 push 前會先 `git pull --rebase`。若你同時在本機修改並推送 `data/` 內的檔案，可能發生衝突；建議不要手動編輯 `data/`，需要調整時改程式後重跑。

### 想重算全部歷史

刪除 `data/taifex_raw.csv` 與 `data/taifex_calc.csv` 並 commit，再手動 Run workflow 並設定 `backfill_days`。

## 7. 公開或私人 repo

| | Public | Private |
|---|---|---|
| Actions 額度 | 無限 | 每月 2,000 分鐘免費（本工作每次約 1 分鐘，足夠） |
| raw 連結直接讀 `latest.json` | 可以（網站卡片可直接使用） | 需 token 或由網站後端讀取 |

資料全部來自期交所公開資料並自行計算，不含第三方網站加工資料；網站要直接讀 `latest.json` 時用 public 即可。

## 8. 本機測試

```bash
pip install -r requirements.txt
python foreign_cost_tracker.py --data-dir test_data --backfill-days 30
```

用 `--data-dir` 指定其他目錄，避免影響正式資料；`test_data/` 不要 commit。

## 免責聲明

本專案數據與計算結果僅供參考，不構成任何投資建議，投資人應自行判斷並承擔風險。資料以臺灣期貨交易所公告為準，使用者須遵守各資料來源網站之使用規範。
