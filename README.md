# Garmin FIT 跑步分析資料

把 Garmin Connect 匯出的 Original FIT 活動檔轉成固定格式的 Excel：`跑步分析資料 v1.0`。

目前 App 版本：`v1.4.1`  
目前 Excel 格式版本：`跑步分析資料 v1.0`

這個專案的目的不是只看單次活動，而是長期累積一致格式的跑步資料，之後可以比較鞋款、天氣、心率、功率、Stamina 消耗、跑姿指標與訓練效果。

## 資料夾

```text
FIT/       放 Garmin Original FIT 檔
EXCEL/     轉出的 Excel 檔
config/    下拉選單設定
```

主要檔案：

```text
fit_to_excel.py                  FIT 轉 Excel 主程式
inspect_fit.py                   檢查 FIT 欄位用的小工具
config/dropdown_options.json     Excel 下拉選單設定檔
requirements.txt                 Python 套件需求
```

## 安裝

建議在專案目錄執行：

```bash
pip install -r requirements.txt
```

需要的主要套件：

```text
garmin-fit-sdk
openpyxl
```

## 應用程式模式

macOS 可直接雙擊：

```text
跑步分析資料轉檔.command
```

Windows 可直接雙擊：

```text
跑步分析資料轉檔.bat
```

啟動檔會自動建立 `.venv`、安裝需求套件，並啟動本機網頁應用：

```text
http://127.0.0.1:8765
```

使用方式：

```text
1. 開啟應用後按「從電腦選擇 FIT 檔」，或先把 FIT 檔放進 FIT/ 再從清單選擇
2. 視需要填鞋款、課表、訓練目的、補給、備註等資料
3. 按「轉成 Excel」
4. 轉檔完成後會顯示活動摘要
5. 可以開啟 Excel、下載 Excel，或直接開啟 EXCEL/ 資料夾
```

`訓練目的` 可多選；macOS 按 `Command`，Windows 按 `Ctrl`，再點選多個項目。輸出到 Excel 時會合併成同一欄。

如果使用檔案選取按鈕，App 會先把該檔複製到 `FIT/`，再進行轉檔。

應用程式也有「下拉選單設定」頁面，可以直接修改活動資訊裡的鞋款、課表類型、訓練目的與 Garmin 主觀感受選項。每行一個選項，儲存後會更新：

```text
config/dropdown_options.json
```

```text
原本清單模式仍可使用：
1. 把 Garmin Original FIT 檔放進 FIT/
2. 從最近清單選擇 FIT 檔
3. 視需要填鞋款、課表、補給、備註等資料
4. 按「轉成 Excel」
5. 產出的檔案會放在 EXCEL/
```

FIT 清單預設只顯示 `FIT/` 裡最近 30 個檔案；檔案很多時建議直接使用檔案選取按鈕。

如果轉檔失敗，App 會用較容易理解的方式提示常見原因，例如 FIT 檔格式不正確、檔案權限問題、天氣查詢逾時，或 FIT 裡沒有可用的每公里分段資料。

## 基本用法

把 FIT 檔放到 `FIT/` 後執行：

```bash
python3 fit_to_excel.py FIT/20260703_ACTIVITY.fit --max-hr 173 --critical-power 315
```

預設會用 FIT 裡的活動開始時間與 GPS 起點座標，向 Open-Meteo 查詢歷史天氣，並自動填入氣溫、濕度、風向、風速。

預設會輸出到：

```text
EXCEL/跑步分析資料 v1.0_20260703_ACTIVITY.xlsx
```

## 自動抓天氣

自動抓天氣預設已開啟：

```bash
python3 fit_to_excel.py FIT/20260703_ACTIVITY.fit \
  --max-hr 173 \
  --critical-power 315
```

會自動填入：

```text
天氣氣溫(°C)
濕度(%)
風向
風速
```

注意：自動抓天氣會把活動時間與起點座標送到 Open-Meteo。若不想送出位置資料，可以加上 `--no-fetch-weather`，之後在 Excel 手動填寫。

## 常用參數

```bash
python3 fit_to_excel.py FIT/20260703_ACTIVITY.fit \
  --max-hr 173 \
  --critical-power 315 \
  --recovery-time-hr 59 \
  --shoe "EVO SL" \
  --workout-type "Recovery Run（恢復跑）" \
  --training-focus "Recovery" \
  --fueling "跑前咖啡，跑中無補給"
```

也可以用互動模式逐項輸入：

```bash
python3 fit_to_excel.py FIT/20260703_ACTIVITY.fit --interactive
```

## Excel 內容

輸出的工作簿包含：

```text
活動資訊
每公里數據
選項
圖表
```

`每公里數據` 是 v1.0 的固定主表，目前維持 18 欄：

```text
公里
距離(m)
時間(秒)
配速(分:秒/km)
平均心率
平均心率%
最高心率
平均步頻(spm)
平均功率(W)
平均功率%
垂直振幅(mm)
垂直比(%)
觸地時間(ms)
步幅(mm)
溫度(°C)
Stamina 起
Stamina 末
爬升(m)
```

`平均心率%` 會用活動資訊裡的最大心率計算。  
`平均功率%` 會用活動資訊裡的 Critical Power 計算。

最大心率與 Critical Power 會優先從 FIT 的 zone 設定自動帶入；如果 FIT 裡沒有這些欄位，也可以用參數或在 Excel 手動填寫。

## Garmin 欄位

目前會從 FIT 自動帶入：

```text
Garmin 主觀感受
最大心率
Critical Power
Training Effect (Aerobic)
Training Effect (Anaerobic)
Training Load
Stamina 起 / 末
```

Stamina 來自 Garmin FIT 裡尚未由 SDK 命名的 record 欄位，目前已確認：

```text
record 137 / 138
session 205 / 206 / 207
```

`Recovery Time (hr)` 目前保留為手動或參數輸入，因為目前這些 FIT 檔裡沒有確認到公開命名欄位。

## 修改下拉選單

下拉選單設定在：

```text
config/dropdown_options.json
```

可以修改：

```json
{
  "shoes": [],
  "workout_types": [],
  "training_focus": [],
  "garmin_rpe": []
}
```

例如新增鞋款，只要編輯：

```json
"shoes": [
  "Boston 13 Green",
  "Boston 13 Blue",
  "EVO SL",
  "Rebel v5",
  "Nimbus 28",
  "New Shoe"
]
```

下次轉檔時，Excel 的下拉選單就會更新。

## 指定輸出檔

預設輸出到 `EXCEL/`。如果要指定檔案：

```bash
python3 fit_to_excel.py FIT/20260703_ACTIVITY.fit -o EXCEL/custom.xlsx
```

## 檢查 FIT 欄位

如果想看 FIT 裡有哪些訊息與未知欄位：

```bash
python3 inspect_fit.py FIT/20260703_ACTIVITY.fit --json fit_inspection.json
```

## 版本規則

`跑步分析資料 v1.0` 的主表欄位應維持穩定，方便長期累積後做比較。

如果未來要改 `每公里數據` 的欄位順序、欄位名稱或計算邏輯，建議升成新版本，例如：

```text
跑步分析資料 v1.1
```

這樣一年後資料才不會混在一起。
