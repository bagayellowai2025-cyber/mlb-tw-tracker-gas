1.爬蟲與資料擷取 (Data Fetching & Scraping)
系統會先確認球員的基礎識別碼，再兵分兩路去撈取「狀態資訊」與「賽季數據」。
- API 補漏機制： 若清單中只有球員名字而沒有網址，系統會先呼叫 MLB 官方 API 查詢 Player ID 並自動補齊網址，確保後續爬蟲有目標可抓。
- 精準定位數據： scrapeDetailedStats 實作了很聰明的判斷邏輯。它會根據球員目前的層級（MLB 還是 MiLB），去尋找網頁表格中對應年份的 Regular Season 或 MLB/MiLB Stats 行，避免抓到春訓或生涯總計的無效數據。

2.資料庫更新與近況比對 (Database Update & Delta Calc)
抓回來的資料不僅要填入表格，系統還會建立「歷史快照」，用來計算近五天內的數據變化（Delta），讓報告更具洞察力。
- 歷史快照庫 (Stats_History)： 將當天的數據序列化成 JSON 隱藏儲存，避免試算表欄位過度膨脹。
- 投球局數 (IP) 演算法： 棒球的局數進位是 0.1, 0.2 接著跳整數，這裡特別針對 calcIPDiff 做了處理（先轉成出局數，相減後再轉回局數），這是非常專業的細節處理。

```mermaid
graph TD
    classDef api fill:#fef7e0,stroke:#fbbc04,stroke-width:2px,color:#b06000;
    classDef scrape fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px,color:#174ea6;
    classDef sheet fill:#e6f4ea,stroke:#1e8e3e,stroke-width:2px,color:#0d652d;
    classDef calc fill:#fce8e6,stroke:#d93025,stroke-width:2px,color:#b31412;

    A[Read Player Roster<br/>from Sheet] --> B{Missing ID or URLs?}
    
    B -- Yes --> C[Call MLB API to search player]:::api
    C --> D[Generate & save MLB/MiLB URLs]
    
    B -- No --> E[Data Scraping Branch]
    D --> E
    
    E --> F[Scrape Player Profile<br/>Status, Level, Team, Moves]:::scrape
    E --> G[Scrape Season Stats<br/>Auto-detect Minor/Major League]:::scrape
    
    F --> H[(Update Basic Info & Highlights<br/>in Database)]:::sheet
    
    G --> I[(Save Daily Stats Snapshot<br/>to Hidden History Log)]:::sheet
    I --> J[Retrieve Closest Past Record<br/>within last 1-5 Days]
    
    J --> K{Has the player<br/>played recently?}
    
    K -- No / No Record --> L[(Log 'No Game' or 'Pending')]:::sheet
    K -- Yes --> M[Calculate Stat Changes / Delta<br/>incl. Innings Pitched conversion]:::calc
    
    M --> N[(Update Stats & Apply<br/>Positive/Negative Color Codes)]:::sheet
```
