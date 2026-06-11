1.郵件寄送與防錯管理 (Email Management & Distribution)
將整理好的數據轉化為報表，並導入嚴格的「防錯與審查」機制，確保每次的群發都安全無虞。
- 即時防錯 (onEdit)： 這是系統最強健的設計之一。當你在試算表輸入信箱時，系統就已經即時糾正了 @gamil.com 這類手誤，大幅降低發信 API 報錯的機率。
- 安全攔截與審查： 寄送前會區分「合格收件人」與「格式錯誤者」。遇到錯誤時，系統並不會直接當機，而是會將錯誤者剔除，並發送信件通知管理員進場維護，確保其他正常訂閱者的權益不受影響。

```mermaid
graph TD
    classDef check fill:#e8f0fe,stroke:#1a73e8,stroke-width:2px,color:#174ea6;
    classDef mail fill:#e6f4ea,stroke:#1e8e3e,stroke-width:2px,color:#0d652d;
    classDef error fill:#fce8e6,stroke:#d93025,stroke-width:2px,color:#b31412;
    classDef sheet fill:#fef7e0,stroke:#fbbc04,stroke-width:2px,color:#b06000;

    A[Read Email Management Config] --> B{Check Subscription Mode<br/>Daily vs. Immediate}:::check
    
    B -- Active --> C{Validate Email Format<br/>& System Status}:::check
    
    C -- Invalid Format --> D[Move to Error Blocklist]:::error
    C -- Valid Format --> E[Add to BCC Send List]
    
    E --> F[Generate HTML Report<br/>Stats & Recent Transactions]
    D --> F
    
    F --> G{Evaluate Final Mailing List}
    
    G -- Only Errors Exist --> H[Block Send & Alert Admin<br/>about formatting issues]:::error
    G -- No Subscribers --> I[Send Admin-Only Tracking Report]:::mail
    G -- Valid Subscribers Exist --> J[Group BCC Send to Subscribers]:::mail
    
    J --> K[(Update Send Time &<br/>Status Checkmarks)]:::sheet
    J --> L[Send Delivery Summary<br/>& HTML Preview to Admin]:::mail
```
