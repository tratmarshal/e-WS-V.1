graph TD
    %% Define Elements
    User([👤 ผู้ใช้งาน / ตำรวจศาล])
    LIFF[📱 LINE LIFF App <br/> Frontend: HTML5/JS]
    LINE[🟢 LINE Platform <br/> Login & Profile API]
    GAS[⚡ Google Apps Script <br/> API Gateway / Backend]
    Sheet[(📊 Google Sheets <br/> Database / Central Store)]

    %% Connections
    User -->|เปิดเมนูผ่าน LINE| LIFF
    LIFF -->|1. ตรวจสอบสิทธิ์ (Token/Profile)| LINE
    LIFF <-->|2. HTTP POST (JSON) / 5. ตอบกลับข้อมูล| GAS
    GAS <-->|3. อ่าน/เขียน ข้อมูล / 4. คืนค่าผลลัพธ์| Sheet
    LIFF -->|6. Render UI อัปเดตหน้าจอ| User

    %% Styling
    style User fill:#f8fafc,stroke:#94a3b8,stroke-width:2px,color:#0f172a
    style LIFF fill:#1e293b,stroke:#3b82f6,stroke-width:2px,color:#fff
    style LINE fill:#00b900,stroke:#009900,stroke-width:2px,color:#fff
    style GAS fill:#0f172a,stroke:#f59e0b,stroke-width:2px,color:#fff
    style Sheet fill:#064e3b,stroke:#10b981,stroke-width:2px,color:#fff

flowchart TD
    %% Authentication
    Start([🚀 เริ่มต้นเข้าใช้งานผ่าน LINE]) --> Auth1[ระบบเรียก liff.init() & liff.getProfile()]
    Auth1 --> Auth2{ตรวจสอบสิทธิ์ในฐานข้อมูล}
    
    Auth2 -->|Unauthorized| Reject[🚫 หน้าจอแสดง User ID <br/> แจ้งแอดมินเพื่อลงทะเบียน]
    Auth2 -->|Authorized| Menu{เลือกเมนูการทำงาน}

    %% App A: Search
    Menu -->|🔍 A. หน้าค้นหาหมายจับ| S1[เลือกโหมดค้นหา <br/> เลขบัตร ปชช. / ชื่อ-สกุล]
    S1 --> S2[ระบบส่งคำขอ action: 'search' ไปยัง GAS]
    S2 --> S3[แสดงผลข้อมูลหมายจับ <br/> สถานะ, ข้อหา, เลขคดี]

    %% App B: Management
    Menu -->|📝 B. หน้าจัดการรายการ| M1[ค้นหาหมายจับค้างเก่า]
    M1 --> M2[ติ๊กเลือกหมายจับที่ยัง 'ไม่เพิกถอน']
    M2 --> M3[กรอกฟอร์ม: <br/> เหตุที่สิ้นผล & เลือกผู้พิพากษา]
    M3 --> M4[ทบทวนความถูกต้อง <br/> Confirm Summary Box]
    M4 --> M5[ยืนยันการบันทึก <br/> action: 'addProcess']
    
    M5 --> End([✅ บันทึกรายการใหม่สำเร็จ <br/> เข้าสู่ลูปสถานะภารกิจ])

    %% Styling
    style Start fill:#2563eb,stroke:#1e40af,stroke-width:2px,color:#fff
    style Reject fill:#ef4444,stroke:#b91c1c,stroke-width:2px,color:#fff
    style End fill:#10b981,stroke:#047857,stroke-width:2px,color:#fff
    style Menu fill:#8b5cf6,stroke:#6d28d9,stroke-width:2px,color:#fff

    stateDiagram-v2
    [*] --> เสนอศาล : 1. บันทึกเสนอรายงานสิ้นผล (addProcess)
    
    state เสนอศาล {
        [*] --> รอพิจารณา
    }
    
    เสนอศาล --> รอส่งต่อ : 2. ผู้พิพากษาลงนามแล้ว \n(กดปุ่ม "เพิกถอนหมายจับ" / markRevoked)
    
    state รอส่งต่อ {
        [*] --> เตรียมเอกสารส่งคืน
    }
    
    รอส่งต่อ --> ส่งต่อแล้ว : 3. จัดส่งสำนวนคืนหน่วยงานแล้ว \n(กดปุ่ม "ส่งต่อสำนวน" / markForwarded)
    
    ส่งต่อแล้ว --> [*] : ย้ายข้อมูลไปยัง 'คลังประวัติ' (History Tab)

flowchart TD
    %% Process Start
    Trigger[User สั่งบันทึก/ดึงข้อมูล] --> Req[ฟังก์ชัน callGas() ทำงาน]
    Req --> Call{พยายามส่ง Request \n ไปยัง Backend}
    
    %% Success Path
    Call -->|✅ สำเร็จ| Success[ประมวลผลสำเร็จ & \n อัปเดต UI]
    
    %% Fail Path
    Call -->|❌ ล้มเหลว / เน็ตหลุด| CheckRetry{จำนวนครั้งที่ล้มเหลว \n < 3 ครั้ง หรือไม่?}
    
    %% Retry Mechanism
    CheckRetry -->|ใช่| Wait[⏳ หน่วงเวลา Exponential Backoff]
    Wait --> Call
    
    %% Hard Fail
    CheckRetry -->|ไม่ใช่| ErrorUI[⚠️ แสดง Error Modal \n ให้ผู้ใช้ทราบ]
    ErrorUI --> Action[ผู้ใช้กดปุ่ม 'ลองใหม่อีกครั้ง']
    Action --> Req

    %% Background Sync
    Load[เมื่อโหลดหน้า Manage ใหม่] --> Pending{มีคิว Sync \n ค้างอยู่หรือไม่?}
    Pending -->|มีค้างอยู่| BGSync[🔄 triggerBackgroundSync() \n ทำงานอัตโนมัติ]
    BGSync --> Call

    %% Styling
    style Success fill:#10b981,stroke:#047857,stroke-width:2px,color:#fff
    style ErrorUI fill:#ef4444,stroke:#b91c1c,stroke-width:2px,color:#fff
    style Wait fill:#f59e0b,stroke:#b45309,stroke-width:2px,color:#fff
