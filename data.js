// Live data compiled from 5_PN_Job_Relationship_rev00-test.xlsx & 1-List of each sub-assembly.xlsx
const EXCEL_PROJECTS = [
    {
        "id": "JLP CART-51026",
        "customer": "COHU",
        "project_code": "JLP CART",
        "part_number": "8337427001",
        "description": "JLP CART",
        "status": "ontime",
        "team_leader": "NATTHAVUT TORUANG",
        "member_1": "KESORN",
        "job_no": "51026",
        "mc_number": "-",
        "progress": 0,
        "qty": 35,
        "est_hours": 2.0,
        "sub_assemblies": [
            {
                "pn": "0",
                "name": "Parts Receiving",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "1",
                "name": "Stamping PN (แสตมป์พาร์ทนัมเบอร์)",
                "th_desc": "แสตมป์พาร์ท",
                "progress": 0
            },
            {
                "pn": "2",
                "name": "Etching PN (เอจชิ่งพาร์ทนัมเบอร์)",
                "th_desc": "เอจชิ่งพาร์ท",
                "progress": 0
            },
            {
                "pn": "3",
                "name": "Pin installation-Base Plate (ติดตั้งพินที่แผ่น Base Plate)",
                "th_desc": "ติดพิน Base",
                "progress": 0
            },
            {
                "pn": "4",
                "name": "Pin installation-Top Plate (ติดตั้งพินที่แผ่น Top Plate)",
                "th_desc": "ติดพิน Top",
                "progress": 0
            },
            {
                "pn": "5",
                "name": "Pin installation-Side Plate (ติดตั้งพินที่แผ่น Top Plate)",
                "th_desc": "ติดพิน Side",
                "progress": 0
            },
            {
                "pn": "6",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ]
    },
    {
        "id": "LH & RH FRAME-49702",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578603",
        "description": "LH & RH FRAME",
        "status": "ontime",
        "team_leader": "SUWAT JANSIRI",
        "member_1": "-",
        "job_no": "49702",
        "mc_number": "#25-27",
        "progress": 25,
        "qty": 6,
        "est_hours": 0.1,
        "sub_assemblies": [
            {
                "pn": "53578603",
                "name": "LH Frame (Support LH)",
                "th_desc": "",
                "progress": 25
            }
        ]
    },
    {
        "id": "LH & RH FRAME-49703",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578703",
        "description": "LH & RH FRAME",
        "status": "ontime",
        "team_leader": "SUWAT JANSIRI",
        "member_1": "-",
        "job_no": "49703",
        "mc_number": "#25-27",
        "progress": 25,
        "qty": 6,
        "est_hours": 0.1,
        "sub_assemblies": [
            {
                "pn": "53578703",
                "name": "RH Frame (Support RH)",
                "th_desc": "",
                "progress": 25
            }
        ]
    },
    {
        "id": "PRX-49313-49317",
        "customer": "ULC",
        "project_code": "PRX",
        "part_number": "PRX250-REEL ASSY",
        "description": "PRX250 Standard Module Assembly",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49313 - 49317",
        "mc_number": "-",
        "progress": 0,
        "qty": 20,
        "est_hours": 19.5,
        "sub_assemblies": [
            {
                "pn": "PRX250-REEL ASSY",
                "name": "PRX250-REEL ASSY",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-CTL&MONITOR",
                "name": "PRX250-CTL&MONITOR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-POWER DRIVE",
                "name": "PRX250-POWER DRIVE",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-LAUNCH TUBE",
                "name": "PRX250-LAUNCH TUBE",
                "th_desc": "",
                "progress": 0
            }
        ]
    },
    {
        "id": "PRX-49288-49706",
        "customer": "ULC",
        "project_code": "PRX",
        "part_number": "PRX250-ENCODER",
        "description": "PRX250 Standard Module Assembly",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "SUMOLTHA BANCHUEN",
        "job_no": "49288 - 49706",
        "mc_number": "-",
        "progress": 0,
        "qty": 10,
        "est_hours": 19.5,
        "sub_assemblies": [
            {
                "pn": "PRX250-ENCODER",
                "name": "PRX250-ENCODER",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-PENDANT",
                "name": "PRX250-PENDANT",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRXM-001-00-000",
                "name": "PRXM-001-00-000",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRXM-003-00-000",
                "name": "PRXM-003-00-000",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRXM-004-00-000",
                "name": "PRXM-004-00-000",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRXM-005-00-000",
                "name": "PRXM-005-00-000",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-ED-CABLE-SHOR",
                "name": "SHORT  ENCODER CABLE  ASSEMBLY",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-SPRINGS-812",
                "name": "8-12\" CENTERING SPRING ASSEMBLY",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-CONTROL CABLE",
                "name": "PRX250-CONTROL CABLE",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-103-CAMS",
                "name": "PRX250-103-CAMS",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-ED-CABLE-LONG",
                "name": "PRX250-ED-CABLE LONG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-SPRING_4-6\"",
                "name": "PRX250-SPRING 4-6\"",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-SPRING-6-8\"",
                "name": "PRX250-SPRING 6-8\"",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRC-036",
                "name": "LAUNCH TUBE TONGUE",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-206-LT-ASSY",
                "name": "PRX250-206-LT-ASSY",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "PRX250-ED-CABLE-LONG",
                "name": "LONG ENCODER CABLE  ASSEMBLY",
                "th_desc": "",
                "progress": 0
            }
        ]
    },
    {
        "id": "ACBOX-49286",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "THANANCHAI THATHUMNUK",
        "member_1": "-",
        "job_no": "49286",
        "mc_number": "ACBOX#91-95",
        "progress": 15,
        "qty": 5,
        "est_hours": 40.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 100
            },
            {
                "pn": "CABLE-1",
                "name": "Cable PN 05-21-24313-01",
                "th_desc": "สายไฟ 1",
                "progress": 0
            },
            {
                "pn": "CABLE-2",
                "name": "Cable PN 05-21-24313-02",
                "th_desc": "สายไฟ 2",
                "progress": 0
            },
            {
                "pn": "CABLE-3",
                "name": "Cable PN 05-21-24313-03",
                "th_desc": "สายไฟ 3",
                "progress": 0
            },
            {
                "pn": "CABLE-4",
                "name": "Cable PN 05-21-24314-01",
                "th_desc": "สายไฟ 4",
                "progress": 0
            },
            {
                "pn": "CABLE-5",
                "name": "Cable PN 05-21-24314-02",
                "th_desc": "สายไฟ 5",
                "progress": 0
            },
            {
                "pn": "CABLE-6",
                "name": "Cable PN 05-21-24314-03",
                "th_desc": "สายไฟ 6",
                "progress": 0
            },
            {
                "pn": "CABLE-7",
                "name": "Cable PN 05-21-20235-01",
                "th_desc": "สายไฟ 7",
                "progress": 0
            },
            {
                "pn": "CABLE-8",
                "name": "Cable PN 05-21-20235-02",
                "th_desc": "สายไฟ 8",
                "progress": 0
            },
            {
                "pn": "CABLE-9",
                "name": "Cable PN 05-21-20235-03",
                "th_desc": "สายไฟ 9",
                "progress": 0
            },
            {
                "pn": "CABLE-10",
                "name": "Cable PN 05-21-20235-04",
                "th_desc": "สายไฟ 10",
                "progress": 0
            },
            {
                "pn": "CABLE-11",
                "name": "Cable PN 05-21-20235-05",
                "th_desc": "สายไฟ 11",
                "progress": 0
            },
            {
                "pn": "CABLE-12",
                "name": "Cable PN 05-21-20235-06",
                "th_desc": "สายไฟ 12",
                "progress": 0
            },
            {
                "pn": "CABLE-13",
                "name": "Cable PN 05-21-20465",
                "th_desc": "สายไฟ 13",
                "progress": 0
            },
            {
                "pn": "CABLE-14",
                "name": "Cable PN 1313018",
                "th_desc": "สายไฟ 14",
                "progress": 0
            },
            {
                "pn": "CABLE-15",
                "name": "Cable PN 05-21-24316",
                "th_desc": "สายไฟ 15",
                "progress": 0
            },
            {
                "pn": "CABLE-16",
                "name": "Cable PN 1313019",
                "th_desc": "สายไฟ 16",
                "progress": 0
            },
            {
                "pn": "SUB-PANEL 01-24-24455",
                "name": "SUB ASSY PN 01-21-24455",
                "th_desc": "ประกอบย่อย",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY",
                "name": "MAIN PN 01-21-24466",
                "th_desc": "ประกอบหลัก",
                "progress": 15
            },
            {
                "pn": "WIRING",
                "name": "SCHEMATIC 19-21-24510",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "FUNCTION CHECK/TEST",
                "th_desc": "ทดสอบเครื่อง",
                "progress": 0
            },
            {
                "pn": "INS",
                "name": "FINAL INSPECTION",
                "th_desc": "ตรวจงาน",
                "progress": 0
            },
            {
                "pn": "PACK",
                "name": "PACKING",
                "th_desc": "แพ็คงาน",
                "progress": 0
            },
            {
                "pn": "WH",
                "name": "STORING",
                "th_desc": "จัดเก็บ",
                "progress": 0
            }
        ]
    }
];
