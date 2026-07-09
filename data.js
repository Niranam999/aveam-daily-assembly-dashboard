// Live data compiled from 5_PN_Job_Relationship_rev00-test.xlsx & 1-List of each sub-assembly.xlsx
const EXCEL_PROJECTS = [
    {
        "id": "COHU-302195778-",
        "customer": "COHU",
        "project_code": "302195778",
        "part_number": "302195778",
        "description": "ASSY, Z-LINK, 2X4, LEFT",
        "status": "ontime",
        "team_leader": "WANLOP",
        "member_1": "ASSAMA",
        "job_no": "",
        "mc_number": "-",
        "progress": 50,
        "qty": 1,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "302195778",
                "name": "ASSY, Z-LINK, 2X4, LEFT",
                "th_desc": "",
                "progress": 50,
                "qty": 1,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-Spare Part-",
        "customer": "COHU",
        "project_code": "Spare Part",
        "part_number": "8334478101",
        "description": "MAIN RING-TRAY FLIP,8334478101",
        "status": "ontime",
        "team_leader": "MACHINE SHOP",
        "member_1": "ASSAMA",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 28,
        "est_hours": 0.1,
        "sub_assemblies": [
            {
                "pn": "8334478101",
                "name": "MAIN RING-TRAY FLIP,8334478101",
                "th_desc": "",
                "progress": 0,
                "qty": 20,
                "std_time": 0.1
            },
            {
                "pn": "8335676001",
                "name": "ASSY, CONNECTOR, MANIFOLD",
                "th_desc": "",
                "progress": 0,
                "qty": 8,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 2,
        "kanban_completed": 0
    },
    {
        "id": "UIC-LH & RH FRAME-49700",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578603",
        "description": "LH & RH FRAME",
        "status": "completed",
        "team_leader": "SUWAT",
        "member_1": "-",
        "job_no": "49700",
        "mc_number": "#22-24",
        "progress": 100,
        "qty": 3,
        "est_hours": 0.1,
        "sub_assemblies": [
            {
                "pn": "53578603",
                "name": "LH Frame (Support LH)",
                "th_desc": "",
                "progress": 100,
                "qty": 3,
                "std_time": 0.1
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 1,
        "kanban_completed": 0
    },
    {
        "id": "UIC-LH & RH FRAME-49701",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578703",
        "description": "LH & RH FRAME",
        "status": "completed",
        "team_leader": "SUWAT",
        "member_1": "-",
        "job_no": "49701",
        "mc_number": "#22-24",
        "progress": 100,
        "qty": 3,
        "est_hours": 0.1,
        "sub_assemblies": [
            {
                "pn": "53578703",
                "name": "RH Frame (Support RH)",
                "th_desc": "",
                "progress": 100,
                "qty": 3,
                "std_time": 0.1
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 1,
        "kanban_completed": 0
    },
    {
        "id": "COHU-JLP CART-51026",
        "customer": "COHU",
        "project_code": "JLP CART",
        "part_number": "8337427001",
        "description": "JLP CART",
        "status": "ontime",
        "team_leader": "NATTHAVUT",
        "member_1": "KESORN",
        "job_no": "51026",
        "mc_number": "-",
        "progress": 25,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-JLP CART-51027",
        "customer": "COHU",
        "project_code": "JLP CART",
        "part_number": "8337427001",
        "description": "JLP CART",
        "status": "ontime",
        "team_leader": "NATTHAVUT",
        "member_1": "KESORN",
        "job_no": "51027",
        "mc_number": "-",
        "progress": 0,
        "qty": 40,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-100051263-",
        "customer": "COHU",
        "project_code": "100051263",
        "part_number": "100051263",
        "description": "MANIFOLD, COLD GAS, CHUCK ASSY",
        "status": "ontime",
        "team_leader": "WANLOP",
        "member_1": "ASSAMA",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 5,
        "est_hours": 2.0,
        "sub_assemblies": [
            {
                "pn": "100051263",
                "name": "MANIFOLD, COLD GAS, CHUCK ASSY",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 2.0
            },
            {
                "pn": "100051263",
                "name": "MANIFOLD, COLD GAS, CHUCK ASSY",
                "th_desc": "",
                "progress": 0,
                "qty": 4,
                "std_time": 2.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 1,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-FEEDER BASKET-49798",
        "customer": "UIC",
        "project_code": "FEEDER BASKET",
        "part_number": "52777805",
        "description": "FEEDER BASKET",
        "status": "ontime",
        "team_leader": "SUWAT",
        "member_1": "SUMOLTHA",
        "job_no": "49798",
        "mc_number": "-",
        "progress": 0,
        "qty": 25,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52777805",
                "name": "Feeder Basket",
                "th_desc": "",
                "progress": 0,
                "qty": 25,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 1,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-JLP CART-51028",
        "customer": "COHU",
        "project_code": "JLP CART",
        "part_number": "8337427001",
        "description": "JLP CART",
        "status": "ontime",
        "team_leader": "NATTHAVUT",
        "member_1": "KESORN",
        "job_no": "51028",
        "mc_number": "-",
        "progress": 0,
        "qty": 20,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 1,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "ULC-PRX-ALL",
        "customer": "ULC",
        "project_code": "PRX",
        "part_number": "PRX250-All Assy",
        "description": "PRX250 Standard Module Assembly",
        "status": "ontime",
        "team_leader": "WANLOP",
        "member_1": "KAIRUNG",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 777,
        "est_hours": 1.125,
        "sub_assemblies": [
            {
                "pn": "PRX250-REEL ASSY",
                "name": "PRX250-REEL ASSY",
                "th_desc": "",
                "progress": 0,
                "qty": 20,
                "std_time": 3.0
            },
            {
                "pn": "PRX250-CTL&MONITOR",
                "name": "PRX250-CTL&MONITOR",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-POWER DRIVE",
                "name": "PRX250-POWER DRIVE",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-ENCODER",
                "name": "PRX250-ENCODER",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 0.5
            },
            {
                "pn": "PRX250-LAUNCH TUBE",
                "name": "PRX250-LAUNCH TUBE",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-PENDANT",
                "name": "PRX250-PENDANT",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 1.0
            },
            {
                "pn": "PRXM-001-00-000",
                "name": "PRXM-001-00-000",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 2.0
            },
            {
                "pn": "PRXM-003-00-000",
                "name": "PRXM-003-00-000",
                "th_desc": "",
                "progress": 0,
                "qty": 20,
                "std_time": 2.0
            },
            {
                "pn": "PRXM-004-00-000",
                "name": "PRXM-004-00-000",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 2.0
            },
            {
                "pn": "PRXM-005-00-000",
                "name": "PRXM-005-00-000",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 2.0
            },
            {
                "pn": "PRX250-ED-CABLE-SHOR",
                "name": "SHORT  ENCODER CABLE  ASSEMBLY",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-SPRINGS-812",
                "name": "8-12\" CENTERING SPRING ASSEMBLY",
                "th_desc": "",
                "progress": 0,
                "qty": 36,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-CONTROL CABLE",
                "name": "PRX250-CONTROL CABLE",
                "th_desc": "",
                "progress": 0,
                "qty": 20,
                "std_time": 0.5
            },
            {
                "pn": "PRX250-103-CAMS",
                "name": "PRX250-103-CAMS",
                "th_desc": "",
                "progress": 0,
                "qty": 10,
                "std_time": 2.0
            },
            {
                "pn": "PRX250-ED-CABLE-LONG",
                "name": "PRX250-ED-CABLE LONG",
                "th_desc": "",
                "progress": 0,
                "qty": 12,
                "std_time": 0.25
            },
            {
                "pn": "PRX250-SPRING_4-6\"",
                "name": "PRX250-SPRING 4-6\"",
                "th_desc": "",
                "progress": 0,
                "qty": 28,
                "std_time": 0.5
            },
            {
                "pn": "PRX250-SPRING-6-8\"",
                "name": "PRX250-SPRING 6-8\"",
                "th_desc": "",
                "progress": 0,
                "qty": 36,
                "std_time": 0.5
            },
            {
                "pn": "PRC-036",
                "name": "LAUNCH TUBE TONGUE",
                "th_desc": "",
                "progress": 0,
                "qty": 500,
                "std_time": 0.0
            },
            {
                "pn": "PRX250-206-LT-ASSY",
                "name": "PRX250-206-LT-ASSY",
                "th_desc": "",
                "progress": 0,
                "qty": 3,
                "std_time": 1.0
            },
            {
                "pn": "PRX250-ED-CABLE-LONG",
                "name": "LONG ENCODER CABLE  ASSEMBLY",
                "th_desc": "",
                "progress": 0,
                "qty": 2,
                "std_time": 0.25
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 20,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-LH & RH FRAME-49702",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578603",
        "description": "LH & RH FRAME",
        "status": "ontime",
        "team_leader": "SUWAT",
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
                "progress": 25,
                "qty": 6,
                "std_time": 0.1
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 1,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-LH & RH FRAME-49703",
        "customer": "UIC",
        "project_code": "LH & RH FRAME",
        "part_number": "53578703",
        "description": "LH & RH FRAME",
        "status": "ontime",
        "team_leader": "SUWAT",
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
                "progress": 25,
                "qty": 6,
                "std_time": 0.1
            }
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 1,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-SCRAP TAPE CUTTER DRIVE KIT-49552",
        "customer": "UIC",
        "project_code": "SCRAP TAPE CUTTER DRIVE KIT",
        "part_number": "52803004",
        "description": "SCRAP TAPE CUTTER DRIVE KIT",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49552",
        "mc_number": "No.026-050",
        "progress": 0,
        "qty": 12,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52803004",
                "name": "Scrap Tape Cutter Drive Kit",
                "th_desc": "",
                "progress": 0,
                "qty": 12,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-SCRAP TAPE CUTTER DRIVE KIT-49841",
        "customer": "UIC",
        "project_code": "SCRAP TAPE CUTTER DRIVE KIT",
        "part_number": "52803004",
        "description": "SCRAP TAPE CUTTER DRIVE KIT",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49841",
        "mc_number": "No.051-075",
        "progress": 0,
        "qty": 13,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52803004",
                "name": "Scrap Tape Cutter Drive Kit",
                "th_desc": "",
                "progress": 0,
                "qty": 13,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-ACBOX-49286",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "THANANCHAI",
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
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-100042141-",
        "customer": "COHU",
        "project_code": "100042141",
        "part_number": "100042141",
        "description": "TS COLD MANIFOLD, 50MM, 4 CHKS",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 1,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "100042141",
                "name": "TS COLD MANIFOLD, 50MM, 4 CHKS",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-Z-LINK REV.D-",
        "customer": "COHU",
        "project_code": "Z-LINK REV.D",
        "part_number": "8308859901",
        "description": "Z-LINK REV.C",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 14,
        "est_hours": 8.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "A-1",
                "name": "Cuting and forming tubing",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-1",
                "name": "Pins installation",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-2",
                "name": "Pivot parts installation",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-3",
                "name": "Big ball installation",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-4",
                "name": "Small ball installation",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-5",
                "name": "Pre-assemblytubing",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-2",
                "name": "Install 2x fittings on COLD GAS CHECK MANIFOLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-3",
                "name": "Install 2x fittings on COLD ZLINK EXIT BLOCK",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-4",
                "name": "Assemble tubing",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-5",
                "name": "Final Assembly",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-6",
                "name": "Cleaning and marking",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-7",
                "name": "Prepare kit hardware",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "A-8",
                "name": "Final Inspection",
                "th_desc": "ตรวจงาน",
                "progress": 0
            },
            {
                "pn": "A-9",
                "name": "Packing",
                "th_desc": "แพ็คงาน",
                "progress": 0
            },
            {
                "pn": "S-1",
                "name": "Store",
                "th_desc": "",
                "progress": 0
            }
        ],
        "kanban_backlog": 3,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-REMOVABLE FEEDER BANK ASSY-",
        "customer": "UIC",
        "project_code": "REMOVABLE FEEDER BANK ASSY",
        "part_number": "52581707",
        "description": "REMOVABLE FEEDER BANK ASSY",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 50,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52581707",
                "name": "REMOVABLE FEEDER BANK ASSY",
                "th_desc": "",
                "progress": 0,
                "qty": 50,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-42511",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "SUWAT",
        "member_1": "-",
        "job_no": "42511",
        "mc_number": "WF#70",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-42513",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "SUWAT",
        "member_1": "-",
        "job_no": "42513",
        "mc_number": "WF#71",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 0,
        "kanban_assigned": 0,
        "kanban_in_progress": 1,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-JLP CART-51089",
        "customer": "COHU",
        "project_code": "JLP CART",
        "part_number": "8337427001",
        "description": "JLP CART",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51089",
        "mc_number": "-",
        "progress": 0,
        "qty": 60,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 2,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-FEEDER TRANSFER CART-",
        "customer": "UIC",
        "project_code": "FEEDER TRANSFER CART",
        "part_number": "49401811",
        "description": "FEEDER TRANSFER CART",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 20,
        "est_hours": 13.0,
        "sub_assemblies": [
            {
                "pn": "49401811",
                "name": "FEEDER TRANSFER CART",
                "th_desc": "",
                "progress": 0,
                "qty": 20,
                "std_time": 13.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-Spare Part-",
        "customer": "Veeco",
        "project_code": "Spare Part",
        "part_number": "01-25-04710",
        "description": "Air Gauge Amplifier",
        "status": "ontime",
        "team_leader": "SUWAT",
        "member_1": "-",
        "job_no": "",
        "mc_number": "Spare Part",
        "progress": 0,
        "qty": 3,
        "est_hours": 12.0,
        "sub_assemblies": [
            {
                "pn": "01-25-04710",
                "name": "ASSY,AIR GAUGE AMPLIFIER; 01-25-04710",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 12.0
            },
            {
                "pn": "05-25-01443",
                "name": "ASSY,CABLE,ILLUM,INTERLOCK,UPPER",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 0.0
            },
            {
                "pn": "05-25-03085",
                "name": "15616-L1,CATON CONNECTOR CORP, ASSY,CABLE; 05-25-0385",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 0.0
            }
        ],
        "kanban_backlog": 3,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-1991194901-",
        "customer": "COHU",
        "project_code": "1991194901",
        "part_number": "1991194901",
        "description": "TS COLD MANIFOLD,40MM, 4 CHKS",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 1,
        "est_hours": 2.0,
        "sub_assemblies": [
            {
                "pn": "1991194901",
                "name": "TS COLD MANIFOLD,40MM, 4 CHKS",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 2.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-PRX-",
        "customer": "COHU",
        "project_code": "PRX",
        "part_number": "1991194901",
        "description": "TS COLD MANIFOLD,40MM, 4 CHKS",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 2,
        "est_hours": 2.0,
        "sub_assemblies": [
            {
                "pn": "1991194901",
                "name": "TS COLD MANIFOLD,40MM, 4 CHKS",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 2.0
            },
            {
                "pn": "8306304101",
                "name": "EXIT BLOCK,COLD Z-LINK",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 0.2
            }
        ],
        "kanban_backlog": 2,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-302195777-",
        "customer": "COHU",
        "project_code": "302195777",
        "part_number": "302195777",
        "description": "ASSY,Z-LINK,2X4,RIGHT",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 1,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "302195777",
                "name": "ASSY,Z-LINK,2X4,RIGHT",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-SCRAP TAPE CUTTER DRIVE KIT-49940",
        "customer": "UIC",
        "project_code": "SCRAP TAPE CUTTER DRIVE KIT",
        "part_number": "52803004",
        "description": "SCRAP TAPE CUTTER DRIVE KIT",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49940",
        "mc_number": "No.076-0100",
        "progress": 0,
        "qty": 25,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52803004",
                "name": "Scrap Tape Cutter Drive Kit",
                "th_desc": "",
                "progress": 0,
                "qty": 13,
                "std_time": 4.0
            },
            {
                "pn": "52803004",
                "name": "Scrap Tape Cutter Drive Kit",
                "th_desc": "",
                "progress": 0,
                "qty": 12,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 2,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-50489",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50489",
        "mc_number": "WF#72",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-50490",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50490",
        "mc_number": "WF#73",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-FEEDER BASKET-50488",
        "customer": "UIC",
        "project_code": "FEEDER BASKET",
        "part_number": "52777805",
        "description": "FEEDER BASKET",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50488",
        "mc_number": "-",
        "progress": 0,
        "qty": 25,
        "est_hours": 4.0,
        "sub_assemblies": [
            {
                "pn": "52777805",
                "name": "Feeder Basket",
                "th_desc": "",
                "progress": 0,
                "qty": 25,
                "std_time": 4.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "UIC-FIXED FEEDER BANK-",
        "customer": "UIC",
        "project_code": "FIXED FEEDER BANK",
        "part_number": "52648303",
        "description": "FIXED FEEDER BANK ASSY",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 50,
        "est_hours": 8.0,
        "sub_assemblies": [
            {
                "pn": "52648303",
                "name": "FIXED FEEDER BANK",
                "th_desc": "",
                "progress": 0,
                "qty": 50,
                "std_time": 8.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-ACBOX-49922",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49922",
        "mc_number": "ACBOX#96-100",
        "progress": 0,
        "qty": 5,
        "est_hours": 40.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
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
                "progress": 0
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-SP-Lamp Cooling-49712",
        "customer": "Veeco",
        "project_code": "SP-Lamp Cooling",
        "part_number": "01-15-09406",
        "description": "Lamp Cooling SW",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49712",
        "mc_number": "Spare Part",
        "progress": 0,
        "qty": 2,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "01-15-09406",
                "name": "Lamp Cooling SW",
                "th_desc": "",
                "progress": 0,
                "qty": 2,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-SP-Lamp Cooling-49713",
        "customer": "Veeco",
        "project_code": "SP-Lamp Cooling",
        "part_number": "01-15-09406",
        "description": "Lamp Cooling SW",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49713",
        "mc_number": "Spare Part",
        "progress": 0,
        "qty": 1,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "01-15-09406",
                "name": "Lamp Cooling SW",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-SP-Lamp Cooling-49937",
        "customer": "Veeco",
        "project_code": "SP-Lamp Cooling",
        "part_number": "01-15-09406",
        "description": "Lamp Cooling SW",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49937",
        "mc_number": "Spare Part",
        "progress": 0,
        "qty": 1,
        "est_hours": 1.0,
        "sub_assemblies": [
            {
                "pn": "01-15-09406",
                "name": "Lamp Cooling SW",
                "th_desc": "",
                "progress": 0,
                "qty": 1,
                "std_time": 1.0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-50491",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50491",
        "mc_number": "WF#74",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-50492",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50492",
        "mc_number": "WF#75",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-ACBOX-49287",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49287",
        "mc_number": "ACBOX#101-105",
        "progress": 0,
        "qty": 5,
        "est_hours": 40.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
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
                "progress": 0
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-50493",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "50493",
        "mc_number": "WF#76",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-51145",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51145",
        "mc_number": "WF#77",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-ACBOX-49923",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "49923",
        "mc_number": "ACBOX#106-110",
        "progress": 0,
        "qty": 5,
        "est_hours": 40.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
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
                "progress": 0
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-51146",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51146",
        "mc_number": "WF#78",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-51147",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51147",
        "mc_number": "WF#79",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-ACBOX-",
        "customer": "Veeco",
        "project_code": "ACBOX",
        "part_number": "01-21-24466",
        "description": "AC Power Box",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "ACBOX#111-115",
        "progress": 0,
        "qty": 10,
        "est_hours": 40.0,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
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
                "progress": 0
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
        ],
        "kanban_backlog": 2,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-JTP CART-",
        "customer": "COHU",
        "project_code": "JTP CART",
        "part_number": "8337428001",
        "description": "JTP CART",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 20,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "COHU-CURE CART 9002-",
        "customer": "COHU",
        "project_code": "CURE CART 9002",
        "part_number": "8337429002",
        "description": "CURE CART, EPX",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "",
        "mc_number": "-",
        "progress": 0,
        "qty": 14,
        "est_hours": 4.0,
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
                "name": "Brake Floor Lock Assy (ประกอบชุดเบรก)",
                "th_desc": "ประกอบเบรก",
                "progress": 0
            },
            {
                "pn": "7",
                "name": "Top Plate Assy (ประกอบชุดท็อปเพลท)",
                "th_desc": "ประกอบท็อป",
                "progress": 0
            },
            {
                "pn": "8",
                "name": "Handle Assy (ประกอบชุดมือจับ)",
                "th_desc": "ประกอบมือจับ",
                "progress": 0
            },
            {
                "pn": "9",
                "name": "Rear Plate Assy (ประกอบชุดฝาหลัง)",
                "th_desc": "ประกอบฝาหลัง",
                "progress": 0
            },
            {
                "pn": "10",
                "name": "Ground Cable Assy (ประกอบสายกราวน์)",
                "th_desc": "ประกอบสายดิน",
                "progress": 0
            },
            {
                "pn": "11",
                "name": "Land yard cable Assy (ประกอบสายแลนยาร์ด)",
                "th_desc": "ประกอบแลนยาร์ด",
                "progress": 0
            },
            {
                "pn": "12",
                "name": "Main Assy (ประกอบชุดเมนแอสเซมบลี้)",
                "th_desc": "ประกอบเมนแอสเซมบลี้",
                "progress": 0
            },
            {
                "pn": "13",
                "name": "Final Inspection (การตรวจสอบขั้นสุดท้าย)",
                "th_desc": "การตรวจสอบขั้นสุดท้าย",
                "progress": 0
            },
            {
                "pn": "14",
                "name": "Packing (การแพ็คงาน)",
                "th_desc": "แพ็คงาน",
                "progress": 0
            }
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-51148",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51148",
        "mc_number": "WF#80",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    },
    {
        "id": "Veeco-WF-51149",
        "customer": "Veeco",
        "project_code": "WF",
        "part_number": "01-25-07546",
        "description": "WF",
        "status": "ontime",
        "team_leader": "WANLOP CHANPHET",
        "member_1": "-",
        "job_no": "51149",
        "mc_number": "WF#81",
        "progress": 0,
        "qty": 1,
        "est_hours": 166.5,
        "sub_assemblies": [
            {
                "pn": "R-1",
                "name": "Receiving material",
                "th_desc": "รับพาร์ท",
                "progress": 0
            },
            {
                "pn": "SUB-F1",
                "name": "01-25-04710; ASSY,AIR GAUGE AMPLIFIER; 01-2",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F2",
                "name": "11-08-09367; VALVE,SOLENOID, 2 WAY, 24VDC,M",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F3",
                "name": "01-06-00506; ASSY,TRANSDUCER, 1500 AIR GAUG",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F4",
                "name": "01-25-00563; ASSY,AIR RESERVOIR,AIR GAUGE;",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F5",
                "name": "01-18-03446; ASSY,PRESSURE DISTRIBUTION MAN",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F6",
                "name": "01-25-04513-02; 01-08-04106-30,AIR GAP, ASSY,A",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F7",
                "name": "01-08-04106-30; ASSY,REFERENCE AIR GAP,STANDAR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-F8",
                "name": "01-25-01399; ASSY,RETICLE REMOTE SWITCH; 01",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#1",
                "name": "01-25-00431; FRAME ASSEMLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "SUB-U1",
                "name": "04-25-06718; ASSY,ILLUM BASE,OPT CLPR WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U2",
                "name": "04-25-06748; ASSY,ALIGN CELL ILLUM,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U3",
                "name": "04-25-04016; ASSY,LAMPCHAMBER, 1.2 KW,LT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U4",
                "name": "04-25-04155-02; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U5",
                "name": "04-25-04015; ASSY,LAMPCHAMBER, 1.2 KW,RT,QU",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U6",
                "name": "04-25-04155-01; ASSY,LAMPCHAMBER, 1.2 KW 3 WIR",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "SUB-U7",
                "name": "04-25-06639; ASSY,EXPOSURE DETECTOR,WFLD",
                "th_desc": "",
                "progress": 0
            },
            {
                "pn": "MAIN-ASSY#2",
                "name": "04-25-06773; UPPER ASSEMBLY & WIRING",
                "th_desc": "เดินสายไฟ",
                "progress": 0
            },
            {
                "pn": "TESTING",
                "name": "ALIGNMENT & FUNCTION TEST",
                "th_desc": "",
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
        ],
        "kanban_backlog": 1,
        "kanban_assigned": 0,
        "kanban_in_progress": 0,
        "kanban_qa": 0,
        "kanban_completed": 0
    }
];
