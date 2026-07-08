@echo off
title AVEAM - Step 1 Compare SO Tracker and Job Relationship
color 0e

echo =====================================================================
echo         [ขั้นตอนที่ 1] เปรียบเทียบข้อมูล SO Tracker กับ Job Relationship
echo =====================================================================
echo.
echo กำลังรันการเปรียบเทียบข้อมูลและจับคู่รหัส SO/PO...
echo.

python "D:\AVEAM\Antigravity2\Test-LiveArtifact\match_and_update.py"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] เกิดข้อผิดพลาดในการรันการเปรียบเทียบ!
    goto end
)

echo.
echo =====================================================================
echo    เปรียบเทียบและไฮไลต์สีเสร็จเรียบร้อย!
echo    - แถวใหม่ใน 1-DashboardPlanner-SO TRACKER.xls จะเป็นสีแดง
echo    - แถวไม่มีคู่ใน 5_PN_Job_Relationship_rev00-test.xlsx จะเป็นสีน้ำเงิน
echo.
echo    กรุณาเปิดตรวจสอบและแก้ไขไฟล์ Excel ทั้งคู่ก่อนรัน Step 2 ครับ
echo =====================================================================
echo.

:end
echo กดปุ่มใดๆ เพื่อปิดหน้าต่างนี้...
pause > nul
