# รัน script นี้หลังปิด Excel แล้ว
# ย้ายไฟล์ที่เหลือใน folder เดิมไป docs/use_cases_source/
$src = "use_caseทั้งหมด\AI Transformation Project Auto.xlsx"
$dst = "docs\use_cases_source\AI Transformation Project Auto.xlsx"
Move-Item -Path $src -Destination $dst -Force
Remove-Item "use_caseทั้งหมด" -Recurse -Force
Write-Host "เรียบร้อย! ลบ folder เดิมแล้ว"
