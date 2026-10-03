$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root 'Ke_hoach_day_va_demo_EcoMicro_Class.docx'
$work = Join-Path $env:TEMP ('ecomicro_docx_' + [guid]::NewGuid().ToString('N'))

New-Item -ItemType Directory -Path $work | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work '_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work 'word') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work 'word\_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $work 'docProps') | Out-Null

function XmlEscape([string]$text) {
    return [System.Security.SecurityElement]::Escape($text)
}

$body = New-Object System.Text.StringBuilder

function Add-Paragraph {
    param(
        [string]$Text = '',
        [string]$Style = 'Normal',
        [string]$Align = '',
        [switch]$Bold,
        [switch]$Italic,
        [int]$Size = 0,
        [string]$Color = '',
        [int]$Before = 0,
        [int]$After = 120,
        [switch]$KeepNext
    )

    $pPr = "<w:pStyle w:val=`"$Style`"/>"
    if ($Align) { $pPr += "<w:jc w:val=`"$Align`"/>" }
    if ($Before -gt 0 -or $After -ge 0) { $pPr += "<w:spacing w:before=`"$Before`" w:after=`"$After`"/>" }
    if ($KeepNext) { $pPr += '<w:keepNext/>' }

    $rPr = ''
    if ($Bold) { $rPr += '<w:b/>' }
    if ($Italic) { $rPr += '<w:i/>' }
    if ($Size -gt 0) { $rPr += "<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/>" }
    if ($Color) { $rPr += "<w:color w:val=`"$Color`"/>" }

    $escaped = XmlEscape $Text
    [void]$body.Append("<w:p><w:pPr>$pPr</w:pPr><w:r><w:rPr>$rPr</w:rPr><w:t xml:space=`"preserve`">$escaped</w:t></w:r></w:p>")
}

function Add-Bullet([string]$Text, [int]$Level = 0) {
    $indent = 360 + (360 * $Level)
    $escaped = XmlEscape $Text
    [void]$body.Append("<w:p><w:pPr><w:pStyle w:val=`"Normal`"/><w:ind w:left=`"$indent`" w:hanging=`"240`"/><w:spacing w:after=`"80`"/></w:pPr><w:r><w:t>•  $escaped</w:t></w:r></w:p>")
}

function Add-Numbered([string]$Text, [int]$Number) {
    $escaped = XmlEscape $Text
    [void]$body.Append("<w:p><w:pPr><w:pStyle w:val=`"Normal`"/><w:ind w:left=`"360`" w:hanging=`"300`"/><w:spacing w:after=`"80`"/></w:pPr><w:r><w:t>$Number.  $escaped</w:t></w:r></w:p>")
}

function Add-PageBreak {
    [void]$body.Append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
}

function Add-Table {
    param([object[]]$Rows, [int[]]$Widths)

    [void]$body.Append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="A6A6A6"/><w:left w:val="single" w:sz="6" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="6" w:color="A6A6A6"/><w:right w:val="single" w:sz="6" w:color="A6A6A6"/><w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>')
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        [void]$body.Append('<w:tr>')
        for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
            $shade = if ($r -eq 0) { '<w:shd w:fill="1F4E78"/>' } else { if ($r % 2 -eq 0) { '<w:shd w:fill="EAF2F8"/>' } else { '' } }
            $bold = if ($r -eq 0) { '<w:b/><w:color w:val="FFFFFF"/>' } else { '' }
            $width = if ($Widths -and $c -lt $Widths.Count) { $Widths[$c] } else { 2400 }
            $cellText = XmlEscape ([string]$Rows[$r][$c])
            [void]$body.Append("<w:tc><w:tcPr><w:tcW w:w=`"$width`" w:type=`"dxa`"/>$shade</w:tcPr><w:p><w:pPr><w:spacing w:after=`"0`"/></w:pPr><w:r><w:rPr>$bold</w:rPr><w:t xml:space=`"preserve`">$cellText</w:t></w:r></w:p></w:tc>")
        }
        [void]$body.Append('</w:tr>')
    }
    [void]$body.Append('</w:tbl><w:p><w:pPr><w:spacing w:after="100"/></w:pPr></w:p>')
}

# Trang bìa
Add-Paragraph 'HỒ SƠ HUẤN LUYỆN & DEMO STEM' 'Title' 'center' -Bold -Size 40 -Color '1F4E78' -Before 900 -After 240
Add-Paragraph 'EcoMicro-Class' 'Subtitle' 'center' -Bold -Size 34 -Color '70AD47' -After 140
Add-Paragraph 'Hệ sinh thái quản lý vi khí hậu và năng lượng học đường tự động tích hợp nguồn năng lượng mặt trời' 'Subtitle' 'center' -Italic -Size 24 -Color '404040' -After 480
Add-Paragraph 'Đối tượng: Nhóm 3 học sinh lớp 6' 'Normal' 'center' -Bold -Size 24 -After 120
Add-Paragraph 'Mục tiêu: Tạo hứng thú – hiểu nguyên lý – chế tạo – đo đạc – phản biện' 'Normal' 'center' -Size 22 -After 700
Add-Paragraph 'Tài liệu triển khai giáo dục STEM và trình diễn nguyên mẫu' 'Normal' 'center' -Italic -Size 20 -Color '666666' -After 120
Add-Paragraph 'Năm học 2026–2027' 'Normal' 'center' -Size 20 -Color '666666'
Add-PageBreak

Add-Paragraph 'MỤC LỤC NỘI DUNG' 'Heading1'
Add-Numbered 'Mục tiêu buổi học đầu tiên' 1
Add-Numbered 'Demo mở đầu: “Lớp học biết suy nghĩ”' 2
Add-Numbered 'Câu hỏi dẫn dắt và khảo sát sở thích học sinh' 3
Add-Numbered 'Kiến thức và kỹ năng học sinh cần học' 4
Add-Numbered 'Lộ trình huấn luyện' 5
Add-Numbered 'Danh sách linh kiện và sơ đồ đấu nối' 6
Add-Numbered 'Kịch bản trình diễn từng phút' 7
Add-Numbered 'Các điểm khoa học cần chuẩn hóa' 8
Add-Numbered 'Hướng dẫn sử dụng code demo' 9
Add-Numbered 'Sản phẩm cần hoàn thiện cho cuộc thi' 10

Add-Paragraph '1. MỤC TIÊU BUỔI HỌC ĐẦU TIÊN' 'Heading1'
Add-Paragraph 'Ba học sinh được xem là chưa biết STEM, điện tử, Arduino hay lập trình. Buổi đầu phải xây nền tảng bằng ngôn ngữ đời thường, giúp các em hiểu từng từ mới rồi mới giới thiệu bài học, sản phẩm và mô hình demo.'
Add-Bullet 'Học sinh hiểu STEM là cách dùng kiến thức để giải quyết một vấn đề trong đời sống.'
Add-Bullet 'Học sinh hiểu cuộc thi gồm những việc gì và vì sao cần làm việc theo đội.'
Add-Bullet 'Học sinh nhận ra một vấn đề thật trong lớp học: nóng, tối, chói, ẩm hoặc lãng phí điện.'
Add-Bullet 'Học sinh hiểu bằng phép ví von: cảm biến là giác quan, bộ điều khiển là bộ não, thiết bị là tay chân.'
Add-Bullet 'Học sinh biết cách Đoán → Xem → Nói lý do → Nghĩ thêm khi thử một tình huống.'
Add-Bullet 'Giáo viên khảo sát sở thích trước khi phân vai cho ba thành viên.'

Add-Paragraph '1.1. Mạch tổ chức buổi đầu' 'Heading2'
Add-Table @(
    @('Thứ tự', 'Nội dung', 'Mục đích'),
    @('1', 'Nói chuyện về đồ vật tự động quen thuộc', 'Bắt đầu từ trải nghiệm của trẻ'),
    @('2', 'Giải thích STEM bằng ví dụ đơn giản', 'Hiểu cách học sẽ sử dụng'),
    @('3', 'Giải thích cuộc thi và ba vòng bằng câu ngắn', 'Biết hành trình sắp tham gia'),
    @('4', 'Dạy trước các từ: vấn đề, giải pháp, mô hình, tự động, cảm biến, dữ liệu', 'Có đủ từ để hiểu bài'),
    @('5', 'Giới thiệu bài học và tìm vấn đề trong lớp', 'Nối kiến thức với trải nghiệm thật'),
    @('6', 'Giới thiệu EcoMicro-Class bằng phép ví von cơ thể người', 'Hiểu sản phẩm ở mức cơ bản'),
    @('7', 'Mở mô hình 3D ở trạng thái bình thường', 'Nhận diện từng bộ phận'),
    @('8', 'Kiểm thử từng cảnh theo Đoán–Xem–Nói lý do–Nghĩ thêm', 'Tập suy nghĩ khoa học'),
    @('9', 'Khảo sát sở thích và giao nhiệm vụ thử', 'Chuẩn bị phân vai linh hoạt')
) @(1000, 4700, 3500)
Add-Paragraph 'Không dùng ngay các từ LDR, DHT22, PIR, PWM, thuật toán hoặc phản hồi chéo. Lần đầu chỉ nói chức năng bằng từ quen thuộc. Tên kỹ thuật được bổ sung dần sau khi học sinh đã hiểu ý nghĩa.' 'Callout' -Bold

Add-Paragraph '2. DEMO MỞ ĐẦU: “LỚP HỌC BIẾT SUY NGHĨ”' 'Heading1'
Add-Paragraph '2.1. Cấu hình mô hình' 'Heading2'
Add-Bullet 'Hai cảm biến ánh sáng: một cảm biến ngoài trời và một cảm biến tại bàn học.'
Add-Bullet 'Một servo kéo rèm; một đèn LED điều chỉnh bằng PWM.'
Add-Bullet 'Một cảm biến DHT22 đo nhiệt độ, độ ẩm; một quạt mini.'
Add-Bullet 'Một cảm biến PIR phát hiện chuyển động; một màn hình OLED hiển thị quyết định.'
Add-Bullet 'Arduino Uno/Nano điều khiển; nguồn USB 5 V trong phiên bản đầu.'

Add-Paragraph '2.2. Bốn tình huống trình diễn' 'Heading2'
Add-Table @(
    @('Tình huống', 'Tác động minh họa', 'Phản ứng của hệ thống', 'Thông điệp khoa học'),
    @('1. Bình thường', 'Ánh sáng và nhiệt độ phù hợp', 'Rèm mở, LED tắt/yếu, quạt tắt', 'Chỉ dùng năng lượng khi cần'),
    @('2. Nắng gắt', 'Chiếu đèn pin vào LDR ngoài trời', 'Rèm đóng; nếu bàn thiếu sáng, LED tăng sáng', 'Các phân hệ phản hồi chéo'),
    @('3. Trời tối', 'Che cảm biến tại bàn', 'LED tự tăng sáng theo mức thiếu hụt', 'Cảm biến phản ứng với thời tiết thật'),
    @('4. Nóng/không người', 'Tăng nhiệt độ mô phỏng hoặc chạy DEMO_MODE', 'Quạt bật khi nóng; LED tắt khi không người', 'Tiện nghi đi cùng tiết kiệm điện')
) @(1600, 2200, 2800, 2600)

Add-Paragraph 'Lưu ý: Không dùng lửa hoặc vật nung nóng khi dạy trẻ. Chế độ DEMO_MODE trong code có thể tạo dữ liệu mô phỏng để quạt và các cơ cấu hoạt động ngay.' 'Callout' -Bold

Add-Paragraph '3. CÂU HỎI DẪN DẮT VÀ KHẢO SÁT SỞ THÍCH' 'Heading1'
Add-Paragraph '3.1. Câu hỏi dẫn dắt' 'Heading2'
$questions = @(
    'Theo các em, đây là mô hình của nơi nào?',
    'Trong lớp học, điều gì khiến các em khó chịu nhất: nóng, tối, chói hay ẩm?',
    'Nếu trời đang nắng mà đột nhiên mưa, thời khóa biểu có biết để bật đèn không?',
    'Làm thế nào để chiếc rèm biết lúc nào cần đóng?',
    'Bộ phận nào giống mắt, bộ phận nào giống bộ não, bộ phận nào giống tay chân?',
    'Nếu rèm đóng khiến phòng tối đi, hệ thống cần làm gì tiếp theo?',
    'Làm sao tránh bật đèn khi trong phòng không có ai?',
    'Nếu chỉ được cải tiến một chức năng, em muốn làm chức năng nào?'
)
for ($i = 0; $i -lt $questions.Count; $i++) { Add-Numbered $questions[$i] ($i + 1) }

Add-Paragraph '3.2. Cách giải thích gần gũi' 'Heading2'
Add-Table @(
    @('Thành phần', 'Cách gọi gần gũi'),
    @('Cảm biến', 'Giác quan – giúp mô hình nhìn và cảm nhận'),
    @('Arduino/ESP32', 'Bộ não – nhận dữ liệu và quyết định'),
    @('Chương trình', 'Các quy tắc suy nghĩ'),
    @('Servo, LED, quạt', 'Tay chân – thực hiện quyết định'),
    @('Pin mặt trời và pin lưu trữ', 'Nguồn cung cấp năng lượng')
) @(3000, 6200)

Add-Paragraph '3.3. Khảo sát sở thích' 'Heading2'
Add-Paragraph 'Cho mỗi em chọn hai nội dung yêu thích: lắp ráp – đấu dây; lập trình; thiết kế mô hình; đo đạc – thí nghiệm; làm poster – biểu đồ; thuyết trình – phản biện. Chỉ phân vai chính sau 2–3 buổi.'
Add-Table @(
    @('Vai trò chính', 'Nhiệm vụ'),
    @('Kỹ sư phần cứng', 'Lắp cảm biến, LED, servo, quạt; kiểm tra dây và nguồn'),
    @('Kỹ sư phần mềm – dữ liệu', 'Viết code, đặt ngưỡng, lọc dữ liệu và vẽ biểu đồ'),
    @('Nhà khoa học – thuyết trình', 'Thiết kế thí nghiệm, ghi kết quả, giải thích và phản biện')
) @(3000, 6200)
Add-Paragraph 'Cả ba học sinh vẫn phải hiểu toàn bộ sản phẩm để có thể trả lời khi Ban Giám khảo hỏi chéo.' 'Callout' -Bold

Add-Paragraph '4. KIẾN THỨC VÀ KỸ NĂNG HỌC SINH CẦN HỌC' 'Heading1'
Add-Paragraph '4.1. Mức bắt buộc' 'Heading2'
$basic = @(
    'Dòng điện, điện áp, cực âm/dương và an toàn với điện áp thấp.',
    'Mạch nối tiếp, song song và cầu chia áp ở mức cơ bản.',
    'Tín hiệu analog khác tín hiệu digital như thế nào.',
    'Biến, hàm, if/else và cách chương trình chạy trong loop().',
    'PWM điều chỉnh công suất trung bình và độ sáng LED.',
    'Servo hoạt động theo góc; cảm biến DHT22, PIR và LDR hoạt động ra sao.',
    'Nhiệt độ khác độ ẩm; pin mặt trời chuyển quang năng thành điện năng.',
    'Lập bảng đo, tính trung bình và so sánh trước–sau.',
    'Quy trình kỹ thuật: vấn đề → ý tưởng → chế tạo → thử nghiệm → cải tiến.',
    'Trả lời phản biện bằng số liệu và bằng chứng.'
)
foreach ($item in $basic) { Add-Bullet $item }

Add-Paragraph '4.2. Mức nâng cao sau khi mô hình chạy ổn định' 'Heading2'
Add-Bullet 'Trung bình động để giảm nhiễu cảm biến.'
Add-Bullet 'Dùng hai ngưỡng bật/tắt (hysteresis) để tránh thiết bị nhấp nháy.'
Add-Bullet 'Máy trạng thái: bình thường, nắng gắt, thiếu sáng, quá nóng.'
Add-Bullet 'Hiệu chuẩn cảm biến; đo điện áp, dòng điện và công suất.'
Add-Bullet 'So sánh năng lượng tạo ra với năng lượng tiêu thụ.'
Add-Bullet 'Ghi dữ liệu theo thời gian và vẽ biểu đồ.'

Add-Paragraph '5. LỘ TRÌNH HUẤN LUYỆN ĐỀ XUẤT' 'Heading1'
Add-Table @(
    @('Giai đoạn', 'Thời lượng', 'Nội dung chính', 'Sản phẩm đầu ra'),
    @('1. Khởi động đội thi', '1 buổi', 'Cuộc thi, bối cảnh, thử thách, giới thiệu giải pháp, demo 3D và khảo sát sở thích', 'Phiếu ý tưởng và dự đoán nhóm'),
    @('2. Học từng khối', '3–4 buổi', 'LDR–LED; servo–rèm; DHT22–quạt; PIR–OLED', 'Bốn mạch nhỏ chạy độc lập'),
    @('3. Tích hợp', '2–3 buổi', 'Ghép phân hệ, ưu tiên logic, nút điều khiển tay', 'Nguyên mẫu hoàn chỉnh'),
    @('4. Thí nghiệm', '2 buổi', 'Đo ánh sáng, nhiệt độ, thời gian và năng lượng', 'Bảng số liệu, biểu đồ'),
    @('5. Hồ sơ–phản biện', '2–3 buổi', 'Poster, bài nói, hỏi chéo, xử lý lỗi', 'Bài trình bày 3–5 phút'),
    @('6. Luyện chung kết', 'Nhiều lượt', 'Chế tạo thử trong giới hạn 90 phút', 'Quy trình làm việc ổn định')
) @(1300, 1200, 4000, 2700)

Add-Paragraph 'Cấu trúc mỗi buổi: demo nhỏ → học nguyên lý → tự lắp → đo kết quả → giải thích bằng lời của học sinh.' 'Callout' -Bold

Add-Paragraph '6. LINH KIỆN VÀ SƠ ĐỒ ĐẤU NỐI' 'Heading1'
Add-Paragraph '6.1. Danh sách linh kiện' 'Heading2'
Add-Bullet '01 Arduino Uno/Nano; 02 LDR; 02 điện trở 10 kΩ.'
Add-Bullet '01 DHT22; 01 PIR HC-SR501; 01 servo SG90.'
Add-Bullet 'LED trắng/dải LED điện áp thấp; 01 quạt mini 5 V.'
Add-Bullet 'MOSFET/module điều khiển cho LED và quạt; diode bảo vệ nếu cần.'
Add-Bullet '01 OLED SSD1306 I²C 0,96 inch; breadboard và dây cắm.'
Add-Bullet 'Nguồn USB 5 V đủ dòng; đồng hồ vạn năng.'
Add-Bullet 'Bìa foam, que gỗ, vải làm rèm; đèn pin để mô phỏng Mặt Trời.'

Add-Paragraph '6.2. Bảng nối chân cho Arduino Uno/Nano' 'Heading2'
Add-Table @(
    @('Thiết bị', 'Chân thiết bị', 'Chân Arduino'),
    @('LDR ngoài trời', 'Điểm giữa cầu chia áp', 'A0'),
    @('LDR tại bàn học', 'Điểm giữa cầu chia áp', 'A1'),
    @('DHT22', 'DATA', 'D2'),
    @('PIR HC-SR501', 'OUT', 'D3'),
    @('LED qua MOSFET', 'IN/GATE', 'D5 – PWM'),
    @('Quạt qua MOSFET', 'IN/GATE', 'D6'),
    @('Servo SG90', 'Signal', 'D9'),
    @('OLED SSD1306', 'SDA / SCL', 'A4 / A5')
) @(3500, 3200, 2500)
Add-Paragraph 'Bắt buộc nối chung GND. Không cấp trực tiếp dòng cho quạt hoặc LED công suất từ chân Arduino. Servo và tải nên dùng nguồn 5 V ngoài đủ dòng.' 'Callout' -Bold

Add-Paragraph '7. KỊCH BẢN BUỔI ĐẦU TỪNG PHÚT' 'Heading1'
Add-Table @(
    @('Thời gian', 'Việc giáo viên thực hiện', 'Câu hỏi cho học sinh'),
    @('0–5 phút', 'Làm quen; hỏi về cửa, đèn hoặc vòi nước tự động', 'Các em từng thấy đồ vật nào tự hoạt động?'),
    @('5–12 phút', 'Giải thích STEM bằng ví dụ lớp học nóng', 'Chúng ta cần tìm hiểu, làm thử và đo những gì?'),
    @('12–18 phút', 'Giải thích cuộc thi và ba vòng bằng câu ngắn', 'Ở mỗi vòng, các em sẽ làm việc gì?'),
    @('18–25 phút', 'Giải thích các từ mới bằng ví dụ quen thuộc', 'Cảm biến giúp máy làm gì? Dữ liệu là gì?'),
    @('25–30 phút', 'Giới thiệu tên bài và ba mục tiêu rất đơn giản', 'Cuối buổi các em muốn biết làm điều gì?'),
    @('30–35 phút', 'Hỏi vấn đề thật trong lớp học', 'Khi nào các em thấy nóng, tối, chói hoặc ẩm?'),
    @('35–40 phút', 'Giới thiệu EcoMicro-Class bằng ví von cơ thể người', 'Đâu là giác quan, bộ não, tay chân và thức ăn?'),
    @('40–45 phút', 'Mở mô hình 3D ở cảnh bình thường; nhận diện từng phần', 'Em nhìn thấy những bộ phận nào?'),
    @('45–52 phút', 'Hướng dẫn một lượt Đoán–Xem–Nói lý do–Nghĩ thêm', 'Theo em điều gì sắp xảy ra?'),
    @('52–65 phút', 'Thử Nắng gắt và Mưa tối, mỗi lần chỉ đổi một tình huống', 'Bộ phận nào thay đổi và vì sao?'),
    @('65–75 phút', 'Nhắc lại từ mới, hỏi sở thích và giao nhiệm vụ quan sát', 'Em nhớ được điều gì và muốn học phần nào?')
) @(1600, 4500, 3100)

Add-Paragraph '8. CÁC ĐIỂM KHOA HỌC CẦN CHUẨN HÓA' 'Heading1'
Add-Paragraph '8.1. Bộ sạc pin mặt trời' 'Heading2'
Add-Paragraph 'Không nên mô tả TP4056 như một bộ sạc pin mặt trời hoàn chỉnh. Điện áp tấm pin thay đổi theo cường độ nắng; phiên bản thực tế cần chọn tấm pin, bộ sạc năng lượng mặt trời, mạch bảo vệ pin và bộ ổn áp phù hợp.'
Add-Paragraph '8.2. Phun sương và nồm ẩm' 'Heading2'
Add-Paragraph 'Phun sương làm tăng độ ẩm và chỉ hỗ trợ làm mát bay hơi khi điều kiện phù hợp. Khi độ ẩm cao, hệ thống phải khóa phun sương; giải pháp phù hợp hơn là thông gió hoặc hút ẩm. Bản demo đầu nên dùng quạt, chưa dùng phun sương.'
Add-Paragraph '8.3. LDR và đơn vị lux' 'Heading2'
Add-Paragraph 'LDR chỉ cho mức sáng tương đối nếu chưa hiệu chuẩn. Nếu hồ sơ tuyên bố đo lux, nhóm cần hiệu chuẩn với thiết bị tham chiếu hoặc dùng cảm biến ánh sáng số như BH1750.'

Add-Paragraph '9. HƯỚNG DẪN SỬ DỤNG CODE DEMO' 'Heading1'
Add-Paragraph 'File chương trình: EcoMicroClass_Demo/EcoMicroClass_Demo.ino.'
Add-Paragraph '9.1. Thư viện cần cài' 'Heading2'
Add-Bullet 'DHT sensor library by Adafruit và Adafruit Unified Sensor.'
Add-Bullet 'Adafruit GFX Library và Adafruit SSD1306.'
Add-Bullet 'Servo và Wire thường có sẵn trong Arduino IDE.'
Add-Paragraph '9.2. Hai chế độ vận hành' 'Heading2'
Add-Bullet 'DEMO_MODE = true: tự động chạy bốn tình huống, mỗi tình huống 7 giây. Dùng trong buổi giới thiệu.'
Add-Bullet 'DEMO_MODE = false: đọc cảm biến thật. Dùng khi học sinh bắt đầu đo đạc và hiệu chuẩn.'
Add-Paragraph '9.3. Điểm nổi bật của chương trình' 'Heading2'
Add-Bullet 'Không dùng delay() trong loop(), nên các khối được cập nhật liên tục.'
Add-Bullet 'LDR dùng trung bình động 10 mẫu để giảm nhiễu.'
Add-Bullet 'Rèm và quạt dùng hai ngưỡng ON/OFF để tránh rung hoặc bật tắt liên tục.'
Add-Bullet 'Độ sáng LED được tính theo mức thiếu sáng tại bàn và giới hạn trong PWM 0–255.'
Add-Bullet 'OLED và Serial Monitor hiển thị cả dữ liệu lẫn quyết định của hệ thống.'

Add-Paragraph '10. SẢN PHẨM CẦN HOÀN THIỆN CHO CUỘC THI' 'Heading1'
Add-Bullet 'Nguyên mẫu hoạt động ổn định và có chế độ điều khiển thủ công dự phòng.'
Add-Bullet 'Sơ đồ khối, sơ đồ mạch, bảng linh kiện và dự toán chi phí.'
Add-Bullet 'Bảng số liệu, biểu đồ và kết luận từ thí nghiệm.'
Add-Bullet 'Mô tả thuật toán bằng lưu đồ và code có chú thích.'
Add-Bullet 'Poster; bài nói 3–5 phút chia lời hợp lý cho ba học sinh.'
Add-Bullet 'Bộ câu hỏi phản biện; kịch bản xử lý khi cảm biến, servo hoặc nguồn gặp lỗi.'
Add-Bullet 'Bài luyện chế tạo 90 phút theo các bộ vật liệu khác nhau.'

Add-Paragraph 'KẾT LUẬN' 'Heading1'
Add-Paragraph 'Trọng tâm của EcoMicro-Class không phải là gắn thật nhiều linh kiện, mà là chứng minh được một chuỗi quyết định hợp lý bằng dữ liệu: môi trường thay đổi → cảm biến ghi nhận → thuật toán đánh giá → thiết bị phản ứng → nhóm đo được hiệu quả. Khi ba học sinh hiểu và tự giải thích được chuỗi này, sản phẩm vừa có tính STEM, vừa có sức thuyết phục trước Ban Giám khảo.'

$sectPr = @'
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1276" w:header="567" w:footer="567" w:gutter="0"/>
  <w:cols w:space="708"/>
  <w:docGrid w:linePitch="360"/>
</w:sectPr>
'@

$documentXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>$($body.ToString())$sectPr</w:body>
</w:document>
"@

$stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Arial"/><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto" w:after="120"/><w:jc w:val="both"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Aptos Display" w:hAnsi="Aptos Display"/><w:b/><w:color w:val="1F4E78"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="548235"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Callout">
    <w:name w:val="Callout"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="100" w:after="140"/><w:ind w:left="300" w:right="300"/><w:shd w:fill="FFF2CC"/><w:pBdr><w:left w:val="single" w:sz="18" w:color="BF9000" w:space="8"/></w:pBdr></w:pPr>
    <w:rPr><w:color w:val="7F6000"/></w:rPr>
  </w:style>
</w:styles>
'@

$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
'@

$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
'@

$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

$now = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$coreXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Kế hoạch dạy và demo EcoMicro-Class</dc:title>
  <dc:subject>Giáo dục STEM cho học sinh lớp 6</dc:subject>
  <dc:creator>EcoMicro-Class Team</dc:creator>
  <cp:keywords>STEM; Arduino; vi khí hậu; năng lượng mặt trời; lớp 6</cp:keywords>
  <dc:description>Kế hoạch huấn luyện, kịch bản demo, linh kiện và hướng dẫn code.</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">$now</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">$now</dcterms:modified>
</cp:coreProperties>
"@

$appXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>
'@

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $work '[Content_Types].xml'), $contentTypes, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work '_rels\.rels'), $rootRels, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work 'word\document.xml'), $documentXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work 'word\styles.xml'), $stylesXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work 'word\_rels\document.xml.rels'), $docRels, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work 'docProps\core.xml'), $coreXml, $utf8NoBom)
[System.IO.File]::WriteAllText((Join-Path $work 'docProps\app.xml'), $appXml, $utf8NoBom)

$zipPath = [System.IO.Path]::ChangeExtension($output, '.zip')
if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force }

Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zipPath -Force
Move-Item -LiteralPath $zipPath -Destination $output
Remove-Item -LiteralPath $work -Recurse -Force

Write-Output "Created: $output"



