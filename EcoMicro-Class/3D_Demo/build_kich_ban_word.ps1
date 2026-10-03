$ErrorActionPreference = 'Stop'
$folder = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $folder 'Bo_kich_ban_demo_3D_EcoMicro_Class.docx'
$temp = Join-Path $env:TEMP ('ecomicro_scenarios_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp,(Join-Path $temp '_rels'),(Join-Path $temp 'word'),(Join-Path $temp 'word\_rels'),(Join-Path $temp 'docProps') | Out-Null

function X([string]$value) { [System.Security.SecurityElement]::Escape($value) }
$body = New-Object System.Text.StringBuilder

function P {
  param([string]$Text='', [string]$Style='Normal', [switch]$Bold, [switch]$Italic, [string]$Align='', [int]$Size=0, [string]$Color='', [int]$After=120)
  $pPr = "<w:pStyle w:val=`"$Style`"/><w:spacing w:after=`"$After`"/>"
  if ($Align) { $pPr += "<w:jc w:val=`"$Align`"/>" }
  $rPr = ''
  if ($Bold) { $rPr += '<w:b/>' }
  if ($Italic) { $rPr += '<w:i/>' }
  if ($Size) { $rPr += "<w:sz w:val=`"$Size`"/><w:szCs w:val=`"$Size`"/>" }
  if ($Color) { $rPr += "<w:color w:val=`"$Color`"/>" }
  [void]$body.Append("<w:p><w:pPr>$pPr</w:pPr><w:r><w:rPr>$rPr</w:rPr><w:t xml:space=`"preserve`">$(X $Text)</w:t></w:r></w:p>")
}

function Bullet([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"420`" w:hanging=`"240`"/><w:spacing w:after=`"70`"/></w:pPr><w:r><w:t xml:space=`"preserve`">•  $(X $Text)</w:t></w:r></w:p>")
}

function Quote([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"420`" w:right=`"300`"/><w:spacing w:before=`"80`" w:after=`"120`"/><w:shd w:fill=`"EAF2F8`"/><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:color=`"2E75B6`" w:space=`"8`"/></w:pBdr></w:pPr><w:r><w:rPr><w:i/><w:color w:val=`"1F4E78`"/></w:rPr><w:t xml:space=`"preserve`">$(X $Text)</w:t></w:r></w:p>")
}

function Callout([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"260`" w:right=`"260`"/><w:spacing w:before=`"80`" w:after=`"120`"/><w:shd w:fill=`"FFF2CC`"/><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:color=`"BF9000`" w:space=`"8`"/></w:pBdr></w:pPr><w:r><w:rPr><w:b/><w:color w:val=`"7F6000`"/></w:rPr><w:t xml:space=`"preserve`">$(X $Text)</w:t></w:r></w:p>")
}

function Table([object[]]$Rows, [int[]]$Widths) {
  [void]$body.Append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="A6A6A6"/><w:left w:val="single" w:sz="6" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="6" w:color="A6A6A6"/><w:right w:val="single" w:sz="6" w:color="A6A6A6"/><w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>')
  for ($r=0; $r -lt $Rows.Count; $r++) {
    [void]$body.Append('<w:tr>')
    for ($c=0; $c -lt $Rows[$r].Count; $c++) {
      $shade = if ($r -eq 0) {'<w:shd w:fill="1F4E78"/>'} elseif ($r % 2 -eq 0) {'<w:shd w:fill="F1F6F9"/>'} else {''}
      $run = if ($r -eq 0) {'<w:b/><w:color w:val="FFFFFF"/>'} else {''}
      $width = if ($c -lt $Widths.Count) {$Widths[$c]} else {2400}
      [void]$body.Append("<w:tc><w:tcPr><w:tcW w:w=`"$width`" w:type=`"dxa`"/>$shade</w:tcPr><w:p><w:pPr><w:spacing w:after=`"0`"/></w:pPr><w:r><w:rPr>$run</w:rPr><w:t xml:space=`"preserve`">$(X ([string]$Rows[$r][$c]))</w:t></w:r></w:p></w:tc>")
    }
    [void]$body.Append('</w:tr>')
  }
  [void]$body.Append('</w:tbl><w:p><w:pPr><w:spacing w:after="100"/></w:pPr></w:p>')
}

function PageBreak { [void]$body.Append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>') }

function Scenario {
  param(
    [string]$Number, [string]$Name, [string]$Goal, [string]$Setup,
    [string[]]$Actions, [string]$Expected, [string[]]$Questions,
    [string]$Conclusion, [string]$Extension=''
  )
  P "KỊCH BẢN $Number. $Name" 'Heading1'
  P 'Mục tiêu' 'Heading2'; P $Goal
  P 'Thiết lập trên mô hình 3D' 'Heading2'; P $Setup
  P 'Thao tác và lời dẫn' 'Heading2'
  foreach ($action in $Actions) { Bullet $action }
  P 'Phản ứng mong đợi' 'Heading2'; Callout $Expected
  P 'Câu hỏi cho học sinh' 'Heading2'
  foreach ($question in $Questions) { Quote $question }
  P 'Kết luận cần chốt' 'Heading2'; P $Conclusion
  if ($Extension) { P 'Mở rộng' 'Heading2'; P $Extension }
}

# Trang bìa
P 'BỘ KỊCH BẢN DEMO 3D' 'Title' -Bold -Align center -Size 42 -Color '1F4E78' -After 180
P 'EcoMicro-Class' 'Title' -Bold -Align center -Size 36 -Color '70AD47' -After 160
P 'Dạy học sinh lớp 6 khám phá hệ thống quản lý vi khí hậu và năng lượng học đường' 'Subtitle' -Italic -Align center -Size 24 -Color '404040' -After 480
P 'Bao gồm: lời thoại giáo viên – thao tác mô hình – câu hỏi – kết luận – hoạt động phản biện' 'Normal' -Align center -Bold -Size 22 -After 700
P 'Sử dụng cùng mô hình: 3D_Demo/index.html' 'Normal' -Align center -Italic -Color '666666'
PageBreak

P 'A. CHUẨN BỊ TRƯỚC BUỔI DẠY' 'Heading1'
Bullet 'Mở file index.html bằng Edge hoặc Chrome; nhấn F11 để trình chiếu toàn màn hình.'
Bullet 'Kiểm tra các nút tình huống, thanh kéo, chế độ tự động/thủ công và thao tác xoay mô hình.'
Bullet 'Chuẩn bị một đèn pin, mẫu LDR, servo hoặc bo Arduino thật nếu muốn cho học sinh xem sau phần 3D.'
Bullet 'Không giới thiệu toàn bộ đáp án trước. Giáo viên để học sinh dự đoán rồi mới chạy tình huống.'
Bullet 'Chuẩn bị giấy ghi ý tưởng và ba phiếu lựa chọn sở thích: lắp ráp, code, thí nghiệm, thiết kế, thuyết trình.'
Callout 'Nguyên tắc buổi đầu: tạo tò mò trước, gọi tên kiến thức sau. Mỗi cảnh đều theo nhịp DỰ ĐOÁN → QUAN SÁT → GIẢI THÍCH → CẢI TIẾN.'

P 'B. KỊCH BẢN MỞ ĐẦU CHO HỌC SINH CHƯA BIẾT GÌ (40–45 PHÚT)' 'Heading1'
P 'Mục tiêu của phần mở đầu' 'Heading2'
P 'Học sinh được giải thích từng khái niệm bằng ví dụ quen thuộc. Các em chưa cần biết Arduino, cảm biến hay lập trình. Giáo viên chỉ mở mô hình 3D sau khi học sinh đã hiểu STEM, hiểu cuộc thi, hiểu bài học hôm nay và hiểu những bộ phận cơ bản của sản phẩm.'
Callout 'Quy tắc ngôn ngữ: nói câu ngắn; mỗi lần chỉ giới thiệu một ý; ưu tiên ví dụ gần gũi; nói từ dễ trước rồi mới gọi tên kỹ thuật.'

P 'Bước 1. Làm quen và bắt đầu từ đồ vật hằng ngày (0–5 phút)' 'Heading2'
Quote 'Chào các em! Hôm nay chúng ta sẽ bắt đầu một hành trình mới. Các em chưa cần biết điện tử hay lập trình. Chúng ta sẽ học từng bước, cùng thử, cùng sai và cùng sửa.'
Quote 'Các em đã từng thấy cửa tự mở ở siêu thị chưa? Đèn cầu thang tự sáng khi có người đi qua chưa? Vì sao chúng biết lúc nào cần hoạt động?'
Bullet 'Cho từng em kể một đồ vật tự động mà em từng thấy.'
Bullet 'Nếu các em chưa biết câu trả lời, giáo viên nói: “Không sao. Cuối buổi hôm nay chúng ta sẽ hiểu cách một đồ vật có thể tự hoạt động.”'

P 'Bước 2. Giải thích STEM bằng ngôn ngữ đơn giản (5–12 phút)' 'Heading2'
Quote 'STEM là cách chúng ta dùng kiến thức để giải quyết một việc trong đời sống. Ví dụ: lớp nóng quá. Ta tìm hiểu vì sao nóng, nghĩ cách làm mát, làm thử một mô hình rồi đo xem có mát hơn không.'
Table @(
  @('Phần','Cách nói với học sinh','Ví dụ rất đơn giản'),
  @('Khoa học','Tìm hiểu vì sao sự việc xảy ra','Vì sao phòng có nắng thì nóng hơn?'),
  @('Công nghệ','Dùng công cụ để hỗ trợ con người','Dùng cảm biến để biết phòng sáng hay tối'),
  @('Kỹ thuật','Nghĩ cách và chế tạo một đồ vật','Làm rèm có thể tự kéo'),
  @('Toán học','Đo, đếm và so sánh','Nhiệt độ trước và sau khi bật quạt')
) @(1800,3700,3700)
Quote 'STEM không phải là học thuộc bốn chữ. STEM là nhìn thấy một vấn đề, nghĩ cách giải quyết, làm thử và kiểm tra xem cách đó có tốt không.'
P 'Kiểm tra nhanh: hỏi “Nếu làm một chiếc quạt nhưng không đo xem phòng có mát hơn không, chúng ta còn thiếu phần nào?” Chấp nhận nhiều cách diễn đạt; mục tiêu là học sinh nói được “thiếu kiểm tra hoặc đo đạc”.'

P 'Bước 3. Giải thích cuộc thi là gì (12–18 phút)' 'Heading2'
Quote 'Chúng ta sẽ chuẩn bị cho một cuộc thi STEM. Cuộc thi giống như một thử thách lớn. Các em sẽ dùng điều đã học để trả lời câu hỏi, làm sản phẩm và kể cho người khác hiểu sản phẩm của mình.'
Table @(
  @('Vòng','Cách giải thích đơn giản'),
  @('Vòng 1','Mỗi bạn trả lời các câu hỏi để thể hiện cách suy nghĩ và vận dụng kiến thức.'),
  @('Vòng 2','Ba bạn cùng làm một sản phẩm, mang đi giới thiệu và trả lời câu hỏi.'),
  @('Vòng 3','Đội nhận một đề bài mới và vật liệu mới, rồi cùng làm một mô hình trong 90 phút.')
) @(2000,7200)
Quote 'Cuộc thi không yêu cầu một bạn phải biết hết mọi thứ. Ba bạn sẽ học cách giúp nhau. Có bạn thích lắp, có bạn thích suy nghĩ, có bạn thích nói. Nhưng cuối cùng cả ba đều cần hiểu sản phẩm.'
Bullet 'Giải thích “Ban Giám khảo” là những thầy cô nghe đội giới thiệu và đặt câu hỏi.'
Bullet 'Giải thích “phản biện” bằng câu dễ: trả lời vì sao nhóm chọn cách này và đưa ra bằng chứng.'

P 'Bước 4. Dạy trước những từ cần dùng (18–25 phút)' 'Heading2'
Table @(
  @('Từ mới','Giải thích cho học sinh lớp 6'),
  @('Vấn đề','Một điều đang gây khó chịu hoặc cần được làm tốt hơn'),
  @('Giải pháp','Cách chúng ta nghĩ ra để xử lý vấn đề'),
  @('Mô hình','Phiên bản nhỏ để làm thử trước khi áp dụng thật'),
  @('Tự động','Thiết bị tự làm việc khi gặp một điều kiện đã được dạy trước'),
  @('Cảm biến','Bộ phận giúp máy nhận biết ánh sáng, nhiệt độ hoặc có người'),
  @('Bộ điều khiển','Bộ phận nhận thông tin rồi quyết định phải làm gì'),
  @('Chương trình','Các câu lệnh con người viết để dạy máy cách xử lý'),
  @('Dữ liệu','Các con số máy đo được, ví dụ 30°C hoặc mức sáng 600')
) @(2300,6900)
Quote 'Máy không tự thông minh như con người. Con người phải gắn cho máy bộ phận để cảm nhận và viết các câu lệnh để dạy máy phải làm gì.'

P 'Bước 5. Giới thiệu bài học hôm nay (25–30 phút)' 'Heading2'
Quote 'Bài học hôm nay có tên: Lớp học biết quan sát và tự điều chỉnh. Chúng ta sẽ tìm hiểu cách một lớp học nhận biết sáng, tối, nóng và có người; sau đó tự điều khiển rèm, đèn và quạt.'
P 'Sau buổi học, các em chỉ cần làm được ba việc:'
Bullet 'Kể được ít nhất hai vấn đề thường gặp trong lớp học.'
Bullet 'Nói được cảm biến giống giác quan, bộ điều khiển giống bộ não, còn rèm–đèn–quạt giống tay chân.'
Bullet 'Nhìn một tình huống đơn giản và dự đoán thiết bị nào nên bật hoặc tắt.'

P 'Bước 6. Tìm vấn đề trong chính lớp học (30–35 phút)' 'Heading2'
Quote 'Trong lớp học của các em, khi nào các em thấy khó chịu: quá nóng, quá tối, nắng chói vào mắt hay không khí quá ẩm?'
Bullet 'Mỗi em chọn một vấn đề và kể một lần mình đã gặp.'
Bullet 'Giáo viên vẽ ba cột đơn giản: Điều xảy ra – Cách đang làm – Điều chưa tiện.'
Quote 'Nếu trời đang nắng rồi bất ngờ có mây đen, ai sẽ biết để bật đèn? Nếu giáo viên đang giảng, ai sẽ liên tục kéo rèm?'
P 'Chốt bằng câu đơn giản: “Chúng ta cần một lớp học có thể nhận biết sự thay đổi và tự hỗ trợ con người.”'

P 'Bước 7. Giới thiệu sản phẩm EcoMicro-Class (35–40 phút)' 'Heading2'
Quote 'Sản phẩm mà chúng ta sẽ cùng tìm hiểu và sau này cùng chế tạo có tên EcoMicro-Class. Đây là mô hình một lớp học nhỏ. Nó có các bộ phận để nhận biết môi trường, một bộ điều khiển để quyết định, và các thiết bị để thực hiện quyết định.'
Table @(
  @('Cách ví von','Bộ phận của mô hình','Việc nó làm'),
  @('Mắt và da','Cảm biến','Nhận biết sáng, tối, nóng, ẩm hoặc có người'),
  @('Bộ não','Bo điều khiển','Nhận số liệu và chọn hành động'),
  @('Tay chân','Rèm, đèn, quạt','Thực hiện hành động'),
  @('Thức ăn/năng lượng','Nguồn điện, pin mặt trời','Cung cấp điện để hệ thống hoạt động')
) @(2200,3000,4000)
P 'Chưa dùng các từ LDR, DHT22, PIR, PWM hoặc hysteresis ở lần giải thích đầu. Những tên này sẽ được học sau khi các em đã hiểu chức năng.'

P 'Bước 8. Mở demo 3D để quan sát có hướng dẫn (40–45 phút)' 'Heading2'
Bullet 'Bây giờ mới mở file index.html và nhấn F11.'
Bullet 'Chỉ mở cảnh “Bình thường”, chưa chạy liên tiếp các cảnh khác.'
Bullet 'Xoay chậm mô hình và chỉ từng bộ phận. Sau mỗi bộ phận, dừng để học sinh nhắc lại chức năng.'
Quote 'Đây là cửa sổ và rèm. Đây là đèn. Đây là quạt. Các chấm sáng là nơi đặt cảm biến. Các con số bên dưới là dữ liệu mà hệ thống đang nhìn thấy.'
Quote 'Bây giờ các em hãy chỉ giúp thầy/cô: đâu là giác quan, đâu là bộ não, đâu là tay chân và đâu là nguồn năng lượng?'
Callout 'Chỉ khi cả ba học sinh trả lời được bốn nhóm bộ phận trên mới chuyển sang phần kiểm thử tình huống.'

P 'Bước 9. Hướng dẫn cách thử một kịch bản (sau phút 45)' 'Heading2'
P 'Không yêu cầu “phân tích” ngay. Dùng bốn câu hỏi cố định, dễ nhớ:'
Table @(
  @('Bước','Câu giáo viên hỏi'),
  @('1. Đoán','Theo em, điều gì sắp xảy ra?'),
  @('2. Xem','Em nhìn thấy bộ phận nào thay đổi?'),
  @('3. Nói lý do','Vì sao nó lại làm như vậy?'),
  @('4. Nghĩ thêm','Em có muốn thay đổi điều gì không?')
) @(2200,7000)
Quote 'Chúng ta sẽ thử từng trường hợp, không cần vội. Trước tiên là lớp học bình thường. Khi đã hiểu, chúng ta mới thử nắng gắt, mưa tối, nóng ẩm và lớp không có người.'
P 'Sau 2–3 tình huống, giáo viên mới hỏi sở thích. Chưa chia vai cố định trong buổi đầu.'

P 'Thứ tự sử dụng các kịch bản theo mức độ' 'Heading2'
Table @(
  @('Mức','Kịch bản','Thời điểm nên dùng'),
  @('Làm quen','1. Bình thường; 2. Nắng gắt; 3. Mưa tối','Buổi đầu, sau khi đã giải thích từ mới'),
  @('Hiểu sâu hơn','4. Nóng ẩm; 5. Không có người','Buổi 2–3'),
  @('Vận dụng','6. Thử thách; 7. Người và thuật toán; 8. Tự tạo kịch bản','Khi học sinh đã hiểu các bộ phận'),
  @('Nâng cao','9. Năng lượng; 10. Lỗi cảm biến','Sau khi đã học đo đạc và phần cứng')
) @(1800,4300,3100)
Callout 'Không chạy hết mười kịch bản trong buổi đầu. Với trẻ mới học, hiểu chắc hai tình huống tốt hơn xem nhanh nhiều tình huống.'

Scenario '1' 'PHÒNG HỌC BÌNH THƯỜNG' `
  'Cho học sinh làm quen với màn hình và biết trạng thái bình thường là gì.' `
  'Nhấn “Bình thường”: nắng 73%, mây 0%, nhiệt độ 27°C, độ ẩm 58%, có 4 người, bật tự động.' `
  @('Chỉ cho học sinh vị trí rèm, đèn, quạt và các con số.','Đọc chậm: nhiệt độ 27°C, độ ẩm 58%, lớp có người.','Yêu cầu học sinh nói thiết bị nào đang bật và thiết bị nào đang tắt.','Giải thích: khi căn phòng đã ổn, hệ thống không cần làm nhiều việc.') `
  'Rèm mở; đèn tắt hoặc chỉ bù sáng rất ít; quạt tắt; trạng thái “Phòng học thoải mái”.' `
  @('Em nhìn thấy rèm đang mở hay đóng?','Quạt có cần chạy không? Vì sao?','Nếu luôn bật tất cả thiết bị thì có lãng phí điện không?') `
  'Khi lớp học đã dễ chịu, hệ thống giữ nguyên và tiết kiệm điện.'

Scenario '2' 'NẮNG GẮT – CHỐNG CHÓI VÀ BÙ SÁNG' `
  'Thể hiện sự phối hợp phản hồi chéo giữa rèm và đèn.' `
  'Nhấn “Nắng gắt”: nắng 96%, nhiệt độ 31°C, có học sinh, bật tự động.' `
  @('Nói: “Bây giờ Mặt Trời chiếu rất mạnh. Theo em rèm sẽ làm gì?”','Nhấn tình huống và chỉ quan sát rèm trước.','Sau khi rèm đóng, hỏi: “Bàn học có tối hơn không? Đèn nên làm gì?”','Quan sát đèn và quạt; đọc phần giải thích trên màn hình.') `
  'Rèm đóng để giảm chói; ánh sáng tại bàn giảm nên LED bù sáng; nếu nhiệt độ vượt 30°C, quạt bật.' `
  @('Rèm đóng giúp chúng ta tránh điều gì?','Rèm đóng làm xuất hiện vấn đề mới nào?','Đèn đã giúp giải quyết vấn đề mới đó như thế nào?') `
  'Các bộ phận cần giúp đỡ lẫn nhau: rèm giảm chói, còn đèn giữ cho bàn học đủ sáng.' `
  'Cho học sinh thử chế độ thủ công: một em điều khiển rèm, một em điều khiển LED, một em đánh giá độ thoải mái.'

Scenario '3' 'MƯA GIÔNG – TRỜI TỐI ĐỘT NGỘT' `
  'Chứng minh ưu điểm của cảm biến thời gian thực so với thời khóa biểu cố định.' `
  'Nhấn “Mưa tối”: nắng 18%, mây 92%, nhiệt độ 25°C, độ ẩm 82%, có 4 người.' `
  @('Đặt câu hỏi: “Đang giữa tiết học nhưng trời đột nhiên tối, thời khóa biểu có biết không?”','Bấm Mưa tối và quan sát đèn.','Tăng/giảm thanh Mây để học sinh thấy thay đổi liên tục.','So sánh quyết định theo thời gian với một lịch bật đèn cố định.') `
  'Rèm mở; LED tăng sáng để bù ánh sáng tự nhiên bị giảm; quạt không nhất thiết bật; hệ thống cảnh báo ẩm cao.' `
  @('Vì sao hệ thống không cần biết hôm nay là thứ mấy?','Nếu không có người trong lớp thì đèn có nên bật không?','Dữ liệu nào giúp hệ thống biết trời đã tối?') `
  'Lịch trình chỉ dự đoán trước; cảm biến phản ứng với điều kiện thật. Giải pháp tốt có thể kết hợp lịch và cảm biến nhưng cảm biến là nguồn quyết định quan trọng.'

Scenario '4' 'NÓNG VÀ ẨM CAO' `
  'Giúp học sinh phân biệt làm mát, thông gió và phun sương; sửa ngộ nhận rằng phun sương luôn có lợi.' `
  'Nhấn “Nóng ẩm”: nhiệt độ 34°C, độ ẩm 84%, nắng 72%, có học sinh.' `
  @('Cho học sinh dự đoán: bật quạt, phun sương hay cả hai.','Nhấn tình huống và đọc phần “Bộ não đang suy luận”.','Kéo độ ẩm từ 50% lên 90%, giữ nguyên nhiệt độ.','Yêu cầu học sinh giải thích vì sao quyết định cần thay đổi theo hai đại lượng.') `
  'Quạt bật; hệ thống ghi rõ độ ẩm cao và không cho phép phun sương.' `
  @('Nếu phun thêm nước khi không khí đã rất ẩm, điều gì xảy ra?','Nhiệt độ và độ ẩm khác nhau như thế nào?','Giải pháp thực tế nào phù hợp hơn khi trời nồm?') `
  'Phun sương làm tăng độ ẩm. Khi độ ẩm cao, nên thông gió hoặc hút ẩm; không được tuyên bố phun sương xử lý nồm.' `
  'Có thể yêu cầu học sinh thiết kế thêm điều kiện khóa: chỉ phun khi nóng và độ ẩm dưới một ngưỡng an toàn.'

Scenario '5' 'KHÔNG CÓ NGƯỜI – TIẾT KIỆM NĂNG LƯỢNG' `
  'Giải thích vai trò của PIR và nguyên tắc không sử dụng năng lượng khi không cần.' `
  'Nhấn “Không có người”: số người 0, ánh sáng thấp, hệ thống tự động.' `
  @('Hỏi trước: “Phòng tối nhưng không có ai, có nên bật đèn không?”','Nhấn tình huống và quan sát học sinh biến mất, LED tắt.','Kéo số người từ 0 lên 1 để kích hoạt lại logic chiếu sáng.','Giải thích cảm biến PIR phát hiện chuyển động, không nhận dạng danh tính.') `
  'Đèn tắt dù phòng thiếu sáng; các thiết bị chỉ hoạt động khi cần thiết. Quạt có thể được cấu hình thông gió bảo vệ thiết bị nhưng bản demo ưu tiên tiết kiệm.' `
  @('PIR có biết người đó là ai không?','Nếu học sinh ngồi im lâu, PIR có thể gặp vấn đề gì?','Làm thế nào tránh đèn tắt nhầm giữa giờ học?') `
  'Cảm biến có giới hạn. Cần bộ đếm thời gian duy trì hiện diện hoặc kết hợp nhiều loại cảm biến để tránh tắt nhầm.' `
  'Liên hệ code Arduino: PRESENCE_HOLD_MS giữ trạng thái có người thêm một khoảng thời gian sau lần phát hiện chuyển động cuối.'

Scenario '6' 'THỬ THÁCH TỔNG HỢP' `
  'Rèn khả năng xử lý nhiều điều kiện đồng thời và xác định thứ tự ưu tiên.' `
  'Nhấn “Thử thách”: nắng 91%, nhiệt độ 35°C, độ ẩm 88%, có 4 người.' `
  @('Không cho xem phản ứng ngay; mỗi em viết chuỗi quyết định riêng.','Nhấn tình huống, sau đó so sánh dự đoán với hệ thống.','Yêu cầu nhóm xếp thứ tự ưu tiên: an toàn, sức khỏe, thị lực, tiết kiệm năng lượng.','Tắt tự động để nhóm tự điều khiển, rồi bật lại để so sánh.') `
  'Rèm đóng chống chói; LED bù sáng; quạt bật; phun sương bị khóa vì độ ẩm quá cao.' `
  @('Nếu nguồn điện không đủ cho mọi thiết bị, thiết bị nào được ưu tiên?','Có quyết định nào giải quyết một vấn đề nhưng làm vấn đề khác xấu đi?','Nhóm sẽ đo những gì để chứng minh quyết định là đúng?') `
  'Một hệ thống thực tế phải xử lý đồng thời nhiều mục tiêu và có thứ tự ưu tiên rõ ràng, không chỉ là nhiều câu lệnh if rời rạc.'

Scenario '7' 'SO SÁNH CON NGƯỜI VÀ THUẬT TOÁN' `
  'Biến học sinh thành người điều khiển, sau đó đánh giá bằng dữ liệu thay vì cảm tính.' `
  'Chọn một tình huống bất kỳ, tắt “Điều khiển tự động”.' `
  @('Học sinh 1 điều khiển rèm; học sinh 2 chỉnh LED và quạt; học sinh 3 ghi số liệu.','Cho nhóm 30 giây tạo trạng thái thoải mái và tiết kiệm.','Ghi ánh sáng tại bàn và công suất tải.','Bật tự động, ghi lại cùng hai chỉ số rồi so sánh.') `
  'Học sinh nhận ra có thể đạt độ sáng tương tự bằng nhiều cách nhưng mức tiêu thụ năng lượng khác nhau.' `
  @('Phương án nào dùng ít điện hơn?','Phương án nào phản ứng nhất quán hơn?','Con người nên có quyền ghi đè hệ thống trong trường hợp nào?') `
  'Tự động hóa hỗ trợ con người; vẫn cần nút điều khiển tay và quy tắc an toàn. Kết luận phải dựa trên số đo, không chỉ dựa vào cảm giác.'

Scenario '8' 'HỌC SINH TỰ TẠO VÀ LƯU KỊCH BẢN' `
  'Khuyến khích tư duy thiết kế thí nghiệm và sở hữu ý tưởng.' `
  'Dùng các thanh kéo tự tạo nắng, mây, nhiệt độ, độ ẩm, số người; nhập tên ở “Lưu kịch bản của tôi”.' `
  @('Mỗi em tạo một tình huống có tên và nêu bối cảnh thực tế.','Trước khi bật tự động, em phải viết dự đoán.','Nhấn Lưu, sau đó mở lại để kiểm tra kết quả có lặp lại không.','Đổi đúng một biến và quan sát kết quả để hình thành thí nghiệm công bằng.') `
  'Kịch bản được lưu trong trình duyệt; nhóm có thể mở lại, so sánh và xóa.' `
  @('Biến độc lập của thí nghiệm là gì?','Biến nào phải giữ nguyên?','Kết quả nào khiến em sửa thuật toán?') `
  'Một thí nghiệm tốt chỉ thay đổi có chủ đích, ghi lại dữ liệu và có khả năng lặp lại.' `
  'Gợi ý tên: “Chiều hè hướng Tây”, “Mưa giông giữa tiết 3”, “Lớp trống giờ ra chơi”, “Nồm ẩm tháng Hai”.'

Scenario '9' 'NĂNG LƯỢNG MẶT TRỜI KHÔNG ĐỦ TẢI' `
  'Giúp học sinh hiểu cân bằng công suất và không tuyên bố hệ thống tự chủ năng lượng khi chưa có số liệu.' `
  'Tạo nắng thấp, mây cao, có người; tăng LED và bật quạt ở chế độ thủ công.' `
  @('Quan sát Solar và Công suất tải.','Ghi hai giá trị, tính phần công suất còn thiếu.','Giảm LED hoặc tắt tải không cần thiết để cân bằng.','Thảo luận vai trò của pin lưu trữ và nguồn USB dự phòng.') `
  'Khi công suất tải lớn hơn công suất solar, màn hình báo đang dùng pin/USB.' `
  @('Pin mặt trời đang tạo ra bao nhiêu watt?','Thiết bị đang tiêu thụ bao nhiêu watt?','Pin lưu trữ giải quyết vấn đề gì và có giới hạn gì?') `
  'Muốn khẳng định tự chủ năng lượng, nhóm phải đo công suất tạo ra, công suất tiêu thụ và thời gian hoạt động; không chỉ gắn một tấm pin vào mô hình.'

Scenario '10' 'LỖI CẢM BIẾN VÀ PHƯƠNG ÁN DỰ PHÒNG' `
  'Chuẩn bị cho phản biện và tình huống thiết bị thật gặp lỗi.' `
  'Dùng chế độ thủ công để mô phỏng cảm biến báo sai hoặc hệ thống tự động không phản ứng.' `
  @('Giả sử cảm biến ánh sáng bị che hoặc đứt dây.','Yêu cầu nhóm nhận biết dấu hiệu bất thường từ số liệu.','Chuyển sang điều khiển tay để duy trì hoạt động.','Đề xuất thông báo lỗi, giới hạn an toàn và cảm biến dự phòng.') `
  'Nhóm không giấu lỗi; biết chuyển sang chế độ an toàn, giải thích nguyên nhân và kế hoạch khắc phục.' `
  @('Làm sao biết giá trị cảm biến là không hợp lý?','Nếu servo kẹt, hệ thống nên tiếp tục ra lệnh hay dừng?','Khi nào cần cho giáo viên quyền điều khiển bằng tay?') `
  'Một sản phẩm tốt không phải sản phẩm không bao giờ lỗi, mà là sản phẩm phát hiện được lỗi và chuyển sang trạng thái an toàn.'

P 'C. PHÂN VAI CHO BA HỌC SINH' 'Heading1'
Table @(
  @('Thành viên','Vai trò khi demo','Câu cần nói được'),
  @('Học sinh 1','Điều khiển mô hình và phần cứng','Cảm biến nào cung cấp dữ liệu; thiết bị nào thực hiện lệnh'),
  @('Học sinh 2','Giải thích thuật toán và số liệu','Vì sao có ngưỡng; vì sao rèm và đèn phối hợp'),
  @('Học sinh 3','Dẫn tình huống và phản biện','Vấn đề thực tế, kết quả đo, giới hạn và cải tiến')
) @(1800,3000,4400)
P 'Sau mỗi lượt luyện tập phải đổi vai. Mục tiêu cuối cùng là cả ba em đều có thể trình bày toàn bộ chuỗi cảm biến → quyết định → hành động → kiểm chứng.'

P 'D. PHIẾU GHI NHANH CHO MỖI KỊCH BẢN' 'Heading1'
Table @(
  @('Nội dung','Học sinh ghi'),
  @('Tên tình huống','................................................................................'),
  @('Dự đoán trước khi chạy','................................................................................'),
  @('Dữ liệu đầu vào','Nắng: ........  Mây: ........  Nhiệt độ: ........  Độ ẩm: ........  Số người: ........'),
  @('Hành động của hệ thống','Rèm: ........  LED: ........  Quạt: ........'),
  @('Kết quả quan sát','................................................................................'),
  @('Điều chưa hợp lý','................................................................................'),
  @('Đề xuất cải tiến','................................................................................')
) @(2600,6600)

P 'E. KẾT THÚC BUỔI HỌC' 'Heading1'
Quote 'Hôm nay em phát hiện điều gì mà trước khi xem mô hình em chưa nghĩ tới?'
Quote 'Nếu được bổ sung một cảm biến hoặc một quy tắc, em sẽ bổ sung gì và vì sao?'
Quote 'Em muốn đảm nhiệm phần nào trong ba buổi tiếp theo?'
P 'Giáo viên chụp lại hoặc ghi tên các kịch bản học sinh đã lưu, tổng hợp lựa chọn sở thích và giao nhiệm vụ nhỏ: mỗi em tìm một vấn đề vi khí hậu thực tế trong lớp học để báo cáo ở buổi sau.'

$sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1276"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
$document = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`"><w:body>$($body.ToString())$sect</w:body></w:document>"

$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Arial"/><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto" w:after="120"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="300" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="180" w:after="90"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="548235"/><w:sz w:val="25"/><w:szCs w:val="25"/></w:rPr></w:style>
</w:styles>
'@
$types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
$docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
$stamp = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
$core = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>Bộ kịch bản demo 3D EcoMicro-Class</dc:title><dc:creator>EcoMicro-Class Team</dc:creator><dc:subject>Giáo dục STEM lớp 6</dc:subject><dcterms:created xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:created><dcterms:modified xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:modified></cp:coreProperties>"
$app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $temp '[Content_Types].xml'),$types,$utf8)
[IO.File]::WriteAllText((Join-Path $temp '_rels\.rels'),$rels,$utf8)
[IO.File]::WriteAllText((Join-Path $temp 'word\document.xml'),$document,$utf8)
[IO.File]::WriteAllText((Join-Path $temp 'word\styles.xml'),$styles,$utf8)
[IO.File]::WriteAllText((Join-Path $temp 'word\_rels\document.xml.rels'),$docRels,$utf8)
[IO.File]::WriteAllText((Join-Path $temp 'docProps\core.xml'),$core,$utf8)
[IO.File]::WriteAllText((Join-Path $temp 'docProps\app.xml'),$app,$utf8)
$zip = [IO.Path]::ChangeExtension($output,'.zip')
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
if (Test-Path $output) { Remove-Item -LiteralPath $output -Force }
Compress-Archive -Path (Join-Path $temp '*') -DestinationPath $zip -Force
Move-Item -LiteralPath $zip -Destination $output
Remove-Item -LiteralPath $temp -Recurse -Force
Write-Output "Created: $output"



