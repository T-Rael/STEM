$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root 'Ke_hoach_chi_tiet_buoi_1_va_dap_an.docx'
$work = Join-Path $env:TEMP ('ecomicro_lesson1_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work,(Join-Path $work '_rels'),(Join-Path $work 'word'),(Join-Path $work 'word\_rels'),(Join-Path $work 'docProps') | Out-Null

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

function Numbered([string]$Text,[int]$Number) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"420`" w:hanging=`"300`"/><w:spacing w:after=`"70`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$Number.  $(X $Text)</w:t></w:r></w:p>")
}

function Teacher([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"420`" w:right=`"260`"/><w:spacing w:before=`"80`" w:after=`"110`"/><w:shd w:fill=`"EAF2F8`"/><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:color=`"2E75B6`" w:space=`"8`"/></w:pBdr></w:pPr><w:r><w:rPr><w:i/><w:color w:val=`"1F4E78`"/></w:rPr><w:t xml:space=`"preserve`">GIÁO VIÊN NÓI: $(X $Text)</w:t></w:r></w:p>")
}

function Answer([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"420`" w:right=`"260`"/><w:spacing w:after=`"100`"/><w:shd w:fill=`"E2F0D9`"/><w:pBdr><w:left w:val=`"single`" w:sz=`"14`" w:color=`"70AD47`" w:space=`"8`"/></w:pBdr></w:pPr><w:r><w:rPr><w:color w:val=`"375623`"/></w:rPr><w:t xml:space=`"preserve`">TRẢ LỜI GỢI Ý: $(X $Text)</w:t></w:r></w:p>")
}

function Callout([string]$Text) {
  [void]$body.Append("<w:p><w:pPr><w:ind w:left=`"260`" w:right=`"260`"/><w:spacing w:before=`"80`" w:after=`"120`"/><w:shd w:fill=`"FFF2CC`"/><w:pBdr><w:left w:val=`"single`" w:sz=`"18`" w:color=`"BF9000`" w:space=`"8`"/></w:pBdr></w:pPr><w:r><w:rPr><w:b/><w:color w:val=`"7F6000`"/></w:rPr><w:t xml:space=`"preserve`">LƯU Ý: $(X $Text)</w:t></w:r></w:p>")
}

function Table([object[]]$Rows,[int[]]$Widths) {
  [void]$body.Append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="A6A6A6"/><w:left w:val="single" w:sz="6" w:color="A6A6A6"/><w:bottom w:val="single" w:sz="6" w:color="A6A6A6"/><w:right w:val="single" w:sz="6" w:color="A6A6A6"/><w:insideH w:val="single" w:sz="4" w:color="D9D9D9"/><w:insideV w:val="single" w:sz="4" w:color="D9D9D9"/></w:tblBorders><w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>')
  for($r=0;$r -lt $Rows.Count;$r++){
    [void]$body.Append('<w:tr>')
    for($c=0;$c -lt $Rows[$r].Count;$c++){
      $shade=if($r -eq 0){'<w:shd w:fill="1F4E78"/>'}elseif($r%2 -eq 0){'<w:shd w:fill="F1F6F9"/>'}else{''}
      $run=if($r -eq 0){'<w:b/><w:color w:val="FFFFFF"/>'}else{''}
      $width=if($c -lt $Widths.Count){$Widths[$c]}else{2400}
      [void]$body.Append("<w:tc><w:tcPr><w:tcW w:w=`"$width`" w:type=`"dxa`"/>$shade</w:tcPr><w:p><w:pPr><w:spacing w:after=`"0`"/></w:pPr><w:r><w:rPr>$run</w:rPr><w:t xml:space=`"preserve`">$(X ([string]$Rows[$r][$c]))</w:t></w:r></w:p></w:tc>")
    }
    [void]$body.Append('</w:tr>')
  }
  [void]$body.Append('</w:tbl><w:p><w:pPr><w:spacing w:after="100"/></w:pPr></w:p>')
}

function QA([string]$Question,[string]$Expected,[string]$Support='') {
  P "Câu hỏi: $Question" 'Question'
  Answer $Expected
  if($Support){ P "Nếu học sinh chưa trả lời được: $Support" 'Support' }
}

function PageBreak { [void]$body.Append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>') }

# Cover
P 'KẾ HOẠCH CHI TIẾT BUỔI 1' 'Title' -Bold -Align center -Size 42 -Color '1F4E78' -After 160
P 'LÀM QUEN VỚI STEM VÀ ECOMICRO-CLASS' 'Title' -Bold -Align center -Size 34 -Color '70AD47' -After 180
P 'Kèm lời thoại giáo viên, câu hỏi và đáp án gợi ý' 'Subtitle' -Italic -Align center -Size 25 -Color '404040' -After 500
P 'Đối tượng: 03 học sinh lớp 6 – chưa có kiến thức trước về STEM, điện tử hoặc lập trình' 'Normal' -Align center -Bold -Size 21 -After 140
P 'Thời lượng chuẩn: 75 phút | Phương án rút gọn: 60 phút' 'Normal' -Align center -Size 21 -Color '666666'
PageBreak

P '1. MỤC TIÊU BUỔI HỌC' 'Heading1'
P 'Buổi đầu không nhằm dạy học sinh đấu mạch hoặc viết code. Mục tiêu là giúp các em hiểu mình sắp học gì, vì sao sản phẩm được tạo ra và biết quan sát một hệ thống tự động ở mức cơ bản.'
P 'Cuối buổi, học sinh có thể:' 'Heading2'
Bullet 'Nói được STEM là dùng kiến thức để giải quyết một vấn đề trong đời sống.'
Bullet 'Kể được ba việc chính của cuộc thi: trả lời câu hỏi, làm sản phẩm và giới thiệu sản phẩm.'
Bullet 'Giải thích được các từ: vấn đề, giải pháp, mô hình, tự động, cảm biến và dữ liệu.'
Bullet 'Nêu được ít nhất hai vấn đề thường gặp trong lớp học.'
Bullet 'Nhận ra cảm biến giống giác quan, bộ điều khiển giống bộ não, rèm–đèn–quạt giống tay chân.'
Bullet 'Thực hiện được bốn bước khi xem demo: Đoán → Xem → Nói lý do → Nghĩ thêm.'
Callout 'Không yêu cầu học sinh nhớ tên LDR, DHT22, PIR, PWM hoặc viết code trong buổi này.'

P '2. CHUẨN BỊ' 'Heading1'
Bullet 'Máy tính nối màn hình hoặc máy chiếu; chuột để học sinh thao tác.'
Bullet 'Mở sẵn file EcoMicro-Class/3D_Demo/index.html nhưng thu nhỏ cửa sổ, chưa cho học sinh xem trước.'
Bullet 'Ba tờ giấy nhỏ và bút để học sinh ghi dự đoán.'
Bullet 'Bảng hoặc giấy lớn, chia ba cột: Điều xảy ra – Cách đang làm – Điều chưa tiện.'
Bullet 'Nếu có: một cảm biến, Arduino hoặc servo thật để cuối buổi cho học sinh cầm xem; không cần đấu mạch.'
P 'Cách bố trí' 'Heading2'
Bullet 'Ba học sinh ngồi gần nhau và nhìn rõ màn hình.'
Bullet 'Giáo viên đứng cạnh màn hình, không đứng chắn mô hình.'
Bullet 'Mỗi lần chỉ cho một học sinh điều khiển chuột; đổi lượt sau mỗi tình huống.'

P '3. TỔNG QUAN TIẾN TRÌNH 75 PHÚT' 'Heading1'
Table @(
  @('Thời gian','Nội dung','Kết quả mong đợi'),
  @('0–5 phút','Làm quen qua đồ vật tự động','Học sinh nói được một ví dụ gần gũi'),
  @('5–12 phút','STEM là gì?','Hiểu STEM là cách giải quyết vấn đề'),
  @('12–18 phút','Cuộc thi diễn ra như thế nào?','Biết ba loại nhiệm vụ chính'),
  @('18–25 phút','Học sáu từ mới','Có vốn từ để hiểu bài'),
  @('25–30 phút','Giới thiệu bài học','Biết mục tiêu của buổi học'),
  @('30–35 phút','Tìm vấn đề trong lớp học','Nêu được nóng, tối, chói, ẩm hoặc lãng phí điện'),
  @('35–40 phút','Giới thiệu EcoMicro-Class','Hiểu bốn nhóm bộ phận'),
  @('40–45 phút','Quan sát mô hình 3D','Nhận diện rèm, đèn, quạt, cảm biến'),
  @('45–52 phút','Kịch bản 1: Bình thường','Biết cách đọc và quan sát cơ bản'),
  @('52–61 phút','Kịch bản 2: Nắng gắt','Hiểu rèm và đèn phối hợp'),
  @('61–68 phút','Kịch bản 3: Mưa tối','Hiểu hệ thống phản ứng với thay đổi'),
  @('68–75 phút','Ôn tập, sở thích và nhiệm vụ','Giáo viên biết mức hiểu và hứng thú của từng em')
) @(1500,4300,3400)

P '4. KỊCH BẢN DẠY CHI TIẾT' 'Heading1'

P 'Hoạt động 1 – Đồ vật tự động quanh em (0–5 phút)' 'Heading2'
Teacher 'Chào các em! Hôm nay chúng ta bắt đầu một hành trình mới. Các em chưa cần biết điện tử hay lập trình. Chúng ta sẽ học từng bước, cùng thử, cùng sai và cùng sửa.'
Teacher 'Các em đã từng thấy cửa ở siêu thị tự mở chưa? Đèn cầu thang tự sáng khi có người đi qua chưa? Hãy kể một đồ vật có thể tự hoạt động mà em từng thấy.'
QA 'Một đồ vật tự động mà em từng thấy là gì?' 'Có thể là cửa tự động, đèn cảm ứng, vòi nước tự chảy, quạt điều hòa tự bật/tắt, robot hút bụi hoặc đèn đường.' 'Gợi ý: “Khi em bước tới cửa siêu thị, cánh cửa có tự mở không?”'
QA 'Theo em, cửa tự động biết có người đến bằng cách nào?' 'Cửa có một bộ phận nhận biết người đến gần. Sau này chúng ta gọi bộ phận đó là cảm biến.' 'Không cần học sinh nói đúng từ “cảm biến”; chấp nhận câu “nó nhìn thấy”, “nó phát hiện” hoặc “có mắt điện tử”.'
Callout 'Không sửa câu chữ quá sớm. Trước tiên công nhận ý đúng, sau đó mới bổ sung tên gọi.'

P 'Hoạt động 2 – STEM là gì? (5–12 phút)' 'Heading2'
Teacher 'STEM là cách chúng ta dùng kiến thức để giải quyết một việc trong đời sống. Ví dụ lớp học nóng quá: ta tìm hiểu vì sao nóng, nghĩ cách làm mát, làm thử một mô hình rồi đo xem có mát hơn không.'
Table @(
  @('Phần','Cách nói đơn giản','Ví dụ'),
  @('Khoa học','Tìm hiểu vì sao sự việc xảy ra','Vì sao có nắng thì phòng nóng hơn?'),
  @('Công nghệ','Dùng công cụ để giúp con người','Dùng một bộ phận để biết phòng sáng hay tối'),
  @('Kỹ thuật','Nghĩ cách rồi làm thử một đồ vật','Làm chiếc rèm có thể tự kéo'),
  @('Toán học','Đo, đếm và so sánh','So sánh nhiệt độ trước và sau khi bật quạt')
) @(1800,3900,3500)
QA 'STEM có phải chỉ là học thuộc bốn chữ không?' 'Không. Điều quan trọng là dùng kiến thức để giải quyết một vấn đề thật.' 'Nhắc lại ví dụ phòng nóng: tìm hiểu – nghĩ cách – làm thử – đo kết quả.'
QA 'Nếu chúng ta làm một chiếc quạt nhưng không kiểm tra phòng có mát hơn không thì còn thiếu điều gì?' 'Thiếu đo đạc và kiểm tra kết quả.' 'Hỏi tiếp: “Làm sao biết chiếc quạt có giúp ích nếu chúng ta không đo hoặc quan sát?”'
QA 'Toán học được dùng ở đâu trong dự án này?' 'Dùng để đọc và so sánh các con số như nhiệt độ, độ ẩm, mức sáng hoặc lượng điện.' 'Cho ví dụ 32°C lớn hơn 28°C nên phòng nóng hơn.'

P 'Hoạt động 3 – Giới thiệu cuộc thi (12–18 phút)' 'Heading2'
Teacher 'Chúng ta sẽ chuẩn bị cho một cuộc thi STEM. Cuộc thi giống như một chuỗi thử thách. Các em sẽ trả lời câu hỏi, cùng nhau làm một sản phẩm và kể cho người khác hiểu sản phẩm hoạt động như thế nào.'
Table @(
  @('Vòng','Cách nói với học sinh'),
  @('Vòng 1','Mỗi bạn trả lời câu hỏi để thể hiện cách suy nghĩ và vận dụng kiến thức.'),
  @('Vòng 2','Hai hoặc ba bạn cùng làm sản phẩm, mang sản phẩm đi giới thiệu và trả lời câu hỏi.'),
  @('Vòng 3','Đội nhận một đề bài mới và vật liệu mới, rồi làm mô hình trong 90 phút.')
) @(2000,7200)
QA 'Cuộc thi có phải chỉ cần làm một sản phẩm đẹp không?' 'Không. Sản phẩm cần hoạt động; đội phải hiểu, thử nghiệm, có số liệu và giải thích được.' 'Hỏi: “Nếu sản phẩm đẹp nhưng không ai biết nó giải quyết việc gì thì đã đủ chưa?”'
QA 'Một bạn có phải biết và làm hết mọi việc không?' 'Không. Ba bạn chia nhau công việc và giúp nhau, nhưng cuối cùng cả ba đều cần hiểu sản phẩm.' 'Cho ví dụ: một bạn lắp, một bạn ghi số liệu, một bạn trình bày; sau đó đổi vai.'
QA 'Ban Giám khảo là ai?' 'Là những thầy cô xem sản phẩm, nghe đội giới thiệu và đặt câu hỏi.' 'Dùng cách gọi ngắn: “những người nghe và đánh giá bài thi”.'
QA 'Phản biện nghĩa là gì?' 'Là giải thích vì sao nhóm chọn cách làm đó và đưa ra bằng chứng để bảo vệ lựa chọn.' 'Không dùng nghĩa “cãi lại”; nói “trả lời có lý do và số liệu”.'

P 'Hoạt động 4 – Sáu từ mới cần biết (18–25 phút)' 'Heading2'
Table @(
  @('Từ','Giải thích đơn giản','Ví dụ'),
  @('Vấn đề','Điều đang gây khó chịu hoặc cần làm tốt hơn','Nắng chiếu gây chói mắt'),
  @('Giải pháp','Cách nghĩ ra để xử lý vấn đề','Đóng rèm để giảm chói'),
  @('Mô hình','Phiên bản nhỏ để làm thử trước','Lớp học thu nhỏ trên máy tính'),
  @('Tự động','Thiết bị tự làm khi gặp điều kiện đã được dạy','Trời tối thì đèn bật'),
  @('Cảm biến','Bộ phận giúp máy nhận biết môi trường','Nhận biết sáng, nóng hoặc có người'),
  @('Dữ liệu','Thông tin hoặc con số đo được','Nhiệt độ 30°C')
) @(1700,4100,3400)
QA '“Lớp quá nóng” là vấn đề hay giải pháp?' 'Đó là vấn đề.' 'Hỏi tiếp: “Bật quạt” là gì? Đáp án: một giải pháp.'
QA 'Mô hình có phải là lớp học thật không?' 'Không. Mô hình là phiên bản nhỏ hoặc mô phỏng để chúng ta thử ý tưởng an toàn và ít tốn kém hơn.' 'So sánh với mô hình ngôi nhà hoặc ô tô đồ chơi.'
QA 'Máy có tự nghĩ giống con người không?' 'Không. Con người phải gắn cảm biến và viết các câu lệnh để dạy máy cách xử lý.' 'Ví dụ: dạy máy “nếu quá nóng thì bật quạt”.'
QA 'Dữ liệu có thể là gì?' 'Có thể là con số nhiệt độ, độ ẩm, mức sáng hoặc thông tin có người/không có người.' 'Chỉ vào nhiệt kế hoặc hỏi nhiệt độ hôm nay.'

P 'Hoạt động 5 – Giới thiệu bài học hôm nay (25–30 phút)' 'Heading2'
Teacher 'Bài học hôm nay có tên là “Lớp học biết quan sát và tự điều chỉnh”. Chúng ta sẽ tìm hiểu cách một lớp học nhận biết sáng, tối, nóng và có người; sau đó tự điều khiển rèm, đèn và quạt.'
Teacher 'Cuối buổi, các em chỉ cần làm được ba việc: kể một vấn đề trong lớp; chỉ ra giác quan, bộ não và tay chân của mô hình; dự đoán một thiết bị nên bật hay tắt.'
QA 'Hôm nay chúng ta có phải viết code ngay không?' 'Không. Hôm nay chúng ta hiểu vấn đề và cách hệ thống hoạt động trước. Code sẽ được học từng bước ở các buổi sau.' 'Khẳng định rằng chưa biết code là hoàn toàn bình thường.'

P 'Hoạt động 6 – Tìm vấn đề trong lớp học (30–35 phút)' 'Heading2'
Teacher 'Trong lớp học, khi nào các em cảm thấy khó chịu: quá nóng, quá tối, nắng chói vào mắt hay không khí quá ẩm? Mỗi bạn hãy kể một lần mình đã gặp.'
Table @(
  @('Điều xảy ra','Cách đang làm','Điều chưa tiện'),
  @('Nắng chiếu vào mắt','Kéo rèm bằng tay','Phải đứng dậy; kéo xong lớp có thể bị tối'),
  @('Lớp tối','Bật đèn','Có thể quên tắt khi không còn người'),
  @('Lớp nóng','Bật quạt','Có thể bật cả khi phòng đã mát hoặc không có người'),
  @('Trời thay đổi bất ngờ','Người tự quan sát và điều chỉnh','Dễ chậm hoặc làm gián đoạn giờ học')
) @(2700,3000,3500)
QA 'Nếu kéo rèm để tránh chói thì vấn đề mới nào có thể xuất hiện?' 'Lớp hoặc bàn học có thể bị tối.' 'Có thể dùng tay che một phần ánh sáng để minh họa.'
QA 'Nếu dùng thời khóa biểu để bật đèn vào đúng giờ thì có gì chưa tốt?' 'Thời khóa biểu không biết hôm đó trời đang sáng, tối, mưa hay nắng.' 'Nêu ví dụ cùng 3 giờ chiều nhưng một hôm nắng, một hôm mưa.'
QA 'Chúng ta cần một lớp học như thế nào?' 'Một lớp học có thể nhận biết môi trường thay đổi và tự hỗ trợ con người.' 'Chấp nhận câu “lớp thông minh”, rồi hỏi thông minh ở điểm nào.'

P 'Hoạt động 7 – Giới thiệu EcoMicro-Class (35–40 phút)' 'Heading2'
Teacher 'Sản phẩm chúng ta sẽ cùng tìm hiểu và sau này cùng chế tạo có tên EcoMicro-Class. Đây là mô hình một lớp học nhỏ. Nó có bộ phận để nhận biết môi trường, một bộ phận để quyết định và các thiết bị để thực hiện quyết định.'
Table @(
  @('Ví von','Trong mô hình','Chức năng'),
  @('Giác quan','Cảm biến','Nhận biết sáng, tối, nóng, ẩm hoặc có người'),
  @('Bộ não','Bộ điều khiển','Nhận dữ liệu và chọn việc cần làm'),
  @('Tay chân','Rèm, đèn, quạt','Thực hiện quyết định'),
  @('Thức ăn/năng lượng','Nguồn điện và pin mặt trời','Cung cấp điện cho hệ thống')
) @(2100,3100,4000)
QA 'Bộ phận nào giống giác quan của con người?' 'Cảm biến, vì nó giúp máy nhận biết môi trường.' 'Hỏi: “Mắt giúp chúng ta biết sáng hay tối; máy dùng gì?”'
QA 'Bộ phận nào giống bộ não?' 'Bộ điều khiển, vì nó nhận thông tin rồi quyết định.' 'Chưa cần nói Arduino nếu học sinh chưa nhớ.'
QA 'Rèm, đèn và quạt giống bộ phận nào của cơ thể?' 'Giống tay chân vì chúng thực hiện việc mà bộ điều khiển yêu cầu.' 'Cho học sinh làm động tác tay kéo rèm.'
QA 'Pin mặt trời có vai trò gì?' 'Nhận ánh sáng Mặt Trời và tạo ra điện để hỗ trợ cấp năng lượng cho hệ thống.' 'Không nói rằng một tấm pin nhỏ chắc chắn đủ chạy toàn bộ hệ thống.'

P 'Hoạt động 8 – Làm quen với demo 3D (40–45 phút)' 'Heading2'
P 'Mở file 3D_Demo/index.html, nhấn F11 và giữ ở cảnh “Bình thường”. Chưa nhấn các cảnh khác.'
Teacher 'Đây là lớp học mô phỏng trên máy tính. Chúng ta nhìn thấy cửa sổ, rèm, đèn, quạt, học sinh và tấm pin mặt trời. Các chấm sáng là vị trí cảm biến. Các con số bên dưới là dữ liệu.'
QA 'Em nhìn thấy những đồ vật nào trong lớp?' 'Cửa sổ, rèm, đèn, quạt, bàn, học sinh, cảm biến và tấm pin mặt trời.' 'Xoay chậm mô hình và chỉ từng khu vực.'
QA 'Đâu là giác quan, đâu là tay chân?' 'Các điểm cảm biến là giác quan; rèm, đèn và quạt là tay chân.' 'Cho học sinh lên chỉ trực tiếp trên màn hình.'
QA 'Con số 27°C cho chúng ta biết điều gì?' 'Cho biết nhiệt độ của phòng là 27 độ C.' 'Nhắc học sinh °C đọc là “độ C”.'
Callout 'Chỉ chuyển sang kịch bản khi cả ba học sinh đều chỉ đúng cảm biến, bộ điều khiển theo khái niệm và thiết bị thực hiện.'

P 'Hoạt động 9 – Kịch bản Bình thường (45–52 phút)' 'Heading2'
Teacher 'Mỗi lần thử, chúng ta làm bốn bước: Đoán – Xem – Nói lý do – Nghĩ thêm. Đây là lớp học bình thường. Theo em, rèm, đèn và quạt nên làm gì?'
QA 'Khi lớp đủ sáng và không nóng, rèm nên mở hay đóng?' 'Rèm có thể mở để sử dụng ánh sáng tự nhiên.' 'Chấp nhận “mở”, rồi hỏi vì sao.'
QA 'Quạt có cần bật không?' 'Không cần nếu nhiệt độ đang dễ chịu.' 'Nhấn mạnh “không cần”, không nói “không bao giờ”.'
QA 'Hệ thống không bật nhiều thiết bị có phải là bị hỏng không?' 'Không. Khi phòng đã ổn, không hoạt động thêm là cách tiết kiệm điện.' 'So sánh với tắt quạt khi đã mát.'
P 'Câu chốt cho học sinh nhắc lại:' 'Heading3'
Answer 'Khi lớp học đã dễ chịu, hệ thống giữ trạng thái phù hợp và không dùng điện một cách lãng phí.'

P 'Hoạt động 10 – Kịch bản Nắng gắt (52–61 phút)' 'Heading2'
P 'Trước khi nhấn nút “Nắng gắt”, yêu cầu mỗi em nói một dự đoán. Sau đó mới nhấn.'
Teacher 'Bây giờ nắng chiếu rất mạnh vào cửa sổ. Theo em, việc gì nên xảy ra trước?'
QA 'Rèm nên làm gì khi nắng gây chói?' 'Rèm nên đóng bớt hoặc đóng lại để giảm ánh sáng chiếu thẳng vào mắt.' 'Dùng từ “giảm chói”, không nói rèm làm phòng mát hoàn toàn.'
QA 'Sau khi rèm đóng, bàn học có thể gặp vấn đề gì?' 'Bàn học có thể bị tối.' 'Cho học sinh nhìn chỉ số ánh sáng tại bàn.'
QA 'Đèn nên làm gì khi bàn học bị tối?' 'Đèn nên sáng lên vừa đủ để bù phần ánh sáng bị thiếu.' 'Nhấn mạnh “vừa đủ”, không nhất thiết sáng 100%.'
QA 'Vì sao rèm đóng mà đèn lại bật?' 'Rèm giúp giảm chói nhưng làm phòng tối hơn; đèn bật để giữ bàn học đủ sáng.' 'Dùng hai bàn tay: một tay che nắng, tay kia tượng trưng cho đèn bù.'
QA 'Nếu nhiệt độ cũng tăng cao thì quạt làm gì?' 'Quạt bật để hỗ trợ làm mát hoặc thông gió.' 'Không khẳng định quạt làm giảm nhiệt độ phòng mạnh; nói quạt giúp con người cảm thấy mát hơn.'
P 'Câu chốt cho học sinh nhắc lại:' 'Heading3'
Answer 'Rèm giảm chói, đèn bù sáng, quạt hỗ trợ khi nóng. Các bộ phận phải phối hợp với nhau.'

P 'Hoạt động 11 – Kịch bản Mưa tối (61–68 phút)' 'Heading2'
Teacher 'Bây giờ trời đang sáng nhưng đột nhiên có mây đen và mưa. Theo em, một thời khóa biểu cố định có biết trời vừa tối đi không?'
QA 'Khi trời tối đột ngột, đèn nên làm gì?' 'Nếu trong lớp có người và bàn học thiếu sáng, đèn nên bật hoặc sáng hơn.' 'Nhắc hai điều kiện: có người và thiếu sáng.'
QA 'Rèm có cần đóng để chống nắng không?' 'Không cần; khi trời tối, rèm có thể mở để tận dụng ánh sáng còn lại.' 'Chấp nhận nếu học sinh nói cần xem tình hình thực tế.'
QA 'Vì sao cảm biến hữu ích hơn việc chỉ bật đèn theo giờ?' 'Cảm biến nhận biết tình trạng sáng/tối đang xảy ra thật, còn giờ cố định không biết thời tiết thay đổi.' 'Dùng ví dụ hai ngày cùng 3 giờ chiều nhưng ánh sáng khác nhau.'
QA 'Nếu lớp không có người thì đèn có nên bật không?' 'Thông thường không nên bật vì không có ai cần ánh sáng, như vậy sẽ lãng phí điện.' 'Nói rằng một số thiết bị an toàn có thể có quy tắc riêng, nhưng chưa cần học hôm nay.'
P 'Câu chốt cho học sinh nhắc lại:' 'Heading3'
Answer 'Cảm biến giúp hệ thống phản ứng với điều đang xảy ra thật, không chỉ làm theo một giờ cố định.'

P 'Hoạt động 12 – Ôn tập và tìm sở thích (68–75 phút)' 'Heading2'
Teacher 'Hôm nay chúng ta chưa chế tạo, nhưng đã học cách nhìn một vấn đề và hiểu một hệ thống tự động. Mỗi bạn hãy nói một điều mới mình vừa biết.'
QA 'STEM là gì theo cách hiểu của em?' 'Là dùng kiến thức để giải quyết một vấn đề thật, làm thử và kiểm tra kết quả.' 'Không yêu cầu học sinh đọc đúng bốn chữ tiếng Anh.'
QA 'Cảm biến làm việc gì?' 'Giúp máy nhận biết môi trường như sáng, tối, nóng hoặc có người.' 'Cho phép trả lời bằng ví von “giác quan của máy”.'
QA 'Bốn bước khi thử một tình huống là gì?' 'Đoán – Xem – Nói lý do – Nghĩ thêm.' 'Giáo viên chỉ bốn ngón tay để gợi nhớ.'
QA 'Em thích phần nào nhất?' 'Không có đáp án đúng. Các lựa chọn có thể là quan sát, điều khiển mô hình, lắp ráp, tìm hiểu cảm biến, viết code, đo số liệu hoặc thuyết trình.' 'Ghi lại nguyên văn sở thích của từng em, chưa phân vai cố định.'
P 'Nhiệm vụ về nhà' 'Heading3'
Bullet 'Quan sát lớp học hoặc phòng ở nhà.'
Bullet 'Ghi lại một lúc em thấy nóng, tối, chói hoặc lãng phí điện.'
Bullet 'Viết một câu: “Em muốn thiết bị tự động làm gì trong tình huống đó?”'

P '5. NGÂN HÀNG CÂU HỎI VÀ ĐÁP ÁN BUỔI ĐẦU' 'Heading1'
Table @(
  @('Câu hỏi','Câu trả lời ngắn mong đợi'),
  @('STEM là gì?','Dùng kiến thức để giải quyết vấn đề trong đời sống, làm thử và kiểm tra.'),
  @('Vấn đề là gì?','Một điều chưa tốt hoặc gây khó chịu cần được giải quyết.'),
  @('Giải pháp là gì?','Cách chúng ta nghĩ ra để xử lý vấn đề.'),
  @('Mô hình để làm gì?','Để thử ý tưởng ở quy mô nhỏ trước khi làm thật.'),
  @('Tự động nghĩa là gì?','Thiết bị tự làm khi gặp điều kiện đã được dạy trước.'),
  @('Cảm biến để làm gì?','Giúp máy nhận biết môi trường.'),
  @('Dữ liệu là gì?','Thông tin hoặc con số đo được.'),
  @('Bộ điều khiển để làm gì?','Nhận dữ liệu và quyết định thiết bị phải làm gì.'),
  @('Chương trình là gì?','Các câu lệnh con người viết để dạy máy cách xử lý.'),
  @('Vì sao cần đo đạc?','Để biết giải pháp có thật sự tốt hơn hay không.'),
  @('Vì sao không bật mọi thiết bị liên tục?','Vì sẽ lãng phí điện và thiết bị nhanh hao mòn.'),
  @('Vì sao kéo rèm rồi lại bật đèn?','Rèm giảm chói nhưng có thể làm tối; đèn bù lại phần thiếu sáng.'),
  @('Vì sao không chỉ dùng thời khóa biểu?','Vì thời tiết có thể thay đổi bất ngờ.'),
  @('PIR có biết người đó là ai không?','Không; loại cảm biến này chủ yếu phát hiện chuyển động/sự hiện diện, không nhận dạng danh tính.'),
  @('Pin mặt trời làm gì?','Chuyển ánh sáng thành điện để hỗ trợ cấp năng lượng.'),
  @('Một bạn có phải biết hết không?','Không cần làm hết, nhưng cả đội cần hiểu sản phẩm và giúp nhau.'),
  @('Nếu trả lời sai thì sao?','Sai là thông tin giúp chúng ta biết cần thử và học thêm ở đâu.'),
  @('Nếu mô hình không chạy thì sao?','Bình tĩnh kiểm tra nguồn, dây, dữ liệu và từng bộ phận; có phương án điều khiển tay.'),
  @('Hôm nay có phải học code không?','Chưa; hôm nay hiểu vấn đề và cách hệ thống hoạt động.'),
  @('Bốn bước thử tình huống là gì?','Đoán – Xem – Nói lý do – Nghĩ thêm.')
) @(3900,5300)

P '6. CÁCH XỬ LÝ CÁC CÂU TRẢ LỜI CHƯA ĐÚNG' 'Heading1'
Table @(
  @('Học sinh nói','Giáo viên không nên nói','Cách phản hồi nên dùng'),
  @('“Máy tự suy nghĩ.”','“Sai rồi.”','“Ý em là máy tự thực hiện. Nhưng trước đó con người phải dạy máy quy tắc nào?”'),
  @('“Nắng thì bật đèn.”','“Không đúng.”','“Nắng có thể làm sáng nhưng cũng gây chói. Nếu đóng rèm thì bàn học thay đổi ra sao?”'),
  @('“Cứ bật quạt cả ngày.”','“Như vậy rất tốn điện.”','“Cách đó giúp tránh nóng. Làm sao chỉ bật khi thật sự cần?”'),
  @('“Dùng camera để biết có người.”','“Không cần camera.”','“Đó là một cách. Camera có vấn đề riêng tư không? Có cảm biến đơn giản hơn không?”'),
  @('“Em không biết.”','“Em phải nghĩ đi.”','“Không sao. Em hãy chọn giữa hai khả năng: rèm mở hay đóng? Vì sao em chọn như vậy?”'),
  @('Im lặng hoặc ngại nói','Gọi trả lời liên tục','Cho 20 giây suy nghĩ, viết ra giấy hoặc nói theo cặp trước.')
) @(2500,2500,4200)
Callout 'Mẫu phản hồi chung: công nhận phần hợp lý → hỏi thêm một câu nhỏ → bổ sung từ đúng. Không biến buổi đầu thành bài kiểm tra thuộc lòng.'

P '7. PHIẾU ĐÁNH GIÁ NHANH CUỐI BUỔI' 'Heading1'
P 'Đánh dấu ✓ nếu học sinh thực hiện được; △ nếu cần gợi ý; để trống nếu chưa thực hiện.'
Table @(
  @('Tiêu chí','Học sinh 1','Học sinh 2','Học sinh 3'),
  @('Nói được STEM dùng để giải quyết vấn đề','□','□','□'),
  @('Phân biệt vấn đề và giải pháp','□','□','□'),
  @('Giải thích được cảm biến là gì','□','□','□'),
  @('Chỉ đúng bộ phận giống bộ não và tay chân','□','□','□'),
  @('Dự đoán đúng ít nhất một tình huống','□','□','□'),
  @('Nói được lý do rèm và đèn phối hợp','□','□','□'),
  @('Nhớ bốn bước thử tình huống','□','□','□'),
  @('Chủ động đặt ít nhất một câu hỏi','□','□','□')
) @(5000,1400,1400,1400)
P 'Ba câu phiếu ra về' 'Heading2'
Numbered 'Hôm nay em hiểu thêm điều gì?' 1
Numbered 'Điều gì em vẫn chưa hiểu?' 2
Numbered 'Buổi sau em muốn được tự làm phần nào?' 3

P '8. PHƯƠNG ÁN RÚT GỌN CÒN 60 PHÚT' 'Heading1'
Table @(
  @('Thời gian','Nội dung'),
  @('0–5 phút','Đồ vật tự động quen thuộc'),
  @('5–12 phút','Giải thích STEM'),
  @('12–17 phút','Giải thích ngắn về cuộc thi'),
  @('17–23 phút','Sáu từ mới'),
  @('23–28 phút','Giới thiệu bài học'),
  @('28–33 phút','Tìm vấn đề trong lớp'),
  @('33–38 phút','Giới thiệu EcoMicro-Class'),
  @('38–44 phút','Quan sát mô hình 3D'),
  @('44–50 phút','Kịch bản Bình thường'),
  @('50–57 phút','Kịch bản Nắng gắt'),
  @('57–60 phút','Nhắc lại và giao nhiệm vụ')
) @(1800,7400)
Callout 'Trong buổi 60 phút, bỏ kịch bản Mưa tối. Không rút ngắn phần giải thích khái niệm để chạy thêm demo.'

P '9. GHI CHÚ SAU BUỔI DẠY' 'Heading1'
P 'Điều học sinh hứng thú nhất: ................................................................................................................'
P 'Từ hoặc khái niệm còn khó: ...................................................................................................................'
P 'Câu hỏi bất ngờ của học sinh: ...............................................................................................................'
P 'Vai trò mỗi em muốn thử: .......................................................................................................................'
P 'Điều cần điều chỉnh cho buổi 2: .............................................................................................................'

$sect='<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1276"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
$document="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:document xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`"><w:body>$($body.ToString())$sect</w:body></w:document>"
$styles=@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:eastAsia="Arial"/><w:sz w:val="23"/><w:szCs w:val="23"/><w:lang w:val="vi-VN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="276" w:lineRule="auto" w:after="120"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="300" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F4E78"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="210" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="548235"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="140" w:after="70"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="5B9BD5"/><w:sz w:val="23"/><w:szCs w:val="23"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Question"><w:name w:val="Question"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="50"/></w:pPr><w:rPr><w:b/><w:color w:val="7030A0"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Support"><w:name w:val="Support"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420"/><w:spacing w:after="100"/></w:pPr><w:rPr><w:i/><w:color w:val="666666"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
</w:styles>
'@
$types='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
$rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
$docRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
$stamp=(Get-Date).ToUniversalTime().ToString('s')+'Z'
$core="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>Kế hoạch chi tiết buổi 1 và đáp án EcoMicro-Class</dc:title><dc:creator>EcoMicro-Class Team</dc:creator><dc:subject>Giáo dục STEM lớp 6</dc:subject><dcterms:created xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:created><dcterms:modified xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:modified></cp:coreProperties>"
$app='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>'
$utf8=New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $work '[Content_Types].xml'),$types,$utf8)
[IO.File]::WriteAllText((Join-Path $work '_rels\.rels'),$rels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'word\document.xml'),$document,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'word\styles.xml'),$styles,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'word\_rels\document.xml.rels'),$docRels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'docProps\core.xml'),$core,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'docProps\app.xml'),$app,$utf8)
$zip=[IO.Path]::ChangeExtension($output,'.zip')
if(Test-Path $zip){Remove-Item -LiteralPath $zip -Force}
if(Test-Path $output){Remove-Item -LiteralPath $output -Force}
Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zip -Force
Move-Item -LiteralPath $zip -Destination $output
Remove-Item -LiteralPath $work -Recurse -Force
Write-Output "Created: $output"

