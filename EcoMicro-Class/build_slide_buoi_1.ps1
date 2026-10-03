$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root 'Bai_thuyet_trinh_Buoi_1_EcoMicro_Class.pptx'
$preview = Join-Path $root '3D_Demo\preview.png'
$work = Join-Path $env:TEMP ('ecomicro_pptx_' + [guid]::NewGuid().ToString('N'))

$dirs = @(
  $work, (Join-Path $work '_rels'), (Join-Path $work 'docProps'),
  (Join-Path $work 'ppt'), (Join-Path $work 'ppt\_rels'),
  (Join-Path $work 'ppt\slides'), (Join-Path $work 'ppt\slides\_rels'),
  (Join-Path $work 'ppt\slideMasters'), (Join-Path $work 'ppt\slideMasters\_rels'),
  (Join-Path $work 'ppt\slideLayouts'), (Join-Path $work 'ppt\slideLayouts\_rels'),
  (Join-Path $work 'ppt\theme'), (Join-Path $work 'ppt\media')
)
New-Item -ItemType Directory -Path $dirs | Out-Null

function X([string]$v) { [System.Security.SecurityElement]::Escape($v) }
function E([double]$inches) { [int64][math]::Round($inches * 914400) }

$script:slides = New-Object System.Collections.ArrayList
$script:sp = $null
$script:shapeId = 1
$script:currentHasImage = $false
$script:slideNumber = 0

function Add-Shape {
  param(
    [double]$X,[double]$Y,[double]$W,[double]$H,
    [string]$Fill='FFFFFF',[string]$Line='none',[string]$Geometry='rect',[double]$Transparency=0
  )
  $id = ++$script:shapeId
  $fillXml = if($Fill -eq 'none'){'<a:noFill/>'}else{
    $alpha = if($Transparency -gt 0){"<a:alpha val=`"$([int](100000-$Transparency*1000))`"/>"}else{''}
    "<a:solidFill><a:srgbClr val=`"$Fill`">$alpha</a:srgbClr></a:solidFill>"
  }
  $lineXml = if($Line -eq 'none'){'<a:ln><a:noFill/></a:ln>'}else{"<a:ln w=`"12700`"><a:solidFill><a:srgbClr val=`"$Line`"/></a:solidFill></a:ln>"}
  [void]$script:sp.Append("<p:sp><p:nvSpPr><p:cNvPr id=`"$id`" name=`"Shape $id`"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"$(E $X)`" y=`"$(E $Y)`"/><a:ext cx=`"$(E $W)`" cy=`"$(E $H)`"/></a:xfrm><a:prstGeom prst=`"$Geometry`"><a:avLst/></a:prstGeom>$fillXml$lineXml</p:spPr></p:sp>")
}

function Add-Text {
  param(
    [double]$X,[double]$Y,[double]$W,[double]$H,[string]$Text,
    [int]$Size=24,[string]$Color='FFFFFF',[switch]$Bold,
    [ValidateSet('l','ctr','r')][string]$Align='l',
    [ValidateSet('t','ctr','b')][string]$VAlign='ctr',
    [string]$Fill='none',[string]$Line='none',[string]$Geometry='rect',
    [double]$Margin=0.10
  )
  $Text = $Text -replace '\\n', "`n"
  $id = ++$script:shapeId
  $fillXml = if($Fill -eq 'none'){'<a:noFill/>'}else{"<a:solidFill><a:srgbClr val=`"$Fill`"/></a:solidFill>"}
  $lineXml = if($Line -eq 'none'){'<a:ln><a:noFill/></a:ln>'}else{"<a:ln w=`"12700`"><a:solidFill><a:srgbClr val=`"$Line`"/></a:solidFill></a:ln>"}
  $boldVal = if($Bold){'1'}else{'0'}
  $paragraphs = New-Object System.Text.StringBuilder
  foreach($line in ($Text -split "`n")){
    [void]$paragraphs.Append("<a:p><a:pPr algn=`"$Align`"/><a:r><a:rPr lang=`"vi-VN`" sz=`"$($Size*100)`" b=`"$boldVal`" dirty=`"0`"><a:solidFill><a:srgbClr val=`"$Color`"/></a:solidFill><a:latin typeface=`"Aptos`"/><a:ea typeface=`"Arial`"/></a:rPr><a:t>$(X $line)</a:t></a:r><a:endParaRPr lang=`"vi-VN`" sz=`"$($Size*100)`"/></a:p>")
  }
  $marginEmu = E $Margin
  [void]$script:sp.Append("<p:sp><p:nvSpPr><p:cNvPr id=`"$id`" name=`"Text $id`"/><p:cNvSpPr txBox=`"1`"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=`"$(E $X)`" y=`"$(E $Y)`"/><a:ext cx=`"$(E $W)`" cy=`"$(E $H)`"/></a:xfrm><a:prstGeom prst=`"$Geometry`"><a:avLst/></a:prstGeom>$fillXml$lineXml</p:spPr><p:txBody><a:bodyPr wrap=`"square`" lIns=`"$marginEmu`" rIns=`"$marginEmu`" tIns=`"$marginEmu`" bIns=`"$marginEmu`" anchor=`"$VAlign`"/><a:lstStyle/>$($paragraphs.ToString())</p:txBody></p:sp>")
}

function Add-Card {
  param([double]$X,[double]$Y,[double]$W,[double]$H,[string]$Title,[string]$Body,[string]$Accent='63E6A4',[int]$TitleSize=22,[int]$BodySize=16)
  Add-Shape $X $Y $W $H '13283B' '29465D' 'roundRect'
  Add-Shape ($X+0.16) ($Y+0.18) 0.07 ($H-0.36) $Accent 'none' 'roundRect'
  Add-Text ($X+0.38) ($Y+0.18) ($W-0.56) 0.46 $Title $TitleSize $Accent -Bold -VAlign t
  Add-Text ($X+0.38) ($Y+0.68) ($W-0.58) ($H-0.82) $Body $BodySize 'D9EAF3' -VAlign t
}

function Add-Header([string]$Title,[string]$Section='BUỔI 1') {
  Add-Shape 0 0 13.333 7.5 '071522' 'none'
  Add-Shape 0 0 13.333 0.12 '63E6A4' 'none'
  Add-Text 0.55 0.32 10.7 0.55 $Title 28 'F4FAFD' -Bold
  Add-Text 11.45 0.35 1.25 0.38 $Section 11 '63E6A4' -Bold -Align ctr -Fill '123247' -Line '2B6076' -Geometry roundRect
  Add-Text 12.65 7.10 0.35 0.2 ([string]$script:slideNumber) 10 '6E8A9A' -Align r
}

function Add-Image {
  param([double]$X,[double]$Y,[double]$W,[double]$H)
  $id = ++$script:shapeId
  $script:currentHasImage = $true
  [void]$script:sp.Append("<p:pic><p:nvPicPr><p:cNvPr id=`"$id`" name=`"Demo 3D`"/><p:cNvPicPr><a:picLocks noChangeAspect=`"1`"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=`"rId2`"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=`"$(E $X)`" y=`"$(E $Y)`"/><a:ext cx=`"$(E $W)`" cy=`"$(E $H)`"/></a:xfrm><a:prstGeom prst=`"roundRect`"><a:avLst/></a:prstGeom><a:ln w=`"19050`"><a:solidFill><a:srgbClr val=`"4FA5BC`"/></a:solidFill></a:ln></p:spPr></p:pic>")
}

function New-Slide {
  param([string]$Title,[scriptblock]$Build,[string]$Section='BUỔI 1',[switch]$NoHeader)
  $script:slideNumber++
  $script:shapeId=1
  $script:currentHasImage=$false
  $script:sp=New-Object System.Text.StringBuilder
  if(-not $NoHeader){ Add-Header $Title $Section }
  & $Build
  $group='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
  $xml="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sld xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"><p:cSld><p:spTree>$group$($script:sp.ToString())</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
  [void]$script:slides.Add([pscustomobject]@{Xml=$xml; HasImage=$script:currentHasImage})
}

# Slide 1
New-Slide '' {
  Add-Shape 0 0 13.333 7.5 '071522' 'none'
  Add-Shape 0 0 0.18 7.5 '63E6A4' 'none'
  Add-Shape 9.75 -0.4 4.2 4.2 '123247' 'none' 'ellipse'
  Add-Shape 10.55 0.35 2.6 2.6 '1A4455' 'none' 'ellipse'
  Add-Text 0.78 1.25 8.8 0.55 'BUỔI 1' 19 '63E6A4' -Bold
  Add-Text 0.72 1.85 10.7 1.35 'LÀM QUEN VỚI STEM' 42 'F5FBFE' -Bold -VAlign t
  Add-Text 0.72 3.12 10.4 0.82 'VÀ ECOMICRO-CLASS' 36 '54D9FF' -Bold -VAlign t
  Add-Text 0.78 4.20 8.7 0.85 'Từ những đồ vật tự động quen thuộc đến một lớp học biết quan sát và tự điều chỉnh' 22 'BED4DF' -VAlign t
  Add-Text 0.78 6.45 4.7 0.38 'Dành cho học sinh lớp 6 – bắt đầu từ số 0' 13 '8FA9B8'
} -NoHeader

# Slide 2
New-Slide 'Hôm nay chúng ta sẽ làm gì?' {
  Add-Card 0.65 1.25 2.85 2.05 '1  HIỂU' 'STEM là gì và cuộc thi có những thử thách nào?' '63E6A4'
  Add-Card 3.75 1.25 2.85 2.05 '2  TÌM' 'Những điều chưa tiện trong lớp học hằng ngày.' '54D9FF'
  Add-Card 6.85 1.25 2.85 2.05 '3  KHÁM PHÁ' 'Một lớp học có giác quan, bộ não và tay chân.' 'FFD166'
  Add-Card 9.95 1.25 2.75 2.05 '4  THỬ' 'Đoán, quan sát và giải thích các tình huống.' 'FF9F43'
  Add-Text 1.05 4.10 11.25 1.2 'Chưa cần biết điện tử. Chưa cần biết lập trình. Chúng ta sẽ học từng bước.' 28 'F4FAFD' -Bold -Align ctr -Fill '102B3D' -Line '244B61' -Geometry roundRect
  Add-Text 1.3 5.70 10.8 0.55 'Mục tiêu hôm nay: hiểu trước – làm sau.' 22 '63E6A4' -Bold -Align ctr
}

# Slide 3
New-Slide 'Khởi động: đồ vật nào biết “tự làm”?' {
  Add-Card 0.75 1.35 3.65 2.35 'CỬA TỰ ĐỘNG' 'Có người đến gần → cửa mở.' '54D9FF' 23 19
  Add-Card 4.85 1.35 3.65 2.35 'ĐÈN CẢM ỨNG' 'Có người đi qua → đèn sáng.' 'FFD166' 23 19
  Add-Card 8.95 1.35 3.65 2.35 'VÒI NƯỚC' 'Đưa tay vào → nước chảy.' '63E6A4' 23 19
  Add-Text 1.0 4.45 11.3 1.1 'CÂU HỎI: Làm sao những đồ vật này biết lúc nào cần hoạt động?' 27 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Line '39718A' -Geometry roundRect
  Add-Text 2.0 5.88 9.3 0.45 'Không biết cũng không sao — cuối buổi chúng ta sẽ trả lời được.' 18 'A9C2D0' -Align ctr
}

# Slide 4
New-Slide 'STEM là gì?' {
  Add-Text 0.8 1.08 11.75 0.68 'STEM là cách dùng kiến thức để giải quyết một việc trong đời sống.' 27 'FFFFFF' -Bold -Align ctr
  Add-Card 0.65 2.0 2.85 2.5 'KHOA HỌC' 'Tìm hiểu vì sao sự việc xảy ra.\n\nVì sao có nắng thì phòng nóng hơn?' '54D9FF' 22 16
  Add-Card 3.75 2.0 2.85 2.5 'CÔNG NGHỆ' 'Dùng công cụ để giúp con người.\n\nDùng bộ phận nhận biết sáng hay tối.' '63E6A4' 22 16
  Add-Card 6.85 2.0 2.85 2.5 'KỸ THUẬT' 'Nghĩ cách rồi làm thử.\n\nLàm rèm có thể tự kéo.' 'FFD166' 22 16
  Add-Card 9.95 2.0 2.75 2.5 'TOÁN HỌC' 'Đo, đếm và so sánh.\n\nSo sánh nhiệt độ trước và sau.' 'FF9F43' 22 16
  Add-Text 1.3 5.25 10.7 0.9 'STEM không phải là học thuộc bốn chữ.' 26 '63E6A4' -Bold -Align ctr -Fill '102B3D' -Geometry roundRect
}

# Slide 5
New-Slide 'STEM giải quyết vấn đề như thế nào?' {
  $xs=@(0.55,3.72,6.89,10.06); $titles=@('1  THẤY','2  NGHĨ','3  LÀM THỬ','4  KIỂM TRA'); $bodies=@('Lớp học quá nóng.','Nghĩ cách làm mát.','Tạo một mô hình nhỏ.','Đo xem có tốt hơn không.'); $colors=@('FF6B6B','FFD166','54D9FF','63E6A4')
  for($i=0;$i -lt 4;$i++){ Add-Card $xs[$i] 2.0 2.72 2.25 $titles[$i] $bodies[$i] $colors[$i] 23 18; if($i -lt 3){Add-Text ($xs[$i]+2.73) 2.72 0.43 0.55 '→' 28 '8FA9B8' -Bold -Align ctr} }
  Add-Text 1.0 5.15 11.3 0.85 'Làm xong chưa phải là hết — cần đo và kiểm tra.' 27 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Line '39718A' -Geometry roundRect
}

# Slide 6
New-Slide 'Cuộc thi STEM có những gì?' {
  Add-Text 0.8 1.0 11.7 0.55 'Cuộc thi là một chuỗi thử thách — không phải một bài kiểm tra duy nhất.' 23 'BFD4DF' -Align ctr
  Add-Card 0.85 1.85 3.55 3.2 'VÒNG 1' 'Mỗi bạn trả lời câu hỏi.\n\nMục tiêu: thể hiện cách suy nghĩ và vận dụng kiến thức.' '54D9FF' 29 18
  Add-Card 4.88 1.85 3.55 3.2 'VÒNG 2' 'Cả đội làm sản phẩm, giới thiệu và trả lời câu hỏi.' '63E6A4' 29 18
  Add-Card 8.90 1.85 3.55 3.2 'VÒNG 3' 'Nhận đề bài và vật liệu mới.\n\nCùng làm mô hình trong 90 phút.' 'FFD166' 29 18
  Add-Text 1.35 5.70 10.65 0.72 'Không cần một bạn biết hết — cả đội học cách giúp nhau.' 25 'FFFFFF' -Bold -Align ctr
}

# Slide 7
New-Slide 'Một sản phẩm đẹp đã đủ chưa?' {
  Add-Text 0.75 1.15 5.7 1.05 'CHƯA ĐỦ' 42 'FF6B6B' -Bold -Align ctr -Fill '2A2028' -Line '6E3B49' -Geometry roundRect
  Add-Text 6.85 1.15 5.7 1.05 'CẦN THÊM' 42 '63E6A4' -Bold -Align ctr -Fill '102D2A' -Line '2C6A57' -Geometry roundRect
  Add-Card 0.75 2.55 5.7 2.9 'CHỈ ĐẸP' '• Không rõ giải quyết việc gì\n• Không biết có hoạt động không\n• Không có số liệu\n• Không giải thích được' 'FF6B6B' 24 20
  Add-Card 6.85 2.55 5.7 2.9 'SẢN PHẨM STEM' '• Có vấn đề rõ ràng\n• Hoạt động được\n• Có thử nghiệm và số liệu\n• Cả đội hiểu và giải thích được' '63E6A4' 24 20
  Add-Text 2.0 6.05 9.3 0.48 'Ban Giám khảo sẽ hỏi: “Vì sao nhóm chọn cách này?”' 21 'FFD166' -Bold -Align ctr
}

# Slide 8
New-Slide 'Ba từ đầu tiên' {
  Add-Card 0.75 1.45 3.75 3.5 'VẤN ĐỀ' 'Một điều chưa tốt hoặc gây khó chịu.\n\nVí dụ: nắng chiếu vào mắt.' 'FF6B6B' 28 20
  Add-Card 4.8 1.45 3.75 3.5 'GIẢI PHÁP' 'Cách nghĩ ra để xử lý vấn đề.\n\nVí dụ: đóng rèm.' '63E6A4' 28 20
  Add-Card 8.85 1.45 3.75 3.5 'MÔ HÌNH' 'Phiên bản nhỏ để làm thử trước khi áp dụng thật.\n\nVí dụ: lớp học 3D.' '54D9FF' 28 20
  Add-Text 1.5 5.60 10.3 0.72 'CÂU HỎI: “Lớp quá nóng” là vấn đề hay giải pháp?' 24 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Geometry roundRect
}

# Slide 9
New-Slide 'Ba từ tiếp theo' {
  Add-Card 0.75 1.45 3.75 3.5 'TỰ ĐỘNG' 'Thiết bị tự làm việc khi gặp điều kiện đã được dạy trước.\n\nTrời tối → bật đèn.' 'FFD166' 28 19
  Add-Card 4.8 1.45 3.75 3.5 'CẢM BIẾN' 'Bộ phận giúp máy nhận biết môi trường.\n\nBiết sáng, nóng hoặc có người.' '54D9FF' 28 19
  Add-Card 8.85 1.45 3.75 3.5 'DỮ LIỆU' 'Thông tin hoặc con số đo được.\n\nVí dụ: nhiệt độ 30°C.' '63E6A4' 28 19
  Add-Text 1.4 5.62 10.5 0.72 'Máy không tự biết mọi thứ — con người phải dạy máy quy tắc.' 24 'FFFFFF' -Bold -Align ctr
}

# Slide 10
New-Slide 'Bài học hôm nay' {
  Add-Text 1.0 1.25 11.35 1.15 'LỚP HỌC BIẾT QUAN SÁT\nVÀ TỰ ĐIỀU CHỈNH' 38 '54D9FF' -Bold -Align ctr -VAlign ctr
  Add-Card 1.05 3.15 3.35 2.0 'MỤC TIÊU 1' 'Kể được một vấn đề trong lớp học.' '63E6A4' 21 18
  Add-Card 4.98 3.15 3.35 2.0 'MỤC TIÊU 2' 'Chỉ ra giác quan, bộ não và tay chân của mô hình.' 'FFD166' 21 18
  Add-Card 8.90 3.15 3.35 2.0 'MỤC TIÊU 3' 'Dự đoán thiết bị nào nên bật hoặc tắt.' 'FF9F43' 21 18
  Add-Text 2.1 5.90 9.2 0.55 'Hôm nay chưa cần viết code.' 25 '63E6A4' -Bold -Align ctr
}

# Slide 11
New-Slide 'Điều gì chưa tiện trong lớp học?' {
  Add-Card 0.70 1.25 3.78 1.65 'NẮNG CHÓI' 'Ánh sáng chiếu thẳng vào mắt.' 'FFD166' 22 17
  Add-Card 4.78 1.25 3.78 1.65 'QUÁ TỐI' 'Khó đọc sách và nhìn bảng.' '54D9FF' 22 17
  Add-Card 8.86 1.25 3.78 1.65 'QUÁ NÓNG' 'Khó tập trung học.' 'FF6B6B' 22 17
  Add-Card 0.70 3.20 3.78 1.65 'QUÁ ẨM' 'Không khí khó chịu, đồ vật dễ ẩm.' '63E6A4' 22 17
  Add-Card 4.78 3.20 3.78 1.65 'QUÊN TẮT ĐIỆN' 'Lãng phí khi không có người.' 'FF9F43' 22 17
  Add-Card 8.86 3.20 3.78 1.65 'THỜI TIẾT ĐỔI' 'Con người phải chỉnh thiết bị liên tục.' 'B28DFF' 22 17
  Add-Text 1.2 5.55 10.9 0.78 'Trong lớp của em, điều nào thường xảy ra nhất?' 26 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Geometry roundRect
}

# Slide 12
New-Slide 'EcoMicro-Class: lớp học có “cơ thể”' {
  Add-Card 0.65 1.4 2.85 3.75 'GIÁC QUAN' 'CẢM BIẾN\n\nNhận biết sáng, tối, nóng, ẩm hoặc có người.' '54D9FF' 24 18
  Add-Text 3.52 2.75 0.35 0.55 '→' 28 '8FA9B8' -Bold -Align ctr
  Add-Card 3.88 1.4 2.85 3.75 'BỘ NÃO' 'BỘ ĐIỀU KHIỂN\n\nNhận dữ liệu và chọn việc cần làm.' '63E6A4' 24 18
  Add-Text 6.75 2.75 0.35 0.55 '→' 28 '8FA9B8' -Bold -Align ctr
  Add-Card 7.10 1.4 2.85 3.75 'TAY CHÂN' 'RÈM · ĐÈN · QUẠT\n\nThực hiện quyết định.' 'FFD166' 24 18
  Add-Card 10.35 1.4 2.35 3.75 'NĂNG LƯỢNG' 'NGUỒN ĐIỆN\nPIN MẶT TRỜI\n\nGiúp hệ thống hoạt động.' 'FF9F43' 22 17
  Add-Text 1.2 5.78 10.9 0.55 'Cảm biến → Bộ điều khiển → Thiết bị thực hiện' 25 'FFFFFF' -Bold -Align ctr
}

# Slide 13
New-Slide 'Hệ thống làm việc theo một chuỗi' {
  Add-Text 0.7 1.2 2.65 1.0 '1\nNHẬN BIẾT' 27 '54D9FF' -Bold -Align ctr -Fill '13283B' -Line '2D6178' -Geometry roundRect
  Add-Text 3.52 1.48 0.55 0.48 '→' 32 '8FA9B8' -Bold -Align ctr
  Add-Text 4.20 1.2 2.65 1.0 '2\nSO SÁNH' 27 '63E6A4' -Bold -Align ctr -Fill '13283B' -Line '2D6A57' -Geometry roundRect
  Add-Text 7.02 1.48 0.55 0.48 '→' 32 '8FA9B8' -Bold -Align ctr
  Add-Text 7.70 1.2 2.65 1.0 '3\nQUYẾT ĐỊNH' 27 'FFD166' -Bold -Align ctr -Fill '13283B' -Line '7A6A37' -Geometry roundRect
  Add-Text 10.52 1.48 0.55 0.48 '→' 32 '8FA9B8' -Bold -Align ctr
  Add-Text 11.20 1.2 1.55 1.0 '4\nLÀM' 27 'FF9F43' -Bold -Align ctr -Fill '13283B' -Line '7A4F32' -Geometry roundRect
  Add-Text 0.9 2.85 11.55 0.9 'Ví dụ: cảm biến thấy trời tối → so với mức cần thiết → quyết định bật đèn → đèn sáng.' 24 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Geometry roundRect
  Add-Text 2.0 4.35 9.4 0.65 'Con người viết các quy tắc để dạy máy.' 28 '63E6A4' -Bold -Align ctr
  Add-Text 2.3 5.35 8.8 0.6 'Máy không tự suy nghĩ giống con người.' 21 'BFD4DF' -Align ctr
}

# Slide 14
New-Slide 'Mô hình 3D chúng ta sẽ khám phá' {
  Add-Image 0.62 1.1 8.55 5.6
  Add-Card 9.48 1.15 3.18 1.15 'HÃY TÌM' 'Cửa sổ và rèm' '54D9FF' 20 16
  Add-Card 9.48 2.55 3.18 1.15 'HÃY TÌM' 'Đèn và quạt' 'FFD166' 20 16
  Add-Card 9.48 3.95 3.18 1.15 'HÃY TÌM' 'Các điểm cảm biến' '63E6A4' 20 16
  Add-Card 9.48 5.35 3.18 1.15 'HÃY TÌM' 'Tấm pin mặt trời' 'FF9F43' 20 16
}

# Slide 15
New-Slide 'Cách thử một tình huống' {
  $xs=@(0.55,3.72,6.89,10.06); $titles=@('1  ĐOÁN','2  XEM','3  NÓI LÝ DO','4  NGHĨ THÊM'); $bodies=@('Theo em, điều gì sắp xảy ra?','Bộ phận nào đã thay đổi?','Vì sao nó lại làm như vậy?','Em muốn thay đổi điều gì?'); $colors=@('FFD166','54D9FF','63E6A4','FF9F43')
  for($i=0;$i -lt 4;$i++){ Add-Card $xs[$i] 1.65 2.72 2.65 $titles[$i] $bodies[$i] $colors[$i] 23 18; if($i -lt 3){Add-Text ($xs[$i]+2.73) 2.65 0.43 0.55 '→' 28 '8FA9B8' -Bold -Align ctr} }
  Add-Text 1.25 5.25 10.85 0.95 'Luôn đoán trước khi nhấn nút.' 31 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Line '39718A' -Geometry roundRect
}

# Slide 16
New-Slide 'Tình huống 1: Phòng học bình thường' {
  Add-Card 0.75 1.35 3.7 3.65 'ĐIỀU KIỆN' '• Đủ ánh sáng\n• Nhiệt độ dễ chịu\n• Có học sinh trong lớp' '54D9FF' 25 20
  Add-Card 4.82 1.35 3.7 3.65 'HỆ THỐNG' '• Rèm mở\n• Đèn tắt hoặc sáng rất ít\n• Quạt tắt' '63E6A4' 25 20
  Add-Card 8.88 1.35 3.7 3.65 'VÌ SAO?' 'Phòng đã dễ chịu nên không cần dùng nhiều điện.' 'FFD166' 25 20
  Add-Text 1.15 5.65 11.0 0.75 'Không hoạt động thêm cũng có thể là một quyết định thông minh.' 23 'FFFFFF' -Bold -Align ctr
}

# Slide 17
New-Slide 'Tình huống 2: Nắng gắt' {
  Add-Text 0.7 1.18 3.0 0.8 'NẮNG GẮT' 30 'FFD166' -Bold -Align ctr -Fill '3A3420' -Geometry roundRect
  Add-Text 3.82 1.35 0.55 0.48 '→' 32 '8FA9B8' -Bold -Align ctr
  Add-Text 4.5 1.18 3.0 0.8 'RÈM ĐÓNG' 30 '54D9FF' -Bold -Align ctr -Fill '173243' -Geometry roundRect
  Add-Text 7.62 1.35 0.55 0.48 '→' 32 '8FA9B8' -Bold -Align ctr
  Add-Text 8.3 1.18 4.25 0.8 'BÀN CÓ THỂ TỐI' 30 'FF9F43' -Bold -Align ctr -Fill '3A2922' -Geometry roundRect
  Add-Card 1.05 2.65 5.25 2.45 'RÈM GIÚP GÌ?' 'Giảm ánh sáng chiếu thẳng vào mắt và giảm chói.' '54D9FF' 24 20
  Add-Card 7.02 2.65 5.25 2.45 'ĐÈN GIÚP GÌ?' 'Bù lại phần ánh sáng bị thiếu tại bàn học.' '63E6A4' 24 20
  Add-Text 1.0 5.65 11.3 0.72 'Rèm và đèn phải phối hợp — không làm việc riêng lẻ.' 25 'FFFFFF' -Bold -Align ctr
}

# Slide 18
New-Slide 'Tình huống 3: Mưa tối bất ngờ' {
  Add-Card 0.75 1.35 3.7 3.65 'ĐIỀU GÌ XẢY RA?' 'Mây che Mặt Trời. Lớp đang sáng bỗng tối đi.' '8FA9B8' 24 20
  Add-Card 4.82 1.35 3.7 3.65 'HỆ THỐNG LÀM GÌ?' 'Nếu có người và bàn thiếu sáng, đèn sáng hơn.' '54D9FF' 24 20
  Add-Card 8.88 1.35 3.7 3.65 'VÌ SAO HỮU ÍCH?' 'Cảm biến biết điều đang xảy ra thật, không chỉ làm theo giờ.' '63E6A4' 24 20
  Add-Text 0.95 5.65 11.45 0.74 'Cùng 3 giờ chiều: hôm nay có thể nắng, ngày mai có thể mưa.' 24 'FFD166' -Bold -Align ctr
}

# Slide 19
New-Slide 'Thử tài nhanh' {
  Add-Text 0.85 1.10 11.7 0.55 'Hãy trả lời bằng lời của em — không cần đọc thuộc lòng.' 21 'AFC7D4' -Align ctr
  Add-Card 0.75 1.95 3.75 1.25 'CÂU 1' 'STEM giúp chúng ta làm gì?' '54D9FF' 20 17
  Add-Card 4.80 1.95 3.75 1.25 'CÂU 2' 'Cảm biến giúp máy làm gì?' '63E6A4' 20 17
  Add-Card 8.85 1.95 3.75 1.25 'CÂU 3' 'Dữ liệu là gì?' 'FFD166' 20 17
  Add-Card 1.80 3.70 4.55 1.35 'CÂU 4' 'Vì sao rèm đóng nhưng đèn lại sáng?' 'FF9F43' 20 17
  Add-Card 7.00 3.70 4.55 1.35 'CÂU 5' 'Bốn bước thử tình huống là gì?' 'B28DFF' 20 17
  Add-Text 2.0 5.80 9.3 0.58 'Suy nghĩ 20 giây trước khi trả lời.' 23 'FFFFFF' -Bold -Align ctr
}

# Slide 20
New-Slide 'Đáp án gợi ý' {
  Add-Card 0.68 1.12 3.86 1.45 '1  STEM' 'Dùng kiến thức để giải quyết vấn đề, làm thử và kiểm tra.' '54D9FF' 19 15
  Add-Card 4.74 1.12 3.86 1.45 '2  CẢM BIẾN' 'Giúp máy nhận biết môi trường.' '63E6A4' 19 15
  Add-Card 8.80 1.12 3.86 1.45 '3  DỮ LIỆU' 'Thông tin hoặc con số đo được.' 'FFD166' 19 15
  Add-Card 1.55 3.08 4.85 1.65 '4  RÈM VÀ ĐÈN' 'Rèm giảm chói nhưng làm tối; đèn bù phần thiếu sáng.' 'FF9F43' 19 16
  Add-Card 6.92 3.08 4.85 1.65 '5  BỐN BƯỚC' 'Đoán → Xem → Nói lý do → Nghĩ thêm.' 'B28DFF' 19 16
  Add-Text 1.2 5.55 10.95 0.82 'Có thể dùng từ khác — miễn là ý đúng.' 27 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Geometry roundRect
}

# Slide 21
New-Slide 'Kết thúc buổi 1' {
  Add-Text 1.0 1.05 11.35 0.75 'HÔM NAY EM ĐÃ BIẾT…' 31 '63E6A4' -Bold -Align ctr
  Add-Card 0.75 2.0 3.75 2.25 'MỘT' 'STEM bắt đầu từ một vấn đề thật.' '54D9FF' 23 18
  Add-Card 4.80 2.0 3.75 2.25 'HAI' 'Máy cần cảm biến, bộ điều khiển và thiết bị thực hiện.' 'FFD166' 23 18
  Add-Card 8.85 2.0 3.75 2.25 'BA' 'Trước khi thử: Đoán – Xem – Nói lý do – Nghĩ thêm.' 'FF9F43' 23 18
  Add-Text 1.0 4.85 11.35 0.72 'NHIỆM VỤ Ở NHÀ' 25 'FFFFFF' -Bold -Align ctr -Fill '15354A' -Geometry roundRect
  Add-Text 1.2 5.72 10.95 0.75 'Tìm một lúc em thấy nóng, tối, chói hoặc lãng phí điện. Viết một câu: “Em muốn thiết bị tự động làm gì?”' 20 'C9DCE6' -Align ctr
}

# Package parts
$utf8 = New-Object System.Text.UTF8Encoding($false)
for($i=0;$i -lt $slides.Count;$i++){
  $n=$i+1
  [IO.File]::WriteAllText((Join-Path $work "ppt\slides\slide$n.xml"),$slides[$i].Xml,$utf8)
  $relImage = if($slides[$i].HasImage){'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/preview.png"/>'}else{''}
  $slideRels="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout`" Target=`"../slideLayouts/slideLayout1.xml`"/>$relImage</Relationships>"
  [IO.File]::WriteAllText((Join-Path $work "ppt\slides\_rels\slide$n.xml.rels"),$slideRels,$utf8)
}
if(Test-Path $preview){Copy-Item -LiteralPath $preview -Destination (Join-Path $work 'ppt\media\preview.png')}

$slideIds=New-Object System.Text.StringBuilder
$presRels=New-Object System.Text.StringBuilder
[void]$presRels.Append('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
for($i=0;$i -lt $slides.Count;$i++){
  $n=$i+1; $rid=$n+1; $sid=255+$n
  [void]$slideIds.Append("<p:sldId id=`"$sid`" r:id=`"rId$rid`"/>")
  [void]$presRels.Append("<Relationship Id=`"rId$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$n.xml`"/>")
}
$presentation="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:presentation xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"><p:sldMasterIdLst><p:sldMasterId id=`"2147483648`" r:id=`"rId1`"/></p:sldMasterIdLst><p:sldIdLst>$($slideIds.ToString())</p:sldIdLst><p:sldSz cx=`"12192000`" cy=`"6858000`" type=`"screen16x9`"/><p:notesSz cx=`"6858000`" cy=`"9144000`"/><p:defaultTextStyle><a:defPPr><a:defRPr lang=`"vi-VN`"/></a:defPPr></p:defaultTextStyle></p:presentation>"
$presentationRels="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$($presRels.ToString())</Relationships>"

$group='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
$master="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sldMaster xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"><p:cSld><p:spTree>$group</p:spTree></p:cSld><p:clrMap accent1=`"63E6A4`" accent2=`"54D9FF`" accent3=`"FFD166`" accent4=`"FF9F43`" accent5=`"B28DFF`" accent6=`"FF6B6B`" bg1=`"071522`" bg2=`"13283B`" folHlink=`"B28DFF`" hlink=`"54D9FF`" tx1=`"FFFFFF`" tx2=`"BFD4DF`"/><p:sldLayoutIdLst><p:sldLayoutId id=`"1`" r:id=`"rId1`"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz=`"3200`" b=`"1`"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz=`"2000`"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang=`"vi-VN`"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>"
$masterRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>'
$layout="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><p:sldLayout xmlns:a=`"http://schemas.openxmlformats.org/drawingml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`" type=`"blank`" preserve=`"1`"><p:cSld name=`"Blank`"><p:spTree>$group</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
$layoutRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>'
$theme=@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="EcoMicro Theme"><a:themeElements><a:clrScheme name="EcoMicro"><a:dk1><a:srgbClr val="071522"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="13283B"/></a:dk2><a:lt2><a:srgbClr val="BFD4DF"/></a:lt2><a:accent1><a:srgbClr val="63E6A4"/></a:accent1><a:accent2><a:srgbClr val="54D9FF"/></a:accent2><a:accent3><a:srgbClr val="FFD166"/></a:accent3><a:accent4><a:srgbClr val="FF9F43"/></a:accent4><a:accent5><a:srgbClr val="B28DFF"/></a:accent5><a:accent6><a:srgbClr val="FF6B6B"/></a:accent6><a:hlink><a:srgbClr val="54D9FF"/></a:hlink><a:folHlink><a:srgbClr val="B28DFF"/></a:folHlink></a:clrScheme><a:fontScheme name="Aptos"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme><a:fmtScheme name="EcoMicro"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:noFill/></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>
'@

$contentOverrides=New-Object System.Text.StringBuilder
for($i=1;$i -le $slides.Count;$i++){[void]$contentOverrides.Append("<Override PartName=`"/ppt/slides/slide$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>")}
$types="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Default Extension=`"png`" ContentType=`"image/png`"/><Override PartName=`"/ppt/presentation.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml`"/><Override PartName=`"/ppt/slideMasters/slideMaster1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml`"/><Override PartName=`"/ppt/slideLayouts/slideLayout1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml`"/><Override PartName=`"/ppt/theme/theme1.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.theme+xml`"/>$($contentOverrides.ToString())<Override PartName=`"/docProps/core.xml`" ContentType=`"application/vnd.openxmlformats-package.core-properties+xml`"/><Override PartName=`"/docProps/app.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.extended-properties+xml`"/></Types>"
$rootRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
$stamp=(Get-Date).ToUniversalTime().ToString('s')+'Z'
$core="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><cp:coreProperties xmlns:cp=`"http://schemas.openxmlformats.org/package/2006/metadata/core-properties`" xmlns:dc=`"http://purl.org/dc/elements/1.1/`" xmlns:dcterms=`"http://purl.org/dc/terms/`" xmlns:xsi=`"http://www.w3.org/2001/XMLSchema-instance`"><dc:title>Bài thuyết trình Buổi 1 EcoMicro-Class</dc:title><dc:creator>EcoMicro-Class Team</dc:creator><dc:subject>Làm quen với STEM cho học sinh lớp 6</dc:subject><dcterms:created xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:created><dcterms:modified xsi:type=`"dcterms:W3CDTF`">$stamp</dcterms:modified></cp:coreProperties>"
$app="<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Properties xmlns=`"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties`" xmlns:vt=`"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes`"><Application>Microsoft Office PowerPoint</Application><PresentationFormat>On-screen Show (16:9)</PresentationFormat><Slides>$($slides.Count)</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>"

[IO.File]::WriteAllText((Join-Path $work '[Content_Types].xml'),$types,$utf8)
[IO.File]::WriteAllText((Join-Path $work '_rels\.rels'),$rootRels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\presentation.xml'),$presentation,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\_rels\presentation.xml.rels'),$presentationRels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\slideMasters\slideMaster1.xml'),$master,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\slideMasters\_rels\slideMaster1.xml.rels'),$masterRels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\slideLayouts\slideLayout1.xml'),$layout,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\slideLayouts\_rels\slideLayout1.xml.rels'),$layoutRels,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'ppt\theme\theme1.xml'),$theme,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'docProps\core.xml'),$core,$utf8)
[IO.File]::WriteAllText((Join-Path $work 'docProps\app.xml'),$app,$utf8)

$zip=[IO.Path]::ChangeExtension($output,'.zip')
if(Test-Path $zip){Remove-Item -LiteralPath $zip -Force}
if(Test-Path $output){Remove-Item -LiteralPath $output -Force}
Compress-Archive -Path (Join-Path $work '*') -DestinationPath $zip -Force
Move-Item -LiteralPath $zip -Destination $output
Remove-Item -LiteralPath $work -Recurse -Force
Write-Output "Created: $output ($($slides.Count) slides)"


