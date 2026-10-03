# EcoMicro-Class – hướng dẫn chạy code demo

## 1. Đấu dây Arduino Uno/Nano

| Thiết bị | Chân thiết bị | Chân Arduino |
|---|---|---|
| LDR ngoài trời | Ngõ ra cầu chia áp | A0 |
| LDR tại bàn | Ngõ ra cầu chia áp | A1 |
| DHT22 | DATA | D2 |
| PIR HC-SR501 | OUT | D3 |
| LED qua MOSFET | IN/GATE | D5 (PWM) |
| Quạt qua MOSFET | IN/GATE | D6 |
| Servo SG90 | Signal | D9 |
| OLED SSD1306 | SDA | A4 |
| OLED SSD1306 | SCL | A5 |

Tất cả module phải nối chung GND. Servo, quạt và LED nên dùng nguồn 5 V ngoài đủ dòng; không lấy dòng tải trực tiếp từ chân Arduino. Nối GND nguồn ngoài với GND Arduino.

Mỗi LDR tạo cầu chia áp: `5V — LDR — điểm đo A0/A1 — điện trở 10 kΩ — GND`. Với cách này, số ADC tăng khi ánh sáng tăng.

## 2. Cài thư viện

Trong Arduino IDE, mở **Tools → Manage Libraries** và cài:

- DHT sensor library by Adafruit
- Adafruit Unified Sensor
- Adafruit GFX Library
- Adafruit SSD1306

`Servo` và `Wire` thường đã có sẵn.

## 3. Chạy phần trình diễn tự động

Mở `EcoMicroClass_Demo.ino`, giữ:

```cpp
const bool DEMO_MODE = true;
```

Mỗi 7 giây, mô hình lần lượt diễn bốn tình huống:

1. Phòng học thoải mái.
2. Nắng gắt: rèm đóng và LED bù sáng.
3. Phòng nóng: quạt bật.
4. Không có người: LED tắt để tiết kiệm điện.

Mở **Serial Monitor**, chọn `115200 baud` để xem số liệu và quyết định của hệ thống.

## 4. Chạy bằng cảm biến thật

Đổi thành:

```cpp
const bool DEMO_MODE = false;
```

Sau đó dùng Serial Monitor ghi lại giá trị thực tế của hai LDR trong các điều kiện sáng/tối và sửa các ngưỡng ở đầu chương trình.

## 5. Nếu hoạt động ngược

- LDR tăng khi trời tối: đảo vị trí LDR và điện trở trong cầu chia áp, hoặc đổi logic/ngưỡng trong code.
- Servo kéo rèm ngược: đổi `CURTAIN_OPEN_ANGLE` và `CURTAIN_CLOSED_ANGLE`.
- OLED không hiện: thử địa chỉ `0x3D` thay cho `0x3C`.
- Arduino tự khởi động lại khi servo chạy: servo đang kéo dòng quá lớn; dùng nguồn 5 V ngoài và nối chung GND.

