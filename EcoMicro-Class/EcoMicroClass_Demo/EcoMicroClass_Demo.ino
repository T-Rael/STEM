/*
  EcoMicro-Class - Chuong trinh demo cho hoc sinh lop 6

  Chuc nang:
  1. Nang gat -> dong rem.
  2. Sau khi dong rem, neu ban hoc thieu sang -> tang LED bu sang.
  3. Nhiet do cao -> bat quat.
  4. Khong co nguoi trong PRESENCE_HOLD_MS -> tat LED de tiet kiem dien.
  5. OLED va Serial hien thi trang thai de demo truc quan.

  Thu vien can cai trong Arduino IDE (Library Manager):
  - DHT sensor library by Adafruit
  - Adafruit Unified Sensor
  - Adafruit GFX Library
  - Adafruit SSD1306
  Servo.h co san trong Arduino IDE.

  Luu y an toan:
  - Khong noi quat/LED cong suat truc tiep vao chan Arduino.
  - Dung MOSFET/module driver va noi chung GND.
  - Nguon demo: USB 5 V. Chua ghep pin mat troi o buoi hoc dau.
*/

#include <Servo.h>
#include <DHT.h>
#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>

// ===================== CHON CHE DO =====================
// true: tu dong dien 4 kich ban, khong can tac dong cam bien.
// false: doc cam bien that.
const bool DEMO_MODE = true;

// ===================== KHAI BAO CHAN ===================
const byte PIN_LDR_OUTSIDE = A0; // LDR ngoai troi: LDR len 5V, R 10k xuong GND
const byte PIN_LDR_DESK    = A1; // LDR tai ban hoc
const byte PIN_DHT         = 2;
const byte PIN_PIR         = 3;
const byte PIN_LED         = 5;  // Chan PWM, dieu khien LED qua MOSFET
const byte PIN_FAN         = 6;  // Dieu khien quat qua MOSFET/module
const byte PIN_SERVO       = 9;

#define DHTTYPE DHT22
DHT dht(PIN_DHT, DHTTYPE);
Servo curtainServo;
Adafruit_SSD1306 display(128, 64, &Wire, -1);

// ===================== NGUONG THU NGHIEM ===============
// Voi cach mac LDR ghi o tren: so ADC cao hon nghia la sang hon.
// Can do thuc te va chinh lai cac nguong nay cho mo hinh cua minh.
const int SUN_STRONG_ON  = 780; // Vuot muc nay: nang gat, dong rem
const int SUN_STRONG_OFF = 700; // Thap hon muc nay: het nang gat, mo rem
const int DESK_TARGET    = 600; // Muc sang mong muon tai ban (gia tri tuong doi)

const float TEMP_FAN_ON  = 30.0; // Bat quat
const float TEMP_FAN_OFF = 28.5; // Tat quat - tao hysteresis, tranh nhap nhay

const int CURTAIN_OPEN_ANGLE   = 15;
const int CURTAIN_CLOSED_ANGLE = 95;
const unsigned long PRESENCE_HOLD_MS = 15000UL; // Demo: giu trang thai co nguoi 15 s

// ===================== BIEN TRANG THAI ==================
int outsideLight = 0;
int deskLight = 0;
float temperatureC = 0;
float humidityPct = 0;
bool occupied = true;
bool curtainClosed = false;
bool fanOn = false;
bool oledReady = false;
int ledPwm = 0;

unsigned long lastSensorRead = 0;
unsigned long lastDisplayUpdate = 0;
unsigned long lastMotionTime = 0;
unsigned long demoStartTime = 0;

// Trung binh dong 10 mau de LDR bot nhieu.
const byte FILTER_SIZE = 10;
int outsideSamples[FILTER_SIZE] = {0};
int deskSamples[FILTER_SIZE] = {0};
byte sampleIndex = 0;
long outsideTotal = 0;
long deskTotal = 0;

// ===================== HAM HO TRO =======================
void addLightSample(int outsideRaw, int deskRaw) {
  outsideTotal -= outsideSamples[sampleIndex];
  deskTotal -= deskSamples[sampleIndex];

  outsideSamples[sampleIndex] = outsideRaw;
  deskSamples[sampleIndex] = deskRaw;

  outsideTotal += outsideRaw;
  deskTotal += deskRaw;
  sampleIndex = (sampleIndex + 1) % FILTER_SIZE;

  outsideLight = outsideTotal / FILTER_SIZE;
  deskLight = deskTotal / FILTER_SIZE;
}

void readRealSensors() {
  addLightSample(analogRead(PIN_LDR_OUTSIDE), analogRead(PIN_LDR_DESK));

  if (digitalRead(PIN_PIR) == HIGH) {
    lastMotionTime = millis();
  }
  occupied = (millis() - lastMotionTime) < PRESENCE_HOLD_MS;

  // DHT22 khong nen doc nhanh hon khoang 2 giay.
  if (millis() - lastSensorRead >= 2000UL) {
    lastSensorRead = millis();
    float newTemp = dht.readTemperature();
    float newHumidity = dht.readHumidity();
    if (!isnan(newTemp)) temperatureC = newTemp;
    if (!isnan(newHumidity)) humidityPct = newHumidity;
  }
}

void createAutomaticDemo() {
  // Moi kich ban keo dai 7 giay; chu ky lap lai sau 28 giay.
  unsigned long phase = ((millis() - demoStartTime) / 7000UL) % 4;

  switch (phase) {
    case 0: // Lop hoc thoai mai
      outsideLight = 620;
      deskLight = 650;
      temperatureC = 27.0;
      humidityPct = 58.0;
      occupied = true;
      break;

    case 1: // Nang gat, dong rem; ban hoc thieu sang nen LED bu sang
      outsideLight = 900;
      deskLight = 350;
      temperatureC = 29.0;
      humidityPct = 60.0;
      occupied = true;
      break;

    case 2: // Phong nong, quat bat
      outsideLight = 650;
      deskLight = 430;
      temperatureC = 32.0;
      humidityPct = 66.0;
      occupied = true;
      break;

    default: // Khong co nguoi, tat LED de tiet kiem
      outsideLight = 300;
      deskLight = 250;
      temperatureC = 27.0;
      humidityPct = 60.0;
      occupied = false;
      break;
  }
}

void decideCurtain() {
  // Hai nguong ON/OFF khac nhau giup servo khong rung gan nguong.
  if (!curtainClosed && outsideLight >= SUN_STRONG_ON) {
    curtainClosed = true;
  } else if (curtainClosed && outsideLight <= SUN_STRONG_OFF) {
    curtainClosed = false;
  }
}

void decideLighting() {
  if (!occupied) {
    ledPwm = 0;
    return;
  }

  // Sai lech cang lon thi LED cang sang. constrain giu PWM trong 0..255.
  int lightDeficit = DESK_TARGET - deskLight;
  ledPwm = constrain(map(lightDeficit, 0, DESK_TARGET, 0, 255), 0, 255);
}

void decideFan() {
  if (!fanOn && temperatureC >= TEMP_FAN_ON) {
    fanOn = true;
  } else if (fanOn && temperatureC <= TEMP_FAN_OFF) {
    fanOn = false;
  }
}

void applyOutputs() {
  curtainServo.write(curtainClosed ? CURTAIN_CLOSED_ANGLE : CURTAIN_OPEN_ANGLE);
  analogWrite(PIN_LED, ledPwm);
  digitalWrite(PIN_FAN, fanOn ? HIGH : LOW);
}

const char* mainMessage() {
  if (!occupied) return "TIET KIEM DIEN";
  if (temperatureC >= TEMP_FAN_ON) return "PHONG DANG NONG";
  if (curtainClosed && ledPwm > 0) return "REM + BU SANG";
  if (curtainClosed) return "DONG REM CHONG CHOI";
  if (ledPwm > 0) return "DANG BU SANG";
  return "PHONG THOAI MAI";
}

void showOnOled() {
  if (!oledReady) return;

  display.clearDisplay();
  display.setTextColor(SSD1306_WHITE);
  display.setTextSize(1);

  display.setCursor(0, 0);
  display.print(DEMO_MODE ? "ECOMICRO [DEMO]" : "ECOMICRO [SENSOR]");

  display.setCursor(0, 14);
  display.print("Ngoai:"); display.print(outsideLight);
  display.print(" Ban:"); display.print(deskLight);

  display.setCursor(0, 26);
  display.print("T:"); display.print(temperatureC, 1); display.print("C ");
  display.print("H:"); display.print(humidityPct, 0); display.print("%");

  display.setCursor(0, 38);
  display.print("Rem:"); display.print(curtainClosed ? "DONG" : "MO");
  display.print(" LED:"); display.print(ledPwm);

  display.setCursor(0, 50);
  display.print(mainMessage());
  display.display();
}

void showOnSerial() {
  Serial.print("Ngoai="); Serial.print(outsideLight);
  Serial.print(" | Ban="); Serial.print(deskLight);
  Serial.print(" | T="); Serial.print(temperatureC, 1);
  Serial.print("C | H="); Serial.print(humidityPct, 0);
  Serial.print("% | Co nguoi="); Serial.print(occupied ? "CO" : "KHONG");
  Serial.print(" | Rem="); Serial.print(curtainClosed ? "DONG" : "MO");
  Serial.print(" | LED PWM="); Serial.print(ledPwm);
  Serial.print(" | Quat="); Serial.print(fanOn ? "BAT" : "TAT");
  Serial.print(" | "); Serial.println(mainMessage());
}

void setup() {
  Serial.begin(115200);
  pinMode(PIN_PIR, INPUT);
  pinMode(PIN_LED, OUTPUT);
  pinMode(PIN_FAN, OUTPUT);

  dht.begin();
  curtainServo.attach(PIN_SERVO);
  curtainServo.write(CURTAIN_OPEN_ANGLE);

  oledReady = display.begin(SSD1306_SWITCHCAPVCC, 0x3C);
  if (!oledReady) {
    Serial.println("Khong tim thay OLED. He thong van tiep tuc chay.");
  } else {
    display.clearDisplay();
    display.setTextColor(SSD1306_WHITE);
    display.setTextSize(1);
    display.setCursor(0, 20);
    display.println("EcoMicro-Class");
    display.println("Dang khoi dong...");
    display.display();
  }

  lastMotionTime = millis();
  demoStartTime = millis();
  delay(1200);
}

void loop() {
  if (DEMO_MODE) {
    createAutomaticDemo();
  } else {
    readRealSensors();
  }

  decideCurtain();
  decideLighting();
  decideFan();
  applyOutputs();

  if (millis() - lastDisplayUpdate >= 500UL) {
    lastDisplayUpdate = millis();
    showOnOled();
    showOnSerial();
  }

  // Khong dung delay() trong loop: cac bo phan duoc cap nhat lien tuc.
}
