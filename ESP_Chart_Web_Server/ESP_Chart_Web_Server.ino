// Librerias Requeridas
  #include <WiFi.h>
  #include <ESPAsyncWebServer.h>
  #include <SPIFFS.h>
  #include <Adafruit_INA219.h>
  #ifdef U8X8_HAVE_HW_SPI
  #include <SPI.h>
  #endif
  #ifdef U8X8_HAVE_HW_I2C
  #include <Wire.h>
  #endif

Adafruit_INA219 ina219;
//U8G2_SSD1306_128X64_NONAME_F_HW_I2C u8g2(U8G2_R0, /* reset=*/ U8X8_PIN_NONE);

//Lectura puertos analogicos
const int VPin = 34;
const int IPin = 35;
// Variables para almacenamiento de las lecturas
int VValue = 0;
int IValue = 0;

// Credenciales de la red WiFi
const char* ssid = "CAMARGO";
const char* password = "DC1234788444";

// Creacion AsyncWebServer object en port 80

AsyncWebServer server(80);
static float ref_current_mA = 0;
static float ref_loadvoltage = 0;

static float shuntvoltage = 0;
static float busvoltage = 0;
static float current_mA = 0;
static float loadvoltage = 0;
static float power = 0;

static uint32_t pos = 0;

String voltage() {
  // Lectura de Voltaje
  float v = loadvoltage;
    return String(v);
  }

String current() {
  // Lectura Corriente
  float i = current_mA;
    return String(i);
  }

void setup(){
  // Puerto serial para comunicacion con el software de instrumentacion
  Serial.begin(115200);
  Wire.begin();
  //u8g2.begin(); 
  ina219.begin();
  
  bool status; 
  // Inicializacion SPIFFS
  if(!SPIFFS.begin()){
    return; }

  // Conexion a la reed WiFi
  WiFi.begin(ssid, password);
  while (WiFi.status() != WL_CONNECTED) {
    delay(1000); }

  // ESP32 Local IP Address
  Serial.println(WiFi.localIP());

  // Ruta conexion con servidor / pagina web
  server.on("/", HTTP_GET, [](AsyncWebServerRequest *request){
    request->send(SPIFFS, "/index.html");
  });
  server.on("/voltage", HTTP_GET, [](AsyncWebServerRequest *request){
    request->send_P(200, "text/plain", voltage().c_str());
  });
   server.on("/current", HTTP_GET, [](AsyncWebServerRequest *request){
    request->send_P(200, "text/plain", current().c_str());
  });
    // Start server
  server.begin();
}



void loop(){
  char *textbuf;

  pos++;

  shuntvoltage = ina219.getShuntVoltage_mV();
  busvoltage = ina219.getBusVoltage_V();
  current_mA = ina219.getCurrent_mA();
  loadvoltage = busvoltage + (shuntvoltage / 1000);
  power = busvoltage * current_mA;


  if (pos < 10) {
    ref_current_mA += current_mA;
    ref_loadvoltage += loadvoltage;
  } else if (pos == 10) {
    ref_current_mA /= 10;
    ref_loadvoltage /= 10;
  } else {
    current_mA -= ref_current_mA;
    loadvoltage -= ref_loadvoltage;
  }

  Serial.print(abs(loadvoltage));
  Serial.print("&");
  Serial.println(abs(current_mA));
  Serial.print("&");
  Serial.println(abs(shuntvoltage));
  Serial.print("&");
  Serial.println(abs(busvoltage));
  
  delay(500);
}
