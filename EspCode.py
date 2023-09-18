import socket
from machine import ADC, Pin
from time import sleep

inputPin = ADC(Pin(34))          #Pin usado para ADC; Default ATTN_0dB y WIDTH 12 bits
inputPin.atten(ADC.ATTN_11DB)    #Configur para un rango de voltaje de 0v a 3,6v
valuePin = 0                     #Variable para almacenar lectura

def do_connect():
    import network
    wlan = network.WLAN(network.STA_IF)
    wlan.active(True)
    if not wlan.isconnected():
        print('connecting to network...')
        wlan.connect('CAMARGO', 'DC148103802')
        while not wlan.isconnected():
            pass
    print('network config:', wlan.ifconfig())


do_connect()
print ('conected to network')



'''
while True:
    ValuePin = inputPin.read()
    Serial.write(ValuePin)
    sleep(0.2)
'''