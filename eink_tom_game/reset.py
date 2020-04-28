import epd2in7b                               
from PIL import Image,ImageDraw,ImageFont     
import time
import RPi.GPIO as GPIO
import requests
#from gpiozero import Button      #alternate way to access buttons            
#btn1 = Button(5)                              
#btn2 = Button(6)                              
#btn3 = Button(13)                             
#btn4 = Button(19)                             

GPIO.setmode(GPIO.BCM)

key1 = 5
key2 = 6
key3 = 13
key4 = 19

GPIO.setup(key1, GPIO.IN, pull_up_down=GPIO.PUD_UP)
GPIO.setup(key2, GPIO.IN, pull_up_down=GPIO.PUD_UP)
GPIO.setup(key3, GPIO.IN, pull_up_down=GPIO.PUD_UP)
GPIO.setup(key4, GPIO.IN, pull_up_down=GPIO.PUD_UP)

def disp_gift(epd,f1,f2):
    size1=264
    size2=176
    size = (size1,size2)
    red = Image.open(f1)
    black = Image.open(f2)
    red = red.resize(size,Image.LANCZOS).convert('1')
    black = black.resize(size,Image.LANCZOS).convert('1')
    
    epd.display(epd.getbuffer(red), epd.getbuffer(black))

def main():
    
    #requests.post('http://192.168.1.139/gift/5')
    epd = epd2in7b.EPD()                          # get the display object and assing to epd
    epd.init()                                    # initialize the display

    f1 = 'intro_b.png'
    f2 = 'intro_red.png'

    
    #f1 = '/home/pi/Desktop/pi/code/2dot7_eink_python3/Tom_eink_1_red.bmp'
    #f2 = '/home/pi/Desktop/pi/code/2dot7_eink_python3/Tom_eink_1_black.bmp'
    f1 = 'off_b.png'
    f2 = 'off_r.png'
    
    disp_gift(epd,f1,f2)
    #time.sleep(2)
    #print('sleep')


if __name__ == '__main__':
    main()
