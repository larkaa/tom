#!/usr/bin/python3


from flask import Flask
from flask import request
import epd7in5
from PIL import Image,ImageDraw,ImageFont
import time

app = Flask(__name__)

epd = epd7in5.EPD();
epd.init() 

@app.route('/gift/<gift_id>', methods = ['POST'])
def user(gift_id):
    
    if request.method == 'POST':
        gid = int(gift_id)
        
        if gid==0:
            
            f='questions_final.jpg'
            show_img_bw(epd,f)
            #print('total time {}'.format(time.time() - t))
            return '200'
        elif gid ==1:
            f='homer.jpg'
            show_img_bw(epd,f)
            return '200'
        elif gid ==2:
            f='eat_dust.png'
            show_img_bw(epd,f)
            return '200'
        elif gid == 3:
            f='oops_final.jpg'
            show_img_bw(epd,f)
            return '200'
        elif gid ==4:
            f='yay.png'
            show_img_bw(epd,f)
            return '200'
        elif gid ==5:
            f='test_1.jpg'
            show_img_bw(epd,f)
            return '200'
        elif gid ==6:
            f='test_2.jpg'
            show_img_bw(epd,f)
            return '200'
        else:
            return '404'
        
def show_img_bw(epd,f):
    t = time.time()
    img = Image.open(f)#.convert('1')
    size = (640,384)
    print(img.size)
    #size_input = img.shape
    if img.size != size:
        print('resize')
        img = img.resize(size,Image.LANCZOS)
    print('image load {}'.format(time.time() - t))
    #t= time.time()
    #img = img.convert('1')
    #print('image convert {}'.format(time.time() - t))
    t= time.time()
    buf = epd.getbuffer(img)
    #buf = getbuffer(img)
    print('buf time {}'.format(time.time()-t))
    t = time.time()
    epd.display(buf)
    
    print('image display {}'.format(time.time() - t))
    
    
def show_img_bwr(epd,red,black):
    size1=264
    size2=176
    size = (size1,size2)
    red = Image.open(red)
    black = Image.open(black)
    red = red.resize(size,Image.LANCZOS).convert('1')
    black = black.resize(size,Image.LANCZOS).convert('1')

    epd.display(epd.getbuffer(red), epd.getbuffer(black)) 
def getbuffer( image_monocolor, width = 640, height = 384):
    buf = [0x00] * (width * height // 4)
    #image_monocolor = image.convert('1')
    imwidth, imheight = image_monocolor.size
    pixels = image_monocolor.load()
    print("imwidth = ", imwidth, "imheight = ", imheight)
    if(imwidth == width and imheight == height):
        for y in range(imheight):
            for x in range(imwidth):
                # Set the bits for the column of pixels at the current position.
                if pixels[x, y] < 64:           # black
                    buf[(x + y * width) // 4] &= ~(0xC0 >> (x % 4 * 2))
                elif pixels[x, y] < 192:     # convert gray to red
                    buf[(x + y * width) // 4] &= ~(0xC0 >> (x % 4 * 2))
                    buf[(x + y * width) // 4] |= 0x40 >> (x % 4 * 2)
                else:                           # white
                    buf[(x + y * width) // 4] |= 0xC0 >> (x % 4 * 2)
    elif(imwidth == height and imheight == width):
        for y in range(imheight):
            for x in range(imwidth):
                newx = y
                newy = height - x - 1                    
                if pixels[x, y] < 64:           # black
                    buf[(newx + newy*width) // 4] &= ~(0xC0 >> (y % 4 * 2))
                elif pixels[x, y] < 192:     # convert gray to red
                    buf[(newx + newy*width) // 4] &= ~(0xC0 >> (y % 4 * 2))
                    buf[(newx + newy*width) // 4] |= 0x40 >> (y % 4 * 2)
                else:                           # white
                    buf[(newx + newy*width) // 4] |= 0xC0 >> (y % 4 * 2)
    return buf  
    
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=80, debug = True)




