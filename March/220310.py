# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 10:45:24 2022

@author: 高崎元希
"""

import serial
import os
import tkinter
import pyautogui
from PIL import Image, ImageTk
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import glob

######################################
#初期設定
######################################

#シリアル通信設定
COM="COM6"
bitRate=9600
ser = serial.Serial(COM, bitRate)

#キャプチャ設定
ration = 2

######################################
#ステージ制御関数
######################################

#完了チェック
def BorR():
    ser.write(bytes(b"!:\r\n"))
    line = ser.readline()

    while(line == b'B\r\n'):
        ser.write(bytes(b"!:\r\n"))
        line = ser.readline()

######################################
#原点復帰

def MoveZero(axis):
    if(axis == "x"):
        ser.write(b"H:2\r\n")
    elif(axis == "z"):
        ser.write(b"H:1\r\n")
    line0 = ser.readline()
    print(line0)
    
######################################
#絶対移動
def AbsoluteMove(amount):
    cmd = ("A:"+str(1)+"+P"+str(amount)+"\r\n").encode()
    print(cmd)
    ser.write(cmd)
    ser.readline()
    ser.write(b"G:\r\n")
    ser.readline()
    
#####################################
#相対移動

def RelativeMove(amount):
    cmd = ("M:"+str(1)+"+P"+str(amount)+"\r\n").encode()
    print(cmd)
    ser.write(cmd)
    ser.readline()
    ser.write(b"G:\r\n")
    ser.readline()
    
#####################################
#現在位置取得
    
def GetPosition():
    ser.write(b"Q:\r\n")
    out = ser.readline().decode()
    print(out)

#####################################
#カメラ制御関数
#####################################

def start(event):
    global start_x, start_y
    
    canvas1.delete("rect1")
    
    canvas1.create_rectangle(event.x,event.y,event.x+1,event.y+1,outline="red",tag="rect1")
    
    start_x = event.x
    start_y = event.y
    
def drawing(event):
    
    if event.x < 0:
        end_x = 0
    else:
        end_x = min(img_resized.width,event.x)
    if event.y < 0:
        end_y = 0
    else:
        end_y = min(img_resized.width,event.y)
        
    canvas1.coords("rect1",start_x,start_y,end_x,end_y)
    
def release_action(event):
    
    start_x,start_y,end_x,end_y = [round(n*ration) for n in canvas1.coords("rect1")]
    
    pyautogui.alert("start_x:" + str(start_x) + "\n" + "start_y:" + str(start_y) + "\n" +
                    "end_x:" + str(end_x) + "\n" + "end_y:" + str(end_y))
    
#####################################
