# -*- coding: utf-8 -*-
"""
Created on Thu Mar 10 11:00:27 2022

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
def AbsoluteMove(axis,amount):
    if(axis == "x"):
        axis = str(2)
    elif(axis == "z"):
        axis = str(1)
    cmd = ("A:"+axis+"+P"+str(amount)+"\r\n").encode()
    print(cmd)
    ser.write(cmd)
    ser.readline()
    ser.write(b"G:\r\n")
    ser.readline()
    
#####################################
#相対移動

def RelativeMove(axis,amount):
    if(axis == "x"):
        axis = str(2)
    elif(axis == "z"):
        axis = str(1)
    cmd = ("M:"+axis+"+P"+str(amount)+"\r\n").encode()
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
#内部情報の表示
    
def ShowInfo():
    cmd = ("?:S2\r\n").encode()
    ser.write(cmd)
    out = ser.readline().decode()
    print("Px = " + out)

    cmd2 = ("?:S1\r\n").encode()
    ser.write(cmd2)
    out2 = ser.readline().decode()
    print("Pz = " + out2)
    
#####################################
#分割数の変更
    
def ChangePara(para):
    """
    if(axis == "x"):
        axis = str(2)
    elif(axis == "z"):
        axis = str(1)
    """
    cmd = ("S:"+str(1)+str(para)+"\r\n").encode()
    ser.write(cmd)
    ser.readline()


BorR()
ShowInfo()
ser.close()