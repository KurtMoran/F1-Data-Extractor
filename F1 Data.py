import pytesseract
import os
import cv2
import openpyxl
from numpy import interp
import numpy as np

# Set Dir
os.chdir(r"C:\Users\kurtm\OneDrive\Projects\Python\F1 Data") #Directory of the video
os.getcwd() #Gets current working directory

# Set up Excel file
wb = openpyxl.Workbook()
sheet = wb.active
colNum = 1

#Video to frames
videocap = cv2.VideoCapture('LEC.mp4')
success,image = videocap.read()
frame = 1

#Reference Colors
yellowRef = [0, 211, 254]
blackRef = [0, 0, 0]
redRef = [4, 29, 255]

while success:

	#cv2.imwrite("LEC%d.png" % frame, image) #Creates .png of every frame (used for troubleshooting)
	videocap.set(1,frame)
	success,image = videocap.read()	
	ret, currentFrame = videocap.read()
	
	print('New Frame: ', success)
	
	try: 
		colorTest = currentFrame[2150, 1750]
	except:
		break
	
	if (np.all(colorTest == yellowRef) or np.all(colorTest == blackRef) or np.all(colorTest == redRef)): #if yellow or session finished
	
		roiColor = currentFrame[2595:3382, 902:1774] # [y1:y2, x1:x2]
		roiVelocity = currentFrame[2810:2950, 1170:1520]
		roiLapT = currentFrame[2317:2412, 1750:2050]
		roiRPM = currentFrame[3070:3177, 1121:1589]
		roiGear = currentFrame[3428:3533, 1366:1482]
		roiLapN = currentFrame[3628:3730, 174:318]
		
	else: #if Green
	
		roiColor = currentFrame[2501:3305, 893:1782]
		roiVelocity = currentFrame[2718:2863, 1180:1494]
		roiLapT = currentFrame[2154:2236, 1750:2041]
		roiRPM = currentFrame[2978:3087, 1072:1623]
		roiGear = currentFrame[3338:3467, 1363:1483]
		roiLapN = currentFrame[3628:3730, 174:318]

	greenMin = np.array([0, 116, 16], np.uint8) #BGR not RGB
	greenMax = np.array([62, 177, 88], np.uint8)
	redMin = np.array([4, 4, 124], np.uint8)
	redMax = np.array([60, 41, 219], np.uint8)

	dst = cv2.inRange(roiColor, greenMin, greenMax)
	green = cv2.countNonZero(dst)
	
	print ('Amount of green pixels = ' + str(green))
	cv2.imshow("Color", dst)
	cv2.waitKey(0)
	
	dst = cv2.inRange(roiColor, redMin, redMax)
	red = cv2.countNonZero(dst)
	
	
	if red > 0: red = 1 #Brake is binary
	
	throttleMap = interp(green,[0,12250],[0,100])    
	print('Mapped Throttle = ' + str(throttleMap))
	print('red = ' + str(red))

	pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract' #tesseract path

	##### Velocity
	print('Velocity = ' + pytesseract.image_to_string(roiVelocity, lang='F1r', config='--psm 6')) ##config='--psm 6' Assumes a single uniform block of text.
	Velocity = (pytesseract.image_to_string(roiVelocity, lang='F1r', config='--psm 6'))
	cv2.imshow("Velocity", roiVelocity)
	cv2.waitKey(0)

	##### Lap time
	print('Lap time = ' + pytesseract.image_to_string(roiLapT, lang='lapTime', config='--psm 6'))
	lapTime= (pytesseract.image_to_string(roiLapT, lang='lapTime', config='--psm 6'))
	cv2.imshow("Lap Time", roiLapT)
	cv2.waitKey(0)
	
	##### RPM
	print('RPM = ' + pytesseract.image_to_string(roiRPM, lang='F1r', config='--psm 6'))
	RPM = (pytesseract.image_to_string(roiRPM, lang='F1r', config='--psm 6',))
	cv2.imshow("RPM", roiRPM)
	cv2.waitKey(0)
	
	##### Gear
	print('Gear = ' + pytesseract.image_to_string(roiGear, lang='F1r', config='--psm 6'))
	gear = (pytesseract.image_to_string(roiGear, lang='F1r', config='--psm 6'))
	cv2.imshow("Gear", roiGear)
	cv2.waitKey(0)
	
	##### Lap Num  ## only during race not qual
	print('Lap Num = ' + pytesseract.image_to_string(roiLapN, lang='F1r', config='--psm 6'))
	lapNum= (pytesseract.image_to_string(roiLapN, lang='F1r', config='--psm 6'))
	cv2.imshow("Lap Num", roiLapN)
	cv2.waitKey(0)
	
	print('')
	
	## write to Excel File
	
	try:
		sheet['A'+ str(colNum)].value = float(throttleMap)
	except:
		sheet['A'+ str(colNum)].value = throttleMap
		
	try:
		sheet['B'+ str(colNum)].value = int(red)
	except:
		sheet['B'+ str(colNum)].value = red
	
	try:
		sheet['C'+ str(colNum)].value = int(RPM)
	except:
		sheet['C'+ str(colNum)].value = RPM
		
	try:
		sheet['D'+ str(colNum)].value = int(gear)
	except:
		sheet['D'+ str(colNum)].value = gear

	try:
		sheet['E'+ str(colNum)].value = int(Velocity)
	except:
		sheet['E'+ str(colNum)].value = Velocity
		
	try:
		sheet['F'+ str(colNum)].value = int(lapNum)
	except:
		sheet['F'+ str(colNum)].value = lapNum
		
	sheet['G'+ str(colNum)].value = lapTime

	colNum += 1
	#os.remove("LEC%d.png" % frame)
	frame += 1

wb.save('LEC.xlsx')
wb.close


