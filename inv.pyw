import time
import threading
import tkinter
import win32api
import win32com.client as wincl
from PIL import ImageGrab

class Inv:
	def __init__(self):
		self.data_lock = threading.Lock()
		self.voice_lock = threading.Lock()
		self.voice = wincl.Dispatch("SAPI.SpVoice")
		self.the_really_long_loop = True
		self.reminded_user = False
		self.main_loop_run = False
		self.x1 = 0
		self.y1 = 0
		self.x2 = 0
		self.y2 = 0
		self.prev_time = 0
		self.prev_img=ImageGrab.grab()
		try:
			file = open("data.txt","r")
			self.data=file.read().split("\n")
			file.close()
		except:
			self.data = ["",""]
		self.current_time = 0
		self.prev_time = 0
		
	def getSeconds(self,seconds):
		self.seconds = seconds
	
	def getUpperLeft(self):
		self.voice.speak("Getting first point in...")
		for i in range(3,0,-1):
			self.voice.speak(i)
		self.x1,self.y1 = win32api.GetCursorPos()
		self.voice.speak("Captured")
	
	def getLowerRight(self):
		self.voice.speak("Getting second point in...")
		for i in range(3,0,-1):
			self.voice.speak(i)
		self.x2,self.y2 = win32api.GetCursorPos()
		self.voice.speak("Captured")
	
	def printImage(self):
		if self.hasGottenPoints():
			ImageGrab.grab([self.x1,self.y1,self.x2,self.y2]).show()
	
	def hasGottenPoints(self):
		return (self.x1!=0 and self.y1!=0 and self.x2!=0 and self.y2!=0)
		
	def setSecond(self, input):
		try:
			self.data[0] = str(int(input))
			file = open("data.txt",'w')
			file.write(self.data[0]+"\n")
			file.write(self.data[1])
			file.close()
		except:
			alert = tkinter.Tk()
			alert.title("Alert")
			alert.geometry("200x40")
			alert.resizable(width=False, height=False)
			messageLabel=tkinter.Label(alert,text="Please enter a number only")
			messageLabel.pack(pady=10)
			alert.mainloop()
		
	def setText(self, input):
		self.data[1] = input
		file = open("data.txt",'w')
		file.write(self.data[0]+"\n")
		file.write(self.data[1])
		file.close()
	
	def getData(self):
		return self.data
	
	def setMainLoopRun(self):
		if(self.hasGottenPoints()):
			self.main_loop_run = not self.main_loop_run
	
	def run(self):
		while self.the_really_long_loop:
			while self.main_loop_run:
				self.current_time = int(time.time())
				self.current_img = ImageGrab.grab([self.x1,self.y1,self.x2,self.y2])
				if self.current_img == self.prev_img and self.current_time-self.prev_time == int(self.getData()[0]) and not self.reminded_user:
					self.voice.speak(self.getData()[1])
					self.reminded_user=True
				elif self.current_img != self.prev_img:
					self.reminded_user=False
					self.prev_time = self.current_time
				self.prev_img = self.current_img
			else:
				time.sleep(.5)
			

def initUI():
	def changeMainLoopRunButton():
		if(remind.hasGottenPoints()):
			if(remind.main_loop_run):
				remind.setMainLoopRun()
				setMainLoopRunButton['text']="Off"
			else:
				remind.setMainLoopRun()
				setMainLoopRunButton['text']="On"
	
	#setup
	tk.title("Remind Me")
	tk.geometry("393x90")
	tk.resizable(width=False, height=False)
	
	#point frame
	pointFrame = tkinter.Frame(tk)
	pointFrame.pack()
	pointFrame.place(x=1, y=0)
	getUpperLeftButton = tkinter.Button(pointFrame, text="Get First Point", command=remind.getUpperLeft,height=2,width=15)
	getUpperLeftButton.pack(side=tkinter.TOP,ipady=2)
	getLowerRightButton = tkinter.Button(pointFrame, text="Get Second Point", command=remind.getLowerRight,height=2,width=15)
	getLowerRightButton.pack(side=tkinter.TOP,ipady=2)
	
	#entry frame
	entryFrame = tkinter.Frame(tk)
	entryFrame.pack()
	entryFrame.place(x=117,y=0)
	changeSecondEntry = tkinter.Entry(entryFrame)
	changeSecondEntry.insert(0,remind.getData()[0]) 
	changeSecondEntry.pack(side=tkinter.TOP,fill=tkinter.X)
	changeSecondButton = tkinter.Button(entryFrame, text="Change Seconds in Between", command=lambda: remind.setSecond(changeSecondEntry.get()))
	changeSecondButton.pack(side=tkinter.TOP,fill=tkinter.X)
	changeTextEntry = tkinter.Entry(entryFrame)
	changeTextEntry.insert(0,remind.getData()[1])
	changeTextEntry.pack(side=tkinter.TOP,fill=tkinter.X)
	changeTextButton = tkinter.Button(entryFrame, text="Change Reminder Message", command=lambda: remind.setText(changeTextEntry.get()))
	changeTextButton.pack(side=tkinter.TOP,fill=tkinter.X)
	
	#onOff frame
	onOffFrame = tkinter.Frame(tk,width=17)
	onOffFrame.pack()
	onOffFrame.place(x=278,y=0)
	setMainLoopRunButton = tkinter.Button(onOffFrame, text="Off", command=lambda: changeMainLoopRunButton(),height=2,width=15)
	setMainLoopRunButton.pack(side=tkinter.TOP,ipady=2)
	
	#check area frame
	checkAreaFrame = tkinter.Frame(tk)
	checkAreaFrame.pack()
	checkAreaFrame.place(x=278,y=45)
	printImageButton = tkinter.Button(checkAreaFrame, text="Check Area", command=remind.printImage,height=2,width=15)
	printImageButton.pack(side=tkinter.TOP,ipady=2)

def end():
	remind.the_really_long_loop=False
	tk.destroy()

remind=Inv()
remind_thread = threading.Thread(target=remind.run)
remind_thread.daemon=True
remind_thread.start()

tk = tkinter.Tk()
initUI()
tk.protocol("WM_DELETE_WINDOW", end)
tk.mainloop()
