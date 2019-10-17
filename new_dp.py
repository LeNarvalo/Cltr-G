import win32gui, time, shutil, os, pythoncom 
import win32com.client
import urllib.parse
from pynput.keyboard import Key, Controller, Listener
from win32com.shell import shell, shellcon
from tkinter import *


clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}' #Cf. base de registre
clavier = Controller()
chemin = os.path.expanduser('~')
window_dir = ''
#keys = []
#waiting = False

master = Tk()
master.configure(bg="white")
master.geometry("0x0+0+-50")
'''time.sleep(1)'''
 
def windowEnumerationHandler(hwnd, top_windows):
	top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))
 	
def explorer_fileselection():
	global clsid, window_dir
	pythoncom.CoInitialize()
	files = []	
	window = win32gui.GetForegroundWindow()
	shellwindows = win32com.client.Dispatch(clsid)
	for window in range(shellwindows.Count):
		try:
			window_URL = urllib.parse.unquote(shellwindows[window].LocationURL,encoding='ISO 8859-1')
			#print("SHELL LOCATION :",shellwindows[window].LocationURL)
			window_dir = window_URL.split("///")[1].replace("/", "\\")
			clavier.press(Key.alt)
			clavier.press(Key.tab)
			clavier.release(Key.alt)
			clavier.release(Key.tab)
			time.sleep(0.5)
			name_folder = win32gui.GetWindowText(win32gui.GetForegroundWindow())
			#print("name_folder",name_folder)
			#print("1.HERE BUG?")
			if chemin in window_dir:
				if name_folder == "Musique":
					name_folder = "Music"
				elif name_folder == "Téléchargements":
					name_folder = "Downloads"
				elif name_folder == "Images":
					name_folder = "Pictures"
			#print("2.HERE BUG?")			
			if name_folder in window_dir:
				#print("3.HERE BUG?")
				selected_files = shellwindows[window].Document.SelectedItems()
				#print("4.HERE BUG?")
				for file in range(selected_files.Count):					
					files.append(selected_files.Item(file).Path)
				#print(files)
				return files
			else:
				os._exit(1)
				master.quit()
				
		except:
			continue	
			

files_to_moved = explorer_fileselection()
	
clavier.press(Key.ctrl)
clavier.press(Key.shift)
clavier.press('n')
clavier.release(Key.ctrl)
clavier.release(Key.shift)
clavier.release('n')

waiting = True

#print("Ctrl+Maj+N new folder created")

def sound():
	winsound.MessageBeep()

def on_press(key):
	 #global keys
	#keys.append(key)
	##print(keys)
	if key == Key.enter:
		time.sleep(1)
		timer = time.time()
		path = window_dir		
		folders = []
		# r=root, d=directories, f = files
		for r, d, f in os.walk(path):
			for folder in d:
				folders.append(os.path.join(r, folder))
		delta = 999
		for f in folders:		
			if timer - os.stat(f).st_mtime < delta:
				delta = timer - os.stat(f).st_mtime
				newFolder = f
		listener.stop()
		#if win32ui.MessageBox("Si vous faites retour-arrère (Ctrl+Z) vos fichiers déplacés seront définitivement perdus !\nVoulez-vous continuez ?", "Avertissement", win32con.MB_YESNOCANCEL) == win32con.IDYES:
		for file in files_to_moved:
			shutil.move(file, newFolder+'\\'+file[::-1][0:file[::-1].index('\\')][::-1])
		#	os._exit(1)
		os._exit(1)
		master.quit()
		
#def on_release(key):
#	global keys
#	while len(keys)>0:
#		keys.pop()
		
def listen():
	global listener
	listener = Listener(
						on_press = on_press,
						#on_release = on_release,
						)
	listener.start()
	

listen()

while True:	
	master.lower()
	master.attributes('-topmost', 0)
	master.mainloop()
	continue
