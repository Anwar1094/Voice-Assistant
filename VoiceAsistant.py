# importing necessary modules
from customtkinter import *
from win32com.client import Dispatch
import speech_recognition as recognize
from datetime import datetime
from PIL import Image, ImageGrab
from AppOpener import open as Open, close as Close
from webbrowser import open as opensite
from playsound import playsound
# import openai
# from Key import Key
from multiprocessing import Process
import cv2
import sounddevice as sd
from scipy.io.wavfile import write
from plyer import notification
import schedule
from time import sleep


# Method to start countdown
def Timer(t):
    import time
    hours = 0
    while(t):
        mins, secs = divmod(t, 60)
        if mins > 60:
            hours, mins = divmod(mins, 60)
        timer = '{:02d}:{:02d}:{:02d}'.format(hours, mins, secs)
        print(timer, end="\r")
        time.sleep(1)
        t-=1     

# # Method for reminder
done = False
def Remind(note):
        notification.notify(
            title = "Alert",
            message = f"Reminder {note}, Sir!",
            app_icon = None,
            timeout = 1,
        )
        global done
        done = True

# Method to schedule remainder
def run(t, note):
    schedule.every().day.at(t).do(lambda: Remind(note))
    while True:
        schedule.run_pending()
        sleep(1)
        global done
        if done == True:
            break
# Main Class
class VoiceAssistant(CTk):
    attempt = 1
    id = 0
    # Constructor
    def __init__(self):
        # Setting up customtkinter window: adding button, label and heading
        super().__init__()
        self._set_appearance_mode("light")
        self.width, self.height = 400, 350
        self.geometry(f"{self.width}x{self.height}")
        self.resizable(0,0)
        self.title("Voice Assistant")
        self.eval('tk::PlaceWindow . center')
        self.wm_iconbitmap(r"Bin\AiLogo.ico")
        self.config(background= '#ffffff')
        self.mic = CTkButton(self, text="", command=self.Start, bg_color="#ffffff", fg_color="#ffffff", hover_color="#ebebeb", image=CTkImage(Image.open("Bin/voice.png"), size=(200, 200)))
        self.mic.place(x=90, y=100)
        CTkLabel(self, text="", image=CTkImage(Image.open("Bin/PythonAIText.png"), size=(250, 100)), font=("Lucida Handwriting", 40)).place(x=75, y=5)
        self.label = CTkLabel(self, text="", bg_color="#ffffff", text_color="black" , font=("Lucida Handwriting", 20))
        self.label.place(x=140, y=310)
        # Import api key
        # key = Key()
        # openai.api_key = key
    
    # Method to starting the listening of voice assistant
    def Start(self):
        self.label.configure(text="Listening....")
        self.mic.configure(image=CTkImage(Image.open("Bin/Active.png"), size=(200, 200)))
        self.after(10 , lambda: self.Assist())
    
    # Method to setup voice assistant and response 
    run = True
    def Assist(self):
        self.speak = Dispatch('SAPI.SpVoice').Speak
        if self.run == True:
            self.speak("Start Listening Sir..")
            self.run = False
        self.rec = recognize.Recognizer()
        # Setting up Microphone to listen to the user
        def Listen():
            with recognize.Microphone() as source:
                data = self.rec.listen(source)
            return self.rec.recognize_google(data).lower()
        try:
            self.text = Listen()
            print("Command: " + self.text.title())
            # Custom response to inputs
            if self.text == "Hello Python AI".lower():
                self.speak("Hii Sir! This is Python AI! How can I help you?")
            elif "time" in self.text:
                self.speak(f"Sir! It's {datetime.now().strftime("%H")} hour { datetime.now().strftime("%M")} minutes and {int(datetime.now().strftime("%S"))} seconds right now!")
            elif "date" in self.text:
                months = {1:"January", 2:"February", 3:"March", 4:"April", 5:"May", 6:"June", 7:"July", 8:"August", 9:"September", 10:"October", 11:"November", 12:"December"}
                self.speak(f"Sir! IT's {datetime.now().day} {months[datetime.now().month]} {datetime.now().year} today!")
                print(f"Sir! IT's {datetime.now().day} {months[datetime.now().month]} {datetime.now().year} today!")
            elif "goodbye" in self.text or "good bye" in self.text:
                self.speak(f"Good Bye Sir!")
                self.quit()
            elif "from web browser open" in self.text:
                websites = [
                    ["youtube", "https://youtube.com"], ["google", "https://google.com"], ["hackerrank", "https://hackerrank.com"], ["wikipedia", "https://wikipedia.com"],
                    ["canva", "https://canva.com"],["linkedin", "https://linkedin.com"],["gmail", "https://gmail.com"],["github", "https://github.com"],
                    ["invertis university", "https://invertisuniversity.ac.in"],["leetcode", "https://leetcode.com"],["openai", "https://openai.com"]
                ]
                for site in websites:
                    if "".join(self.text.split("from web browser open")).strip() == site[0]:
                        self.speak(f"Opening {site[0]} on web browser, Sir")
                        opensite(f"{site[1]}")
            elif "open" in self.text:
                appName = "".join(self.text.split("open")).strip()
                try:
                    self.speak(f"Opening {appName}")
                    Open(appName)
                except Exception as e:
                    self.speak(f"The app {appName} is not found on your system, Sir!")
            elif "close" in self.text:
                appName = "".join(self.text.split("close")).strip()
                try:
                    self.speak(f"Closing {appName}...")
                    Close(appName)
                except Exception as e:
                    self.speak(f"The app {appName} is not running, Sir!")
            elif "play the music" in self.text:
                self.speak("Playing music")
                self.MusicPlayer()
            elif "stop the music" in self.text:
                self.p.terminate()
                self.speak("Sure Sir")
            elif "start video recording" in self.text:
                self.speak("start video recording")
                self.Recording()
            elif "start audio recording" in self.text:
                self.speak("start audio recording")
                self.AudioRecording()
            elif "take a picture" in self.text:
                self.speak("Please Smile Sir...")
                self.takePic()
            elif "change the music" in self.text:
                self.speak("changing the song")
                try:
                    self.p.terminate()
                except:pass
                self.MusicPlayer()
            # elif "using ai" in self.text:
            #     self.AI() 
            elif "take screenshot" in self.text:
                self.speak("Taking screenshot in 5 seconds")
                self.after(5000, self.Screenshot)
                self.speak("Done Sir")
            elif "remind me" in self.text:
                t = []
                for data in self.text.split("remind me at "):
                    if data != "":
                        for i in data.split():
                            t.append(i)
                note = ""
                for item in t[1:]:
                    note += item+' '
                if ":" not in t[0]:
                    t[0] = t[0][:2] + ":"+t[0][2:]
                if t[0][1] == ":":
                    t[0] = "0"+t[0]
                Process(target=run, args=((t[0]), note)).start()
                self.speak("Sure Sir!")
            elif "start countdown" in self.text:
                if ("seconds" in self.text) and ("hour" and "minute" not in self.text):
                    for t in self.text.split():
                        if t.isnumeric():
                            tp = Process(target=Timer, args=(int(t),))
                            tp.start()
                elif ("minute" in self.text) and ("hour" and "seconds" not in self.text):
                    for t in self.text.split():
                        if t.isnumeric():
                            tp = Process(target=Timer, args=(int(t)*60,))
                            tp.start()
                elif ("hour" in self.text) and "minute" and "seconds" not in self.text:
                    for t in self.text.split():
                        if t.isnumeric():
                            tp = Process(target=Timer, args=(int(t)*60*60,))
                            tp.start()
                elif ("minute" and "seconds" in self.text) and "hour" not in self.text:
                    nums = []
                    for t in self.text.split():
                        if t.isnumeric():
                            nums.append(int(t))
                    tp = Process(target=Timer, args=(nums[0]*60+nums[1],))
                    tp.start()
                elif "hour" and "minute" and "seconds" in self.text:
                    nums = []
                    for t in self.text.split():
                        if t.isnumeric():
                            nums.append(int(t))
                    tp = Process(target=Timer, args=((nums[0]*60*60)+(nums[1]*60)+nums[2],))
                    tp.start()
                self.speak("Sure Sir!")
            self.attempt = 0
        except recognize.UnknownValueError:
            self.speak('Sorry! I could not understand Sir')
            if self.attempt == 3:
                self.speak("Good Bye Sir!")
                self.quit()
            self.attempt += 1
        except recognize.RequestError:
            self.speak('Sorry! Currently some services are not available Sir')
            quit()
        except Exception as e:
            print(e)

        self.label.configure(text="")
        try:
            if "stop listening" not in self.text:
                self.after(100, self.Start)
            else:
                self.mic.configure(image=CTkImage(Image.open("Bin/voice.png"), size=(200, 200)))
                self.speak('Sure Sir!')
                self.run = True
        except:
            self.after(100, self.Start)

    # Method to take screenshot
    def Screenshot(self):
        scrshot = ImageGrab.grab()
        scrshot.save("screenshot.jpg")
        scrshot.close()

    # Method to play the music
    def MusicPlayer(self):
        paths = os.listdir("Music")
        try:
            self.p = Process(target=playsound, args=(f"./Music/{paths[self.id]}",)) 
            self.p.start()
            self.id += 1
        except:
            self.p.terminate()
            self.speak("Song Finished")

    # Method to take a pic from webcam
    def takePic(self):
        cam = cv2.VideoCapture(0)
        result, image = cam.read()
        if result:
            cv2.imshow("Image", image)
            cv2.imwrite("Image.png", image)
            cv2.waitKey(0)
            cv2.destroyWindow("Image")
        else:
            print("fail")

    # Method to record video
    def Recording(self):
        video = cv2.VideoCapture(0)
        result = ""
        for i in range(1000):
            if not os.path.exists(f"Recorded Video ({i}).avi"):
                result = cv2.VideoWriter(f"Recorded Video ({i}).avi", cv2.VideoWriter_fourcc(*'MJPG'), 30, (int(video.get(3)), int(video.get(4))))
                break
        while True:
            ret, frame = video.read()
            result.write(frame)
            cv2.imshow('frame', frame)
            if cv2.waitKey(1) & 0xFF == ord('q'): 
                break
        video.release()
        result.release()
        cv2.destroyAllWindows()
    
    # Method to record voice
    def AudioRecording(self):
        fs = 44100
        seconds = 10
        myrecording = sd.rec(int(seconds * fs), samplerate=fs, channels=2)
        sd.wait()
        for i in range(100):
            if not os.path.exist(f'output ({i}).wav'):
                write(f'output {i}.wav', fs, myrecording)
                break
    
    # Method to connect Python AI to openai
    # def AI(self):
    #     self.response = openai.Completion.create(
    #         model = "gpt-3.5-turbo-instruct",
    #         prompt="Write a program to print even number",
    #         temperature=0.7,
    #         max_tokens=256,
    #         top_p=1,
    #         frequency_penalty=0,
    #         presence_penalty=0
    #     )
    #     print(self.response)

    #  Method to play a music

if __name__ == "__main__":
    vAssist = VoiceAssistant()
    vAssist.mainloop()