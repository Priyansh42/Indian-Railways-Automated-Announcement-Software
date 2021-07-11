import os
import glob
import gtts
from gtts import gTTS
import pandas as pd
from pydub import AudioSegment, audio_segment

def generateSkeleton():
    audio=AudioSegment.from_mp3('Indian-Railways_Announcement.mp3')

    ############### HINDI ###############

    # Part-1 start music
    start=0000
    finish=1550
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-1.mp3',format="mp3")

    # Part-2 kripya dhyan...gadi sankhiya
    start=1800
    finish=4150
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-2.mp3',format="mp3")

    # Part-3 train no and name

    # Part-4 from city

    # Part-5 se chalkar
    start=8950
    finish=9750
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-5.mp3',format="mp3")

    # Part-6 via city

    # Part-7 ke raaste
    start=11100
    finish=11850
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-7.mp3',format="mp3")

    # Part-8 to city

    # Part-9 ko jane wali...platform sankhiya
    start=12650
    finish=15850
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-9.mp3',format="mp3")

    # Part-10 platform no

    # Part-11 par aa rahi hai...end music
    start=16400
    finish=19500
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-11.mp3',format="mp3")


    ############### ENGLISH ###############

    # Part-12 attention please...train no
    start=36950
    finish=39800
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-12.mp3',format="mp3")

    # Part-13 train no and name

    # Part-14 from
    start=44200
    finish=44750
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-14.mp3',format="mp3")

    # Part-15 from city

    # Part-16 to
    start=45600
    finish=46150
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-16.mp3',format="mp3")

    # Part-17 to city

    # Part-18 via
    start=47100
    finish=47650
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-18.mp3',format="mp3")

    # Part-19 via city
 
    # Part-20 is arriving...platform no
    start=49000
    finish=52100
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-20.mp3',format="mp3")

    # Part-21 platform no

    # Part-22 thank you...end music
    start=52800
    finish=54800
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-22.mp3',format="mp3")


    ############### GUJARATI ###############

    # Part-23 krupa karine...gadi sankhiya
    start=19600
    finish=22350
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-23.mp3',format="mp3")

    # Part-24 train no and name

    # Part-25 from

    # Part-26 thi...via
    start=27200
    finish=28450
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-26.mp3',format="mp3")

    # Part-27 via

    # Part-28 to

    # Part-29 sudhi...platform no
    start=30800
    finish=33700
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-29.mp3',format="mp3")

    # Part-30 platform no

    # Part-31 par avi rahi che...end music
    start=34200
    finish=36900
    audioProcessed=audio[start:finish]
    audioProcessed.export('Part-31.mp3',format="mp3")


def textToSpeechHindi(text, filename):
    mytext=str(text)
    language='hi'
    myobj=gTTS(text=mytext,lang=language,slow=False)
    myobj.save(filename)


def textToSpeechEnglish(text, filename):
    mytext=str(text)
    language='en'
    myobj=gTTS(text=mytext,lang=language,slow=False)
    myobj.save(filename)    


def textToSpeechGujarati(text, filename):
    mytext=str(text)
    language='gu'
    myobj=gTTS(text=mytext,lang=language,slow=False)
    myobj.save(filename)


def mergeAudios(audios):
    combined=AudioSegment.empty()
    for audio in audios:
        combined+=AudioSegment.from_mp3(audio)
    return combined


def generateAnnouncement(filename):
    excelFile=pd.read_excel(filename)
    print(excelFile)
    for index, item in excelFile.iterrows():

        ############### HINDI ###############

        # Part-3 train no and name
        textToSpeechHindi(item['Train_No']+"  "+item['Train_Name'],'Part-3.mp3')

        # Part-4 from city
        textToSpeechHindi(item['From'],'Part-4.mp3')

        # Part-6 via city
        textToSpeechHindi(item['Via'],'Part-6.mp3')

        # Part-8 to city
        textToSpeechHindi(item['To'],'Part-8.mp3')

        # Part-10 platform no
        textToSpeechHindi(item['Platform_No'],'Part-10.mp3')


        ############### ENGLISH ###############

        # Part-13 train no and name
        textToSpeechEnglish(item['Train_No']+"  "+item['Train_Name'],'Part-13.mp3')

        # Part-15 from city
        textToSpeechEnglish(item['From'],'Part-15.mp3')
        
        # Part-17 to city
        textToSpeechEnglish(item['To'],'Part-17.mp3')

        # Part-19 via city
        textToSpeechEnglish(item['Via'],'Part-19.mp3')

        # Part-21 platform no
        textToSpeechEnglish(item['Platform_No'],'Part-21.mp3')


        ############### GUJARATI ###############

        # Part-24 train no and name
        textToSpeechGujarati(item['Train_No']+"  "+item['Train_Name'],'Part-24.mp3')

        # Part-25 from city
        textToSpeechGujarati(item['From'],'Part-25.mp3')
        
        # Part-27 via city
        textToSpeechGujarati(item['Via'],'Part-27.mp3')

        # Part-28 to city
        textToSpeechGujarati(item['To'],'Part-28.mp3')

        # Part-30 platform no
        textToSpeechGujarati(item['Platform_No'],'Part-30.mp3')

        audios=[f"Part-{i}.mp3" for i in range(1,32)]

        announcement=mergeAudios(audios)
        announcement.export(f"Announcement_{item['Train_No']}.mp3",format="mp3")


if __name__ == "__main__":
    for filename in glob.glob("E:\Indian Railways Announcement Software/Announcement*"):
        os.remove(filename) 
    print("\n")
    print("Generating Skeleton...Wait for a while...")
    generateSkeleton()
    print("Skeleton Generated Successfully.")
    print("Generating Announcement...Wait for a while...")
    generateAnnouncement("Train-Announcement_List.xlsx")
    for i in range(1,32):
        os.remove(f"Part-{i}.mp3")
    print("Announcement Generated Successfully.")  
    print("\n")      