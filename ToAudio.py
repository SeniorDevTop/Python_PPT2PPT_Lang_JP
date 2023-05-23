from comtypes.client import CreateObject
from comtypes.gen import SpeechLib

engine = CreateObject("SAPI.SpVoice")
stream = CreateObject("SAPI.SpFileStream")

def main(strTxtFilePath, strSaveAudioPath):
    stream.Open(strSaveAudioPath, SpeechLib.SSFMCreateForWrite)
    engine.AudioOutputStream = stream
    f = open(strTxtFilePath, 'r', encoding="utf-8")
    theText = f.read()
    f.close()
    engine.speak(theText)
    stream.Close()