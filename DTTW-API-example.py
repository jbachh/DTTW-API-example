# -*- coding: utf-8 -*-
"""
Calls out emini S&P futures quotes. 

Connects to the PPro API. Depends on standard installed Windows fe/male voice.

Python 3.8. Compile like so

<python -m nuitka --mingw64 --standalone la_rana.py>
or <python -m nuitka --mingw64 la_rana.py>

voice: https://mail.python.org/pipermail/tutor/2008-June/062384.html

Created on Thu Dec 10 01:38:13 2020.

@author: Joop
"""
from atexit import register
from socket import socket, AF_INET, SOCK_STREAM, SOCK_DGRAM
from threading import Thread, Event
from winsound import Beep
from win32com.client import Dispatch
import sys
sys.tracebacklimit = 0

SYMBOL = 'ES\\M21.CM'  # change according to front-month
UP = False


def cleanup(sock, speaker):
    """
    Close everything at exit.

    """
    if speaker.Status.RunningState == 2:
        speaker.Speak("", 1 | 2)
    speaker.Speak('exiting')
    try:
        sock.close()
        turn_port('off')
    except Exception as e:
        print(e)


def turn_port(mode):
    """
    Open or close the UDP port.

    Parameters
    ----------
    mode : str
        On or off for open or close.

    Returns
    -------
    None.

    """
    if mode == 'on':
        s = socket(AF_INET, SOCK_STREAM)
        s.connect(("localhost", 8080))
        req = f"GET /Register?symbol={SYMBOL}&feedtype=L1 HTTP/1.1\r\n".encode()
        req += b"Host: localhost:8080 \r\nConnection: close\r\n\r\n"
        s.sendall(req)
        s.close()

    s = socket(AF_INET, SOCK_STREAM)
    s.connect(("localhost", 8080))
    req = f"GET /SetOutput?symbol={SYMBOL}&".encode()
    # mode 'on' or 'off'
    req += f"region=1&feedtype=L1&output=4135&status={mode} HTTP/1.1\r\n"\
        .encode()
    req += b"Host: localhost:8080 \r\nConnection: close\r\n\r\n"
    s.sendall(req)
    s.close()

    if mode == 'off':
        s = socket(AF_INET, SOCK_STREAM)
        s.connect(("localhost", 8080))
        req = f"GET /Deregister?symbol={SYMBOL}&feedtype=L1 HTTP/1.1\r\n".encode()
        req += b"Host: localhost:8080 \r\nConnection: close\r\n\r\n"
        s.sendall(req)
        s.close()


def get_first_bid():
    """
    Return current bid, ask, and midpoint.

    Returns
    -------
    last_bid : float
        current bid.

    """
    s = socket(AF_INET, SOCK_STREAM)
    s.connect(("localhost", 8080))
    request = f"GET /GetLv1?symbol={SYMBOL} HTTP/1.1\r\n".encode()
    request = request + b"Host: localhost:8080 \r\nConnection: close\r\n\r\n"
    s.sendall(request)
    data = s.recv(1024).decode()
    s.close()
    last_bid = float(data[data.find('BidPrice')+10:data.find('BidPrice')+17])
    return last_bid - last_bid % 0.5


def voice(speaker, price, side):
    """
    Render voice.

    """
    if price.is_integer():
        price = str(int(price) % 100)
    else:
        price = str(price % 100)
    if speaker.Status.RunningState == 2:
        speaker.Speak("", 1 | 2)
    speaker.Speak(f"<pitch absmiddle='{0}'/>{price} {side}", 9)  # 9=async


def beep(event):
    """
    Play beep.

    """
    while True:
        event.wait()
        if UP:
            Beep(512, 112)
        else:
            Beep(256, 96)
        event.clear()


def main():
    """
    Run main thread.

    Returns
    -------
    None.

    """
    print("Starting la Rana.. Press Ctrl-C to close.")

    speaker = Dispatch("SAPI.SpVoice")
    speakers = [speaker.GetVoices()[i].GetDescription() for i in range(
        len(speaker.GetVoices()))]
    index = [idx for idx, s in enumerate(speakers) if 'Zira' in s][0]
    speaker.Voice = speaker.GetVoices()[index]
    """
    Make sure that you have standard voice installed
    https://www.trishtech.com/2015/09/choose-text-to-speech-voice-in
    -windows-10/  ,
    if not (like I did on my foreign windows) install the US English voice.
    """
    sock = socket(AF_INET, SOCK_DGRAM)
    sock.bind(("", 4135))
    sock.settimeout(5)  # seconds

    register(cleanup, sock, speaker)  # execute at exit

    last_bid = get_first_bid()
    last_beep = last_bid
    last_voice = round(last_bid/2.5) * 2.5
    up_voice_trigger = last_voice + 2.5
    down_voice_trigger = last_voice - 2.5
    speaker.Speak(  # 9 is flag for async
        f'ES trading at {last_bid % 100}.', 9)
    turn_port('on')
    beep_event = Event()
    Thread(target=beep, args=(beep_event,), daemon=True).start()
    global UP

    while True:
        data = sock.recv(512)
        try:  # sometimes bid price index is different; skip
            bid = float(data[84:91])
        except Exception as e:
            print(e)
            continue

        if bid >= up_voice_trigger:
            last_voice = bid - bid % 2.5
            up_voice_trigger = last_voice + 2.5
            down_voice_trigger = last_voice - 2.5
            voice(speaker, last_voice, 'up')  # async
        elif bid <= down_voice_trigger:
            last_voice = bid - bid%2.5+2.5 if bid < down_voice_trigger else bid
            up_voice_trigger = last_voice + 2.5
            down_voice_trigger = last_voice - 2.5
            voice(speaker, last_voice, 'down')  # async
        if bid - last_bid >= 0.5:
            last_bid = bid - bid % .5
            if bid != last_beep:
                UP = True
                beep_event.set()
                last_beep = last_bid
        elif last_bid - bid >= 0.5:
            last_bid = bid + bid % .5
            if bid != last_beep:
                UP = False
                beep_event.set()
                last_beep = last_bid


if __name__ == "__main__":
    main()
