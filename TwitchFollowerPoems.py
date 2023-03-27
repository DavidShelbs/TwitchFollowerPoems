import os
import asyncio
import requests
import threading
import pythoncom
import PySimpleGUI as sg
from dotenv import load_dotenv
import win32com.client as wincom
from twitchAPI.twitch import Twitch
from twitchAPI.helper import first
from twitchAPI.eventsub import EventSub

load_dotenv()
# parser = argparse.ArgumentParser(description='Twitch Limerick Creator')
# parser.add_argument('--eventsub-url', help='The URL of the HTTPS reverse proxy server (ngrStart Program, NGINX, etc.)', required=True)
# args = parser.parse_args()

# EVENTSUB_URL=args.eventsub_url

layout = [
    [sg.Text('Reverse Proxy URL:'), sg.InputText(key='_input_'), sg.Text(key='_output_', visible=False)],
    [sg.Button('Start Program', key='_ok_'), sg.Button('Cancel', key='_cancel_'), sg.Push(), sg.Text('Status: Not Connected', key='_status_text_', justification='center')]
]

# sg.theme('DarkAmber')
window = sg.Window("Twitch Follower Poems", layout)

# EVENTSUB_URL = input('Enter the URL of the HTTPS reverse proxy server (ngrStart Program, NGINX, etc.): ')
TARGET_USERNAME=os.getenv('TARGET_USERNAME')
APP_ID=os.getenv('APP_ID')
APP_SECRET=os.getenv('APP_SECRET')
OPENAI_URL=os.getenv('OPENAI_URL')
OPENAI_API_KEY=os.getenv('OPENAI_API_KEY')

def run_in_thread(speak_id, limerick, follower_user_name):
    # Initialize
    pythoncom.CoInitialize()
    # Get instance from the id
    speak = wincom.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(speak_id, pythoncom.IID_IDispatch))

    speak.Speak(f'{follower_user_name} just followed the stream! {limerick}')

async def create_sessions(EVENTSUB_URL):
    # create the api instance and get the ID of the target user
    twitch = await Twitch(APP_ID, APP_SECRET)
    user = await first(twitch.get_users(logins=TARGET_USERNAME))
    
    # basic setup, will run on port 8080 and a reverse proxy takes care of the https and certificate
    event_sub = EventSub(EVENTSUB_URL, APP_ID, 9696, twitch)
    return twitch, user, event_sub

async def on_follow(data: dict):
    # our event happend, lets do things with the data we got!
    follower_user_name = data['event']['user_login']
    headers = {'Content-Type': 'application/json; charset=utf-8', 'Authorization': f'Bearer {OPENAI_API_KEY}'}
    data = {'model': 'gpt-3.5-turbo','messages': [{'role': 'user', 'content': f'Create limerick for someone named {follower_user_name} without using gendered terms.'}]}

    response = requests.post(OPENAI_URL, headers=headers, json=data)
    limerick = response.json()['choices'][0]['message']['content']
    print(f'Limerick: {limerick}')

    # Initialize
    pythoncom.CoInitialize()
    # Get instance
    speak = wincom.Dispatch("SAPI.SpVoice")
    # Create id
    speak_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, speak)
    # Pass the id to the new thread
    thread = threading.Thread(target=run_in_thread, kwargs={'speak_id': speak_id, 'limerick': limerick, 'follower_user_name': follower_user_name})
    thread.start()
    # Wait for child to finish
    thread.join()

async def eventsub(EVENTSUB_URL, window):
    window['_status_text_'].update('Status: Connecting...')
    window.refresh()

    # create the api instance and get the ID of the target user
    twitch = await Twitch(APP_ID, APP_SECRET)
    user = await first(twitch.get_users(logins=TARGET_USERNAME))
    
    # basic setup, will run on port 8080 and a reverse proxy takes care of the https and certificate
    event_sub = EventSub(EVENTSUB_URL, APP_ID, 9696, twitch)

    # unsubscribe from all old events that might still be there
    # this will ensure we have a clean slate
    await event_sub.unsubscribe_all()
    
    # start the eventsub client
    event_sub.start()
    
    # subscribing to the desired eventsub hoStart Program for our user
    # the given function will be called every time this event is triggered
    # eventsub will run in its own process
    await event_sub.listen_channel_follow_v2(user.id, user.id, on_follow)

    window['_status_text_'].update('Status: Connected')
    window.refresh()

async def gui_window_loop():
    eventsub_task = None
    while True:
        event, values = window.read(timeout=100)
        if event in (sg.WIN_CLOSED, '_cancel_'):
            if eventsub_task and not eventsub_task.done():
                eventsub_task.cancel()
            window.close()
            os._exit(0)
        elif event in ('_ok_'):
            EVENTSUB_URL = values['_input_']
            window['_input_'].update(visible=False)
            window['_output_'].update(EVENTSUB_URL, visible=True)
            window['_ok_'].update(disabled=True)
            window.refresh()
            eventsub_task = asyncio.create_task(eventsub(EVENTSUB_URL, window))
            await eventsub_task
        elif event in (sg.TIMEOUT_EVENT):
            pass

async def main():
    main_gui = asyncio.create_task(gui_window_loop())
    await main_gui

asyncio.run(main())