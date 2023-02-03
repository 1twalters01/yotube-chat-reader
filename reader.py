import pytchat
import time
import json

link = "https://www.youtube.com/watch?v=MtQFR8xIfVI"
video_id = link.replace('https://www.youtube.com/watch?v=','')


chat = pytchat.create(video_id=video_id)
while chat.is_alive():
    try:
        for c in chat.get().sync_items():
            if c.type == 'superChat':
                print(f"{c.datetime} [{c.author.name}]- {c.message}")
    except AttributeError:
        print("Error occurred. Restarting...")
        chat = pytchat.create(video_id="uIx8l2xlYVY")
