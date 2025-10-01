import winsound

def play_end_sound():
    """終了時のポーン音"""
    winsound.MessageBeep(winsound.MB_ICONASTERISK)

def play_stop_sound():
    """停止時の音"""
    winsound.MessageBeep(winsound.MB_ICONHAND)
