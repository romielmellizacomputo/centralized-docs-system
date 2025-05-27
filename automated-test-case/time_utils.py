import datetime
import pytz

def get_current_times():
    ph_tz = pytz.timezone('Asia/Manila')
    ug_tz = pytz.timezone('Africa/Kampala')
    now_utc = datetime.datetime.utcnow().replace(tzinfo=pytz.utc)

    now_ph = now_utc.astimezone(ph_tz)
    now_ug = now_utc.astimezone(ug_tz)

    ph_time_str = now_ph.strftime('%B %d, %Y %I:%M %p')
    ug_time_str = now_ug.strftime('%B %d, %Y %I:%M %p')

    return ph_time_str, ug_time_str
