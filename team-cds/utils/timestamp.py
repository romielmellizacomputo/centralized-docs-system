from datetime import datetime
import pytz

def generate_timestamp_string():
    now = datetime.now()

    time_zone_eat = pytz.timezone('Africa/Nairobi')
    time_zone_pht = pytz.timezone('Asia/Manila')

    formatted_date_eat = now.astimezone(time_zone_eat).strftime('%B %d, %Y')
    formatted_date_pht = now.astimezone(time_zone_pht).strftime('%B %d, %Y')

    formatted_eat = now.astimezone(time_zone_eat).strftime('%I:%M:%S %p')
    formatted_pht = now.astimezone(time_zone_pht).strftime('%I:%M:%S %p')

    return f"Sync on {formatted_date_eat}, {formatted_eat} (UG) / {formatted_date_pht}, {formatted_pht} (PH)"
